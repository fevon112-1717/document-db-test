(function(root){
  'use strict';
  const core=root.M365Core, Base=core.Database, enc=encodeURIComponent;
  const unescape=v=>typeof v==='string' && /^'[\s]*[=+@-]|^'[\t\r\n]/.test(v)?v.slice(1):v;
  class ListDatabase extends Base {
    constructor(config,graph,user){
      super({...config},graph,user); this.remote=graph; this.columns=[];
      this.graph=async(path,options)=>{
        if(path===this.tablePath+'/rows/add'){
          let committed=0;
          try {
            for(const row of options.body.values){
              await this.remote(this.listPath+'/items',{method:'POST',body:{fields:this.fields(row)}});committed++;
            }
          }catch(e){if(committed)e.uncertain=true;throw e;}
          return {};
        }
        return this.remote(path,options);
      };
    }
    get listPath(){return `/sites/${enc(this.config.siteId)}/lists/${enc(this.config.listId)}`;}
    get tablePath(){return this.listPath;}
    get itemPath(){return this.listPath;}
    async pages(path){
      const result=[],seen=new Set();
      while(path){
        if(seen.has(path))throw Error('重複分頁，已停止讀取。');seen.add(path);
        const p=await this.remote(path);
        if(!Array.isArray(p.value))throw Error('清單回傳格式不符。');
        result.push(...p.value);if(result.length>100000)throw Error('超過測試版讀取上限。');path=p['@odata.nextLink'];
      }
      return result;
    }
    async connect(){
      const c=this.config, list=await this.remote(this.listPath);
      const actual=new URL(list.webUrl), expected=new URL(c.listUrl);
      if(actual.origin!==expected.origin || actual.pathname.replace(/\/(AllItems\.aspx)?$/i,'').toLowerCase()!==expected.pathname.replace(/\/(AllItems\.aspx)?$/i,'').toLowerCase())throw Error('清單位置與 Document_DB 不符，已停止。');
      const columns=await this.pages(this.listPath+'/columns');
      this.columns=core.HEADERS.map((name,i)=>{
        const found=columns.filter(col=>col.displayName===name);
        if(found.length!==1)throw Error('請建立且勿重複此清單欄位：'+name);
        const col=found[0];
        if(col.readOnly || ([4,7].includes(i)?!col.dateTime:!col.text))throw Error(name+' 欄位類型須為'+([4,7].includes(i)?'日期和時間':'文字')+'，且可編輯。');
        if(i===4 && col.dateTime.format!=='dateOnly')throw Error('檔案日期請設為僅日期。');
        if(i===7 && col.dateTime.format!=='dateTime')throw Error('上傳時間請包含時間。');
        return col;
      });
      if(columns.some(col=>col.required && !col.readOnly && !col.hidden && !this.columns.some(x=>x.name===col.name) && col.defaultValue===undefined))throw Error('清單另有必填欄位，請先取消額外必填設定。');
      const folder=await this.remote(`/drives/${enc(c.driveId)}/items/${enc(c.uploadFolderItemId)}`);
      const url=new URL(folder.webUrl);
      if(!folder.folder || url.hostname!==c.sharePointHost || decodeURIComponent(url.pathname).replace(/\/$/,'')!==c.expectedFolderPath)throw Error('附件資料夾不是指定的 E_test。');
      this.item={parentReference:{driveId:c.driveId,id:folder.id}};
      // 明確授權的 E_test 下，各版本使用不同子資料夾。
      this.config.uploadFolderItemId='';
      await this.read();
      return {fileName:'Document_DB（List 版）',webUrl:list.webUrl};
    }
    fields(row){
      return Object.fromEntries(this.columns.map((col,i)=>{
        let v=unescape(row[i]);
        if(i===4)v=v===''?null:core.dateValue(v)+'T00:00:00Z';
        if(i===7)v=v===''?null:new Date(core.dateValue(v,true).replace(' ','T')+'+08:00').toISOString();
        if(col.text && v.length>(col.text.maxLength||255))throw Error(col.displayName+' 超過清單欄位字數限制。');
        return [col.name,v];
      }));
    }
    async read(){
      const items=await this.pages(this.listPath+'/items?$expand=fields&$top=200'),ids=new Set();
      return items.map((item,index)=>{
        const f=item.fields||{}, row=this.columns.map((col,i)=>{
          const v=f[col.name]??'';
          if(i===4)return v?String(v).slice(0,10):'';
          if(i===7 && v){const d=new Date(v);if(!Number.isFinite(d.valueOf()))throw Error('上傳時間格式無效。');return new Date(d.valueOf()+8*3600000).toISOString().slice(0,19).replace('T',' ');}
          return String(v);
        });
        const doc=core.toDoc({index,values:[row]});
        if(!doc)return null;
        if(ids.has(doc.docId))throw Error('清單文件編號重複。');ids.add(doc.docId);
        return {...doc,_itemId:item.id,_etag:item.eTag||f['@odata.etag']};
      }).filter(Boolean);
    }
    async operation(documents,file){
      // 先驗證所有欄位長度，避免附件上傳後才發現資料無法寫入。
      for(const d of documents)this.fields(['DOC'+ '0'.repeat(32),d.docName||'',d.category||'其他',d.location||'',core.dateSerial(d.fileDate),d.fileUrl?.startsWith('blob:')?'':d.fileUrl||'',this.user.name,46000]);
      return super.operation(documents,file);
    }
    async change(docId,updated,expected){
      if(!expected?._etag || !expected?._itemId)throw Error('缺少資料版本，請重新查詢。');
      const current=(await this.read()).find(d=>d.docId===docId);
      if(!current || current._etag!==expected._etag || current._itemId!==expected._itemId)throw Error('記錄已修改或刪除，請重新查詢。');
      const path=this.listPath+'/items/'+enc(current._itemId);
      let fields;
      if(updated){
        const row=current._raw.slice();
        if(!updated.docName?.trim() || !updated.location?.trim())throw Error('請填寫文件名稱與收藏位置。');
        row.splice(1,4,updated.docName.trim(),updated.category?.trim()||'其他',updated.location.trim(),core.dateSerial(updated.fileDate));
        const all=this.fields(row);fields=Object.fromEntries(this.columns.slice(1,5).map(c=>[c.name,all[c.name]]));
      }
      try{await this.remote(path+(updated?'/fields':''),{method:updated?'PATCH':'DELETE',headers:{'If-Match':expected._etag},...(updated?{body:fields}:{})});}
      catch(e){if(e.status===412)throw Error('其他人剛修改了這筆資料，請重新查詢。');throw e;}
      const actual=(await this.read()).find(d=>d.docId===docId);
      if(updated ? !actual || actual.docName!==updated.docName.trim() || actual.category!==(updated.category?.trim()||'其他') || actual.location!==updated.location.trim() || actual.fileDate!==(updated.fileDate||'--') : !!actual)throw Error('寫入後讀取結果尚未確認，請重新查詢，勿重複送出。');
      return {success:true,message:updated?'已更新清單並重新讀取確認。':'已刪除清單記錄；附件仍保留。'};
    }
  }
  core.Database=ListDatabase;
})(globalThis);
