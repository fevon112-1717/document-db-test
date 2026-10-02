(function (root) {
  'use strict';
  const HEADERS = ['文件編號','文件名稱','分類','收藏位置','檔案日期','雲端檔案連結','上傳者姓名','上傳時間'];
  const DAY = 86400000, EPOCH = Date.UTC(1899, 11, 30);
  const enc = s => encodeURIComponent(String(s));
  const str = v => v === null || v === undefined ? '' : String(v);
  const blank = v => v === null || v === undefined || v === '';
  const uuid = () => 'DOC' + root.crypto.randomUUID().replaceAll('-', '');
  const safeCell = value => /^[\s]*[=+@-]|^[\t\r\n]/.test(str(value)) ? "'" + str(value) : str(value);
  function dateValue(value, timestamp = false) {
    if (blank(value) || value === '--') return timestamp ? '' : '--';
    if (typeof value === 'number') {
      if (!Number.isFinite(value)) throw Error('Excel 日期數值無效。');
      return new Date(EPOCH + Math.round(value * DAY / 1000) * 1000).toISOString().slice(0, timestamp ? 19 : 10).replace('T',' ');
    }
    const v = str(value).trim();
    if (/^\d{4}-\d{2}-\d{2}(?:[ T].*)?$/.test(v)) return v.replace('T',' ').slice(0,timestamp ? 19 : 10);
    if (/^\d{4}\/\d{1,2}\/\d{1,2}/.test(v)) {
      const [y,m,d] = v.split(/[\/ ]/); return `${y}-${m.padStart(2,'0')}-${d.padStart(2,'0')}`;
    }
    throw Error('Excel 包含無法辨識的日期，請改成真正的日期或 yyyy-mm-dd 格式。');
  }
  function dateSerial(value) {
    if (!value || value === '--') return '';
    if (!/^\d{4}-\d{2}-\d{2}$/.test(value)) throw Error('檔案日期格式須為 yyyy-mm-dd。');
    const d = new Date(value + 'T00:00:00Z');
    if (!Number.isFinite(d.valueOf()) || d.toISOString().slice(0,10) !== value) throw Error('檔案日期不是有效日期。');
    return (d.valueOf() - EPOCH) / DAY;
  }
  function taipeiNow() {
    const parts = Object.fromEntries(new Intl.DateTimeFormat('en-CA', {timeZone:'Asia/Taipei',year:'numeric',month:'2-digit',day:'2-digit',hour:'2-digit',minute:'2-digit',second:'2-digit',hourCycle:'h23'}).formatToParts(new Date()).map(p=>[p.type,p.value]));
    return `${parts.year}-${parts.month}-${parts.day} ${parts.hour}:${parts.minute}:${parts.second}`;
  }
  function timestampSerial() { return (Date.parse(taipeiNow().replace(' ','T')+'Z') - EPOCH)/DAY; }
  function safeUrl(value, allowBlank = true) {
    if (!value && allowBlank) return '';
    const url = new URL(str(value));
    if (url.protocol !== 'https:' || url.username || url.password) throw Error('附件連結只接受 HTTPS 網址。');
    if (url.href.length > 2048) throw Error('附件連結過長。');
    return url.href;
  }
  function textField(value, label, required = false, max = 500) {
    const v = str(value).trim();
    if (required && !v) throw Error(`請填寫${label}。`);
    if (v.length > max) throw Error(`${label}最多 ${max} 個字元。`);
    return v;
  }
  function meta(value) {
    return { docName:textField(value.docName,'文件名稱',true), category:textField(value.category || '其他','分類',true,100), location:textField(value.location,'收藏位置',true), fileDate:dateSerial(value.fileDate) };
  }
  function toDoc(row) {
    const a = row.values?.[0];
    if (!Array.isArray(a) || a.length !== 8) throw Error('Document_DB 的資料列必須包含八個欄位。');
    if (a.every(blank)) return null; // 範本空白列、已刪除的空白列都不計數。
    if (!str(a[0]).trim()) throw Error('Excel 有缺少文件編號的資料列，請先修正後再操作。');
    return {docId:str(a[0]),docName:str(a[1]),category:str(a[2]),location:str(a[3]),fileDate:dateValue(a[4]),fileUrl:str(a[5]),uploaderName:str(a[6]),uploadTime:dateValue(a[7],true),_index:row.index,_raw:a.slice()};
  }
  function csvCell(v) { return '"' + safeCell(v).replaceAll('"','""') + '"'; }
  function csvBase64(docs) {
    const text = '\uFEFF' + HEADERS.join(',') + '\r\n' + docs.map(d=>[d.docId,d.docName,d.category,d.location,d.fileDate === '--' ? '' : d.fileDate,d.fileUrl,d.uploaderName,d.uploadTime].map(csvCell).join(',')).join('\r\n');
    let binary=''; for (const b of new TextEncoder().encode(text)) binary+=String.fromCharCode(b);
    return root.btoa(binary);
  }
  function shareId(url) {
    let binary=''; for (const b of new TextEncoder().encode(url)) binary+=String.fromCharCode(b);
    return 'u!' + root.btoa(binary).replace(/=+$/,'').replaceAll('/','_').replaceAll('+','-');
  }
  function messageFor(status, code) {
    if (status === 401) return '登入已失效，請登出後重新登入。';
    if (status === 403) return 'Microsoft 拒絕存取。請 IT 檢查應用程式授權、檔案與附件資料夾權限。';
    if (status === 404) return '找不到檔案、資料夾或 Document_DB Table，請確認設定。';
    if (status === 409 || status === 423) return 'Excel 正在被其他操作使用或發生衝突。請稍後重新載入確認。';
    if (status === 429) return 'Microsoft 暫時限制請求次數，請稍後再試。';
    return `Microsoft Graph 操作失敗（${status}${code ? ' / '+code : ''}）。`;
  }
  function createGraph(getToken, fetcher = root.fetch.bind(root), sleep = ms => new Promise(r=>setTimeout(r,ms))) {
    return async function graph(path, {method='GET',body,headers={}} = {}) {
      const url = new URL(path.startsWith('https:') ? path : 'https://graph.microsoft.com/v1.0'+path);
      if (url.origin !== 'https://graph.microsoft.com' || !url.pathname.startsWith('/v1.0/')) throw Error('拒絕將登入憑證傳到非 Microsoft Graph 網址。');
      for (let attempt=0; attempt<3; attempt++) {
        const token = await getToken();
        const controller = new AbortController(), timer=setTimeout(()=>controller.abort(),45000);
        let response;
        try {
          response = await fetcher(url.href, {method,headers:{Authorization:'Bearer '+token,...(body !== undefined && !(body instanceof Blob) ? {'Content-Type':'application/json'} : {}),...headers},body:body === undefined ? undefined : body instanceof Blob ? body : JSON.stringify(body),signal:controller.signal,cache:'no-store',redirect:'error'});
        } catch (cause) {
          const e = Error(method === 'GET' ? '無法連線 Microsoft，請確認網路後重試。' : '未收到寫入結果。請先重新查詢清單 確認，勿重複送出。');
          e.uncertain = method !== 'GET'; e.cause = cause; throw e;
        } finally { clearTimeout(timer); }
        if (method === 'GET' && [429,502,503,504].includes(response.status) && attempt<2) {
          const retry = Number(response.headers.get('Retry-After'));
          await sleep(Math.min(Number.isFinite(retry) && retry > 0 ? retry*1000 : 1000*2**attempt,15000)); continue;
        }
        if (!response.ok) {
          const detail=await response.json().catch(()=>({}));
          const e=Error(messageFor(response.status,detail.error?.code)); e.status=response.status; e.code=detail.error?.code; e.uncertain=method !== 'GET' && response.status>=500; throw e;
        }
        if (response.status === 204) return null;
        return response.json();
      }
    };
  }
  class Database {
    constructor(config, graph, user) { this.config=config; this.graph=graph; this.user=user; this.item=null; this.folder=null; this.busy=false; this.staged=new Map(); this.pending=null; }
    get itemPath() { if(!this.item) throw Error('尚未連線 Excel。'); return `/drives/${enc(this.item.parentReference.driveId)}/items/${enc(this.item.id)}`; }
    get tablePath() { return `${this.itemPath}/workbook/tables/${enc(this.config.tableName)}`; }
    async connect() {
      const c=this.config;
      let item = c.driveId && c.workbookItemId ? await this.graph(`/drives/${enc(c.driveId)}/items/${enc(c.workbookItemId)}`) : await this.graph(`/shares/${shareId(c.workbookShareUrl)}/driveItem`);
      item=item.remoteItem || item;
      if (!item.id || !item.parentReference?.driveId || !item.file) throw Error('分享連結未解析成有效的 Excel 檔案。請 IT 提供 driveId 與 workbookItemId。');
      if (new URL(item.webUrl).hostname !== c.sharePointHost) throw Error('檔案不在設定的公司 SharePoint。');
      if (item.name !== c.expectedWorkbookName) throw Error(`檔案名稱不符：預期 ${c.expectedWorkbookName}，實際為 ${item.name}。`);
      this.item=item;
      const table=await this.graph(this.tablePath);
      const header=await this.graph(this.tablePath+'/headerRowRange');
      const worksheet=await this.graph(this.tablePath+'/worksheet');
      if(table.name !== c.tableName || worksheet.name !== c.worksheetName || JSON.stringify(header.values?.[0]) !== JSON.stringify(HEADERS)) throw Error('工作表、Table 名稱或八個欄位順序不符，未進行任何寫入。');
      await this.read();
      return {fileName:item.name,webUrl:item.webUrl,tableName:table.name,driveId:item.parentReference.driveId,itemId:item.id};
    }
    async read() {
      let path=this.tablePath+'/rows?$top=500&$skip=0', offset=0; const docs=[], ids=new Set(), visited=new Set();
      while(path) {
        if(visited.has(path)) throw Error('Microsoft 回傳重複的分頁連結，已停止讀取。'); visited.add(path);
        const page=await this.graph(path);
        if(!Array.isArray(page.value)) throw Error('Microsoft 未回傳預期的資料列。');
        for(const row of page.value) {
          if(!Number.isInteger(row.index) || row.index<0) throw Error('資料列索引無效。');
          const d=toDoc(row); if(!d)continue;
          if(ids.has(d.docId)) throw Error(`文件編號重複：${d.docId}。請先在 Excel 修正。`);
          ids.add(d.docId); docs.push(d);
        }
        offset+=page.value.length;
        if(offset>100000)throw Error('資料量超過本版讀取上限，請評估正式資料庫。');
        path=page['@odata.nextLink'] || (page.value.length===500 ? this.tablePath+`/rows?$top=500&$skip=${offset}` : null);
      }
      return docs;
    }
    async searchDocuments(keyword='',category='',start='',end='') {
      if(start && end && start>end)throw Error('起始日期不可晚於結束日期。');
      const key=str(keyword).trim().toLowerCase();
      return (await this.read()).filter(d=>(!key || [d.docId,d.docName,d.location].some(v=>v.toLowerCase().includes(key))) && (!category || d.category===category) && ((!start && !end) || (d.fileDate!=='--' && (!start || d.fileDate>=start) && (!end || d.fileDate<=end)))).sort((a,b)=>b.uploadTime.localeCompare(a.uploadTime));
    }
    async getCategoryStats() {
      const docs=await this.read(), cats=new Map(), today=taipeiNow().slice(0,10);
      for(const d of docs)cats.set(d.category || '其他',(cats.get(d.category || '其他')||0)+1);
      return {totalFiles:docs.length,todayCount:docs.filter(d=>d.uploadTime.startsWith(today)).length,uniqueCategories:cats.size,chartData:{labels:[...cats.keys()],values:[...cats.values()]}};
    }
    async exportAuditList(...filters) { return {fileName:'文件盤點_'+taipeiNow().replace(/[- :]/g,'')+'.csv',base64Data:csvBase64(await this.searchDocuments(...filters))}; }
    async exclusive(fn) {
      if(this.busy)throw Error('上一筆操作仍在處理中，請勿重複送出。');
      this.busy=true;
      try { return root.navigator?.locks ? await root.navigator.locks.request('DocumentDB:'+this.itemPath,fn) : await fn(); }
      finally {this.busy=false;}
    }
    checkFile(file) {
      if(!(file instanceof Blob) || file.size===0)throw Error('附件為空或無效。');
      if(file.size>this.config.maxFileBytes)throw Error(`每個附件上限 ${this.config.maxFileBytes/1024/1024} MB。`);
    }
    decodeFile(input) {
      if(!input)return null;
      if(input instanceof Blob){this.checkFile(input);return input;}
      const match=/^data:([^;,]*);base64,([\s\S]+)$/.exec(input.base64Str || '');
      if(!match || match[2].length > Math.ceil(this.config.maxFileBytes/3)*4+8)throw Error('附件格式無效或超過大小限制。');
      const bytes=Uint8Array.from(root.atob(match[2]),c=>c.charCodeAt(0));
      const file=new File([bytes],str(input.name)||'attachment',{type:match[1]||'application/octet-stream'}); this.checkFile(file); return file;
    }
    async folderItem() {
      if(this.folder)return this.folder;
      const drive=`/drives/${enc(this.item.parentReference.driveId)}`;
      if(this.config.uploadFolderItemId) {
        const f=await this.graph(`${drive}/items/${enc(this.config.uploadFolderItemId)}`);
        if(!f.folder)throw Error('附件 Item ID 不是資料夾。'); return this.folder=f;
      }
      const parent=this.item.parentReference.id;
      if(!parent)throw Error('找不到 Excel 上層資料夾，請填入附件資料夾 Item ID。');
      const name=this.config.uploadFolderName;
      if(!name || /["*:<>?\/\\|]/.test(name))throw Error('附件資料夾名稱無效。');
      const path=`${drive}/items/${enc(parent)}:/${enc(name)}`;
      try {const f=await this.graph(path);if(!f.folder)throw Error('附件資料夾名稱被檔案占用。');return this.folder=f;}
      catch(e){if(e.status!==404)throw e;}
      try {return this.folder=await this.graph(`${drive}/items/${enc(parent)}/children`,{method:'POST',body:{name,folder:{},'@microsoft.graph.conflictBehavior':'fail'}});}
      catch(e){if(e.status===409){const f=await this.graph(path);if(f.folder)return this.folder=f;}throw e;}
    }
    async upload(file, id) {
      this.checkFile(file); const f=await this.folderItem();
      const name=(str(file.name)||'attachment').normalize('NFC').replace(/["*:<>?\/\\|\x00-\x1f]/g,'_').replace(/[. ]+$/g,'').slice(-140)||'attachment';
      const path=`/drives/${enc(this.item.parentReference.driveId)}/items/${enc(f.id)}:/${enc(id+'_'+name)}`;
      // 同一筆重試沿用同一個 UUID 路徑，不會覆寫其他文件的同名附件。
      try {const item=await this.graph(path+'/content',{method:'PUT',body:file,headers:{'Content-Type':file.type||'application/octet-stream'}}); return safeUrl(item.webUrl,false);}
      catch(e) { if(e.uncertain) {try{const item=await this.graph(path);if(item.size===file.size)return safeUrl(item.webUrl,false);}catch{} } throw e; }
    }
    async uploadScanFilesBatch(files) {
      if(!files.length || files.length>this.config.maxBatchRows)throw Error('附件數量超過限制或沒有附件。');
      const decoded=files.map(f=>this.decodeFile(f));
      if(decoded.reduce((n,f)=>n+f.size,0)>this.config.maxBatchBytes)throw Error('一批附件總大小上限 30 MB。');
      this.discardDraftAttachments();
      return decoded.map(file=>{const url=URL.createObjectURL(file);this.staged.set(url,file);return {name:file.name,url};});
    }
    discardDraftAttachments() {for(const url of this.staged.keys())URL.revokeObjectURL(url);this.staged.clear();}
    async operation(documents, singleFile=null) {
      if(!documents.length || documents.length>this.config.maxBatchRows)throw Error(`每批須為 1 至 ${this.config.maxBatchRows} 筆。`);
      const normalized=documents.map(meta); // 所有列先驗證，避免半批寫入。
      const files=documents.map((d,i)=>i===0 && singleFile ? this.decodeFile(singleFile) : this.staged.get(d.fileUrl)||null);
      files.filter(Boolean).forEach(f=>this.checkFile(f));
      if(files.reduce((n,f)=>n+(f?.size||0),0)>this.config.maxBatchBytes)throw Error('一批附件總大小上限 30 MB。');
      const urls=documents.map((d,i)=>files[i] ? '' : safeUrl(d.fileUrl || ''));
      const fileHashes=await Promise.all(files.map(async f=>f ? Array.from(new Uint8Array(await root.crypto.subtle.digest('SHA-256',await f.arrayBuffer()))).map(b=>b.toString(16).padStart(2,'0')).join('') : ''));
      const key=JSON.stringify({normalized,urls,fileHashes});
      if(this.pending && this.pending.key!==key)throw Error('前一筆尚未確認完成。請先以原資料重試確認，或重新載入並檢查清單。');
      const op=this.pending || (this.pending={key,rows:normalized.map((m,i)=>[uuid(),safeCell(m.docName),safeCell(m.category),safeCell(m.location),m.fileDate,urls[i],safeCell(this.user.name),timestampSerial()]),uncertain:false});
      try {
        const existing=await this.read(), byId=new Map(existing.map(d=>[d.docId,d]));
        const found=op.rows.filter(r=>byId.has(r[0]));
        if(found.length===op.rows.length){this.pending=null;return op.rows;}
        if(found.length || op.uncertain)throw Error('前次寫入結果尚未完整確認。請先檢查清單，勿重新送出或修改這批資料。');
        for(let i=0;i<files.length;i++)if(files[i] && !op.rows[i][5])op.rows[i][5]=await this.upload(files[i],op.rows[i][0]);
        try {
          await this.graph(this.tablePath+'/rows/add',{method:'POST',body:{index:null,values:op.rows}});
        } catch(e) {op.uncertain=!!e.uncertain;throw e;}
        // POST 已成功，即使驗證讀取失敗也不可再次新增。
        op.uncertain=true;
        const after=new Set((await this.read()).map(d=>d.docId));
        if(!op.rows.every(r=>after.has(r[0])))throw Error('Microsoft 已回覆寫入，但重新讀取尚未確認全部資料。請先檢查清單。');
        this.pending=null;return op.rows;
      } catch(e) {
        const uploaded=op.rows.some(r=>r[5] && files[op.rows.indexOf(r)]);
        if(!op.uncertain && !uploaded)this.pending=null;
        if(uploaded)e.message+=' 附件可能已存入 Document_DB_Attachments，請保留本頁後以相同資料重試；取消時請檢查未建檔的附件。';
        throw e;
      }
    }
    async uploadAndSaveDocument(metadata,file) {
      return this.exclusive(async()=>{const rows=await this.operation([metadata],file);return {success:true,docId:rows[0][0],message:'文件已寫入 SharePoint 清單，並完成重新讀取確認。'};});
    }
    async batchImportDocuments(documents) {
      return this.exclusive(async()=>{const rows=await this.operation(documents);this.discardDraftAttachments();return {success:true,importedCount:rows.length};});
    }
    async change(docId,updated,expected) {
      if(!expected?._raw)throw Error('缺少原始資料，請重新查詢後再操作。');
      const docs=await this.read(), current=docs.find(d=>d.docId===docId);
      if(!current)throw Error('找不到文件，可能已被刪除。請重新查詢。');
      if(JSON.stringify(current._raw)!==JSON.stringify(expected._raw))throw Error('文件已被其他人修改，請重新查詢後再編輯。');
      const path=this.tablePath+`/rows/itemAt(index=${current._index})/range`;
      const range=await this.graph(path);
      if(JSON.stringify(range.values?.[0])!==JSON.stringify(current._raw))throw Error('資料列位置或內容已變動，請重新查詢後再操作。');
      const address=str(range.address).split('!').pop();
      if(!/^A\d+:H\d+$/.test(address))throw Error('Excel 資料範圍不是預期的 A:H 欄，已停止寫入。');
      const writePath=`${this.itemPath}/workbook/worksheets/${enc(this.config.worksheetName)}/range(address='${address}')`;
      const row=current._raw.slice();
      if(updated){const m=meta(updated);row.splice(1,4,safeCell(m.docName),safeCell(m.category),safeCell(m.location),m.fileDate);}
      else row.fill(''); // 清空記錄而不移動其他列，保留附件供公司自行管理。
      const formats=[['@','@','@','@','yyyy-mm-dd','@','@','yyyy-mm-dd hh:mm:ss']];
      try{await this.graph(writePath,{method:'PATCH',body:{values:[row],numberFormat:formats}});}
      catch(e){if(!e.uncertain)throw e;}
      const verify=await this.graph(path), actual=verify.values?.[0];
      // Excel 會移除文字防公式用的前置單引號。
      const canonical=a=>a?.map(v=>typeof v==='string' && /^'[=+@-]/.test(v)?v.slice(1):v);
      if(JSON.stringify(canonical(actual))!==JSON.stringify(canonical(row)))throw Error('儲存结果未確認或同一筆資料又被修改，請重新查詢清單。');
      return {success:true,message:updated ? '已更新 SharePoint 清單。' : '已刪除清單 文件記錄；SharePoint 附件仍保留。'};
    }
    updateDocument(id,updated,expected) {return this.exclusive(()=>this.change(id,updated,expected));}
    deleteDocument(id,expected) {return this.exclusive(()=>this.change(id,null,expected));}
  }
  root.M365Core={HEADERS,Database,createGraph,dateSerial,dateValue,taipeiNow,safeCell,safeUrl,toDoc,csvBase64,shareId};
})(globalThis);
