(function(root){
  'use strict';
  let auth=null, account=null, db=null;
  const c=root.M365_CONFIG;
  function configErrors(){
    const errors=[];
    const guid=/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
    if(!guid.test(c.tenantId))errors.push('目錄（租用戶）ID');
    if(!guid.test(c.clientId))errors.push('應用程式（用戶端）ID');
    if(!c.redirectUri)errors.push('登入返回網址');
    else {try{const u=new URL(c.redirectUri);if(u.origin!==location.origin || (u.protocol!=='https:' && !(u.protocol==='http:' && ['localhost','127.0.0.1'].includes(u.hostname))))errors.push('與目前網站同來源的 HTTPS 返回網址（本機 localhost 可用 HTTP）');}catch{errors.push('有效的登入返回網址');}}
    if(location.protocol==='file:')errors.push('請透過網站網址開啟，不能直接雙擊 HTML 登入');
    if(!((c.driveId && c.workbookItemId)||c.workbookShareUrl))errors.push('Excel 分享連結或檔案 ID');
    return errors;
  }
  function status(text){document.getElementById('configStatus').textContent=text;}
  async function init(){
    installDiagnostics();
    const missing=configErrors();
    if(missing.length){status('尚未連線。請先完成：'+missing.join('、')+'。');document.getElementById('m365LoginButton').disabled=true;return;}
    if(!root.msal){status('Microsoft 登入元件載入失敗，請確認完整解壓縮所有檔案。');return;}
    auth=new root.msal.PublicClientApplication({auth:{clientId:c.clientId,authority:'https://login.microsoftonline.com/'+c.tenantId,redirectUri:c.redirectUri,postLogoutRedirectUri:c.redirectUri,navigateToLoginRequestUrl:false},cache:{cacheLocation:'sessionStorage'}});
    try{
      await auth.initialize();
      const result=await auth.handleRedirectPromise();
      account=result?.account || auth.getActiveAccount() || auth.getAllAccounts().find(a=>a.tenantId===c.tenantId);
      if(account){auth.setActiveAccount(account);await connect();}
      else status('設定已填入，請使用公司 Microsoft 365 帳號登入。');
    }catch(e){status('連線未完成：'+e.message);root.toggleLoading?.(false);}
  }
  async function token(){
    if(!auth || !account)throw Error('請先使用公司帳號登入。');
    try{return (await auth.acquireTokenSilent({account,scopes:c.scopes})).accessToken;}
    catch{throw Error('登入或授權需要重新確認，請登出後再按「使用公司帳號登入」。');}
  }
  async function connect(){
    root.toggleLoading?.(true,'正在驗證 SharePoint 檔案與八個欄位...');
    const user={name:account.name || account.username,username:account.username,role:'公司成員'};
    db=new root.M365Core.Database(c,root.M365Core.createGraph(token),user);
    try{const info=await db.connect();root.onM365Connected(user,info);status('已連線 '+info.fileName);}
    catch(e){db=null;throw e;}
    finally{root.toggleLoading?.(false);}
  }
  async function signIn(){
    if(configErrors().length)throw Error('請先完成 config.js 內的公司設定。');
    if(!auth)throw Error('登入元件尚未準備好，請重新載入網頁。');
    await auth.loginRedirect({scopes:c.scopes,prompt:'select_account'});
  }
  async function signOut(){
    if(db?.busy)throw Error('寫入尚未完成，請稍候再登出。');
    db?.discardDraftAttachments();db=null;
    if(auth)await auth.logoutRedirect({account,postLogoutRedirectUri:c.redirectUri});
  }
  function installDiagnostics(){
    if(document.getElementById('diagnosticD2'))return;
    const box=document.createElement('div');box.id='diagnosticD2';
    box.style.cssText='margin-top:12px;text-align:left;color:#334155;font-size:14px';
    const button=document.createElement('button');button.type='button';button.textContent='執行唯讀連線檢查（D2）';
    button.style.cssText='padding:10px;background:#e2e8f0;border-radius:8px;color:#0f172a;cursor:pointer';
    const output=document.createElement('textarea');output.readOnly=true;output.hidden=true;
    output.setAttribute('aria-label','連線檢查報告');output.style.cssText='width:100%;height:280px;margin-top:10px;padding:8px;font-size:12px;color:#0f172a;background:white';
    box.append(button,output);document.getElementById('configStatus').after(box);
    button.onclick=async()=>{
      button.disabled=true;output.hidden=false;output.value='正在檢查，約需數秒……';
      try{output.value=await runDiagnostics();}finally{button.disabled=false;}
    };
  }
  async function runDiagnostics(){
    const lines=['連線檢查 D2（僅讀取資訊，不修改文件）','時間：'+new Date().toISOString(),
      '應用程式：'+c.clientId,'租用戶：'+c.tenantId,'返回網址：'+c.redirectUri,
      '目前頁面：'+location.origin+location.pathname,'要求權限：'+c.scopes.join(', '),
      'driveId：'+c.driveId,'Excel itemId：'+c.workbookItemId,
      'E_test itemId：'+c.expectedParentFolderItemId];
    if(!auth || !account)return lines.concat('尚未登入：請先按公司帳號登入，再執行檢查。').join('\n');
    lines.push('登入帳號：'+account.username,'帳號租用戶相符：'+(account.tenantId===c.tenantId));
    let result;
    try{result=await auth.acquireTokenSilent({account,scopes:c.scopes,forceRefresh:true});}
    catch(e){return lines.concat('重新取得登入授權失敗：'+String(e.errorCode||'unknown').replace(/[^a-zA-Z0-9_.-]/g,''),'請重新登入後再試。').join('\n');}
    // Report the MSAL result scopes, never the access token or raw authentication response.
    lines.push('登入元件回報的權限：'+(result.scopes||[]).join(', '));
    const enc=encodeURIComponent, base='/drives/'+enc(c.driveId)+'/items/';
    const probes=[['E_test 資料夾',base+enc(c.expectedParentFolderItemId),c.expectedParentFolderItemId],
      ['Excel（檔案 ID）',base+enc(c.workbookItemId),c.workbookItemId],
      ['Excel（E_test 內檔名）',base+enc(c.expectedParentFolderItemId)+':/'+enc(c.expectedWorkbookName),c.workbookItemId]];
    if(c.scopes.includes('Lists.SelectedOperations.Selected'))probes.push(['Document_DB 清單','/sites/'+enc(c.siteId)+'/lists/'+enc(c.listId),c.listId]);
    for(const [label,path,expectedId] of probes){
      const url='https://graph.microsoft.com/v1.0'+path+'?$select=id,name';
      lines.push('\n'+label,'GET '+url);
      const controller=new AbortController(),timer=setTimeout(()=>controller.abort(),15000);
      try{
        const response=await root.fetch(url,{method:'GET',headers:{Authorization:'Bearer '+result.accessToken},signal:controller.signal,cache:'no-store',redirect:'error'});
        const data=await response.json().catch(()=>({}));
        lines.push('HTTP '+response.status);
        if(response.ok)lines.push('回傳 ID：'+String(data.id||''),'與設定 ID 相符：'+(data.id===expectedId));
        else lines.push('錯誤代碼：'+String(data.error?.code||'unknown').replace(/[^a-zA-Z0-9_.-]/g,'').slice(0,80));
        const requestId=String(response.headers.get('request-id')||data.error?.innerError?.['request-id']||'').replace(/[^a-zA-Z0-9-]/g,'').slice(0,80);
        if(requestId)lines.push('查詢代碼：'+requestId);
      }catch(e){lines.push(e.name==='AbortError'?'請求逾時':'網路請求失敗（未取得 HTTP 回覆）');}
      finally{clearTimeout(timer);}
    }
    lines.push('\n此報告不含存取權杖、密碼或文件內容。HTTP 成功僅表示可讀取資訊，尚未測試寫入及 Excel 表格 API。');
    return lines.join('\n');
  }

  const operations=['searchDocuments','getCategoryStats','exportAuditList','uploadAndSaveDocument','uploadScanFilesBatch','batchImportDocuments','updateDocument','deleteDocument'];
  function chain(success=()=>{},failure=()=>{}){
    const runner={withSuccessHandler:fn=>chain(fn,failure),withFailureHandler:fn=>chain(success,fn)};
    for(const name of operations)runner[name]=(...args)=>{
      Promise.resolve().then(()=>{if(!db)throw Error('尚未連接 SharePoint，資料未儲存。');return db[name](...args);}).then(success).catch(failure);
    };
    return runner;
  }
  root.appApi={get run(){return chain();}};
  root.M365App={runDiagnostics,init,signIn,signOut,configErrors,discardDraftAttachments:()=>db?.discardDraftAttachments()};
})(globalThis);
