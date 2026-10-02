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
  const operations=['searchDocuments','getCategoryStats','exportAuditList','uploadAndSaveDocument','uploadScanFilesBatch','batchImportDocuments','updateDocument','deleteDocument'];
  function chain(success=()=>{},failure=()=>{}){
    const runner={withSuccessHandler:fn=>chain(fn,failure),withFailureHandler:fn=>chain(success,fn)};
    for(const name of operations)runner[name]=(...args)=>{
      Promise.resolve().then(()=>{if(!db)throw Error('尚未連接 SharePoint，資料未儲存。');return db[name](...args);}).then(success).catch(failure);
    };
    return runner;
  }
  root.appApi={get run(){return chain();}};
  root.M365App={init,signIn,signOut,configErrors,discardDraftAttachments:()=>db?.discardDraftAttachments()};
})(globalThis);
