    // =======================================
    // Microsoft 365：資料存取由 m365-app.js 與 m365-core.js 處理
    // =======================================

    // =======================================
    // 全域變數狀態大腦
    // =======================================
    let currentUser = null;
    let allDocuments = [];
    let doughnutChartInstance = null;
    let barChartInstance = null;
    let searchTimeout = null;
    let pendingDrafts = [];
    let editOriginal = null;
    let deleteOriginal = null;
    let deleteTargetId = null; 

    // 分類櫃位映射設定 (預設初始對應)
    function readCategoryStorage(){try{return localStorage.getItem('DocumentDB_M365_CategoryPreferences_'+M365_CONFIG.tenantId);}catch{return null;}}
    let categoryLocationMap = Object.assign(Object.create(null), {
      '用印': 'A1',
      '合約': 'B1',
      '簽呈': 'C1',
      '請購': 'D1',
      '驗收': 'E1',
      '專利': 'F1',
      '法務': 'G1',
      '其他': 'H1'
    });

    // 讀取此瀏覽器的個人分類偏好；文件資料不存 localStorage
    if (readCategoryStorage()) {
      try {
        categoryLocationMap = Object.assign(Object.create(null), categoryLocationMap, JSON.parse(readCategoryStorage()));
      } catch (e) {
        console.error("無法加載儲存的櫃位設定", e);
      }
    }

    // =======================================
    // 1. 訊息提示系統 (Toast)
    // =======================================
    function showToast(message, type = 'info') {
      const container = document.getElementById('toastContainer');
      const toast = document.createElement('div');
      
      const bgColors = {
        success: 'bg-emerald-500',
        error: 'bg-rose-500',
        info: 'bg-blue-500'
      };
      
      const icons = {
        success: 'fa-circle-check',
        error: 'fa-circle-xmark',
        info: 'fa-circle-info'
      };

      toast.className = `flex items-center gap-3 px-4 py-3 rounded-xl shadow-xl text-white ${bgColors[type]} transition-all duration-300 transform translate-x-12 opacity-0 text-sm font-semibold`;
      const icon=document.createElement('i');icon.className='fa-solid '+icons[type];
      const span=document.createElement('span');span.textContent=message;toast.append(icon,span);
      const status=document.getElementById('operationStatus');
      if(status){status.textContent=message;status.classList.remove('hidden');}
      
      container.appendChild(toast);
      
      setTimeout(() => {
        toast.classList.remove('translate-x-12', 'opacity-0');
      }, 50);

      setTimeout(() => {
        toast.classList.add('translate-x-12', 'opacity-0');
        setTimeout(() => toast.remove(), 300);
      }, 3000);
    }

    // 控制全螢幕遮罩
    function toggleLoading(show, text = "系統處理中，請稍候...") {
      const overlay = document.getElementById('loadingOverlay');
      const textElem = document.getElementById('loadingText');
      if (show) {
        textElem.textContent = text;
        overlay.classList.remove('hidden');
      } else {
        overlay.classList.add('hidden');
      }
    }

    // =======================================
    // 2. 登入、登出與權限控制
    // =======================================
    async function handleLogin(event) {
      event.preventDefault();
      try { await M365App.signIn(); } catch(e) { showToast(e.message,'error'); }
    }
    async function handleLogout() {
      try { await M365App.signOut(); } catch(e) { showToast(e.message,'error'); }
    }
    function onM365Connected(user, info) {
      currentUser=user;
      document.getElementById('userDisplayName').textContent=user.name;
      document.getElementById('userRoleBadge').textContent='依 SharePoint 權限';
      document.getElementById('btnExport').disabled=false;
      document.getElementById('btn-settingsTab').classList.remove('hidden');
      document.getElementById('loginScreen').classList.add('hidden');
      document.getElementById('mainApp').classList.remove('hidden');
      const link=document.getElementById('workbookLink');link.href=info.webUrl;link.textContent=info.fileName;
      populateSettingsForm();updateAllCategoryDropdowns();fetchDocuments();
    }

    function populateSettingsForm(){
      const container=document.getElementById('settingsFormContainer');container.replaceChildren();
      const defaults=['用印','合約','簽呈','請購','驗收','專利','法務','其他'];
      Object.entries(categoryLocationMap).forEach(([cat,loc])=>{
        const row=document.createElement('div');row.className='flex gap-3 items-center bg-slate-50 p-3 rounded-lg';
        const label=document.createElement('span');label.className='flex-1 text-sm';label.textContent=cat;
        const input=document.createElement('input');input.className='w-1/3 border rounded p-2 text-sm';input.value=loc;input.dataset.category=cat;input.required=true;
        row.append(label,input);
        if(!defaults.includes(cat)){const b=document.createElement('button');b.type='button';b.textContent='移除';b.className='text-rose-600 text-sm';b.onclick=()=>deleteCustomCategory(cat);row.append(b);}
        container.append(row);
      });
    }

    function addNewCategory() {
      const nameInput = document.getElementById('newCategoryName');
      const locInput = document.getElementById('newCategoryLocation');
      const name = nameInput.value.trim();
      const loc = locInput.value.trim().toUpperCase();

      if (!name || !loc) {
        showToast("請完整填寫分類名稱與預設櫃位！", "error");
        return;
      }

      if (categoryLocationMap[name]) {
        showToast(`分類「${name}」已存在！`, "error");
        return;
      }

      // 新增至全域對應大腦
      categoryLocationMap[name] = loc;
      try { localStorage.setItem('DocumentDB_M365_CategoryPreferences_' + M365_CONFIG.tenantId, JSON.stringify(categoryLocationMap)); } catch { showToast('此瀏覽器無法保存個人分類設定。','error'); return; }
      
      // 清空輸入框
      nameInput.value = "";
      locInput.value = "";

      showToast(`成功新增自訂分類「${name}」對應櫃位「${loc}」！`, "success");
      populateSettingsForm();
      updateAllCategoryDropdowns();
    }

    // 刪除自訂分類
    function deleteCustomCategory(catName) {
      // 避免 iframe 內建原生 confirm 被擋，採直接確認，或彈窗提示
      const isConfirmed = confirm(`確定要刪除「${catName}」這個分類嗎？\n刪除後，該分類將不再出現在下拉選單中。`);
      if (!isConfirmed) return;

      delete categoryLocationMap[catName];
      try { localStorage.setItem('DocumentDB_M365_CategoryPreferences_' + M365_CONFIG.tenantId, JSON.stringify(categoryLocationMap)); } catch { showToast('此瀏覽器無法保存個人分類設定。','error'); return; }
      showToast(`自訂分類「${catName}」已刪除`, 'info');
      populateSettingsForm();
      updateAllCategoryDropdowns();
    }

    // 儲存櫃位變更
    function handleSettingsSave(event) {
      event.preventDefault();
      
      // 遍歷所有動態產生的輸入框取得修改後的櫃位
      const inputs = document.querySelectorAll('#settingsFormContainer input[data-category]');
      inputs.forEach(input => {
        const cat = input.getAttribute('data-category');
        const loc = input.value.trim().toUpperCase();
        if (cat && loc) {
          categoryLocationMap[cat] = loc;
        }
      });

      // 保存至 localStorage 中實現持久化
      try { localStorage.setItem('DocumentDB_M365_CategoryPreferences_' + M365_CONFIG.tenantId, JSON.stringify(categoryLocationMap)); } catch { showToast('此瀏覽器無法保存個人分類設定。','error'); return; }
      
      // 同步更新當前單筆新增的位置
      const singleCat = document.getElementById('singleCategory').value;
      document.getElementById('singleLocation').value = categoryLocationMap[singleCat] || "";

      showToast("此瀏覽器的個人預設已儲存。", "success");
      updateAllCategoryDropdowns();
    }

    // 動態同步更新全系統所有關聯分類的下拉選單
    function updateAllCategoryDropdowns(){
      const categories=[...new Set([...Object.keys(categoryLocationMap),...allDocuments.map(d=>d.category)])];
      for(const id of ['filterCategory','singleCategory','editCategory']){
        const select=document.getElementById(id);if(!select)continue;
        const old=select.value;select.replaceChildren();
        if(id==='filterCategory')select.add(new Option('全部分類',''));
        for(const cat of categories)select.add(new Option(cat,cat));
        if(categories.includes(old))select.value=old;
        else if(id==='filterCategory')select.value='';
      }
    }

    // 單筆新增類別更動時觸發自動櫃位帶入
    window.handleSingleCategoryChange = function(category) {
      if (categoryLocationMap[category]) {
        document.getElementById('singleLocation').value = categoryLocationMap[category];
      }
    };

    // =======================================
    // 4. 檔案進階搜尋與渲染控制 (Grid Card)
    // =======================================
    function switchTab(tabId) {
      document.querySelectorAll('.tab-content').forEach(el => el.classList.add('hidden'));
      document.querySelectorAll('header button').forEach(el => {
        if (el.id && el.id.startsWith('btn-')) {
          el.classList.remove('border-blue-500', 'text-blue-400');
          el.classList.add('border-transparent', 'text-slate-400');
        }
      });

      document.getElementById(tabId).classList.remove('hidden');
      const activeBtn = document.getElementById('btn-' + tabId);
      if (activeBtn) {
        activeBtn.classList.remove('border-transparent', 'text-slate-400');
        activeBtn.classList.add('border-blue-500', 'text-blue-400');
      }

      if (tabId === 'statsTab') {
        fetchAndRenderStats();
      } else if (tabId === 'settingsTab') {
        populateSettingsForm();
      }
    }

    // 搜尋防抖
    function debouncedSearch() {
      clearTimeout(searchTimeout);
      searchTimeout = setTimeout(() => {
        fetchDocuments();
      }, 450);
    }

    // 清除所有進階篩選條件
    function clearSearch() {
      document.getElementById('searchInput').value = "";
      document.getElementById('filterCategory').value = "";
      document.getElementById('filterStartDate').value = "";
      document.getElementById('filterEndDate').value = "";
      fetchDocuments();
    }

    // 核心數據讀取 (進階多條件呼叫)
    function fetchDocuments() {
      const keyword = document.getElementById('searchInput').value;
      const category = document.getElementById('filterCategory').value;
      const startDate = document.getElementById('filterStartDate').value;
      const endDate = document.getElementById('filterEndDate').value;

      toggleLoading(true, "正在篩選關聯檔案...");

      appApi.run
        .withSuccessHandler(function(data) {
          toggleLoading(false);
          allDocuments = data;
          updateAllCategoryDropdowns();
          renderDocumentGrid(data);
        })
        .withFailureHandler(function(error) {
          toggleLoading(false);
          showToast(error.message, 'error');
        })
        .searchDocuments(keyword, category, startDate, endDate);
    }

    // 渲染響應式卡片
    function renderDocumentGrid(docs) {
      const grid=document.getElementById('documentGrid'), empty=document.getElementById('noDataView');
      grid.replaceChildren();grid.classList.toggle('hidden',!docs.length);empty.classList.toggle('hidden',!!docs.length);
      document.getElementById('searchCount').textContent=docs.length;
      for(const doc of docs){
        const card=document.createElement('article');
        card.className='bg-white rounded-xl shadow-sm border border-slate-100 overflow-hidden flex flex-col justify-between';
        card.innerHTML=`<div class="px-5 py-3 bg-slate-50 text-xs text-slate-500 break-all" data-field="docId"></div>
          <div class="p-5 space-y-3"><h4 class="font-bold text-lg break-words" data-field="docName"></h4>
          <span class="text-xs bg-blue-50 text-blue-700 rounded px-2 py-1" data-field="category"></span>
          <div class="bg-amber-50 rounded-lg p-3 text-sm"><p>收藏位置：<strong data-field="location"></strong></p><p>檔案日期：<span data-field="fileDate"></span></p></div>
          <p class="text-xs text-slate-500">上傳者：<span data-field="uploaderName"></span></p><p class="text-xs text-slate-500" data-field="uploadTime"></p></div>
          <div class="px-5 py-3 border-t flex gap-2"><button class="attachment px-3 py-2 rounded bg-blue-600 text-white text-xs">開啟附件</button><button class="edit px-3 py-2 rounded bg-slate-100 text-xs">編輯</button><button class="delete px-3 py-2 rounded bg-rose-50 text-rose-600 text-xs">刪除記錄</button></div>`;
        card.querySelectorAll('[data-field]').forEach(el=>el.textContent=doc[el.dataset.field]);
        const attachment=card.querySelector('.attachment');attachment.disabled=!doc.fileUrl;if(!doc.fileUrl)attachment.textContent='無附件';
        attachment.addEventListener('click',()=>previewAttachment(doc.fileUrl,doc.docName));
        card.querySelector('.edit').addEventListener('click',()=>openEditModal(doc.docId,doc.docName,doc.category,doc.location,doc.fileDate));
        card.querySelector('.delete').addEventListener('click',()=>confirmDeleteDocument(doc.docId));
        grid.appendChild(card);
      }
    }

    function downloadBase64CSV(base64Data, fileName) {
      try {
        const cleanBase64 = base64Data.replace(/\s/g, '');
        const byteCharacters = atob(cleanBase64);
        const byteNumbers = new Array(byteCharacters.length);
        
        for (let i = 0; i < byteCharacters.length; i++) {
          byteNumbers[i] = byteCharacters.charCodeAt(i);
        }
        
        const byteArray = new Uint8Array(byteNumbers);
        
        let finalArray = byteArray;
        const hasBom = byteArray[0] === 0xEF && byteArray[1] === 0xBB && byteArray[2] === 0xBF;
        if (!hasBom) {
          const bom = new Uint8Array([0xEF, 0xBB, 0xBF]);
          finalArray = new Uint8Array(bom.length + byteArray.length);
          finalArray.set(bom, 0);
          finalArray.set(byteArray, bom.length);
        }
        
        const blob = new Blob([finalArray], { type: 'text/csv;charset=utf-8;' });
        
        const link = document.createElement("a");
        const url = URL.createObjectURL(blob);
        link.href = url;
        link.download = fileName;
        document.body.appendChild(link);
        link.click();
        
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
        
        showToast("盤點清單下載成功（已置入繁體中文相容規格）", 'success');
      } catch (err) {
        showToast("解碼下載檔案失敗: " + err.message, 'error');
      }
    }

    // 導出盤點
    function handleExport() {


      const keyword = document.getElementById('searchInput').value;
      const category = document.getElementById('filterCategory').value;
      const startDate = document.getElementById('filterStartDate').value;
      const endDate = document.getElementById('filterEndDate').value;

      toggleLoading(true, "正根據當前篩選條件生成盤點清單...");

      appApi.run
        .withSuccessHandler(function(response) {
          toggleLoading(false);
          if (response && response.base64Data) {
            downloadBase64CSV(response.base64Data, response.fileName);
          }
        })
        .withFailureHandler(function(error) {
          toggleLoading(false);
          showToast("導出失敗: " + error.message, 'error');
        })
        .exportAuditList(keyword, category, startDate, endDate);
    }

    // =======================================
    // 6. 附件極速快顯線上預覽大腦 (支援 Base64)
    // =======================================
    function previewAttachment(fileUrl,docName){
      try { window.open(M365Core.safeUrl(fileUrl,false),'_blank','noopener,noreferrer'); }
      catch(e){showToast(e.message,'error');}
    }

    function closePreviewModal() {
      document.getElementById('previewModal').classList.add('hidden');
      document.getElementById('previewImage').src = "";
      document.getElementById('previewIframe').src = "";
    }

    // =======================================
    // 7. 模式 A：單筆檔案上傳控制
    // =======================================
    function updateFileNameDisplay(input, targetId) {
      const textElem = document.getElementById(targetId);
      if (input.files && input.files.length > 0) {
        textElem.textContent = input.files[0].name;
      } else {
        textElem.textContent = "點擊選取附件（單檔上限 10 MB）";
      }
    }

    function resetSingleForm() {
      document.getElementById('singleUploadForm').reset();
      document.getElementById('singleFileNameText').textContent = "點擊選取附件（單檔上限 10 MB）";
      // 重設表單時，同步帶入當前預設分類的設定櫃位
      const cat = document.getElementById('singleCategory').value;
      document.getElementById('singleLocation').value = categoryLocationMap[cat] || "";
    }

    function handleSingleUpload(event) {
      event.preventDefault();
      
      const metaData = {
        docName: document.getElementById('singleDocName').value,
        category: document.getElementById('singleCategory').value,
        location: document.getElementById('singleLocation').value,
        fileDate: document.getElementById('singleFileDate').value,
        uploaderName: currentUser.name
      };

      const fileInput = document.getElementById('singleFile');
      
      toggleLoading(true, "正在上傳附件與寫入 SharePoint 清單...");

      if (fileInput.files && fileInput.files.length > 0) {
        const file = fileInput.files[0];
        if(file.size > M365_CONFIG.maxFileBytes){toggleLoading(false);showToast('每個附件上限 10 MB。','error');return;}
        executeSingleUpload(metaData,file);
      } else {
        executeSingleUpload(metaData, null);
      }
    }

    function executeSingleUpload(metaData, base64File) {
      appApi.run
        .withSuccessHandler(function(response) {
          toggleLoading(false);
          if (response.success) {
            showToast(response.message, 'success');
            resetSingleForm();
            switchTab('searchTab');
            fetchDocuments();
          }
        })
        .withFailureHandler(function(error) {
          toggleLoading(false);
          showToast(error.message, 'error');
        })
        .uploadAndSaveDocument(metaData, base64File);
    }

    // =======================================
    // 8. 模式 B：多檔案批次建檔 (與 SharePoint 附件建檔)
    // =======================================
    function switchSubTab(subTabId) {
      document.getElementById('singleUploadView').classList.add('hidden');
      document.getElementById('bulkUploadView').classList.add('hidden');
      
      document.getElementById('btn-singleUpload').className = "py-2 px-4 text-sm font-semibold border-b-2 border-transparent text-slate-500 hover:text-slate-800 flex items-center gap-1.5 transition";
      document.getElementById('btn-bulkUpload').className = "py-2 px-4 text-sm font-semibold border-b-2 border-transparent text-slate-500 hover:text-slate-800 flex items-center gap-1.5 transition";

      if (subTabId === 'singleUpload') {
        document.getElementById('singleUploadView').classList.remove('hidden');
        document.getElementById('btn-singleUpload').className = "py-2 px-4 text-sm font-semibold border-b-2 border-blue-500 text-blue-600 flex items-center gap-1.5 transition";
      } else {
        document.getElementById('bulkUploadView').classList.remove('hidden');
        document.getElementById('btn-bulkUpload').className = "py-2 px-4 text-sm font-semibold border-b-2 border-blue-500 text-blue-600 flex items-center gap-1.5 transition";
      }
    }

    // --- CSV 檔案二進位安全載入與 RFC 4180 標準解析演算法 ---
    let parsedCsvDocs = [];

    function handleCSVSelect(input) {
      if (!input.files || input.files.length === 0) return;
      const file = input.files[0];
      if(file.size > 2*1024*1024){showToast('CSV 上限 2 MB，請分批匯入。','error');return;}
      document.getElementById('csvNameText').textContent = file.name;

      const encoding = document.getElementById('csvEncoding').value;
      toggleLoading(true, `正在以 ${encoding} 編碼格式讀取並解析繁體中文 CSV...`);

      const reader = new FileReader();
      
      reader.onload = function(e) {
        try {
          const arrayBuffer = e.target.result;
          const decoder = new TextDecoder(encoding);
          const decodedText = decoder.decode(arrayBuffer);
          
          parseCSVStandard(decodedText);
          toggleLoading(false);
        } catch (err) {
          toggleLoading(false);
          showToast("解碼讀取 CSV 失敗，請確認檔案結構與編碼：" + err.message, "error");
        }
      };
      
      reader.onerror=()=>{toggleLoading(false);showToast('無法讀取 CSV。','error');};
      reader.readAsArrayBuffer(file);
    }

    function parseCSVStandard(text){
      parsedCsvDocs=[];
      document.getElementById('csvPreviewArea').classList.add('hidden');
      const lines=[],row=[];let value='',quoted=false;
      text=text.replace(/^\uFEFF/,'');
      for(let i=0;i<text.length;i++){
        const ch=text[i];
        if(ch==='"'){if(quoted && text[i+1]==='"'){value+='"';i++;}else quoted=!quoted;}
        else if(ch===',' && !quoted){row.push(value);value='';}
        else if((ch==='\n'||ch==='\r') && !quoted){if(ch==='\r' && text[i+1]==='\n')i++;row.push(value);lines.push(row.splice(0));value='';}
        else value+=ch;
      }
      if(quoted)throw Error('CSV 引號未完整閉合。');
      if(value || row.length){row.push(value);lines.push(row);}
      if(lines.length<2)throw Error('CSV 需要標題與至少一筆資料。');
      const headers=lines[0].map(s=>s.trim());
      const index=label=>headers.indexOf(label);
      const name=index('文件名稱'),cat=index('分類'),location=index('收藏位置')>=0?index('收藏位置'):index('實體收藏位置'),date=index('檔案日期'),url=index('雲端檔案連結');
      if(name<0||cat<0||location<0)throw Error('CSV 必須有「文件名稱、分類、收藏位置」標題，可另加檔案日期與雲端檔案連結。');
      const docs=[];
      for(let i=1;i<lines.length;i++){
        const cells=lines[i];if(cells.every(s=>!s.trim()))continue;
        if(cells.length!==headers.length)throw Error(`CSV 第 ${i+1} 列欄位數量不符。`);
        const d={docName:cells[name].trim(),category:cells[cat].trim()||'其他',location:cells[location].trim(),fileDate:date>=0?cells[date].trim():'',fileUrl:url>=0?cells[url].trim():''};
        if(!d.docName||!d.location)throw Error(`CSV 第 ${i+1} 列缺少文件名稱或收藏位置。`);
        M365Core.dateSerial(d.fileDate);M365Core.safeUrl(d.fileUrl);
        docs.push(d);
      }
      if(!docs.length || docs.length>M365_CONFIG.maxBatchRows)throw Error('每次 CSV 匯入須為 1 至 100 筆。');
      parsedCsvDocs=docs;
      const tbody=document.getElementById('csvPreviewTableBody');tbody.replaceChildren();
      docs.slice(0,10).forEach(d=>{const tr=document.createElement('tr');for(const v of [d.docName,d.category,d.location]){const td=document.createElement('td');td.className='p-2 border-b';td.textContent=v;tr.append(td);}tbody.append(tr);});
      document.getElementById('csvCountText').textContent=docs.length;
      document.getElementById('csvPreviewArea').classList.remove('hidden');
      showToast(`已檢查 ${docs.length} 筆。匯入將產生新編號並記錄目前登入者，不保留 CSV 原編號與上傳者。`,'info');
    }

    function submitCSVImport() {
      if (parsedCsvDocs.length === 0) return;
      toggleLoading(true, "正在批次寫入 SharePoint 清單，請勿重複送出...");

      appApi.run
        .withSuccessHandler(function(response) {
          toggleLoading(false);
          if (response.success) {
            showToast(`匯入成功！共導入 ${response.importedCount} 筆資料`, 'success');
            document.getElementById('csvFileInput').value = "";
            document.getElementById('csvNameText').textContent = "選取並上傳 CSV 檔案";
            document.getElementById('csvPreviewArea').classList.add('hidden');
            parsedCsvDocs = [];

            switchTab('searchTab');
            fetchDocuments();
          }
        })
        .withFailureHandler(function(error) {
          toggleLoading(false);
          showToast("批次匯入失敗: " + error.message, 'error');
        })
        .batchImportDocuments(parsedCsvDocs, currentUser.name);
    }

    // --- 附件草稿處理 ---
    async function handleScannedFiles(input,isCamera=false){
      if(!input.files?.length)return;
      try{
        const files=Array.from(input.files);
        toggleLoading(true,'正在準備附件草稿，尚未上傳...');
        appApi.run.withSuccessHandler(files=>{toggleLoading(false);generateDraftEditCards(files);})
          .withFailureHandler(e=>{toggleLoading(false);showToast(e.message,'error');}).uploadScanFilesBatch(files);
      }catch(e){toggleLoading(false);showToast(e.message,'error');}
    }

    // 暫存編輯卡片中的類別更動時自動更新該草稿的收藏櫃位
    window.updateDraftLocation = function(draftId, category) {
      if (categoryLocationMap[category]) {
        document.getElementById(`${draftId}_location`).value = categoryLocationMap[category];
      }
    };

    // 生成暫存草稿編輯卡片 (整合全動態分類 options)
    function generateDraftEditCards(files){
      pendingDrafts=[];const container=document.getElementById('batchDraftContainer');container.replaceChildren();
      const categories=Object.keys(categoryLocationMap), defaultCat=categories.includes('合約')?'合約':categories[0];
      for(let i=0;i<Math.max(files.length,5);i++){
        const id='draft_'+i,file=files[i];pendingDrafts.push(id);
        const card=document.createElement('div');card.className='bg-white p-5 rounded-xl border space-y-3';
        card.innerHTML=`<p class="text-xs text-slate-500">草稿 ${i+1} · 確認建檔後才上傳</p><input type="hidden" id="${id}_fileUrl">
          <label class="block text-xs">文件名稱<input class="block border rounded p-2 w-full" id="${id}_name"></label>
          <label class="block text-xs">分類<select class="block border rounded p-2 w-full" id="${id}_category"></select></label>
          <label class="block text-xs">檔案日期<input type="date" class="block border rounded p-2 w-full" id="${id}_fileDate"></label>
          <label class="block text-xs">收藏位置<input class="block border rounded p-2 w-full" id="${id}_location"></label>`;
        container.append(card);
        document.getElementById(id+'_fileUrl').value=file?.url||'';
        document.getElementById(id+'_name').value=file?.name?.replace(/\.[^.]+$/,'')||'';
        const select=document.getElementById(id+'_category');categories.forEach(cat=>select.add(new Option(cat,cat)));select.value=defaultCat;
        select.addEventListener('change',()=>updateDraftLocation(id,select.value));
        document.getElementById(id+'_location').value=categoryLocationMap[defaultCat]||'';
      }
      document.getElementById('draftCount').textContent=pendingDrafts.length;
      document.getElementById('batchDraftArea').classList.remove('hidden');
    }

    function submitBatchDrafts() {
      const documentsArray = [];

      for (let i = 0; i < pendingDrafts.length; i++) {
        const id = pendingDrafts[i];
        const docName = document.getElementById(`${id}_name`).value.trim();
        const category = document.getElementById(`${id}_category`).value;
        const fileDate = document.getElementById(`${id}_fileDate`).value || "--";
        const location = document.getElementById(`${id}_location`).value.trim();
        const fileUrl = document.getElementById(`${id}_fileUrl`).value;

        if ((docName || fileUrl) && (!docName || !location)) { showToast('有草稿缺少文件名稱或收藏位置，請補齊後再建檔。','error'); return; }
        if (docName && location) {
          documentsArray.push({
            docName: docName,
            category: category,
            fileDate: fileDate,
            location: location,
            fileUrl: fileUrl
          });
        }
      }

      if (documentsArray.length === 0) {
        showToast("請至少填入 1 筆包含「文件名稱」與「實體收藏位置」的草稿資料。", 'error');
        return;
      }

      toggleLoading(true, "批次寫入資料庫建檔中，請稍後...");

      appApi.run
        .withSuccessHandler(function(response) {
          toggleLoading(false);
          if (response.success) {
            showToast(`草稿批次建檔成功！共新增 ${response.importedCount} 筆文件`, 'success');
            
            document.getElementById('scanFilesInput').value = "";
            document.getElementById('cameraFilesInput').value = "";
            document.getElementById('batchDraftArea').classList.add('hidden');
            pendingDrafts = [];

            switchTab('searchTab');
            fetchDocuments();
          }
        })
        .withFailureHandler(function(error) {
          toggleLoading(false);
          showToast(error.message, 'error');
        })
        .batchImportDocuments(documentsArray, currentUser.name);
    }

    // =======================================
    // 9. 編輯與物理刪除機制 (重構為 Custom Modal)
    // =======================================
    function openEditModal(docId, docName, category, location, fileDate) {
      editOriginal = structuredClone(allDocuments.find(d=>d.docId===docId));
      document.getElementById('editDocId').value = docId;
      document.getElementById('editDocName').value = docName;
      
      // 更新動態分類選單
      updateAllCategoryDropdowns();
      document.getElementById('editCategory').value = category;
      
      document.getElementById('editLocation').value = location;
      document.getElementById('editFileDate').value = fileDate === "--" ? "" : fileDate;

      document.getElementById('editModal').classList.remove('hidden');
    }

    function closeEditModal() {
      document.getElementById('editModal').classList.add('hidden');
    }

    function handleEditSubmit(event) {
      event.preventDefault();
      
      const docId = document.getElementById('editDocId').value;
      const updatedData = {
        docName: document.getElementById('editDocName').value,
        category: document.getElementById('editCategory').value,
        location: document.getElementById('editLocation').value,
        fileDate: document.getElementById('editFileDate').value
      };

      toggleLoading(true, "正在更新資料庫屬性...");

      appApi.run
        .withSuccessHandler(function(response) {
          toggleLoading(false);
          if (response.success) {
            showToast(response.message, 'success');
            closeEditModal();
            fetchDocuments();
          }
        })
        .withFailureHandler(function(error) {
          toggleLoading(false);
          showToast(error.message, 'error');
        })
        .updateDocument(docId, updatedData, editOriginal);
    }

    function confirmDeleteDocument(docId) {
      deleteOriginal = structuredClone(allDocuments.find(d=>d.docId===docId));
      deleteTargetId = docId;
      document.getElementById('deleteTargetIdDisplay').textContent = docId;
      document.getElementById('deleteConfirmModal').classList.remove('hidden');
    }

    function closeDeleteConfirmModal() {
      document.getElementById('deleteConfirmModal').classList.add('hidden');
      deleteTargetId = null;
    }

    function executeDeleteDocument() {
      if (!deleteTargetId) return;
      const docId = deleteTargetId;
      
      closeDeleteConfirmModal();
      toggleLoading(true, "正在刪除清單 文件記錄...");

      appApi.run
        .withSuccessHandler(function(response) {
          toggleLoading(false);
          if (response.success) {
            showToast(response.message, 'success');
            fetchDocuments();
          }
        })
        .withFailureHandler(function(error) {
          toggleLoading(false);
          showToast(error.message, 'error');
        })
        .deleteDocument(docId, deleteOriginal);
    }

    // --- 批次暫存取消重來控制 ---
    function triggerCancelBatchDrafts() {
      document.getElementById('cancelBatchConfirmModal').classList.remove('hidden');
    }

    function closeCancelBatchConfirmModal() {
      document.getElementById('cancelBatchConfirmModal').classList.add('hidden');
    }

    function executeCancelBatchDrafts() {
      closeCancelBatchConfirmModal();
      M365App.discardDraftAttachments();
      
      // 清除照片與檔案上傳 input
      document.getElementById('scanFilesInput').value = "";
      document.getElementById('cameraFilesInput').value = "";
      
      // 隱藏批次暫存編輯區
      document.getElementById('batchDraftArea').classList.add('hidden');
      
      // 重設全域變數與草稿卡片
      pendingDrafts = [];
      document.getElementById('batchDraftContainer').innerHTML = "";
      
      showToast("已成功取消並清空暫存草稿檔案，您可以重新拍照或上傳檔案。", "info");
    }

    // =======================================
    // 10. 檔案分類統計 (Chart.js + 點擊快篩聯動)
    // =======================================
    function fetchAndRenderStats() {
      appApi.run
        .withSuccessHandler(function(data) {
          document.getElementById('statTotal').textContent = data.totalFiles;
          document.getElementById('statToday').textContent = data.todayCount;
          document.getElementById('statCategories').textContent = data.uniqueCategories;

          renderCharts(data.chartData);
        })
        .withFailureHandler(function(error) {
          showToast("載入統計資訊失敗: " + error.message, 'error');
        })
        .getCategoryStats();
    }

    function renderCharts(chartData) {
      const colors = [
        '#2563eb', // blue
        '#ef4444', // red
        '#f59e0b', // amber
        '#10b981', // emerald
        '#06b6d4', // cyan
        '#8b5cf6', // purple
        '#64748b', // slate
        '#cbd5e1'  // secondary
      ];

      const ctxDoughnut = document.getElementById('doughnutChart').getContext('2d');
      if (doughnutChartInstance) {
        doughnutChartInstance.destroy();
      }

      doughnutChartInstance = new Chart(ctxDoughnut, {
        type: 'doughnut',
        data: {
          labels: chartData.labels,
          datasets: [{
            data: chartData.values,
            backgroundColor: colors.slice(0, chartData.labels.length),
            borderWidth: 2,
            borderColor: '#ffffff'
          }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            legend: {
              position: 'bottom',
              labels: {
                boxWidth: 12,
                font: { size: 12, weight: 'bold' }
              }
            }
          },
          onClick: (event, elements) => {
            if (elements.length > 0) {
              const clickedIndex = elements[0].index;
              const category = chartData.labels[clickedIndex];
              triggerChartFastFilter(category);
            }
          }
        }
      });

      const ctxBar = document.getElementById('barChart').getContext('2d');
      if (barChartInstance) {
        barChartInstance.destroy();
      }

      barChartInstance = new Chart(ctxBar, {
        type: 'bar',
        data: {
          labels: chartData.labels,
          datasets: [{
            label: '各分類文件累計數',
            data: chartData.values,
            backgroundColor: 'rgba(37, 99, 235, 0.85)',
            hoverBackgroundColor: 'rgba(37, 99, 235, 1)',
            borderRadius: 6,
            borderWidth: 0
          }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          scales: {
            y: {
              beginAtZero: true,
              ticks: { precision: 0 }
            }
          },
          plugins: {
            legend: { display: false }
          },
          onClick: (event, elements) => {
            if (elements.length > 0) {
              const clickedIndex = elements[0].index;
              const category = chartData.labels[clickedIndex];
              triggerChartFastFilter(category);
            }
          }
        }
      });
    }

    function triggerChartFastFilter(categoryName) {
      document.getElementById('filterCategory').value = categoryName;
      switchTab('searchTab');
      fetchDocuments();
      showToast(`已自動為您切換至搜尋分頁並篩選「${categoryName}」！`, 'success');
    }
M365App.init();
