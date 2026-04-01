<script>
  const GAS_URL = "https://script.google.com/macros/s/AKfycbyqRrQGSF9pJuDsM1ueh5d9XyeOTL-RXHGIDrLMmEqXbisD_aQUF_W3aaCeiOmPVoRK/exec"; // 記得要重新新建部署並設為「任何人」

  // 封裝 Fetch，模擬 google.script.run 的行為
  async function callGAS(functionName, args = []) {
      const url = `${GAS_URL}?action=${functionName}`;
    
      // 判斷是讀取 (GET) 還是 寫入 (POST)
      // GAS 的 Web App 對 POST 有 CORS 限制，通常建議簡單操作都用 GET 傳參數
      // 或是將資料封裝成 Base64/JSON 傳入
      try {
          const response = await fetch(url, {
              method: 'POST', // 使用 POST 才能傳送較大的資料 (如 TinyMCE 內容)
              body: JSON.stringify({
                  method: functionName,
                  data: args
              })
          });
          return await response.json();
      } catch (e) {
          console.error("GAS連線失敗:", e);
          throw e;
      }
  }
  let commentLookup = {};
  // 阻止 Bootstrap 搶焦點
  document.addEventListener('focusin', function (e) {
    if (e.target.closest('.tox-tinymce, .tox-tinymce-aux, .moxman-window, .tam-assetmanager-root') !== null) {
      e.stopImmediatePropagation();
    }
  });

  let user = JSON.parse(localStorage.getItem('tb_user'));
  let currentThreadId = null;
  let appData = { plans: [], threads: [], announcements: [], versions: [], currentThread: null, myData: {} };
  let pendingTarget = null;
  let onMsgClose = null;
  
  const ITEMS_PER_PAGE = 12;
  const FORUM_PER_PAGE = 10;
  const HOME_PER_PAGE = 5;

  let curPage = { plans: 1, forum: 1, profile_plans: 1, profile_archived: 1, profile_saved: 1, profile_threads: 1, profile_replies: 1, announcements: 1, versions: 1 };
  
  let filterState = { category: 'all', grades: [], planSearch: '', forumSearch: '' };
  let profileFilterState = { category: 'all', grades: [] }; 
  let profileTab = 'plans'; 

  const initParams = JSON.parse(<?= initParams ?>);
  const appUrl = "<?= appUrl ?>"; 

  const msgModalEl = document.getElementById('msgModal');
  msgModalEl.addEventListener('hidden.bs.modal', () => { if (onMsgClose) { onMsgClose(); onMsgClose = null; } });
  
  function showMsg(title, text, type = 'info', callback = null) {
    onMsgClose = callback; document.getElementById('msg-title').textContent = title; document.getElementById('msg-body').innerHTML = text;
    const header = document.getElementById('msg-header'); header.className = 'modal-header text-white ' + (type==='success'?'bg-success':type==='error'?'bg-danger':type==='warning'?'bg-warning text-dark':'bg-secondary');
    header.querySelector('.btn-close').className = 'btn-close ' + (type==='warning'?'':'btn-close-white'); new bootstrap.Modal(msgModalEl).show();
  }
  function onFailure(error) { 
      toggleLoading(false);
      showMsg('系統錯誤', error.message, 'error'); 
  }
  
  function toggleLoading(show) {
      document.getElementById('loading-overlay').classList.toggle('hidden', !show);
  }

  let onConfirmAction = null;
  const confirmModalEl = new bootstrap.Modal(document.getElementById('confirmModal'));
  function showConfirm(title, text, callback) {
    document.getElementById('confirm-title').textContent = title; document.getElementById('confirm-body').innerHTML = text; onConfirmAction = callback; confirmModalEl.show();
  }
  document.getElementById('btn-confirm-yes').addEventListener('click', function() { confirmModalEl.hide(); if(onConfirmAction) onConfirmAction(); });

  function forceAuth() { showMsg('權限提示', '此內容僅限會員觀看。<br>請先登入或註冊會員。', 'warning', () => { showAuthModal('login'); }); }

  function initGradeChecks(containerId, onChangeFunc) {
    const container = document.getElementById(containerId);
    let html = '';
    for(let i=1; i<=6; i++) { html += `<div class="form-check form-check-inline"><input class="form-check-input" type="checkbox" value="${i}" onchange="${onChangeFunc}()"> <label class="form-check-label">${i}年級</label></div>`; }
    container.innerHTML = html;
  }

  function initGradeFilterButtons(containerId, stateObj, renderFunc) {
      const container = document.getElementById(containerId);
      if(!container) return;
      let html = '';
      for(let i=1; i<=6; i++) {
          html += `<button type="button" class="btn btn-outline-secondary" onclick="toggleGradeFilter('${containerId}', ${i})">${i}</button>`;
      }
      container.innerHTML = html;
      
      window.toggleGradeFilter = function(cId, grade) {
          const isProfile = cId === 'profile-grade-filter';
          const state = isProfile ? profileFilterState : filterState;
          const idx = state.grades.indexOf(grade);
          if (idx === -1) state.grades.push(grade); else state.grades.splice(idx, 1);
          
          const btns = document.querySelectorAll(`#${cId} button`);
          btns.forEach((btn, index) => {
              const g = index + 1;
              if (state.grades.includes(g)) {
                  btn.classList.remove('btn-outline-secondary');
                  btn.classList.add('btn-secondary');
              } else {
                  btn.classList.add('btn-outline-secondary');
                  btn.classList.remove('btn-secondary');
              }
          });
          
          if(isProfile) { curPage['profile_' + profileTab] = 1; renderProfileList(); } 
          else { curPage.plans = 1; renderPlans(); }
      };
  }

  function cleanUrl() { if (window.history.replaceState) google.script.history.replace(null, {}, null); }
  document.getElementById('viewPlanModal').addEventListener('hidden.bs.modal', cleanUrl);

  function initTinyMCE() {
    tinymce.init({
      selector: '#plan-content, #thread-content, #comment-input',
      menubar: false,
      plugins: 'lists link',
      toolbar: 'bold italic underline | bullist numlist | link removeformat',
      height: 200,
      setup: function (editor) { editor.on('change', function () { editor.save(); }); }
    });
  }

  window.onload = function() { 
  updateAuthUI(); 
  initGradeChecks('plan-grades-check', 'void(0)');
  initGradeFilterButtons('plan-grade-filter', filterState, renderPlans);
  initGradeFilterButtons('profile-grade-filter', profileFilterState, renderProfileList);
  initTinyMCE();
  
  // ★ 這行一定要有，預覽功能才會啟動
  bindRefPreviews();

  if (initParams.view && initParams.id) {
       if (initParams.view === 'thread_detail') {
          if (!user) { pendingTarget = initParams; forceAuth(); return; }
          switchView('forum'); loadThreadDetail(initParams.id);
       } else if (initParams.view === 'plan_detail') {
          switchView('plans'); jumpToPlan(initParams.id); 
       }
    } else { switchView('home'); }
  };

  function copyShareLink(view, id) {
    let shareTitle = "";
    let shareText = "";
    const link = `${appUrl}?view=${view}&id=${id}`;

    // 根據 view 類型抓取當前頁面的標題
    if (view === 'plan_detail') {
      // 試著從 appData 或是 Modal 中抓標題
      const plan = appData.plans.find(p => p.id == id);
      shareTitle = plan ? plan.title : "優秀教案資源";
      shareText = `【碳減活寶桌遊教案分享】\n推薦一個很棒的資源：${shareTitle}\n點擊連結查看詳情：\n${link}`;
    } else if (view === 'thread_detail') {
      // 從討論區列表或是詳情頁抓標題
      const thread = appData.threads.find(t => t.id == id);
      shareTitle = thread ? thread.title : "精彩討論內容";
      shareText = `【碳減活寶桌遊教案討論區】\n快來看看這則熱門討論：${shareTitle}\n加入對話：\n${link}`;
    } else {
      // 預設（首頁或其他）
      shareText = `歡迎來到碳減活寶桌遊教案資源網：\n${link}`;
    }

    // 執行複製
    navigator.clipboard.writeText(shareText).then(() => {
      showMsg('分享連結已複製', '包含標題與網址的分享訊息已存至剪貼簿，您可以直接貼上傳給他人！', 'success');
    }).catch(err => {
      console.error('複製失敗', err);
      // 備用方案：如果瀏覽器不支援，至少給出網址
      showMsg('複製失敗', `請手動複製連結：<br>${link}`, 'warning');
    });
  }
  function switchView(viewName) {
    // --- 新增：自動收合漢堡選單 ---
    const navbarCollapse = document.getElementById('navbarNav');
    if (navbarCollapse.classList.contains('show')) {
      // 使用 Bootstrap 原生的 Collapse 實例來收合
      const bsCollapse = bootstrap.Collapse.getInstance(navbarCollapse) || new bootstrap.Collapse(navbarCollapse);
      bsCollapse.hide();
    }
    // 強制移除可能殘留的 Modal 鎖定狀態
   document.body.classList.remove('modal-open');
   document.body.style.overflow = '';
   document.body.style.paddingRight = '';
   if ((viewName === 'forum' || viewName === 'profile') && !user) { forceAuth(); return; } ['home', 'plans', 'forum', 'profile'].forEach(v => { document.getElementById('view-'+v).classList.add('hidden'); document.getElementById('nav-'+v).classList.remove('active'); }); document.getElementById('view-'+viewName).classList.remove('hidden'); document.getElementById('nav-'+viewName).classList.add('active'); cleanUrl(); if(viewName==='home') loadHome(); if(viewName==='plans') loadPlans(); if(viewName==='forum') { document.getElementById('forum-list-view').classList.remove('hidden'); document.getElementById('forum-detail-view').classList.add('hidden'); loadForum(); } if(viewName==='profile') loadProfile(); }
  function getActionMenu(type, id, currentStatus, inModal=false) { const isArchived = currentStatus === 'archived'; const isClosed = currentStatus === 'closed'; let extraItems = ''; if(type === 'thread') extraItems += `<li><button class="dropdown-item" type="button" onclick="updateStatus('${type}', '${id}', '${isClosed?'active':'closed'}')">${isClosed?'重新開啟':'關閉討論'}</button></li>`; const editOnClick = inModal ? `openEdit('${type}', '${id}', true)` : `openEdit('${type}', '${id}')`; return `<div class="dropdown d-inline-block ms-2"><button class="btn btn-sm btn-link text-secondary p-0" type="button" data-bs-toggle="dropdown"><i class="bi bi-three-dots-vertical"></i></button><ul class="dropdown-menu dropdown-menu-end"><li><button class="dropdown-item" type="button" onclick="${editOnClick}">編輯</button></li><li><button class="dropdown-item" type="button" onclick="updateStatus('${type}', '${id}', '${isArchived?'active':'archived'}')">${isArchived?'取消典藏':'典藏'}</button></li>${extraItems}<li><hr class="dropdown-divider"></li><li><button class="dropdown-item text-danger" type="button" onclick="updateStatus('${type}', '${id}', 'deleted')">刪除</button></li></ul></div>`; }
  function renderPagination(totalItems, currentPage, targetId, onPageChange, pageSize) { const totalPages = Math.ceil(totalItems / pageSize); const container = document.getElementById(targetId); if (totalPages <= 1) { container.innerHTML = ''; return; } let html = '<ul class="pagination pagination-sm">'; html += `<li class="page-item ${currentPage === 1 ? 'disabled' : ''}"><button class="page-link" onclick="${onPageChange}(${currentPage - 1})">上一頁</button></li>`; for (let i = 1; i <= totalPages; i++) { html += `<li class="page-item ${i === currentPage ? 'active' : ''}"><button class="page-link" onclick="${onPageChange}(${i})">${i}</button></li>`; } html += `<li class="page-item ${currentPage === totalPages ? 'disabled' : ''}"><button class="page-link" onclick="${onPageChange}(${currentPage + 1})">下一頁</button></li></ul>`; container.innerHTML = html; }
  
  function loadHome() { google.script.run.withSuccessHandler(res => { if(res.error) return onFailure(new Error(res.error)); appData.announcements = res.announcements; appData.versions = res.versions; renderAnnouncements(); renderVersions(); }).withFailureHandler(onFailure).getData('home'); }
  function showAnnounce(id) { const a = appData.announcements.find(x => x.id == id); if(!a) return; document.getElementById('ann-title').textContent = a.title; document.getElementById('ann-date').textContent = a.startStr; function linkify(text) { return text.replace(/(https?:\/\/[^\s]+)/g, '<a href="$1" target="_blank">$1</a>'); } let contentHtml = linkify(esc(a.body)).replace(/\n/g, '<br>'); if(a.imgsRaw) { const urls = a.imgsRaw.split(',').map(u=>u.trim()).filter(u=>u); contentHtml += '<div class="mt-3">' + urls.map(u => `<img src="${u}" class="img-fluid rounded mb-2" style="max-height:300px;">`).join('') + '</div>'; } document.getElementById('ann-body').innerHTML = contentHtml; new bootstrap.Modal(document.getElementById('announceModal')).show(); }
  function renderAnnouncements() { const list = document.getElementById('home-announce-list'); const start = (curPage.announcements - 1) * HOME_PER_PAGE; const items = appData.announcements.slice(start, start + HOME_PER_PAGE); renderPagination(appData.announcements.length, curPage.announcements, 'announce-pagination', 'changeAnnouncePage', HOME_PER_PAGE); if (items.length === 0) list.innerHTML = '<div class="p-3 text-center">目前無公告</div>'; else list.innerHTML = items.map(a => `<div class="list-group-item list-group-item-action p-3 cursor-pointer ${a.pin?'pinned':''}" onclick="showAnnounce('${a.id}')"><div class="d-flex w-100 justify-content-between"><h6 class="mb-1 title text-truncate">${a.pin?'<span class="badge bg-danger me-1">置頂</span>':''} ${esc(a.title)}</h6><small class="text-muted">${a.startStr}</small></div></div>`).join(''); }
  function changeAnnouncePage(p) { curPage.announcements = p; renderAnnouncements(); }
  function renderVersions() {
    const list = document.getElementById('home-version-list');
    const start = (curPage.versions - 1) * HOME_PER_PAGE;
    const items = (appData.versions || []).slice(start, start + HOME_PER_PAGE);
    renderPagination((appData.versions || []).length, curPage.versions, 'version-pagination', 'changeVersionPage', HOME_PER_PAGE);
  
    if (items.length === 0) {
      list.innerHTML = '<p class="text-muted p-3">尚無更新紀錄</p>';
    } else {
      let vHtml = '<table class="table table-hover mb-0"><thead><tr><th>版本</th><th>日期</th><th>內容</th></tr></thead><tbody>';
      vHtml += items.map(v => {
        // --- 核心修正：格式化日期，只取前 10 碼 (YYYY/MM/DD) ---
        let dateDisplay = "";
        if (v.date === 'Coming Soon') {
          dateDisplay = '<span class="badge bg-warning text-dark">Coming Soon</span>';
        } else if (v.date) {
          // 即使後端傳來的是 2026/02/10 00:00:00，這裡也會只留下 2026/02/10
          dateDisplay = v.date.toString().substring(0, 10);
        }
        
        const verBadgeClass = (v.date === 'Coming Soon') ? 'bg-secondary' : 'bg-success';
        
        return `<tr>
          <td><span class="badge ${verBadgeClass}">${esc(v.version)}</span></td>
          <td style="white-space: nowrap;">${dateDisplay}</td>
          <td>${esc(v.content)}</td>
        </tr>`;
      }).join('');
      vHtml += '</tbody></table>';
      list.innerHTML = vHtml;
    }
  }
  function changeVersionPage(p) { curPage.versions = p; renderVersions(); }
  function loadPlans() { const btnAdd = document.getElementById('btn-add-plan'); if(user) { btnAdd.classList.remove('hidden'); document.getElementById('guest-plan-hint').classList.add('hidden'); } else { btnAdd.classList.add('hidden'); document.getElementById('guest-plan-hint').classList.remove('hidden'); } const list = document.getElementById('plans-list'); if(list.innerHTML.trim() === '') list.innerHTML = '<div class="text-center w-100"><div class="spinner-border text-success"></div></div>'; google.script.run.withSuccessHandler(res => { if(res.error) return onFailure(new Error(res.error)); appData.plans = res.plans; renderPlanStats(res.stats); filterPlans('all', document.querySelector('.btn-group button:first-child')); }).withFailureHandler(onFailure).getData('plans', user ? user.username : null); }
  function renderPlanStats(stats) { const div = document.getElementById('plan-stats-bar'); if(!stats) { div.classList.add('hidden'); return; } div.classList.remove('hidden'); div.innerHTML = `<div class="col-md-6"><div class="card stat-card stat-card-blue h-100 shadow-sm p-3"><i class="bi bi-file-earmark-word stat-icon"></i><div class="d-flex align-items-center mb-2"><h4 class="mb-0 fw-bold">教案</h4></div><div class="row text-center mt-2"><div class="col-6 border-end border-white border-opacity-25"><small class="text-white-50">公開</small><div class="stat-num">${stats.wordPub}</div></div><div class="col-6"><small class="text-white-50">會員</small><div class="stat-num">${stats.wordMem}</div></div></div></div></div><div class="col-md-6"><div class="card stat-card stat-card-red h-100 shadow-sm p-3"><i class="bi bi-file-earmark-slides stat-icon"></i><div class="d-flex align-items-center mb-2"><h4 class="mb-0 fw-bold">簡報</h4></div><div class="row text-center mt-2"><div class="col-6 border-end border-white border-opacity-25"><small class="text-white-50">公開</small><div class="stat-num">${stats.pptPub}</div></div><div class="col-6"><small class="text-white-50">會員</small><div class="stat-num">${stats.pptMem}</div></div></div></div></div>`; }
  function filterPlans(type, btn) { const group = btn.parentElement; Array.from(group.children).forEach(b => { const txt = b.textContent.trim(); b.className = 'btn'; if (txt === '全部') b.classList.add('btn-outline-secondary'); else if (txt === '教案') b.classList.add('btn-outline-word'); else if (txt === '簡報') b.classList.add('btn-outline-ppt'); }); if (type === 'all') btn.className = 'btn btn-secondary active'; else if (type === '教案') btn.className = 'btn btn-word active'; else if (type === '簡報') btn.className = 'btn btn-ppt active'; filterState.category = type; curPage.plans = 1; renderPlans(); }
  function setProfileFilter(cat, btn) { const group = btn.parentElement; Array.from(group.children).forEach(b => { const txt = b.textContent.trim(); b.className = 'btn'; if (txt === '全部') b.classList.add('btn-outline-secondary'); else if (txt === '教案') b.classList.add('btn-outline-word'); else if (txt === '簡報') b.classList.add('btn-outline-ppt'); }); if (cat === 'all') btn.className = 'btn btn-secondary active'; else if (cat === '教案') btn.className = 'btn btn-word active'; else if (cat === '簡報') btn.className = 'btn btn-ppt active'; profileFilterCat = cat; curPage['profile_' + profileTab] = 1; renderProfileList(); }
  function handlePlanSearch(val) { filterState.planSearch = val.toLowerCase(); curPage.plans = 1; renderPlans(); }

  // ★ 嚴格驗證：檔案上傳 ★
  function submitPlan() { 
      const grades = Array.from(document.querySelectorAll('#plan-grades-check input:checked')).map(c=>parseInt(c.value)); 
      const title = document.getElementById('plan-title').value.trim();
      const link = document.getElementById('plan-link').value.trim();
      
      const editor = tinymce.get('plan-content');
      const content = editor ? editor.getContent() : ''; 
      // 取得純文字 (濾掉HTML標籤) 並去除 &nbsp; 
      const textContent = editor ? editor.getContent({format: 'text'}).replace(/\u00a0/g, ' ').trim() : '';

      // 檢查 1: 標題必填
      if (title.length === 0) return showMsg('資料不完整', '請輸入<b>標題</b>', 'warning');
      
      // 檢查 2: 年級必選
      if (grades.length === 0) return showMsg('資料不完整', '請至少勾選一個<b>適用年級</b>', 'warning');
      
      // 檢查 3: 內容簡介 OR 連結 擇一 (內容若有圖片也算有內容)
      const hasContent = textContent.length > 0 || content.includes('<img') || content.includes('<iframe');
      if (!hasContent && link.length === 0) return showMsg('資料不完整', '「內容簡介」與「雲端連結」請至少填寫一項', 'warning');

      const id = document.getElementById('plan-id').value; 
      const data = { author: user.username, type: id?'edit':'new', access: document.getElementById('plan-access').value, category: document.getElementById('plan-category').value, title: title, content: content, link: link, grades: grades }; 
      const handler = id ? 'editPostContent' : 'createPlan'; const args = id ? [id, data, user.username] : [data]; 
      toggleLoading(true);
      google.script.run.withSuccessHandler(res => { toggleLoading(false); bootstrap.Modal.getInstance(document.getElementById('planModal')).hide(); showMsg('成功', id?'已更新':'已發布', 'success'); loadPlans(); document.getElementById('plan-id').value = ''; document.getElementById('plan-modal-title').textContent = '上傳檔案'; }).withFailureHandler(onFailure)[handler](...args); 
  }

  // ★ 嚴格驗證：發起討論 ★
  function submitThread() { 
      const id = document.getElementById('thread-id').value; 
      const title = document.getElementById('thread-title').value.trim();
      const content = tinymce.get('thread-content').getContent(); 
      
      if (title.length === 0) return showMsg('資料不完整', '請輸入<b>討論標題</b>', 'warning');

      const data = { author: user.username, title: title, content: content }; 
      const handler = id ? 'editThreadContent' : 'createThread'; const args = id ? [id, data, user.username] : [data]; 
      toggleLoading(true);
      google.script.run.withSuccessHandler(() => { toggleLoading(false); bootstrap.Modal.getInstance(document.getElementById('threadModal')).hide(); if(id && currentThreadId) loadThreadDetail(currentThreadId); else loadForum(); document.getElementById('thread-id').value = ''; }).withFailureHandler(onFailure)[handler](...args); 
  }

  // ★ 嚴格驗證：留言 ★
  function submitComment() {
      const editor = tinymce.get('comment-input');
      const content = editor.getContent();
      
      // ★ 關鍵：檢查這裡有沒有正確取得標籤內容 ★
      // 我們從 reply-indicator 的文字裡抓，或是設定一個全域變數存它
      let replyTag = "";
      const indicator = document.getElementById('reply-indicator');
      if (!indicator.classList.contains('hidden')) {
          // 從 "正在回覆 B1-1 ..." 中擷取出 "B1-1"
          const match = indicator.textContent.match(/B\d+(?:-\d+)?/);
          if (match) replyTag = match[0];
      }
  
      const data = { 
          threadId: currentThreadId, 
          author: user.username, 
          content: content, 
          parentFloor: replyToFloor, 
          replyTo: replyTag  // ★ 確保這裡把 B1 或 B1-1 傳過去
      };
  
      toggleLoading(true);
      google.script.run.withSuccessHandler(() => {
          toggleLoading(false);
          editor.setContent('');
          replyToFloor = 0;
          indicator.classList.add('hidden');
          loadThreadDetail(currentThreadId);
      }).withFailureHandler(onFailure).createComment(data);
  }

  function renderPlans() {
      try {
        const container = document.getElementById('plans-list'); 

        // 1. 執行過濾邏輯
        let displayPlans = appData.plans.filter(p => { 
            if (filterState.category !== 'all' && p.category !== filterState.category) return false; 
            if (filterState.grades.length > 0) { 
                const hasGrade = p.grades && p.grades.some(g => filterState.grades.includes(g)); 
                if (!hasGrade) return false; 
            }
            if (filterState.planSearch) { 
                const term = filterState.planSearch; 
                const match = p.title.toLowerCase().includes(term) || 
                              p.content.toLowerCase().includes(term) || 
                              p.authorName.toLowerCase().includes(term); 
                if (!match) return false; 
            }
            return true; 
        });

        // 2. 分頁計算
        const start = (curPage.plans - 1) * ITEMS_PER_PAGE; 
        const paginatedPlans = displayPlans.slice(start, start + ITEMS_PER_PAGE); 
        renderPagination(displayPlans.length, curPage.plans, 'plans-pagination', 'changePlanPage', ITEMS_PER_PAGE);

        // 3. 空資料處理
        if(paginatedPlans.length === 0) { 
            container.innerHTML = '<p class="text-center text-muted w-100 py-5">暫無符合條件的檔案</p>'; 
            return; 
        }

        // 4. 渲染卡片內容
        container.innerHTML = paginatedPlans.map(p => {
          const isMine = user && user.username === p.author; 
          const isSaved = user && user.saved && user.saved.includes(String(p.id)); 

          // 標籤樣式判斷
          const catClass = p.category === '簡報' ? 'badge-ppt' : 'badge-word'; 
          const accClass = p.access === 'public' ? 'badge-public' : 'badge-member';
          const typeBadge = `<span class="badge ${catClass} me-1">${p.category}</span>`;
          const accessBadge = `<span class="badge ${accClass}">${p.access==='public'?'公開':'會員'}</span>`;

          // 年級標籤
          let gradeTags = ''; 
          if(p.grades && p.grades.length > 0) { 
              gradeTags = '<div class="mt-1 small text-secondary">適用：' + p.grades.map(g=>g+'年級').join('、') + '</div>'; 
          }

          // ★ 新增：收藏數顯示
          let btns = ''; 
          const sCount = p.saveCount || 0;
          const starIcon = isSaved ? 'bi-star-fill text-white' : 'bi-star';
          const btnClass = isSaved ? 'btn-warning' : 'btn-outline-warning';
          const labelText = isSaved ? '已收藏' : '收藏';
          
          if (p.link) {
              btns += `<a href="${p.link}" target="_blank" class="btn btn-sm btn-outline-success">開啟檔案連結</a>`;
          }
          
          if (user) {
              btns += `
                  <button class="btn btn-sm ${btnClass} ms-2 d-inline-flex align-items-center" onclick="toggleSave('${p.id}', this)">
                      <i class="bi ${starIcon} me-1"></i>
                      <span class="save-label me-1">${labelText}</span>
                      <span class="save-count badge bg-light text-dark ms-1" style="font-size: 0.7rem;">${sCount > 0 ? sCount : ''}</span>
                  </button>
              `;
          }
          
          btns += `<button class="btn btn-sm btn-outline-primary ms-2" onclick="copyShareLink('plan_detail', '${p.id}')">分享</button>`;

          // 右上角動作選單
          const menu = isMine ? getActionMenu('post', p.id, p.status) : ''; 

          return `
            <div class="col-md-6 col-lg-4">
              <div class="card h-100 shadow-sm">
                <div class="card-body">
                  <div class="d-flex justify-content-between mb-2">
                    <div>${typeBadge}${accessBadge}</div>
                    <div class="d-flex align-items-center">
                      <small class="text-muted me-1">${formatDate(p.timestamp)}</small>
                      ${menu}
                    </div>
                  </div>
                  <div class="cursor-pointer" onclick="jumpToPlan('${p.id}')">
                    <h5 class="card-title text-truncate fw-bold">${esc(p.title)}</h5>
                    <h6 class="card-subtitle mb-2 text-muted small">作者: ${esc(p.authorName)}</h6>
                    ${gradeTags}
                    <div class="card-text mt-2 content-preview text-secondary" style="font-size: 0.9rem;">${p.content}</div>
                  </div>
                  <div class="mt-3 d-flex align-items-center">${btns}</div>
                </div>
              </div>
            </div>`;
        }).join('');
      } catch(e) { 
          console.error("Render Plans Error:", e);
          onFailure(e); 
      }
  }
  function changePlanPage(page) { curPage.plans = page; renderPlans(); }

  function jumpToPlan(id) {
    const sid = String(id);
    let p = null;
    let isFromMyArchive = false;

    // 1. 優先從「我的典藏」搜尋，如果在這裡找到，代表絕對是作者本人
    if (appData.myData && appData.myData.archived) {
        p = appData.myData.archived.find(x => String(x.id) === sid);
        if (p) isFromMyArchive = true;
    }

    // 2. 如果不是典藏，再找其他地方
    if (!p) {
        p = appData.plans.find(x => String(x.id) === sid);
    }
    
    if (!p && appData.myData) {
        p = (appData.myData.plans || []).find(x => String(x.id) === sid) ||
            (appData.myData.saved || []).find(x => String(x.id) === sid);
    }

    if (p) {
        // --- 核心邏輯修正 ---
        // 如果是從「我的典藏」找出來的，isMine 直接設為 true
        // 否則進行帳號比對
        const isMine = isFromMyArchive || (user && String(user.username) === String(p.author));

        // 只有「不是我的」且「狀態是典藏」才擋住
        if (String(p.status) === 'archived' && !isMine) { 
            showMsg('無法觀看', '此貼文已被作者典藏，暫時無法查看內容。', 'warning'); 
            return; 
        }
        
        if (p.access === 'member' && !user) { 
            pendingTarget = { view: 'plan_detail', id: sid }; 
            forceAuth(); 
            return; 
        }

        // 渲染 Modal
        const menu = isMine ? getActionMenu('post', p.id, p.status, true) : '';
        const catClass = p.category === '簡報' ? 'badge-ppt' : 'badge-word';
        document.getElementById('view-plan-title').textContent = p.title;
        document.getElementById('view-plan-actions').innerHTML = menu;
        document.getElementById('view-plan-badge').innerHTML = `<span class="badge ${catClass}">${p.category}</span>`;
        let gradeTxt = p.grades && p.grades.length > 0 ? ' | 適用：' + p.grades.map(g=>g+'年級').join('、') : '';
        document.getElementById('view-plan-grades').textContent = gradeTxt;
        document.getElementById('view-plan-meta').textContent = `作者: ${p.authorName || (isMine ? user.nickname : '未知')} | 發布於: ${formatDate(p.timestamp)}`;
        document.getElementById('view-plan-content').innerHTML = p.content;
        document.getElementById('view-plan-link').innerHTML = p.link ? `<a href="${p.link}" target="_blank" class="btn btn-success w-100">前往下載</a>` : '';
        document.getElementById('view-plan-share-btn').onclick = () => copyShareLink('plan_detail', sid);
        new bootstrap.Modal(document.getElementById('viewPlanModal')).show();
    } else {
        // 如果真的找不到，才去後端抓
        toggleLoading(true);
        google.script.run.withSuccessHandler(res => {
            toggleLoading(false);
            if(res.error) return showMsg('錯誤', res.error, 'error');
            const pRes = res.plan;
            const isMineRes = user && (String(user.username) === String(pRes.author));
            if (String(pRes.status) === 'archived' && !isMineRes) {
                showMsg('無法觀看', '此貼文已被作者典藏。', 'warning');
                return;
            }
            if (p.access === 'member' && !user) { pendingTarget = { view: 'plan_detail', id: id }; forceAuth(); return; }
            const menu = isMine ? getActionMenu('post', p.id, p.status, true) : '';
            const catClass = p.category === '簡報' ? 'badge-ppt' : 'badge-word';
            document.getElementById('view-plan-title').textContent = p.title;
            document.getElementById('view-plan-actions').innerHTML = menu;
            document.getElementById('view-plan-badge').innerHTML = `<span class="badge ${catClass}">${p.category}</span>`;
            let gradeTxt = p.grades && p.grades.length > 0 ? ' | 適用：' + p.grades.map(g=>g+'年級').join('、') : '';
            document.getElementById('view-plan-grades').textContent = gradeTxt;
            document.getElementById('view-plan-meta').textContent = `作者: ${p.authorName} | 發布於: ${formatDate(p.timestamp)}`;
            document.getElementById('view-plan-content').innerHTML = p.content;
            document.getElementById('view-plan-link').innerHTML = p.link ? `<a href="${p.link}" target="_blank" class="btn btn-success w-100">前往下載</a>` : '';
            document.getElementById('view-plan-share-btn').onclick = () => copyShareLink('plan_detail', id);
            new bootstrap.Modal(document.getElementById('viewPlanModal')).show();
        }).withFailureHandler(onFailure).getData('plan_detail', sid);
    }
  }

  function jumpToThread(id) { if(!user) { pendingTarget = { view: 'thread_detail', id: id }; forceAuth(); return; } switchView('forum'); loadThreadDetail(id); }
  function loadForum() { const list = document.getElementById('forum-threads'); if(list.innerHTML.trim() === '') list.innerHTML = '載入中...'; google.script.run.withSuccessHandler(res => { if(res.error) return onFailure(new Error(res.error)); appData.threads = res.threads; renderForumList(); }).withFailureHandler(onFailure).getData('forum', user ? user.username : null); }
  
  function handleForumSearch(val) { filterState.forumSearch = val.toLowerCase(); curPage.forum = 1; renderForumList(); }

  function renderForumList() {
      const list = document.getElementById('forum-threads'); 
      let displayThreads = appData.threads;
      if (filterState.forumSearch) { 
          const term = filterState.forumSearch; 
          displayThreads = displayThreads.filter(t => t.title.toLowerCase().includes(term) || t.content.toLowerCase().includes(term) || t.authorName. toLowerCase().includes(term)); 
      }
      const start = (curPage.forum - 1) * FORUM_PER_PAGE; 
      const items = displayThreads.slice(start, start + FORUM_PER_PAGE); 
      renderPagination(displayThreads.length, curPage.forum, 'forum-pagination', 'changeForumPage', FORUM_PER_PAGE);

      if(items.length === 0) {
          list.innerHTML = '<p class="text-center text-muted">暫無討論</p>'; 
      } else {
          list.innerHTML = items.map((t, index) => {
              const serial = start + index + 1;

              // 1. 定義標籤內容
              const archiveBadge = t.status === 'archived' ? '<span class="badge bg-warning text-dark ms-1">已典藏</span>' : '';
              const closedBadge = t.status === 'closed' ? '<span class="badge bg-danger ms-1"><i class="bi bi-lock-fill"></i>  此討論區目前僅供瀏覽</span>' : '';

              // 2. 點讚狀態
              const isLiked = user && t.likes && t.likes.includes(user.username);
              const heartClass = isLiked ? 'bi-heart-fill text-danger' : 'bi-heart';

              return `
                <div class="list-group-item list-group-item-action p-3 cursor-pointer" onclick="loadThreadDetail('${t.id}')">
                  <div class="d-flex w-100 justify-content-between">
                    <h5 class="mb-1 text-success fw-bold">
                      <span class="text-muted me-2">#${serial}</span>${esc(t.title)}
                      ${archiveBadge} ${closedBadge}  </h5>
                    <small class="text-muted">${formatDate(t.timestamp)}</small>
                  </div>
                  <p class="mb-1 text-secondary text-truncate" style="max-width: 90%;">${esc(t.content.replace(/<[^>]*>?/gm, ''))}</p>

                  <div class="d-flex justify-content-between align-items-center mt-2">
                    <small class="text-muted">作者: ${esc(t.authorName)}</small>
                    <div class="d-flex gap-3">
                      <span class="small text-secondary">
                        <i class="bi bi-chat-dots"></i> ${t.commentCount || 0}
                      </span>
                      <span class="small text-secondary">
                        <i class="bi ${heartClass}"></i> ${t.likeCount || 0}
                      </span>
                    </div>
                  </div>
                </div>`;
          }).join('');
      }
  }
  function changeForumPage(page) { curPage.forum = page; renderForumList(); }

  // ★ 補上缺失的函式：渲染留言外觀 (支援樓中樓) ★
  function renderCommentItem(c, isSub) {
      const isMine = user && user.username === c.author;
      const isArchived = c.status === 'archived';
  
      // --- 關鍵修正：內容遮蔽邏輯 ---
      let displayContent = c.content;
      if (isArchived && !isMine) {
          displayContent = '<i class="text-muted">（此留言已由作者典藏，暫時無法查看內容）</i>';
      }
  
      const cMenu = isMine ? getActionMenu('comment', c.id, c.status) : ''; 
      const cArchiveBadge = isArchived ? '<span class="badge bg-warning text-dark me-2">已典藏</span>' : '';
      
      const floorDisplay = isSub ? `B${c.floor}-${c.subFloor}` : `B${c.floor}`;
      const likes = c.likes || [];
      const likeCount = likes.length > 0 ? likes.length : '';
      const isLiked = user && likes.includes(user.username);
      
      const containerClass = isSub ? 'sub-comment' : 'card mb-3 border-0 bg-white shadow-sm';
      const bodyClass = isSub ? '' : 'card-body p-3';
  
      return `
        <div class="${containerClass}" style="${isArchived && !isMine ? 'opacity: 0.7;' : ''}">
          <div class="${bodyClass}">
              <div class="d-flex justify-content-between mb-1">
                  <div>
                      <span class="floor-tag">${floorDisplay}</span>
                      <strong class="me-2">${esc(c.authorName)}</strong> 
                      ${cArchiveBadge}
                      <small class="text-muted ms-2">${formatDate(c.timestamp)}</small>
                  </div>
                  <div>${cMenu}</div>
              </div>
              <div class="mb-2" style="overflow-wrap: break-word;">${displayContent}</div>
              
              ${(isArchived && !isMine) ? '' : `
              <div class="d-flex align-items-center">
                  <button class="btn-like ${isLiked ? 'liked' : ''} me-3" onclick="toggleLike('${c.id}', this)">
                      <i class="bi ${isLiked ? 'bi-heart-fill' : 'bi-heart'}"></i><span class="like-count">${likeCount}</span>
                  </button>
                  <span class="reply-ref" onclick="replyTo(${c.floor}, ${c.subFloor})"><i class="bi bi-reply-fill"></i> 回覆</span>
              </div>
              `}
          </div>
        </div>`;
  }

  // ★ 修改 1：載入留言時，把資料存進 lookup 表 ★
  function loadThreadDetail(tid) { 
      // 進入時重置狀態
      currentThreadId = tid; 
      replyToFloor = 0;

      toggleLoading(true); 
      document.getElementById('reply-indicator').classList.add('hidden');
      document.getElementById('forum-list-view').classList.add('hidden'); 
      document.getElementById('forum-detail-view').classList.remove('hidden'); 

      const area = document.getElementById('detail-comments'); 
      area.innerHTML = '';

      google.script.run.withSuccessHandler(res => { 
          toggleLoading(false); 
          if(res.error) { showMsg('錯誤', res.error, 'error'); return; } 

          try { 
              const t = res.thread; 
              appData.currentThread = t; 
              const isMine = user && user.username === t.author; 

              // 1. 判斷愛心狀態
              const tLikes = t.likes || [];
              const isThreadLiked = user && tLikes.includes(user.username);
              const tHeartClass = isThreadLiked ? 'btn-danger' : 'btn-outline-danger';
              const tHeartIcon = isThreadLiked ? 'bi-heart-fill' : 'bi-heart';

              // 2. 渲染標題區 (記得補上 event.stopPropagation() 防止點擊衝突)
              document.getElementById('detail-title').innerHTML = `
                  ${esc(t.title)} 
                  <button class="btn btn-sm ${tHeartClass} ms-2" onclick="event.stopPropagation(); toggleThreadLike('${t.id}', this)">
                      <i class="bi ${tHeartIcon}"></i> 
                      <span class="t-like-count">${t.likeCount || 0}</span>
                  </button>
              `;

              document.getElementById('detail-author').textContent = t.authorName; 
              document.getElementById('detail-date').textContent = formatDate(t.timestamp); 
              document.getElementById('detail-content').innerHTML = t.content;
              document.getElementById('detail-actions').innerHTML = isMine ? getActionMenu('thread', t.id, t.status) : '';

              // 狀態處理（關閉討論區）
              const statusBadgeArea = document.getElementById('detail-status-badge');
              if (statusBadgeArea) {
                  statusBadgeArea.innerHTML = t.status === 'closed' ? '<span class="badge bg-secondary ms-2"><i class="bi bi-lock-fill"></i> 此討論區目前僅供瀏覽</ span>' : '';
              }

              const commentSection = document.getElementById('comment-section');
              const closedMsg = document.getElementById('thread-closed-msg');

              if (t.status === 'closed') {
                  if (commentSection) commentSection.classList.add('hidden');
                  if (closedMsg) {
                      closedMsg.innerHTML = '<i class="bi bi-info-circle-fill"></i> 此討論區已被作者關閉，目前僅供瀏覽。';
                      closedMsg.classList.remove('hidden');
                  }
              } else {
                  if (commentSection) commentSection.classList.remove('hidden');
                  if (closedMsg) closedMsg.classList.add('hidden');
              }

              // --- 處理留言查詢表與渲染 ---
              commentLookup = {};
              res.comments.forEach(c => {
                  commentLookup[c.floor + '-' + c.subFloor] = c;
              });

              let commentGroups = {}; 
              res.comments.forEach(c => {
                  const floor = c.floor;
                  if (!commentGroups[floor]) commentGroups[floor] = { main: null, subs: [] };
                  if (c.subFloor === 0) commentGroups[floor].main = c;
                  else commentGroups[floor].subs.push(c);
              });

              let html = '';
              Object.keys(commentGroups).sort((a,b)=>a-b).forEach(floorKey => {
                  const group = commentGroups[floorKey];
                  if(group.main) {
                      html += renderCommentItem(group.main, false); 
                      group.subs.sort((a,b)=>a.subFloor-b.subFloor).forEach(sub => {
                          html += renderCommentItem(sub, true);
                      });
                  }
              });
              area.innerHTML = html; 

              // 啟動預覽監聽
              bindRefPreviews();

          } catch(e) { 
              console.error("Detail Logic Error:", e);
              onFailure(e); 
          } 
      }).withFailureHandler(onFailure).getData('thread_detail', tid); 
  }
  
  function openEdit(type, id, closeViewModal=false) { 
    if(closeViewModal) bootstrap.Modal.getInstance(document.getElementById('viewPlanModal')).hide(); 
    let data=null; 
    if(type==='post') { 
        if(appData.myData && (appData.myData.plans || appData.myData.archived)) 
            data = (appData.myData.plans || []).find(p=>p.id==id) || (appData.myData.archived || []).find(p=>p.id==id); 
        if(!data) data = appData.plans.find(p=>p.id==id); 
    } else if(type==='thread') data=(appData.currentThread&&appData.currentThread.id==id)?appData.currentThread:appData.threads.find(t=>t.id==id); 
    if(!data && type!=='comment') return showMsg('錯誤','找不到資料','error'); 
    
    if(type==='post'){ document.getElementById('plan-modal-title').textContent='編輯檔案'; document.getElementById('plan-id').value=id; document.getElementById('plan-title').value=data.title; 
    tinymce.get('plan-content').setContent(data.content); 
    document.getElementById('plan-link').value=data.link; document.getElementById('plan-category').value=data.category||'教案'; document.getElementById('plan-access').value=data.access||'member'; document.querySelectorAll('#plan-grades-check input').forEach(c => c.checked = false); if(data.grades) data.grades.forEach(g => { let ck=document.querySelector(`#plan-grades-check input[value="${g}"]`); if(ck) ck.checked=true; }); new bootstrap.Modal(document.getElementById('planModal')).show(); } 
    else if(type==='thread'){ document.getElementById('thread-modal-title').textContent='編輯討論'; document.getElementById('thread-id').value=id; document.getElementById('thread-title').value=data.title; 
    tinymce.get('thread-content').setContent(data.content); 
    new bootstrap.Modal(document.getElementById('threadModal')).show(); } 
    else if(type==='comment'){ const newContent=prompt("編輯留言內容:"); if(newContent) google.script.run.withSuccessHandler(()=>loadThreadDetail(currentThreadId)).editCommentContent(id,newContent,user.username); } 
  }

  function openPlanModal() { document.getElementById('plan-modal-title').textContent = '上傳檔案'; document.getElementById('plan-id').value = ''; document.getElementById('plan-title').value = ''; 
  tinymce.get('plan-content').setContent(''); 
  document.getElementById('plan-link').value = ''; document.querySelectorAll('#plan-grades-check input').forEach(c=>c.checked=false); new bootstrap.Modal(document.getElementById('planModal')).show(); }
  function openThreadModal() { document.getElementById('thread-modal-title').textContent = '發起討論'; document.getElementById('thread-id').value = ''; document.getElementById('thread-title').value = ''; 
  tinymce.get('thread-content').setContent(''); 
  new bootstrap.Modal(document.getElementById('threadModal')).show(); }
  
  function updateStatus(type, id, newStatus) { const msg = newStatus === 'deleted' ? '確定要刪除嗎？刪除後無法復原。' : (newStatus === 'closed' ? '確定要關閉討論嗎？關閉後將無法回應。' : '確定要變更狀態嗎？'); showConfirm('操作確認', msg, () => { const handlers = {'post': 'updatePostStatus', 'thread': 'updateThreadStatus', 'comment': 'updateCommentStatus'}; google.script.run.withSuccessHandler(res => { if(res.status === 'success') { showMsg('成功', '狀態已更新', 'success'); if(type === 'post') { loadPlans(); if(!document.getElementById('view-profile').classList.contains('hidden')) loadProfile(); } else if(type === 'thread') { if(newStatus==='deleted') backToForumList(); else if(currentThreadId) loadThreadDetail(currentThreadId); else loadForum(); } else if(type === 'comment') loadThreadDetail(currentThreadId); } else onFailure(new Error(res.message)); })[handlers[type]](id, newStatus, user.username); }); }
  function loadProfile() {
      if (!user) return forceAuth();

      document.getElementById('profile-nickname').textContent = user.nickname;
      document.getElementById('profile-username').textContent = '@' + user.username;

      // 顯示載入中狀態
      document.getElementById('profile-list-container').innerHTML = '<div class="text-center py-5"><div class="spinner-border text-success"></div></  div>';

      google.script.run.withSuccessHandler(res => {
          if(res.error) return onFailure(new Error(res.error));

          // 1. 重要：將抓回來的資料完整存入全域變數，供後續 renderProfileList 使用
          appData.myData = res; 

          // 2. 渲染碳值與等級進度條
          const points = res.carbonPoints || 0;
          const info = getLevelInfo(points);
          const nextTarget = info.next === 'MAX' ? points : info.next;
          const progress = info.next === 'MAX' ? 100 : Math.min(100, (points / info.next) * 100);

          const carbonHtml = `
              <div id="carbon-status-box-inner" 
                   onclick="showScoreDetails()" 
                   style="cursor: pointer;" 
                   title="點擊查看分數明細"
                   class="mt-3 p-3 rounded bg-white shadow-sm border-start border-4 border-success">

                  <div class="d-flex justify-content-between align-items-center mb-1">
                      <span class="fw-bold" style="color: ${info.color}">LV.${info.lv} ${info.name}</span>
                      <span class="badge" style="background-color: ${info.color}; color: white;">${points} CP</span>
                  </div>
                  <div class="progress" style="height: 10px; background-color: #e9ecef; border-radius: 5px; overflow: hidden;">
                      <div class="progress-bar progress-bar-striped progress-bar-animated" 
                           style="width: ${progress}%; background-color: ${info.color}; transition: width 1s ease-in-out;"></div>
                  </div>
                  <div class="d-flex justify-content-between mt-1">
                      <small class="text-muted" style="font-size: 0.7rem;">
                          ${info.next === 'MAX' ? '巔峰等級' : `下一級: ${info.next} CP`}
                      </small>
                      <small class="text-muted" style="font-size: 0.7rem;">
                          <i class="bi bi-search"></i> 點擊查看明細
                      </small>
                  </div>
              </div>
          `;

          const cardBody = document.querySelector('#view-profile .card-body');
          const oldBox = document.getElementById('carbon-status-box');
          if(oldBox) oldBox.remove();
          const div = document.createElement('div');
          div.id = 'carbon-status-box';
          div.innerHTML = carbonHtml;
          cardBody.insertBefore(div, cardBody.querySelector('.text-start'));

          // 3. 更新側邊欄統計數字
          document.getElementById('count-share').textContent = res.stats.share;
          document.getElementById('count-archive').textContent = res.stats.archive;
          document.getElementById('count-saved').textContent = res.stats.saved;
          document.getElementById('count-threads').textContent = res.stats.threads;
          document.getElementById('count-replies').textContent = res.stats.replies;

          // 4. ★ 關鍵：抓完資料後，立即呼叫渲染列表 ★
          renderProfileList();

      }).getData('profile', user.username);
  }
  function changeProfileTab(tabName) { profileTab = tabName; const showFilter = (tabName === 'plans' || tabName === 'archived' || tabName === 'saved'); document.getElementById('profile-filter-container').classList.toggle('hidden', !showFilter); profileFilterCat = 'all'; const btns = document.querySelectorAll('#profile-filter-container button'); btns.forEach(b => { b.className = 'btn btn-outline-secondary'; if(b.textContent === '全部') b.className = 'btn btn-secondary active'; else if(b.textContent === '教案') b.classList.add('btn-outline-word'); else if(b.textContent === '簡報') b.classList.add('btn-outline-ppt'); }); renderProfileList(); }
  function setProfileFilter(cat, btn) {
      const group = btn.parentElement;
      // 處理按鈕樣式切換
      Array.from(group.children).forEach(b => {
          const txt = b.textContent.trim();
          b.className = 'btn';
          if (txt === '全部') b.classList.add('btn-outline-secondary');
          else if (txt === '教案') b.classList.add('btn-outline-word');
          else if (txt === '簡報') b.classList.add('btn-outline-ppt');
      });

      // 設定高亮樣式
      if (cat === 'all') btn.className = 'btn btn-secondary active';
      else if (cat === '教案') btn.className = 'btn btn-word active';
      else if (cat === '簡報') btn.className = 'btn btn-ppt active';

      // ★ 關鍵修正：同步更新篩選狀態物件
      profileFilterState.category = cat; 

      // 重置分頁並重新渲染
      curPage['profile_' + profileTab] = 1;
      renderProfileList();
  }
  function renderProfileList() {
      const container = document.getElementById('profile-list-container');
      const pgId = 'profile-pagination-container';

      if (!appData.myData) return;

      let items = [];
      if (profileTab === 'plans') items = appData.myData.plans || [];
      else if (profileTab === 'archived') items = appData.myData.archived || [];
      else if (profileTab === 'saved') items = appData.myData.saved || [];
      else if (profileTab === 'threads') items = appData.myData.threads || [];
      else if (profileTab === 'replies') items = appData.myData.replies || [];

      // --- 修正後的過濾區塊 ---
      if (['plans', 'archived', 'saved'].includes(profileTab)) {
          // 過濾類別 (簡報/教案)
          if (profileFilterState.category && profileFilterState.category !== 'all') {
              items = items.filter(i => i.category === profileFilterState.category);
          }
          // 過濾年級
          if (profileFilterState.grades && profileFilterState.grades.length > 0) {
              items = items.filter(i => i.grades && i.grades.some(g => profileFilterState.grades.includes(g)));
          }
      }

      const cPage = curPage['profile_' + profileTab] || 1;
      const start = (cPage - 1) * ITEMS_PER_PAGE;
      const paginatedItems = items.slice(start, start + ITEMS_PER_PAGE);

      renderPagination(items.length, cPage, pgId, 'changeProfilePage', ITEMS_PER_PAGE);

      if (paginatedItems.length === 0) {
          container.innerHTML = '<div class="text-center py-5 text-muted"><i class="bi bi-inbox fs-1 d-block mb-2"></i>目前尚無資料</div>';
          return;
      }

      container.innerHTML = '<div class="list-group list-group-flush border rounded">' + paginatedItems.map(i => {
          if (profileTab === 'replies') {
              return `
                  <div class="list-group-item list-group-item-action cursor-pointer" onclick="jumpToThread('${i.threadId}')">
                      <div class="d-flex w-100 justify-content-between">
                          <h6 class="mb-1 text-primary text-truncate">回覆：${esc(i.threadTitle)}</h6>
                      </div>
                      <p class="mb-0 small text-secondary text-truncate">${esc(i.content.replace(/<[^>]*>?/gm, ''))}</p>
                  </div>`;
          }

          const clickFunc = (profileTab === 'threads') ? 'jumpToThread' : 'jumpToPlan';
          const badgeClass = i.category === '簡報' ? 'badge-ppt' : 'badge-word';
          const badge = i.category ? `<span class="badge ${badgeClass} me-2">${i.category}</span>` : '';

          return `
              <div class="list-group-item list-group-item-action d-flex justify-content-between align-items-center">
                  <div class="flex-grow-1 cursor-pointer fw-bold" onclick="${clickFunc}('${i.id}')">
                      ${badge}${esc(i.title)}
                  </div>
              </div>`;
      }).join('') + '</div>';
  }
  function changeProfilePage(p) { curPage['profile_' + profileTab] = p; renderProfileList(); }
  function submitProfileEdit() {
    // 1. 取得基本欄位值
    const nickVal = document.getElementById('edit-nick').value.trim();
    const passVal = document.getElementById('edit-pass').value.trim();
    const notifyVal = document.getElementById('edit-notify').checked;
  
    // 2. 判斷 Email 是否處於編輯模式 (如果 email-edit-group 沒被隱藏，就是正在改)
    const isEmailEditing = !document.getElementById('email-edit-group').classList.contains('hidden');
    const emailInput = document.getElementById('edit-email');
    const vCodeInput = document.getElementById('edit-vcode');
    const emailVal = emailInput ? emailInput.value.trim() : "";
    const vCodeVal = vCodeInput ? vCodeInput.value.trim() : "";
  
    // 3. 準備傳送給後端的資料物件
    const data = {
      nickname: nickVal,
      password: passVal, // 留空則後端不更新
      notify: notifyVal
    };
  
    // 4. Email 修改邏輯驗證
    if (isEmailEditing) {
      // 如果輸入的信箱跟原本不一樣，且不是要刪除信箱(填空)
      if (emailVal !== user.email) {
        if (emailVal !== "" && !vCodeVal) {
          return showMsg('提示', '請先完成電子郵件驗證碼驗證', 'warning');
        }
        data.email = emailVal;
        data.vCode = vCodeVal;
      }
    }
  
    // 5. 執行後端更新
    toggleLoading(true);
    google.script.run.withSuccessHandler(res => {
      toggleLoading(false);
      if (res.status === 'success') {
        // 統一風格的成功提示
        showMsg('更新成功', '您的資料與設定已成功儲存。', 'success', () => {
          // 更新成功後必須 reload，確保 localStorage 與全域變數 user 取得最新狀態 (包含 email 與 notify)
          location.reload();
        });
      } else {
        // 統一風格的錯誤提示 (例如驗證碼錯誤)
        showMsg('更新失敗', res.message || '請檢查您的輸入內容或驗證碼', 'error');
      }
    }).withFailureHandler(err => {
      toggleLoading(false);
      showMsg('系統錯誤', '連線失敗：' + err.toString(), 'error');
    }).updateProfileSettings(user.username, data);
  }
  // 修改原本的 openEditProfile
  // 在 index.html 的 <script> 區塊中替換 openEditProfile 函式
  // --- 修改 1：解決編輯資料按鈕與 UI 顯示問題 ---
  function openEditProfile() {
    if (!user) return forceAuth();
  
    // 1. 基本資料填入
    document.getElementById('edit-nick').value = user.nickname || '';
    document.getElementById('edit-pass').value = ''; 
    document.getElementById('edit-notify').checked = (user.notify === true);
  
    const verifiedGroup = document.getElementById('email-verified-group');
    const editGroup = document.getElementById('email-edit-group');
    const displayInput = document.getElementById('display-email');
    const vCodeInput = document.getElementById('edit-vcode');
    const cancelBtn = document.getElementById('btn-cancel-email-edit');
  
    vCodeInput.classList.add('hidden'); 
  
    // 2. 判斷顯示哪一組
    if (user.email && user.email.trim() !== "") {
      verifiedGroup.classList.remove('hidden');
      editGroup.classList.add('hidden');
      displayInput.value = user.email;
      cancelBtn.classList.remove('hidden'); // 如果本來就有信箱，允許取消
    } else {
      verifiedGroup.classList.add('hidden');
      editGroup.classList.remove('hidden');
      document.getElementById('edit-email').value = '';
      cancelBtn.classList.add('hidden'); // 沒信箱時不需要取消按鈕
    }
  
    // 3. 顯示彈窗
    const modalEl = document.getElementById('profileEditModal');
    bootstrap.Modal.getOrCreateInstance(modalEl).show();
  }
  
  // 點擊「更改」
  function enableEmailEdit() {
    document.getElementById('email-verified-group').classList.add('hidden');
    document.getElementById('email-edit-group').classList.remove('hidden');
    document.getElementById('btn-cancel-email-edit').classList.remove('hidden');
    document.getElementById('edit-email').focus();
  }
  
  // 點擊「取消」
  function cancelEmailEdit() {
    if (user.email) {
      document.getElementById('email-verified-group').classList.remove('hidden');
      document.getElementById('email-edit-group').classList.add('hidden');

      // 順便清空輸入框與驗證碼，避免殘留
      document.getElementById('edit-email').value = "";
      const vCodeInput = document.getElementById('edit-vcode');
      vCodeInput.value = "";
      vCodeInput.classList.add('hidden');
    }
  }
  function deleteEmail() {
    showConfirm('刪除信箱', '確定要解除電子郵件綁定嗎？刪除後將無法接收系統通知。', () => {
      const data = { 
        nickname: document.getElementById('edit-nick').value, 
        email: "",    // 傳送空字串代表刪除
        notify: false // 刪除信箱後自動關閉通知
      };

      toggleLoading(true);
      google.script.run.withSuccessHandler(res => {
        toggleLoading(false);
        if(res.status === 'success') {
          // --- 核心修正處：同步更新本地數據 ---
          user.email = "";
          user.notify = false;
          localStorage.setItem('tb_user', JSON.stringify(user));

          // 1. 更新個人資料頁面的 Email 大字 (如果有顯示的話)
          if(document.getElementById('profile-email-label')) {
             document.getElementById('profile-email-label').textContent = "";
          }

          // 2. 更新通知開關
          document.getElementById('edit-notify').checked = false;

          // 3. 強制切換 UI 狀態：隱藏「已驗證組」，顯示「編輯組」
          document.getElementById('email-verified-group').classList.add('hidden');
          document.getElementById('email-edit-group').classList.remove('hidden');
          document.getElementById('edit-email').value = '';
          document.getElementById('edit-vcode').classList.add('hidden');

          // 4. 顯示統一風格的系統訊息
          showMsg('成功', '電子郵件已解除綁定', 'success');
        } else {
          showMsg('錯誤', res.message, 'error');
        }
      }).updateProfileSettings(user.username, data);
    });
  }
  function toggleSave(pid, btn) {
      if (!user) return forceAuth();

      const countSpan = btn.querySelector('.save-count');
      const labelSpan = btn.querySelector('.save-label');
      const icon = btn.querySelector('i');

      // --- 樂觀更新：立即改變外觀，不等後端 ---
      const isCurrentlySaved = btn.classList.contains('btn-warning');
      let currentCount = parseInt(countSpan.textContent) || 0;

      if (isCurrentlySaved) {
          // 模擬取消收藏
          btn.className = 'btn btn-sm btn-outline-warning ms-2 d-inline-flex align-items-center';
          icon.className = 'bi bi-star me-1';
          labelSpan.textContent = '收藏';
          countSpan.textContent = currentCount > 1 ? currentCount - 1 : '';
      } else {
          // 模擬新增收藏
          btn.className = 'btn btn-sm btn-warning ms-2 d-inline-flex align-items-center';
          icon.className = 'bi bi-star-fill text-white me-1';
          labelSpan.textContent = '已收藏';
          countSpan.textContent = currentCount + 1;
      }

      // 暫時鎖定按鈕防止連點
      btn.style.pointerEvents = 'none';

      google.script.run.withSuccessHandler(res => {
          btn.style.pointerEvents = 'auto'; // 解鎖
          if(res.status === 'success') {
              // 同步正確的伺服器數字（修正多人同時點擊的誤差）
              countSpan.textContent = res.saveCount > 0 ? res.saveCount : '';
              user.saved = res.saved;
              localStorage.setItem('tb_user', JSON.stringify(user));

              const p = appData.plans.find(x => String(x.id) === String(pid));
              if (p) p.saveCount = res.saveCount;
          } else {
              // 如果後端失敗，才刷回列表原本的樣子
              loadPlans();
              showMsg('錯誤', res.message, 'error');
          }
      }).toggleSavePost(user.username, pid);
  }
  function backToForumList() { document.getElementById('forum-list-view').classList.remove('hidden'); document.getElementById('forum-detail-view').classList.add('hidden'); currentThreadId = null; cleanUrl(); }
  function updateAuthUI() { if (user) { document.getElementById('btn-login-nav').classList.add('hidden'); document.getElementById('btn-logout').classList.remove('hidden'); document.getElementById('nav-user-display').textContent = user.nickname; document.getElementById('hero-btns').classList.add('hidden'); } else { document.getElementById('btn-login-nav').classList.remove('hidden'); document.getElementById('btn-logout').classList.add('hidden'); document.getElementById('nav-user-display').textContent = ''; document.getElementById('hero-btns').classList.remove('hidden'); } }
  // --- 修改 3：註冊與登入也移除全頁重整 ---
  function handleLogin() {
    const form = document.querySelector('#form-login');
    toggleLoading(true);
    google.script.run.withSuccessHandler(res => {
      toggleLoading(false);
      if(res.status === 'success') {
        user = res.user;
        localStorage.setItem('tb_user', JSON.stringify(user));
        updateAuthUI();
        bootstrap.Modal.getInstance(document.getElementById('authModal')).hide();
        showMsg('歡迎', `登入成功，${user.nickname}！`, 'success', () => {
          switchView('home');
        });
      } else {
        showMsg('失敗', res.message, 'error');
      }
    }).userLogin({ username: form.username.value, password: form.password.value });
  }

  function handleRegister() {
    const data = {
      username: document.getElementById('reg-user').value,
      password: document.getElementById('reg-pass').value,
      nickname: document.getElementById('reg-nick').value,
      email: document.getElementById('reg-email').value,
      vCode: document.getElementById('reg-vcode').value
    };
    if(data.email && !data.vCode) return showMsg('提示', '請完成信箱驗證', 'warning');

    toggleLoading(true);
    google.script.run.withSuccessHandler(res => {
      toggleLoading(false);
      if(res.status === 'success') {
        bootstrap.Modal.getInstance(document.getElementById('authModal')).hide();
        showMsg('註冊成功', '請使用新帳號登入', 'success', () => {
          showAuthModal('login');
        });
      } else {
        if (res.message && res.message.includes("驗證碼")) {
          triggerInputError('reg-vcode');
        }
        showMsg('註冊失敗', res.message, 'error');
      }
    }).userRegister(data);
  }

  // 發送驗證碼
  function sendVCode(inputId) {
    const email = document.getElementById(inputId).value;
    if(!email || !email.includes('@')) return showMsg('格式錯誤', '請輸入正確的電子郵件', 'warning');

    toggleLoading(true);
    google.script.run.withSuccessHandler(res => {
      toggleLoading(false);
      if(res.status === 'success') {
        showMsg('發送成功', '驗證碼已寄至您的信箱，請於10分鐘內輸入。', 'success');

        // 判斷是註冊還是修改頁面
        const vId = (inputId === 'reg-email') ? 'reg-vcode' : 'edit-vcode';
        const vInput = document.getElementById(vId);

        if (vInput) {
          // --- 核心修正：清空前一次留下來的驗證碼 ---
          vInput.value = ""; 

          // 顯示輸入框
          vInput.classList.remove('hidden');
          vInput.focus(); // 自動聚焦方便使用者輸入
        }
      } else {
        showMsg('發送失敗', res.message, 'error');
      }
    }).sendVerificationCode(email);
  }

  // 小工具：觸發輸入框抖動並清空
  function triggerInputError(elementId) {
    const el = document.getElementById(elementId);
    if (!el) return;

    el.value = ""; // 清空內容
    el.classList.add('shake-error'); // 加入動畫類別
    el.focus();

    // 0.5秒後移除類別，方便下次錯誤時能再次觸發動畫
    setTimeout(() => {
      el.classList.remove('shake-error');
    }, 500);
  }

  // 新增提交個人資料設定
  function submitProfileEdit() {
    const isEmailEditing = !document.getElementById('email-edit-group').classList.contains('hidden');
    const emailInput = document.getElementById('edit-email');
    const vCodeInput = document.getElementById('edit-vcode');
    const emailVal = emailInput.value.trim();
    const vCodeVal = vCodeInput.value.trim();
    const nickVal = document.getElementById('edit-nick').value.trim();
    const notifyVal = document.getElementById('edit-notify').checked;

    const data = { nickname: nickVal, password: document.getElementById('edit-pass').value, notify: notifyVal };

    // --- 關鍵修正：判斷是否真的有改信箱 ---
    if (isEmailEditing) {
      // 如果使用者按了更改，但沒填東西，或者填的跟原本一模一樣 -> 視為不改，保留原本的
      if (emailVal === "" || emailVal === user.email) {
        // 不將 email 放入 data 物件，後端就不會動到這一欄
      } else {
        // 使用者填了新信箱且不為空
        if (!vCodeVal) return showMsg('提示', '請先完成信箱驗證', 'warning');
        data.email = emailVal;
        data.vCode = vCodeVal;
      }
    }

    toggleLoading(true);
    google.script.run.withSuccessHandler(res => {
      toggleLoading(false);
      if (res.status === 'success') {
        user.nickname = nickVal;
        user.notify = notifyVal;
        // 只有在真的成功修改 email 時才更新本地資料
        if (data.email !== undefined) user.email = data.email;
        localStorage.setItem('tb_user', JSON.stringify(user));

        updateAuthUI(); 
        if(document.getElementById('profile-nickname')) document.getElementById('profile-nickname').textContent = nickVal;
        if(document.getElementById('profile-email-label')) document.getElementById('profile-email-label').textContent = user.email || '';

        bootstrap.Modal.getInstance(document.getElementById('profileEditModal')).hide();
        showMsg('成功', '設定已更新', 'success');
      } else {
        // 如果錯誤訊息包含「驗證碼」，就觸發抖動
        if (res.message && res.message.includes("驗證碼")) {
          triggerInputError('edit-vcode');
        }
        showMsg('更新失敗', res.message || '請檢查您的輸入內容', 'error');
      }
    }).updateProfileSettings(user.username, data);
  }
  function logout() { localStorage.removeItem('tb_user'); user = null; updateAuthUI(); switchView('home'); }
  function toggleAuthTab(tab) { document.getElementById('auth-error-msg').classList.add('hidden'); document.getElementById('tab-login').classList.toggle('active', tab==='login'); document.getElementById('tab-register').classList.toggle('active', tab==='register'); document.getElementById('form-login').classList.toggle('hidden', tab!=='login'); document.getElementById('form-register').classList.toggle('hidden', tab!=='register'); }
  function showAuthModal(tab) { toggleAuthTab(tab); new bootstrap.Modal(document.getElementById('authModal')).show(); }
  function esc(s) { return s ? String(s).replace(/</g,'&lt;').replace(/>/g,'&gt;') : ''; }
  function formatDate(ts) { if(!ts) return ''; const d=new Date(ts); return isNaN(d.getTime()) ? '' : `${d.getFullYear()}/${String(d.getMonth()+1).padStart(2,'0')}/${String(d.getDate()).padStart(2,'0')}`; }

  let replyToFloor = 0;
  // ★ 修正版：回覆功能 (使用傳統字串連接，確保變數正確顯示) ★
  function replyTo(floor, subFloor) {
    if(!user) return forceAuth();
    
    replyToFloor = parseInt(floor); 

    const editor = tinymce.get('comment-input');
    const indicator = document.getElementById('reply-indicator');
    
    var tagText = "B" + floor;
    if(subFloor && subFloor > 0) tagText += "-" + subFloor;
    
    // --- 修正處：移除 background-color ---
    var tagHtml = '<span class="ref-link" contenteditable="false" data-floor="' + floor + '" data-sub="' + subFloor + '" style="color:#0d6efd; font-weight:bold; cursor:pointer; text-decoration:none; margin-right:2px;">' + tagText + '</span>&nbsp;';
    
    indicator.textContent = "正在回覆 " + tagText + " ...";
    indicator.classList.remove('hidden');
    
    var currentContent = editor.getContent();
    if (currentContent.trim() === '') {
        editor.setContent('<p>' + tagHtml + '</p>');
    } else {
        editor.setContent(currentContent + '<p>' + tagHtml + '</p>');
    }
    
    editor.focus();
    editor.selection.select(editor.getBody(), true);
    editor.selection.collapse(false);
    
    document.getElementById('comment-section').scrollIntoView({ behavior: 'smooth' });
  }

  // ★ 修正版：留言按讚功能 ★
  function toggleLike(commentId, btn) {
    if(!user) return forceAuth();

    const isLiked = btn.classList.contains('liked');
    const icon = btn.querySelector('i');
    const countSpan = btn.querySelector('.like-count');
    let count = parseInt(countSpan.textContent) || 0;

    // --- 1. 立即更新 UI ---
    btn.classList.toggle('liked');
    icon.className = isLiked ? 'bi bi-heart' : 'bi bi-heart-fill';
    count = isLiked ? Math.max(0, count - 1) : count + 1;
    countSpan.textContent = count > 0 ? count : '';

    // --- 2. 發送至後端 ---
    google.script.run.withSuccessHandler(res => {
        if(res.status !== 'success') {
            // 失敗時彈回原狀
            btn.classList.toggle('liked');
            icon.className = !isLiked ? 'bi bi-heart' : 'bi bi-heart-fill';
            countSpan.textContent = parseInt(countSpan.textContent) + (isLiked ? 1 : -1);
            showMsg('錯誤', res.message, 'error');
        } else {
            // 這裡可以選擇是否呼叫 refreshUserCarbonPoints，但通常後端 toggle 時就會算好了
        }
    }).withFailureHandler(onFailure).toggleCommentLike(commentId, user.username);
  }

  // ★ 預覽小視窗邏輯 ★
  let previewTimeout;
  // ★ 預覽小視窗邏輯：點擊標籤顯示，點擊外面消失 ★
  function bindRefPreviews() {
      const box = document.getElementById('ref-preview-box');
      if (!box) return;

      // 建立一個處理預覽顯示的共用函式
      const showPreview = (target, e) => {
          const floor = target.getAttribute('data-floor');
          const sub = target.getAttribute('data-sub');
          const key = floor + '-' + sub;
          const targetComment = commentLookup[key];

          if (targetComment) {
              const floorTitle = "B" + floor + (parseInt(sub) > 0 ? "-" + sub : "");
              
              // 更新小視窗內容
              box.innerHTML = `
                  <div style="border-bottom: 2px solid #198754; margin-bottom: 8px; padding-bottom: 4px; display: flex; justify-content: space-between; align-items: center;">
                    <strong style="color: #198754;">${floorTitle} ${targetComment.authorName}</strong>
                    <span style="cursor:pointer; font-size:1.2rem;" onclick="document.getElementById('ref-preview-box').style.display='none'">&times;</span>
                  </div>
                  <div class="preview-inner-content">${targetComment.content}</div>
              `;

              // 如果是第一次從外部點擊，才需要計算位置
              // 如果是在視窗內點擊，通常維持原位或微調即可
              if (e) {
                  const rect = target.getBoundingClientRect();
                  box.style.display = 'block';
                  box.style.left = (rect.left + window.scrollX) + 'px';
                  box.style.top = (rect.bottom + window.scrollY + 8) + 'px';
              }
              
              // 檢查新內容裡有沒有圖片，有的話調整寬度
              if (box.querySelector('img')) box.style.width = '350px';
          }
      };

      // 監聽 1：討論區主體內的點擊
      document.getElementById('forum-detail-view').addEventListener('click', function(e) {
          const target = e.target.closest('.ref-link');
          if (target) {
              e.preventDefault();
              e.stopPropagation();
              showPreview(target, e);
          } else {
              // 點擊非連結處，隱藏視窗
              if (!box.contains(e.target)) box.style.display = 'none';
          }
      });

      // 監聽 2：小視窗內部的點擊 (實現一層接一層)
      box.addEventListener('click', function(e) {
          const target = e.target.closest('.ref-link');
          if (target) {
              e.preventDefault();
              e.stopPropagation();
              // 在視窗內點擊，不傳入 e，讓它維持在原本位置更新內容
              showPreview(target, null); 
          }
          // 阻止點擊視窗內部導致視窗關閉
          e.stopPropagation();
      });
  }
  function toggleThreadLike(tid, btn) {
    if(!user) return forceAuth();

    const icon = btn.querySelector('i');
    const countSpan = btn.querySelector('.t-like-count');
    let count = parseInt(countSpan.textContent) || 0;
    const isLiking = icon.classList.contains('bi-heart');

    // --- 1. 立即更新詳情頁 UI (讓使用者感覺很順暢) ---
    if (isLiking) {
      icon.className = 'bi bi-heart-fill';
      btn.classList.replace('btn-outline-danger', 'btn-danger');
      countSpan.textContent = count + 1;
    } else {
      icon.className = 'bi bi-heart';
      btn.classList.replace('btn-danger', 'btn-outline-danger');
      countSpan.textContent = Math.max(0, count - 1);
    }

    // --- 2. 發送至後端 ---
    google.script.run.withSuccessHandler(res => {
      if(res.status === 'success') {
        // ★ 核心修正：同步更新本地數據快取 ★
        const threadInList = appData.threads.find(t => t.id == tid);
        if (threadInList) {
          threadInList.likes = res.likes;
          threadInList.likeCount = res.likes.length;

          // ★ 關鍵：重新渲染列表，這樣回到列表時才會看到最新數字 ★
          renderForumList(); 
        }

        // 同步更新詳情頁的全域變數
        if (appData.currentThread && appData.currentThread.id == tid) {
          appData.currentThread.likes = res.likes;
          appData.currentThread.likeCount = res.likes.length;
        }
      } else {
        showMsg('錯誤', res.message, 'error');
        loadThreadDetail(tid); // 失敗則刷回正確數據
      }
    }).toggleThreadLike(tid, user.username);
  }

  function getLevelInfo(points) {
      // 根據等級調整為由淺入深的綠色系
      if (points >= 8001) return { lv: 5, name: '淨零造物主', color: '#004b23', next: 'MAX' }; // 深沉森林綠
      if (points >= 3001) return { lv: 4, name: '環境守護神', color: '#007f5f', next: 8001 };  // 科技翡翠綠
      if (points >= 1001) return { lv: 3, name: '綠能導師', color: '#2b9348', next: 3001 };    // 標準活力綠
      if (points >= 201)  return { lv: 2, name: '減碳專員', color: '#80b918', next: 1001 };    // 嫩芽草地綠
      return { lv: 1, name: '碳足跡者', color: '#aacc00', next: 201 };                       // 檸檬淺綠色
  }

  function showScoreDetails() {
      if (!user) return;

      const modal = new bootstrap.Modal(document.getElementById('scoreDetailModal'));
      modal.show();

      google.script.run.withSuccessHandler(res => {
          if(res.error) {
              document.getElementById('score-detail-content').innerHTML = "讀取失敗";
              return;
          }

          // 定義計算邏輯與顯示文字
          const rows = [
              { label: '發布簡報 (公開)', count: res.pptActive, price: 50 },
              { label: '發布簡報 (典藏)', count: res.pptArchived, price: 40 },
              { label: '發布教案 (公開)', count: res.planActive, price: 40 },
              { label: '發布教案 (典藏)', count: res.planArchived, price: 32 },
              { label: '討論區參與', count: res.threadCount, price: 30 }
          ];

          // 計算獎勵分 (總分 - 基礎分總和)
          const baseTotal = rows.reduce((sum, r) => sum + (r.count * r.price), 0);
          const rewardScore = res.totalScore - baseTotal;

          let html = `
              <style>
                  .score-grid {
                      display: grid;
                      grid-template-columns: 1fr auto auto auto auto auto;
                      gap: 0 10px;
                      align-items: center;
                      font-family: "Courier New", Courier, monospace; /* 使用等寬字體更整齊 */
                  }
                  .score-grid div { padding: 4px 0; }
                  .text-right { text-align: right; }
                  .text-center { text-align: center; }
              </style>
              <div class="score-grid">
          `;

          rows.forEach(r => {
              html += `
                  <div class="text-secondary">${r.label}</div>
                  <div class="text-right fw-bold">${r.count}</div>
                  <div class="text-center text-muted">×</div>
                  <div class="text-right">${r.price}</div>
                  <div class="text-center text-muted">=</div>
                  <div class="text-right fw-bold">${r.count * r.price}</div>
              `;
          });

          // 加入互動獎勵行
          html += `
                  <div class="text-secondary">互動獎勵 (愛心/收藏)</div>
                  <div></div><div></div><div></div>
                  <div class="text-center text-muted">+</div>
                  <div class="text-right fw-bold">${rewardScore}</div>
              </div>

              <div class="d-flex justify-content-between mt-3 pt-3 border-top fs-5 text-success">
                  <strong>目前總碳值</strong>
                  <strong style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;">${res.totalScore} CP</strong>
              </div>
          `;

          document.getElementById('score-detail-content').innerHTML = html;
      }).getPointDetails(user.username);
  }
</script>