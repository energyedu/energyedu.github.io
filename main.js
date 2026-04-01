const GAS_URL = "https://script.google.com/macros/s/AKfycbyqRrQGSF9pJuDsM1ueh5d9XyeOTL-RXHGIDrLMmEqXbisD_aQUF_W3aaCeiOmPVoRK/exec"; 

/* ==========================================
   核心通訊：封裝 Fetch 模擬 google.script.run
   ========================================== */
async function callGAS(functionName, args = []) {
    const url = `${GAS_URL}?action=${functionName}`;
    try {
        const response = await fetch(url, {
            method: 'POST', 
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

/* ==========================================
   全域變數與初始化狀態
   ========================================== */
let commentLookup = {};
let user = JSON.parse(localStorage.getItem('tb_user'));
let currentThreadId = null;
let replyToFloor = 0;
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

// 解析網址參數 (取代原本 GAS 的 <?= ?>)
const params = new URLSearchParams(window.location.search);
const initParams = Object.fromEntries(params.entries());
const appUrl = window.location.origin + window.location.pathname;

/* ==========================================
   UI 輔助小工具 (訊息、Loading、確認框)
   ========================================== */
function showMsg(title, text, type = 'info', callback = null) {
    onMsgClose = callback;
    const msgModalEl = document.getElementById('msgModal');
    if (!msgModalEl) return;

    document.getElementById('msg-title').textContent = title;
    document.getElementById('msg-body').innerHTML = text;
    
    const header = document.getElementById('msg-header');
    header.className = 'modal-header text-white ' + (type === 'success' ? 'bg-success' : type === 'error' ? 'bg-danger' : type === 'warning' ? 'bg-warning text-dark' : 'bg-secondary');
    
    const closeBtn = header.querySelector('.btn-close');
    if (closeBtn) closeBtn.className = 'btn-close ' + (type === 'warning' ? '' : 'btn-close-white');
    
    bootstrap.Modal.getOrCreateInstance(msgModalEl).show();
}

function onFailure(error) {
    toggleLoading(false);
    showMsg('系統錯誤', error.message || error.toString(), 'error');
}

function toggleLoading(show) {
    const overlay = document.getElementById('loading-overlay');
    if (overlay) overlay.classList.toggle('hidden', !show);
}

let onConfirmAction = null;
window.showConfirm = function(title, text, callback) {
    document.getElementById('confirm-title').textContent = title;
    document.getElementById('confirm-body').innerHTML = text;
    onConfirmAction = callback;
    const confirmModalEl = document.getElementById('confirmModal');
    bootstrap.Modal.getOrCreateInstance(confirmModalEl).show();
};

/* ==========================================
   初始化與路由導向
   ========================================== */
window.onload = function() { 
    // 綁定確認按鈕
    const confirmBtn = document.getElementById('btn-confirm-yes');
    if (confirmBtn) {
        confirmBtn.onclick = function() {
            const modal = bootstrap.Modal.getInstance(document.getElementById('confirmModal'));
            if (modal) modal.hide();
            if (onConfirmAction) onConfirmAction();
        };
    }

    // 監聽訊息視窗關閉
    const msgModalEl = document.getElementById('msgModal');
    if (msgModalEl) {
        msgModalEl.addEventListener('hidden.bs.modal', () => { 
            if (onMsgClose) { onMsgClose(); onMsgClose = null; } 
        });
    }

    updateAuthUI(); 
    initGradeChecks('plan-grades-check'); // 修正傳參
    initGradeFilterButtons('plan-grade-filter', filterState, renderPlans);
    initGradeFilterButtons('profile-grade-filter', profileFilterState, renderProfileList);
    initTinyMCE();
    bindRefPreviews();

    // 路由邏輯
    if (initParams.view && initParams.id) {
        if (initParams.view === 'thread_detail') {
            if (!user) { pendingTarget = initParams; forceAuth(); return; }
            switchView('forum'); loadThreadDetail(initParams.id);
        } else if (initParams.view === 'plan_detail') {
            switchView('plans'); jumpToPlan(initParams.id); 
        }
    } else if (initParams.view) {
        switchView(initParams.view);
    } else { 
        switchView('home'); 
    }
};

// 修改後的 forceAuth，確保在提示框關閉後才彈出登入框
window.forceAuth = function() {
    showMsg('權限提示', '此內容僅限會員觀看。<br>請先登入或註冊會員。', 'warning', () => {
        // 延遲 300 毫秒，等提示框完全消失後再開登入框
        setTimeout(() => {
            window.showAuthModal('login');
        }, 300);
    });
};

/* ==========================================
   核心功能：UI 切換與網址同步
   ========================================== */
window.switchView = function(viewName) {
    const navbarCollapse = document.getElementById('navbarNav');
    if (navbarCollapse && navbarCollapse.classList.contains('show')) {
        bootstrap.Collapse.getOrCreateInstance(navbarCollapse).hide();
    }

    document.body.classList.remove('modal-open');
    document.body.style.overflow = '';

    if ((viewName === 'forum' || viewName === 'profile') && !user) {
        forceAuth(); return;
    }

    ['home', 'plans', 'forum', 'profile'].forEach(v => {
        const el = document.getElementById('view-' + v);
        const nav = document.getElementById('nav-' + v);
        if (el) el.classList.toggle('hidden', v !== viewName);
        if (nav) nav.classList.toggle('active', v === viewName);
    });

    const url = new URL(window.location);
    url.searchParams.set('view', viewName);
    url.searchParams.delete('id');
    window.history.pushState({ view: viewName }, '', url);

    if (viewName === 'home') loadHome();
    if (viewName === 'plans') loadPlans();
    if (viewName === 'forum') {
        document.getElementById('forum-list-view').classList.remove('hidden');
        document.getElementById('forum-detail-view').classList.add('hidden');
        loadForum();
    }
    if (viewName === 'profile') loadProfile();
};

/* ==========================================
   初始化複選框與篩選按鈕
   ========================================== */
function initGradeChecks(containerId) {
    const container = document.getElementById(containerId);
    if (!container) return;
    let html = '';
    for (let i = 1; i <= 6; i++) {
        html += `<div class="form-check form-check-inline">
            <input class="form-check-input" type="checkbox" value="${i}"> 
            <label class="form-check-label">${i}年級</label>
        </div>`;
    }
    container.innerHTML = html;
}

function initGradeFilterButtons(containerId, stateObj, renderFunc) {
    const container = document.getElementById(containerId);
    if (!container) return;
    let html = '';
    for (let i = 1; i <= 6; i++) {
        html += `<button type="button" class="btn btn-outline-secondary" onclick="toggleGradeFilter('${containerId}', ${i})">${i}</button>`;
    }
    container.innerHTML = html;

    window.toggleGradeFilter = function(cId, grade) {
        const isProfile = (cId === 'profile-grade-filter');
        const state = isProfile ? profileFilterState : filterState;
        const idx = state.grades.indexOf(grade);
        if (idx === -1) state.grades.push(grade); else state.grades.splice(idx, 1);

        const btns = document.querySelectorAll(`#${cId} button`);
        btns.forEach((btn, index) => {
            const g = index + 1;
            btn.classList.toggle('btn-secondary', state.grades.includes(g));
            btn.classList.toggle('btn-outline-secondary', !state.grades.includes(g));
        });

        if (isProfile) { curPage['profile_' + profileTab] = 1; renderProfileList(); } 
        else { curPage.plans = 1; renderPlans(); }
    };
}

// --- 核心修正：取代原本的 google.script.history ---
function cleanUrl() {
    // GitHub Pages 環境下，使用標準 Web API 清理網址參數
    const url = new URL(window.location);
    url.searchParams.delete('id'); // 移除 ID 參數，保持網址整潔
    window.history.replaceState({}, '', url);
}

const viewPlanModalEl = document.getElementById('viewPlanModal');
if (viewPlanModalEl) {
    viewPlanModalEl.addEventListener('hidden.bs.modal', cleanUrl);
}

function initTinyMCE() {
    tinymce.init({
        selector: '#plan-content, #thread-content, #comment-input',
        menubar: false,
        plugins: 'lists link',
        toolbar: 'bold italic underline | bullist numlist | link removeformat',
        height: 200,
        setup: function(editor) {
            editor.on('change', function() {
                editor.save();
            });
        }
    });
}

window.onload = function() { 
    updateAuthUI(); 
    // 注意：這裡直接傳入 'void(0)' 字符串在 GitHub 環境下可能會失效
    // 建議 initGradeChecks 內部邏輯要處理好
    initGradeChecks('plan-grades-check', 'void(0)');
    initGradeFilterButtons('plan-grade-filter', filterState, renderPlans);
    initGradeFilterButtons('profile-grade-filter', profileFilterState, renderProfileList);
    initTinyMCE();
    
    // ★ 預覽功能啟動
    bindRefPreviews();

    // 利用我們在第一段改好的 initParams (從 URLSearchParams 來的)
    if (initParams.view && initParams.id) {
        if (initParams.view === 'thread_detail') {
            if (!user) { 
                pendingTarget = initParams; 
                forceAuth(); 
                return; 
            }
            switchView('forum'); 
            loadThreadDetail(initParams.id);
        } else if (initParams.view === 'plan_detail') {
            switchView('plans'); 
            jumpToPlan(initParams.id); 
        }
    } else if (initParams.view) {
        // 如果只有 view 沒有 id
        switchView(initParams.view);
    } else { 
        switchView('home'); 
    }
};

function copyShareLink(view, id) {
    let shareTitle = "";
    let shareText = "";
    // appUrl 是我們在第一段定義的 window.location.origin + window.location.pathname
    const link = `${appUrl}?view=${view}&id=${id}`;

    if (view === 'plan_detail') {
        const plan = appData.plans.find(p => String(p.id) === String(id));
        shareTitle = plan ? plan.title : "優秀教案資源";
        shareText = `【碳減活寶桌遊教案分享】\n推薦一個很棒的資源：${shareTitle}\n點擊連結查看詳情：\n${link}`;
    } else if (view === 'thread_detail') {
        const thread = appData.threads.find(t => String(t.id) === String(id));
        shareTitle = thread ? thread.title : "精彩討論內容";
        shareText = `【碳減活寶桌遊教案討論區】\n快來看看這則熱門討論：${shareTitle}\n加入對話：\n${link}`;
    } else {
        shareText = `歡迎來到碳減活寶桌遊教案資源網：\n${link}`;
    }

    // 執行複製
    if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(shareText).then(() => {
            showMsg('分享連結已複製', '包含標題與網址的分享訊息已存至剪貼簿，您可以直接貼上傳給他人！', 'success');
        }).catch(err => {
            console.error('複製失敗', err);
            showMsg('複製失敗', `請手動複製連結：<br>${link}`, 'warning');
        });
    } else {
        // 針對不支援 clipboard API 的舊瀏覽器或非安全環境 (http)
        showMsg('提示', `您的瀏覽器不支援自動複製，請手動複製網址：<br>${link}`, 'warning');
    }
}

/* ==========================================
   UI 切換與網址同步
   ========================================== */
function switchView(viewName) {
    // 1. 自動收合行動版漢堡選單 (Bootstrap 5)
    const navbarCollapse = document.getElementById('navbarNav');
    if (navbarCollapse && navbarCollapse.classList.contains('show')) {
        const bsCollapse = bootstrap.Collapse.getInstance(navbarCollapse) || new bootstrap.Collapse(navbarCollapse);
        bsCollapse.hide();
    }

    // 2. 強制移除殘留的 Modal 鎖定狀態 (確保頁面可捲動)
    document.body.classList.remove('modal-open');
    document.body.style.overflow = '';
    document.body.style.paddingRight = '';

    // 3. 權限檢查：討論區與個人資料必須登入
    if ((viewName === 'forum' || viewName === 'profile') && !user) {
        forceAuth();
        return;
    }

    // 4. 切換顯示/隱藏區塊
    ['home', 'plans', 'forum', 'profile'].forEach(v => {
        const el = document.getElementById('view-' + v);
        const nav = document.getElementById('nav-' + v);
        if (el) el.classList.toggle('hidden', v !== viewName);
        if (nav) nav.classList.toggle('active', v === viewName);
    });

    // 5. ★ 核心修正：同步更新網址 (GitHub 專用)
    // 這樣使用者按「首頁」時網址變 ?view=home，按「討論區」變 ?view=forum
    const url = new URL(window.location);
    url.searchParams.set('view', viewName);
    url.searchParams.delete('id'); // 切換大頁面時移除舊文章 ID
    window.history.pushState({ view: viewName }, '', url);

    // 6. 載入對應數據
    if (viewName === 'home') loadHome();
    if (viewName === 'plans') loadPlans();
    if (viewName === 'forum') {
        document.getElementById('forum-list-view').classList.remove('hidden');
        document.getElementById('forum-detail-view').classList.add('hidden');
        loadForum();
    }
    if (viewName === 'profile') loadProfile();
}

/* ==========================================
   動作選單 (編輯、典藏、刪除)
   ========================================== */
function getActionMenu(type, id, currentStatus, inModal = false) {
    const isArchived = currentStatus === 'archived';
    const isClosed = currentStatus === 'closed';
    let extraItems = '';

    if (type === 'thread') {
        extraItems += `<li><button class="dropdown-item" type="button" onclick="updateStatus('${type}', '${id}', '${isClosed ? 'active' : 'closed'}')">${isClosed ? '重新開啟' : '關閉討論'}</button></li>`;
    }

    const editOnClick = inModal ? `openEdit('${type}', '${id}', true)` : `openEdit('${type}', '${id}')`;

    return `
    <div class="dropdown d-inline-block ms-2">
      <button class="btn btn-sm btn-link text-secondary p-0" type="button" data-bs-toggle="dropdown">
        <i class="bi bi-three-dots-vertical"></i>
      </button>
      <ul class="dropdown-menu dropdown-menu-end">
        <li><button class="dropdown-item" type="button" onclick="${editOnClick}">編輯</button></li>
        <li><button class="dropdown-item" type="button" onclick="updateStatus('${type}', '${id}', '${isArchived ? 'active' : 'archived'}')">${isArchived ? '取消典藏' : '典藏'}</button></li>
        ${extraItems}
        <li><hr class="dropdown-divider"></li>
        <li><button class="dropdown-item text-danger" type="button" onclick="updateStatus('${type}', '${id}', 'deleted')">刪除</button></li>
      </ul>
    </div>`;
}

/* ==========================================
   分頁渲染組件
   ========================================== */
function renderPagination(totalItems, currentPage, targetId, onPageChange, pageSize) {
    const totalPages = Math.ceil(totalItems / pageSize);
    const container = document.getElementById(targetId);
    if (!container) return;
    
    if (totalPages <= 1) {
        container.innerHTML = '';
        return;
    }

    let html = '<ul class="pagination pagination-sm">';
    // 上一頁
    html += `<li class="page-item ${currentPage === 1 ? 'disabled' : ''}">
             <button class="page-link" onclick="window.${onPageChange}(${currentPage - 1})">上一頁</button></li>`;

    // 頁碼 (在 GitHub 模式下，確保呼叫的是全域 window 下的 function)
    for (let i = 1; i <= totalPages; i++) {
        html += `<li class="page-item ${i === currentPage ? 'active' : ''}">
                 <button class="page-link" onclick="window.${onPageChange}(${i})">${i}</button></li>`;
    }

    // 下一頁
    html += `<li class="page-item ${currentPage === totalPages ? 'disabled' : ''}">
             <button class="page-link" onclick="window.${onPageChange}(${currentPage + 1})">下一頁</button></li>`;
    
    html += '</ul>';
    container.innerHTML = html;
}

/* ==========================================
   首頁數據載入 (公告與版本紀錄)
   ========================================== */
function loadHome() {
    // 使用 callGAS 取代 google.script.run
    callGAS('getData', ['home'])
        .then(res => {
            if (res.error) return onFailure(new Error(res.error));
            
            // 儲存數據到全域變數
            appData.announcements = res.announcements;
            appData.versions = res.versions;
            
            // 執行渲染
            renderAnnouncements();
            renderVersions();
        })
        .catch(onFailure);
}

/* ==========================================
   顯示公告詳情 (showAnnounce)
   ========================================== */
window.showAnnounce = function(id) {
    // 從 appData 中尋找對應 ID 的公告
    const a = appData.announcements.find(x => String(x.id) === String(id));
    if (!a) return;

    // 填入標題與日期
    document.getElementById('ann-title').textContent = a.title;
    document.getElementById('ann-date').textContent = a.startStr;

    // 將內容中的網址轉為可點擊連結的輔助函式
    function linkify(text) {
        return text.replace(/(https?:\/\/[^\s]+)/g, '<a href="$1" target="_blank">$1</a>');
    }

    // 處理內文：HTML 逸出 -> 網址連結化 -> 換行符轉 <br>
    let contentHtml = linkify(esc(a.body)).replace(/\n/g, '<br>');

    // 如果公告有圖片網址 (以逗號分隔)，則生成 img 標籤
    if (a.imgsRaw) {
        const urls = a.imgsRaw.split(',').map(u => u.trim()).filter(u => u);
        contentHtml += '<div class="mt-3">' + 
            urls.map(u => `<img src="${u}" class="img-fluid rounded mb-2" style="max-height:300px;">`).join('') + 
            '</div>';
    }

    // 填入 Modal 並顯示
    document.getElementById('ann-body').innerHTML = contentHtml;
    const announceModalEl = document.getElementById('announceModal');
    const bsModal = bootstrap.Modal.getOrCreateInstance(announceModalEl);
    bsModal.show();
};

/* ==========================================
   公告渲染與顯示
   ========================================== */
function renderAnnouncements() {
    const list = document.getElementById('home-announce-list');
    if (!list) return;

    const start = (curPage.announcements - 1) * HOME_PER_PAGE;
    const items = appData.announcements.slice(start, start + HOME_PER_PAGE);

    renderPagination(appData.announcements.length, curPage.announcements, 'announce-pagination', 'changeAnnouncePage', HOME_PER_PAGE);

    if (items.length === 0) {
        list.innerHTML = '<div class="p-3 text-center">目前無公告</div>';
    } else {
        list.innerHTML = items.map(a => `
            <div class="list-group-item list-group-item-action p-3 cursor-pointer ${a.pin ? 'pinned' : ''}" 
                 onclick="window.showAnnounce('${a.id}')"> <div class="d-flex w-100 justify-content-between">
                    <h6 class="mb-1 title text-truncate">
                        ${a.pin ? '<span class="badge bg-danger me-1">置頂</span>' : ''} ${esc(a.title)}
                    </h6>
                    <small class="text-muted">${a.startStr}</small>
                </div>
            </div>`).join('');
    }
}
// 務必將 showAnnounce 掛載到 window
window.showAnnounce = showAnnounce;

function renderAnnouncements() {
    const list = document.getElementById('home-announce-list');
    if (!list) return;

    const start = (curPage.announcements - 1) * HOME_PER_PAGE;
    const items = appData.announcements.slice(start, start + HOME_PER_PAGE);

    // 渲染分頁按鈕
    renderPagination(appData.announcements.length, curPage.announcements, 'announce-pagination', 'changeAnnouncePage', HOME_PER_PAGE);

    if (items.length === 0) {
        list.innerHTML = '<div class="p-3 text-center">目前無公告</div>';
    } else {
        list.innerHTML = items.map(a => `
            <div class="list-group-item list-group-item-action p-3 cursor-pointer ${a.pin ? 'pinned' : ''}" onclick="showAnnounce('${a.id}')">
                <div class="d-flex w-100 justify-content-between">
                    <h6 class="mb-1 title text-truncate">
                        ${a.pin ? '<span class="badge bg-danger me-1">置頂</span>' : ''} ${esc(a.title)}
                    </h6>
                    <small class="text-muted">${a.startStr}</small>
                </div>
            </div>`).join('');
    }
}

// 註冊分頁切換函式到 window
window.changeAnnouncePage = function(p) {
    curPage.announcements = p;
    renderAnnouncements();
};

/* ==========================================
   版本紀錄渲染
   ========================================== */
function renderVersions() {
    const list = document.getElementById('home-version-list');
    if (!list) return;

    const start = (curPage.versions - 1) * HOME_PER_PAGE;
    const items = (appData.versions || []).slice(start, start + HOME_PER_PAGE);

    renderPagination((appData.versions || []).length, curPage.versions, 'version-pagination', 'changeVersionPage', HOME_PER_PAGE);

    if (items.length === 0) {
        list.innerHTML = '<p class="text-muted p-3">尚無更新紀錄</p>';
    } else {
        let vHtml = '<table class="table table-hover mb-0"><thead><tr><th>版本</th><th>日期</th><th>內容</th></tr></thead><tbody>';
        vHtml += items.map(v => {
            // 格式化日期，只取前 10 碼 (YYYY/MM/DD)
            let dateDisplay = "";
            if (v.date === 'Coming Soon') {
                dateDisplay = '<span class="badge bg-warning text-dark">Coming Soon</span>';
            } else if (v.date) {
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

// 註冊分頁切換函式到 window
window.changeVersionPage = function(p) {
    curPage.versions = p;
    renderVersions();
};

/* ==========================================
   檔案資源數據載入 (Plans)
   ========================================== */
function loadPlans() {
    const btnAdd = document.getElementById('btn-add-plan');
    const guestHint = document.getElementById('guest-plan-hint');
    const list = document.getElementById('plans-list');

    // 1. 權限與 UI 初始狀態
    if (user) {
        if (btnAdd) btnAdd.classList.remove('hidden');
        if (guestHint) guestHint.classList.add('hidden');
    } else {
        if (btnAdd) btnAdd.classList.add('hidden');
        if (guestHint) guestHint.classList.remove('hidden');
    }

    // 2. 顯示載入動畫
    if (list && list.innerHTML.trim() === '') {
        list.innerHTML = '<div class="text-center w-100 py-5"><div class="spinner-border text-success"></div></div>';
    }

    // 3. 呼叫 GAS API
    callGAS('getData', ['plans', user ? user.username : null])
        .then(res => {
            if (res.error) return onFailure(new Error(res.error));
            
            appData.plans = res.plans;
            renderPlanStats(res.stats);
            
            // 預設篩選「全部」
            const firstBtn = document.querySelector('#view-plans .btn-group button:first-child');
            if (firstBtn) {
                filterPlans('all', firstBtn);
            } else {
                renderPlans();
            }
        })
        .catch(onFailure);
}

/* ==========================================
   統計資訊條渲染
   ========================================== */
function renderPlanStats(stats) {
    const div = document.getElementById('plan-stats-bar');
    if (!div) return;

    if (!stats) {
        div.classList.add('hidden');
        return;
    }

    div.classList.remove('hidden');
    div.innerHTML = `
        <div class="col-md-6">
            <div class="card stat-card stat-card-blue h-100 shadow-sm p-3">
                <i class="bi bi-file-earmark-word stat-icon"></i>
                <div class="d-flex align-items-center mb-2">
                    <h4 class="mb-0 fw-bold">教案</h4>
                </div>
                <div class="row text-center mt-2">
                    <div class="col-6 border-end border-white border-opacity-25">
                        <small class="text-white-50">公開</small>
                        <div class="stat-num">${stats.wordPub}</div>
                    </div>
                    <div class="col-6">
                        <small class="text-white-50">會員</small>
                        <div class="stat-num">${stats.wordMem}</div>
                    </div>
                </div>
            </div>
        </div>
        <div class="col-md-6">
            <div class="card stat-card stat-card-red h-100 shadow-sm p-3">
                <i class="bi bi-file-earmark-slides stat-icon"></i>
                <div class="d-flex align-items-center mb-2">
                    <h4 class="mb-0 fw-bold">簡報</h4>
                </div>
                <div class="row text-center mt-2">
                    <div class="col-6 border-end border-white border-opacity-25">
                        <small class="text-white-50">公開</small>
                        <div class="stat-num">${stats.pptPub}</div>
                    </div>
                    <div class="col-6">
                        <small class="text-white-50">會員</small>
                        <div class="stat-num">${stats.pptMem}</div>
                    </div>
                </div>
            </div>
        </div>`;
}

/* ==========================================
   篩選與搜尋邏輯
   ========================================== */
function filterPlans(type, btn) {
    const group = btn.parentElement;
    if (!group) return;

    // 重置按鈕樣式
    Array.from(group.children).forEach(b => {
        const txt = b.textContent.trim();
        b.className = 'btn';
        if (txt === '全部') b.classList.add('btn-outline-secondary');
        else if (txt === '教案') b.classList.add('btn-outline-word');
        else if (txt === '簡報') b.classList.add('btn-outline-ppt');
    });

    // 套用選中樣式
    if (type === 'all') btn.className = 'btn btn-secondary active';
    else if (type === '教案') btn.className = 'btn btn-word active';
    else if (type === '簡報') btn.className = 'btn btn-ppt active';

    filterState.category = type;
    curPage.plans = 1;
    renderPlans();
}

// 註冊搜尋事件
function handlePlanSearch(val) {
    filterState.planSearch = val.toLowerCase();
    curPage.plans = 1;
    renderPlans();
}

// 個人資料頁面的分類篩選 (邏輯與 filterPlans 相似)
function setProfileFilter(cat, btn) {
    const group = btn.parentElement;
    if (!group) return;

    Array.from(group.children).forEach(b => {
        const txt = b.textContent.trim();
        b.className = 'btn';
        if (txt === '全部') b.classList.add('btn-outline-secondary');
        else if (txt === '教案') b.classList.add('btn-outline-word');
        else if (txt === '簡報') b.classList.add('btn-outline-ppt');
    });

    if (cat === 'all') btn.className = 'btn btn-secondary active';
    else if (cat === '教案') btn.className = 'btn btn-word active';
    else if (cat === '簡報') btn.className = 'btn btn-ppt active';

    // 修正：同步更新 profileFilterState 並重新渲染
    profileFilterState.category = cat; 
    curPage['profile_' + profileTab] = 1;
    renderProfileList();
}


/* ==========================================
   檔案發布與編輯 (submitPlan)
   ========================================== */
/* 範例修正：上傳檔案 */
window.submitPlan = function() {
    const grades = Array.from(document.querySelectorAll('#plan-grades-check input:checked')).map(c => parseInt(c.value));
    const title = document.getElementById('plan-title').value.trim();
    const link = document.getElementById('plan-link').value.trim();
    const editor = tinymce.get('plan-content');
    const content = editor ? editor.getContent() : '';
    const textContent = editor ? editor.getContent({format: 'text'}).replace(/\u00a0/g, ' ').trim() : '';

    if (!title) return showMsg('提示', '請輸入標題', 'warning');
    if (grades.length === 0) return showMsg('提示', '請選擇適用年級', 'warning');

    const id = document.getElementById('plan-id').value;
    const data = { author: user.username, access: document.getElementById('plan-access').value, category: document.getElementById('plan-category').value, title, content, link, grades };
    
    toggleLoading(true);
    const method = id ? 'editPostContent' : 'createPlan';
    const args = id ? [id, data, user.username] : [data];

    callGAS(method, args).then(res => {
        toggleLoading(false);
        if (res.status === 'success') {
            bootstrap.Modal.getInstance(document.getElementById('planModal')).hide();
            showMsg('成功', '檔案已儲存', 'success');
            loadPlans();
        } else {
            showMsg('失敗', res.message, 'error');
        }
    }).catch(onFailure);
};

/* ==========================================
   發起或編輯討論 (submitThread)
   ========================================== */
function submitThread() { 
    const id = document.getElementById('thread-id').value; 
    const title = document.getElementById('thread-title').value.trim();
    const content = tinymce.get('thread-content').getContent(); 
    
    if (title.length === 0) return showMsg('資料不完整', '請輸入<b>討論標題</b>', 'warning');

    const data = { author: user.username, title: title, content: content }; 
    const handler = id ? 'editThreadContent' : 'createThread'; 
    const args = id ? [id, data, user.username] : [data]; 

    toggleLoading(true);

    // 取代原本的 google.script.run
    callGAS(handler, args)
        .then(() => {
            toggleLoading(false);
            // 關閉 Modal
            const modalEl = document.getElementById('threadModal');
            const bsModal = bootstrap.Modal.getInstance(modalEl);
            if (bsModal) bsModal.hide();

            // 判斷是要重新載入詳情還是回到列表
            if (id && currentThreadId) {
                loadThreadDetail(currentThreadId);
            } else {
                loadForum();
            }
            
            document.getElementById('thread-id').value = ''; 
        })
        .catch(onFailure);
}

/* ==========================================
   送出留言回覆 (submitComment)
   ========================================== */
function submitComment() {
    const editor = tinymce.get('comment-input');
    const content = editor.getContent();
    
    // 從回覆指示器中擷取標籤 (例如 B1-1)
    let replyTag = "";
    const indicator = document.getElementById('reply-indicator');
    if (indicator && !indicator.classList.contains('hidden')) {
        const match = indicator.textContent.match(/B\d+(?:-\d+)?/);
        if (match) replyTag = match[0];
    }

    const data = { 
        threadId: currentThreadId, 
        author: user.username, 
        content: content, 
        parentFloor: replyToFloor, 
        replyTo: replyTag 
    };

    toggleLoading(true);

    // 取代原本的 google.script.run.createComment(data)
    callGAS('createComment', [data])
        .then(() => {
            toggleLoading(false);
            editor.setContent('');
            replyToFloor = 0;
            if (indicator) indicator.classList.add('hidden');
            
            // 重新載入該討論串以顯示最新留言
            loadThreadDetail(currentThreadId);
        })
        .catch(onFailure);
}

function renderPlans() {
    try {
        const container = document.getElementById('plans-list');
        if (!container) return; // 安全檢查

        // 1. 執行過濾邏輯 (完全保留原本邏輯)
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
        // 注意：這裡呼叫 renderPagination 時，確定第四個參數是全域 function 的名稱字串
        renderPagination(displayPlans.length, curPage.plans, 'plans-pagination', 'changePlanPage', ITEMS_PER_PAGE);

        // 3. 空資料處理
        if (paginatedPlans.length === 0) {
            container.innerHTML = '<p class="text-center text-muted w-100 py-5">暫無符合條件的檔案</p>';
            return;
        }

        // 4. 渲染卡片內容
        container.innerHTML = paginatedPlans.map(p => {
            const isMine = user && user.username === p.author;
            // 確保 saved 是陣列，避免 includes 報錯
            const savedList = (user && user.saved) ? user.saved.map(String) : [];
            const isSaved = savedList.includes(String(p.id));

            // 標籤樣式判斷
            const catClass = p.category === '簡報' ? 'badge-ppt' : 'badge-word';
            const accClass = p.access === 'public' ? 'badge-public' : 'badge-member';
            const typeBadge = `<span class="badge ${catClass} me-1">${p.category}</span>`;
            const accessBadge = `<span class="badge ${accClass}">${p.access === 'public' ? '公開' : '會員'}</span>`;

            // 年級標籤
            let gradeTags = '';
            if (p.grades && p.grades.length > 0) {
                gradeTags = '<div class="mt-1 small text-secondary">適用：' + p.grades.map(g => g + '年級').join('、') + '</div>';
            }

            // 收藏按鈕邏輯
            const sCount = p.saveCount || 0;
            const starIcon = isSaved ? 'bi-star-fill text-white' : 'bi-star';
            const btnClass = isSaved ? 'btn-warning' : 'btn-outline-warning';
            const labelText = isSaved ? '已收藏' : '收藏';

            let btns = '';
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
    } catch (e) {
        console.error("Render Plans Error:", e);
        onFailure(e);
    }
}

// 註冊到全域，確保分頁按鈕點擊時找得到
window.changePlanPage = function(page) {
    curPage.plans = page;
    renderPlans();
};

/* ==========================================
   開啟教案詳細資訊 (jumpToPlan)
   ========================================== */
function jumpToPlan(id) {
    const sid = String(id);
    let p = null;
    let isFromMyArchive = false;

    // 1. 優先從本地數據搜尋 (我的典藏 -> 全部教案 -> 個人資料快取)
    if (appData.myData && appData.myData.archived) {
        p = appData.myData.archived.find(x => String(x.id) === sid);
        if (p) isFromMyArchive = true;
    }

    if (!p) {
        p = appData.plans.find(x => String(x.id) === sid);
    }
    
    if (!p && appData.myData) {
        p = (appData.myData.plans || []).find(x => String(x.id) === sid) ||
            (appData.myData.saved || []).find(x => String(x.id) === sid);
    }

    // 2. 如果本地有資料，直接檢查權限並渲染
    if (p) {
        const isMine = isFromMyArchive || (user && String(user.username) === String(p.author));

        if (String(p.status) === 'archived' && !isMine) { 
            showMsg('無法觀看', '此貼文已被作者典藏，暫時無法查看內容。', 'warning'); 
            return; 
        }
        
        if (p.access === 'member' && !user) { 
            pendingTarget = { view: 'plan_detail', id: sid }; 
            forceAuth(); 
            return; 
        }

        renderPlanDetailModal(p, isMine);
    } else {
        // 3. 本地沒資料，呼叫 GAS API 抓取
        toggleLoading(true);
        callGAS('getData', ['plan_detail', sid])
            .then(res => {
                toggleLoading(false);
                if (res.error) return showMsg('錯誤', res.error, 'error');
                
                const pRes = res.plan;
                const isMineRes = user && (String(user.username) === String(pRes.author));
                
                if (String(pRes.status) === 'archived' && !isMineRes) {
                    showMsg('無法觀看', '此貼文已被作者典藏。', 'warning');
                    return;
                }
                
                if (pRes.access === 'member' && !user) { 
                    pendingTarget = { view: 'plan_detail', id: sid }; 
                    forceAuth(); 
                    return; 
                }

                renderPlanDetailModal(pRes, isMineRes);
            })
            .catch(onFailure);
    }
}

/* ==========================================
   渲染教案詳細資訊彈窗 (Helper)
   ========================================== */
function renderPlanDetailModal(p, isMine) {
    const sid = String(p.id);
    const menu = isMine ? getActionMenu('post', p.id, p.status, true) : '';
    const catClass = p.category === '簡報' ? 'badge-ppt' : 'badge-word';
    
    document.getElementById('view-plan-title').textContent = p.title;
    document.getElementById('view-plan-actions').innerHTML = menu;
    document.getElementById('view-plan-badge').innerHTML = `<span class="badge ${catClass}">${p.category}</span>`;
    
    let gradeTxt = p.grades && p.grades.length > 0 ? ' | 適用：' + p.grades.map(g => g + '年級').join('、') : '';
    document.getElementById('view-plan-grades').textContent = gradeTxt;
    
    const authorDisplayName = p.authorName || (isMine ? user.nickname : '未知');
    document.getElementById('view-plan-meta').textContent = `作者: ${authorDisplayName} | 發布於: ${formatDate(p.timestamp)}`;
    
    document.getElementById('view-plan-content').innerHTML = p.content;
    document.getElementById('view-plan-link').innerHTML = p.link ? `<a href="${p.link}" target="_blank" class="btn btn-success w-100">前往下載</a>` : '';
    
    // 更新分享按鈕點擊事件
    document.getElementById('view-plan-share-btn').onclick = () => copyShareLink('plan_detail', sid);
    
    // 顯示 Modal
    const modalEl = document.getElementById('viewPlanModal');
    new bootstrap.Modal(modalEl).show();

    // ★ 核心修正：同步網址列 (GitHub 專用)
    const url = new URL(window.location);
    url.searchParams.set('view', 'plan_detail');
    url.searchParams.set('id', sid);
    window.history.pushState({ view: 'plan_detail', id: sid }, '', url);
}
  
/* ==========================================
   討論區列表載入與跳轉
   ========================================== */
function jumpToThread(id) { 
    if (!user) { 
        pendingTarget = { view: 'thread_detail', id: id }; 
        forceAuth(); 
        return; 
    } 
    switchView('forum'); 
    loadThreadDetail(id); 
}

function loadForum() { 
    const list = document.getElementById('forum-threads'); 
    if (list && list.innerHTML.trim() === '') {
        list.innerHTML = '<div class="text-center py-5"><div class="spinner-border text-primary"></div><div class="mt-2 text-muted">載入討論中...</div></div>'; 
    }

    // 取代 google.script.run
    callGAS('getData', ['forum', user ? user.username : null])
        .then(res => {
            if (res.error) return onFailure(new Error(res.error)); 
            appData.threads = res.threads; 
            renderForumList(); 
        })
        .catch(onFailure);
}

/* ==========================================
   討論區搜尋與分頁
   ========================================== */
function handleForumSearch(val) { 
    filterState.forumSearch = val.toLowerCase(); 
    curPage.forum = 1; 
    renderForumList(); 
}

window.changeForumPage = function(page) { 
    curPage.forum = page; 
    renderForumList(); 
};

/* ==========================================
   討論區列表渲染 (Forum List)
   ========================================== */
function renderForumList() {
    const list = document.getElementById('forum-threads');
    if (!list) return;

    let displayThreads = appData.threads;
    
    // 1. 搜尋過濾
    if (filterState.forumSearch) { 
        const term = filterState.forumSearch; 
        displayThreads = displayThreads.filter(t => 
            t.title.toLowerCase().includes(term) || 
            t.content.toLowerCase().includes(term) || 
            t.authorName.toLowerCase().includes(term)
        ); 
    }

    // 2. 分頁計算
    const start = (curPage.forum - 1) * FORUM_PER_PAGE; 
    const items = displayThreads.slice(start, start + FORUM_PER_PAGE); 
    
    renderPagination(displayThreads.length, curPage.forum, 'forum-pagination', 'changeForumPage', FORUM_PER_PAGE);

    // 3. 渲染內容
    if (items.length === 0) {
        list.innerHTML = '<p class="text-center text-muted py-5">暫無討論內容</p>'; 
    } else {
        list.innerHTML = items.map((t, index) => {
            const serial = start + index + 1;

            // 定義狀態標籤
            const archiveBadge = t.status === 'archived' ? '<span class="badge bg-warning text-dark ms-1">已典藏</span>' : '';
            const closedBadge = t.status === 'closed' ? '<span class="badge bg-danger ms-1"><i class="bi bi-lock-fill"></i> 此討論區目前僅供瀏覽</span>' : '';

            // 點讚與愛心狀態
            const isLiked = user && t.likes && t.likes.includes(user.username);
            const heartClass = isLiked ? 'bi-heart-fill text-danger' : 'bi-heart';

            // 移除內容中的 HTML 標籤用於預覽
            const plainText = t.content.replace(/<[^>]*>?/gm, '');

            return `
                <div class="list-group-item list-group-item-action p-3 cursor-pointer" onclick="loadThreadDetail('${t.id}')">
                    <div class="d-flex w-100 justify-content-between">
                        <h5 class="mb-1 text-success fw-bold">
                            <span class="text-muted me-2">#${serial}</span>${esc(t.title)}
                            ${archiveBadge} ${closedBadge}
                        </h5>
                        <small class="text-muted">${formatDate(t.timestamp)}</small>
                    </div>
                    <p class="mb-1 text-secondary text-truncate" style="max-width: 90%;">${esc(plainText)}</p>

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

/* ==========================================
   渲染留言外觀 (支援樓中樓)
   ========================================== */
function renderCommentItem(c, isSub) {
    const isMine = user && user.username === c.author;
    const isArchived = c.status === 'archived';

    // 內容遮蔽邏輯：已典藏且不是本人則隱藏內容
    let displayContent = c.content;
    if (isArchived && !isMine) {
        displayContent = '<i class="text-muted">（此留言已由作者典藏，暫時無法查看內容）</i>';
    }

    const cMenu = isMine ? getActionMenu('comment', c.id, c.status) : '';
    const cArchiveBadge = isArchived ? '<span class="badge bg-warning text-dark me-2">已典藏</span>' : '';
    
    // 樓層顯示 (例如 B1 或 B1-1)
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

/* ==========================================
   載入討論串詳情 (loadThreadDetail)
   ========================================== */
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

    // ★ 核心修正：同步網址列 (GitHub 專用)
    const url = new URL(window.location);
    url.searchParams.set('view', 'thread_detail');
    url.searchParams.set('id', tid);
    window.history.pushState({ view: 'thread_detail', id: tid }, '', url);

    // 取代 google.script.run.getData('thread_detail', tid)
    callGAS('getData', ['thread_detail', tid])
        .then(res => {
            toggleLoading(false); 
            if (res.error) { showMsg('錯誤', res.error, 'error'); return; } 

            try { 
                const t = res.thread; 
                appData.currentThread = t; 
                const isMine = user && user.username === t.author; 

                // 1. 判斷愛心狀態
                const tLikes = t.likes || [];
                const isThreadLiked = user && tLikes.includes(user.username);
                const tHeartClass = isThreadLiked ? 'btn-danger' : 'btn-outline-danger';
                const tHeartIcon = isThreadLiked ? 'bi-heart-fill' : 'bi-heart';

                // 2. 渲染標題區
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

                // 狀態處理（關閉討論區判斷）
                const statusBadgeArea = document.getElementById('detail-status-badge');
                if (statusBadgeArea) {
                    statusBadgeArea.innerHTML = t.status === 'closed' ? '<span class="badge bg-secondary ms-2"><i class="bi bi-lock-fill"></i> 此討論區目前僅供瀏覽</span>' : '';
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

                // --- 處理留言查詢表 (Lookup) ---
                commentLookup = {};
                res.comments.forEach(c => {
                    commentLookup[c.floor + '-' + c.subFloor] = c;
                });

                // --- 留言分組處理 (主樓 + 樓中樓) ---
                let commentGroups = {}; 
                res.comments.forEach(c => {
                    const floor = c.floor;
                    if (!commentGroups[floor]) commentGroups[floor] = { main: null, subs: [] };
                    if (c.subFloor === 0) commentGroups[floor].main = c;
                    else commentGroups[floor].subs.push(c);
                });

                let html = '';
                Object.keys(commentGroups).sort((a,b) => a - b).forEach(floorKey => {
                    const group = commentGroups[floorKey];
                    if (group.main) {
                        html += renderCommentItem(group.main, false); 
                        group.subs.sort((a,b) => a.subFloor - b.subFloor).forEach(sub => {
                            html += renderCommentItem(sub, true);
                        });
                    }
                });
                area.innerHTML = html; 

                // ★ 重要：內容渲染完後重新綁定預覽小視窗
                bindRefPreviews();

            } catch (e) { 
                console.error("渲染討論詳情發生錯誤:", e);
                onFailure(e); 
            } 
        })
        .catch(onFailure);
}

/* ==========================================
   編輯表單初始化 (openEdit)
   ========================================== */
function openEdit(type, id, closeViewModal = false) {
    if (closeViewModal) {
        const viewModal = bootstrap.Modal.getInstance(document.getElementById('viewPlanModal'));
        if (viewModal) viewModal.hide();
    }

    let data = null;
    // 從快取中尋找資料
    if (type === 'post') {
        if (appData.myData && (appData.myData.plans || appData.myData.archived)) {
            data = (appData.myData.plans || []).find(p => String(p.id) === String(id)) || 
                   (appData.myData.archived || []).find(p => String(p.id) === String(id));
        }
        if (!data) data = appData.plans.find(p => String(p.id) === String(id));
    } else if (type === 'thread') {
        data = (appData.currentThread && String(appData.currentThread.id) === String(id)) ? 
               appData.currentThread : appData.threads.find(t => String(t.id) === String(id));
    }

    if (!data && type !== 'comment') return showMsg('錯誤', '找不到資料', 'error');

    // 1. 處理教案/簡報編輯
    if (type === 'post') {
        document.getElementById('plan-modal-title').textContent = '編輯檔案';
        document.getElementById('plan-id').value = id;
        document.getElementById('plan-title').value = data.title;
        tinymce.get('plan-content').setContent(data.content);
        document.getElementById('plan-link').value = data.link;
        document.getElementById('plan-category').value = data.category || '教案';
        document.getElementById('plan-access').value = data.access || 'member';
        
        // 重置並勾選年級
        document.querySelectorAll('#plan-grades-check input').forEach(c => c.checked = false);
        if (data.grades) {
            data.grades.forEach(g => {
                let ck = document.querySelector(`#plan-grades-check input[value="${g}"]`);
                if (ck) ck.checked = true;
            });
        }
        new bootstrap.Modal(document.getElementById('planModal')).show();
    } 
    // 2. 處理討論主題編輯
    else if (type === 'thread') {
        document.getElementById('thread-modal-title').textContent = '編輯討論';
        document.getElementById('thread-id').value = id;
        document.getElementById('thread-title').value = data.title;
        tinymce.get('thread-content').setContent(data.content);
        new bootstrap.Modal(document.getElementById('threadModal')).show();
    } 
    // 3. 處理留言編輯 (直接使用 Prompt)
    else if (type === 'comment') {
        const newContent = prompt("編輯留言內容:", "");
        if (newContent) {
            callGAS('editCommentContent', [id, newContent, user.username])
                .then(() => loadThreadDetail(currentThreadId))
                .catch(onFailure);
        }
    }
}

/* ==========================================
   開啟空白表單 (新增)
   ========================================== */
function openPlanModal() {
    document.getElementById('plan-modal-title').textContent = '上傳檔案';
    document.getElementById('plan-id').value = '';
    document.getElementById('plan-title').value = '';
    tinymce.get('plan-content').setContent('');
    document.getElementById('plan-link').value = '';
    document.querySelectorAll('#plan-grades-check input').forEach(c => c.checked = false);
    new bootstrap.Modal(document.getElementById('planModal')).show();
}

function openThreadModal() {
    document.getElementById('thread-modal-title').textContent = '發起討論';
    document.getElementById('thread-id').value = '';
    document.getElementById('thread-title').value = '';
    tinymce.get('thread-content').setContent('');
    new bootstrap.Modal(document.getElementById('threadModal')).show();
}

/* ==========================================
   狀態更新 (典藏/刪除/關閉討論)
   ========================================== */
function updateStatus(type, id, newStatus) {
    const msgMap = {
        'deleted': '確定要刪除嗎？刪除後無法復原。',
        'closed': '確定要關閉討論嗎？關閉後將無法回應。'
    };
    const msg = msgMap[newStatus] || '確定要變更狀態嗎？';

    showConfirm('操作確認', msg, () => {
        const handlers = { 'post': 'updatePostStatus', 'thread': 'updateThreadStatus', 'comment': 'updateCommentStatus' };
        
        callGAS(handlers[type], [id, newStatus, user.username])
            .then(res => {
                if (res.status === 'success') {
                    showMsg('成功', '狀態已更新', 'success');
                    // 根據目前所在頁面決定重新載入邏輯
                    if (type === 'post') {
                        loadPlans();
                        if (!document.getElementById('view-profile').classList.contains('hidden')) loadProfile();
                    } else if (type === 'thread') {
                        if (newStatus === 'deleted') backToForumList();
                        else if (currentThreadId) loadThreadDetail(currentThreadId);
                        else loadForum();
                    } else if (type === 'comment') {
                        loadThreadDetail(currentThreadId);
                    }
                } else {
                    onFailure(new Error(res.message));
                }
            })
            .catch(onFailure);
    });
}

/* ==========================================
   載入個人資料與 Carbon Points (進度條渲染)
   ========================================== */
function loadProfile() {
    if (!user) return forceAuth();

    document.getElementById('profile-nickname').textContent = user.nickname;
    document.getElementById('profile-username').textContent = '@' + user.username;

    const listContainer = document.getElementById('profile-list-container');
    if (listContainer) {
        listContainer.innerHTML = '<div class="text-center py-5"><div class="spinner-border text-success"></div></div>';
    }

    callGAS('getData', ['profile', user.username])
        .then(res => {
            if (res.error) return onFailure(new Error(res.error));

            appData.myData = res;

            // 渲染碳值等級與進度條
            const points = res.carbonPoints || 0;
            const info = getLevelInfo(points);
            const progress = info.next === 'MAX' ? 100 : Math.min(100, (points / info.next) * 100);

            const carbonHtml = `
                <div id="carbon-status-box-inner" onclick="showScoreDetails()" style="cursor: pointer;" title="點擊查看分數明細"
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
                </div>`;

            const cardBody = document.querySelector('#view-profile .card-body');
            if (cardBody) {
                const oldBox = document.getElementById('carbon-status-box');
                if (oldBox) oldBox.remove();
                const div = document.createElement('div');
                div.id = 'carbon-status-box';
                div.innerHTML = carbonHtml;
                // 插在「編輯資料」按鈕之前或統計數字之前
                cardBody.insertBefore(div, cardBody.querySelector('.text-start'));
            }

            // 更新統計數字
            const statsMap = { 'share': 'share', 'archive': 'archive', 'saved': 'saved', 'threads': 'threads', 'replies': 'replies' };
            Object.keys(statsMap).forEach(key => {
                const el = document.getElementById('count-' + key);
                if (el) el.textContent = res.stats[statsMap[key]];
            });

            renderProfileList();
        })
        .catch(onFailure);
}

/* ==========================================
   個人資料頁籤切換 (changeProfileTab)
   ========================================== */
function changeProfileTab(tabName) {
    profileTab = tabName;
    
    // 1. 判斷是否顯示過濾工具列 (僅分享、典藏、珍藏需要)
    const showFilter = ['plans', 'archived', 'saved'].includes(tabName);
    const filterContainer = document.getElementById('profile-filter-container');
    if (filterContainer) filterContainer.classList.toggle('hidden', !showFilter);

    // 2. 重置過濾狀態
    profileFilterState.category = 'all';
    
    // 3. 重置按鈕 UI 樣式
    const btns = document.querySelectorAll('#profile-filter-container button');
    btns.forEach(b => {
        const txt = b.textContent.trim();
        b.className = 'btn btn-outline-secondary';
        if (txt === '全部') {
            b.className = 'btn btn-secondary active';
        } else if (txt === '教案') {
            b.classList.add('btn-outline-word');
        } else if (txt === '簡報') {
            b.classList.add('btn-outline-ppt');
        }
    });

    // 4. 重置分頁並重新渲染
    curPage['profile_' + profileTab] = 1;
    renderProfileList();
}

/* ==========================================
   個人資料列表渲染 (renderProfileList)
   ========================================== */
function renderProfileList() {
    const container = document.getElementById('profile-list-container');
    const pgId = 'profile-pagination-container';

    if (!container || !appData.myData) return;

    // 1. 根據目前頁籤選取資料源
    let items = [];
    if (profileTab === 'plans') items = appData.myData.plans || [];
    else if (profileTab === 'archived') items = appData.myData.archived || [];
    else if (profileTab === 'saved') items = appData.myData.saved || [];
    else if (profileTab === 'threads') items = appData.myData.threads || [];
    else if (profileTab === 'replies') items = appData.myData.replies || [];

    // 2. 執行篩選邏輯 (分類與年級)
    if (['plans', 'archived', 'saved'].includes(profileTab)) {
        // 過濾類別
        if (profileFilterState.category && profileFilterState.category !== 'all') {
            items = items.filter(i => i.category === profileFilterState.category);
        }
        // 過濾年級
        if (profileFilterState.grades && profileFilterState.grades.length > 0) {
            items = items.filter(i => i.grades && i.grades.some(g => profileFilterState.grades.includes(g)));
        }
    }

    // 3. 分頁計算
    const cPage = curPage['profile_' + profileTab] || 1;
    const start = (cPage - 1) * ITEMS_PER_PAGE;
    const paginatedItems = items.slice(start, start + ITEMS_PER_PAGE);

    renderPagination(items.length, cPage, pgId, 'changeProfilePage', ITEMS_PER_PAGE);

    // 4. 空資料處理
    if (paginatedItems.length === 0) {
        container.innerHTML = '<div class="text-center py-5 text-muted"><i class="bi bi-inbox fs-1 d-block mb-2"></i>目前尚無資料</div>';
        return;
    }

    // 5. 渲染 HTML
    container.innerHTML = '<div class="list-group list-group-flush border rounded">' + paginatedItems.map(i => {
        // 特殊處理：我的回覆 (顯示討論標題與縮略內容)
        if (profileTab === 'replies') {
            return `
                <div class="list-group-item list-group-item-action cursor-pointer" onclick="jumpToThread('${i.threadId}')">
                    <div class="d-flex w-100 justify-content-between">
                        <h6 class="mb-1 text-primary text-truncate">回覆：${esc(i.threadTitle)}</h6>
                    </div>
                    <p class="mb-0 small text-secondary text-truncate">${esc(i.content.replace(/<[^>]*>?/gm, ''))}</p>
                </div>`;
        }

        // 一般處理：檔案或討論主題
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

/* ==========================================
   全域分頁回呼註冊
   ========================================== */
window.changeProfilePage = function(p) {
    curPage['profile_' + profileTab] = p;
    renderProfileList();
};

// 建議明確掛載到 window 物件上，確保 HTML 的 onclick 絕對抓得到它
window.changeProfilePage = function(p) {
    curPage['profile_' + profileTab] = p; 
    renderProfileList(); 
};

function submitProfileEdit() {
    // 1. 取得基本欄位值
    const nickVal = document.getElementById('edit-nick').value.trim();
    const passVal = document.getElementById('edit-pass').value.trim();
    const notifyVal = document.getElementById('edit-notify').checked;

    // 2. 判斷 Email 是否處於編輯模式
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
        if (emailVal !== user.email) {
            if (emailVal !== "" && !vCodeVal) {
                return showMsg('提示', '請先完成電子郵件驗證碼驗證', 'warning');
            }
            data.email = emailVal;
            data.vCode = vCodeVal;
        }
    }

    // 5. 執行後端更新 (這裡從 google.script.run 改為 callGAS)
    toggleLoading(true);

    callGAS('updateProfileSettings', [user.username, data])
        .then(res => {
            toggleLoading(false);
            if (res.status === 'success') {
                showMsg('更新成功', '您的資料與設定已成功儲存。', 'success', () => {
                    // 更新本地快取，避免一定要 reload 才能看到變化
                    user.nickname = nickVal;
                    user.notify = notifyVal;
                    if (data.email !== undefined) user.email = data.email;
                    localStorage.setItem('tb_user', JSON.stringify(user));

                    // 關閉 Modal (如果是用 Bootstrap)
                    const modalEl = document.getElementById('profileEditModal');
                    const bsModal = bootstrap.Modal.getInstance(modalEl);
                    if (bsModal) bsModal.hide();

                    // 如果你的頁面其他地方需要即時反應，可以 reload 或呼叫 loadProfile()
                    location.reload(); 
                });
            } else {
                // 如果錯誤訊息包含「驗證碼」，可以額外提示
                if (res.message && res.message.includes("驗證碼")) {
                    triggerInputError('edit-vcode'); // 觸發抖動動畫
                }
                showMsg('更新失敗', res.message || '請檢查您的輸入內容或驗證碼', 'error');
            }
        })
        .catch(err => {
            toggleLoading(false);
            showMsg('系統錯誤', '連線失敗：' + err.toString(), 'error');
        });
}

/* ==========================================
   個人資料編輯視窗初始化 (openEditProfile)
   ========================================== */
window.openEditProfile = function() {
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

  if (vCodeInput) vCodeInput.classList.add('hidden'); 

  // 2. 判斷 Email 顯示狀態
  if (user.email && user.email.trim() !== "") {
    verifiedGroup.classList.remove('hidden');
    editGroup.classList.add('hidden');
    displayInput.value = user.email;
    if (cancelBtn) cancelBtn.classList.remove('hidden');
  } else {
    verifiedGroup.classList.add('hidden');
    editGroup.classList.remove('hidden');
    document.getElementById('edit-email').value = '';
    if (cancelBtn) cancelBtn.classList.add('hidden');
  }

  // 3. 顯示彈窗
  const modalEl = document.getElementById('profileEditModal');
  bootstrap.Modal.getOrCreateInstance(modalEl).show();
};

// 切換至 Email 編輯模式
window.enableEmailEdit = function() {
  document.getElementById('email-verified-group').classList.add('hidden');
  document.getElementById('email-edit-group').classList.remove('hidden');
  document.getElementById('btn-cancel-email-edit').classList.remove('hidden');
  document.getElementById('edit-email').focus();
};

// 取消 Email 編輯模式
window.cancelEmailEdit = function() {
  if (user.email) {
    document.getElementById('email-verified-group').classList.remove('hidden');
    document.getElementById('email-edit-group').classList.add('hidden');
    document.getElementById('edit-email').value = "";
    const vCodeInput = document.getElementById('edit-vcode');
    if (vCodeInput) {
      vCodeInput.value = "";
      vCodeInput.classList.add('hidden');
    }
  }
};

/* ==========================================
   解除 Email 綁定 (deleteEmail)
   ========================================== */
window.deleteEmail = function() {
  showConfirm('刪除信箱', '確定要解除電子郵件綁定嗎？刪除後將無法接收系統通知。', () => {
    const data = { 
      nickname: document.getElementById('edit-nick').value, 
      email: "", 
      notify: false 
    };

    toggleLoading(true);
    // 取代 google.script.run
    callGAS('updateProfileSettings', [user.username, data])
      .then(res => {
        toggleLoading(false);
        if(res.status === 'success') {
          // 同步更新本地數據
          user.email = "";
          user.notify = false;
          localStorage.setItem('tb_user', JSON.stringify(user));

          if(document.getElementById('profile-email-label')) {
            document.getElementById('profile-email-label').textContent = "";
          }
          document.getElementById('edit-notify').checked = false;
          document.getElementById('email-verified-group').classList.add('hidden');
          document.getElementById('email-edit-group').classList.remove('hidden');
          document.getElementById('edit-email').value = '';
          const vCodeInput = document.getElementById('edit-vcode');
          if (vCodeInput) vCodeInput.classList.add('hidden');

          showMsg('成功', '電子郵件已解除綁定', 'success');
        } else {
          showMsg('錯誤', res.message, 'error');
        }
      })
      .catch(onFailure);
  });
};

/* ==========================================
   教案收藏功能 (toggleSave) - 樂觀更新
   ========================================== */
window.toggleSave = function(pid, btn) {
  if (!user) return forceAuth();

  const countSpan = btn.querySelector('.save-count');
  const labelSpan = btn.querySelector('.save-label');
  const icon = btn.querySelector('i');

  // --- 1. 樂觀更新：不等後端，直接變更 UI ---
  const isCurrentlySaved = btn.classList.contains('btn-warning');
  let currentCount = parseInt(countSpan.textContent) || 0;

  if (isCurrentlySaved) {
    btn.className = 'btn btn-sm btn-outline-warning ms-2 d-inline-flex align-items-center';
    icon.className = 'bi bi-star me-1';
    labelSpan.textContent = '收藏';
    countSpan.textContent = currentCount > 1 ? currentCount - 1 : '';
  } else {
    btn.className = 'btn btn-sm btn-warning ms-2 d-inline-flex align-items-center';
    icon.className = 'bi bi-star-fill text-white me-1';
    labelSpan.textContent = '已收藏';
    countSpan.textContent = currentCount + 1;
  }

  btn.style.pointerEvents = 'none'; // 鎖定按鈕

  // --- 2. 發送請求 ---
  callGAS('toggleSavePost', [user.username, pid])
    .then(res => {
      btn.style.pointerEvents = 'auto'; // 解鎖
      if(res.status === 'success') {
        // 同步正確數字
        countSpan.textContent = res.saveCount > 0 ? res.saveCount : '';
        user.saved = res.saved;
        localStorage.setItem('tb_user', JSON.stringify(user));

        const p = appData.plans.find(x => String(x.id) === String(pid));
        if (p) p.saveCount = res.saveCount;
      } else {
        // 失敗則恢復列表狀態
        loadPlans();
        showMsg('錯誤', res.message, 'error');
      }
    })
    .catch(err => {
      btn.style.pointerEvents = 'auto';
      loadPlans(); // 發生網路錯誤也刷回正確狀態
      onFailure(err);
    });
};
  
window.backToForumList = function() {
    // 1. UI 切換：顯示列表，隱藏詳情
    const listView = document.getElementById('forum-list-view');
    const detailView = document.getElementById('forum-detail-view');
    
    if (listView) listView.classList.remove('hidden');
    if (detailView) detailView.classList.add('hidden');

    // 2. 狀態重置
    currentThreadId = null;

    // 3. ★ 核心修正：同步網址列 (GitHub 專用)
    // 取代原本單純的 cleanUrl()，明確將 view 設回 forum 並移除 id
    const url = new URL(window.location);
    url.searchParams.set('view', 'forum');
    url.searchParams.delete('id');
    
    // 使用 pushState 或 replaceState 視你希望保留多少瀏覽紀錄而定
    // 這裡推薦用 pushState，這樣使用者按瀏覽器「上一頁」還能回到剛才的文章
    window.history.pushState({ view: 'forum' }, '', url);

    // 4. 滾動回頂端 (增加使用者體驗)
    window.scrollTo({ top: 0, behavior: 'smooth' });
};

window.updateAuthUI = function() {
    // 1. 再次從 localStorage 確認最新的 user 狀態
    user = JSON.parse(localStorage.getItem('tb_user'));

    // 2. 取得所有相關的 DOM 元件
    const btnLoginNav = document.getElementById('btn-login-nav');
    const btnLogout = document.getElementById('btn-logout');
    const navDisplay = document.getElementById('nav-user-display');
    const heroBtns = document.getElementById('hero-btns');

    // 3. 執行 UI 切換 (加上安全檢查，避免元件不存在時報錯)
    if (user) {
        // 已登入狀態
        if (btnLoginNav) btnLoginNav.classList.add('hidden');
        if (btnLogout) btnLogout.classList.remove('hidden');
        if (navDisplay) navDisplay.textContent = user.nickname || '會員';
        if (heroBtns) heroBtns.classList.add('hidden');
    } else {
        // 未登入狀態
        if (btnLoginNav) btnLoginNav.classList.remove('hidden');
        if (btnLogout) btnLogout.classList.add('hidden');
        if (navDisplay) navDisplay.textContent = '';
        if (heroBtns) heroBtns.classList.remove('hidden');
    }
};

window.handleLogin = function() {
    const form = document.querySelector('#form-login');
    if (!form) return;

    const loginData = { 
        username: form.username.value, 
        password: form.password.value 
    };

    toggleLoading(true);

    // 1. 取代原本的 google.script.run
    callGAS('userLogin', [loginData])
        .then(res => {
            toggleLoading(false);
            if (res.status === 'success') {
                // 2. 儲存使用者資訊到本地
                user = res.user;
                localStorage.setItem('tb_user', JSON.stringify(user));

                // 3. 更新導覽列 UI
                updateAuthUI();

                // 4. 關閉登入彈窗
                const authModalEl = document.getElementById('authModal');
                const bsModal = bootstrap.Modal.getInstance(authModalEl);
                if (bsModal) bsModal.hide();

                // 5. 顯示歡迎訊息並處理跳轉
                showMsg('歡迎', `登入成功，${user.nickname}！`, 'success', () => {
                    // ★ 核心優化：如果原本有想去的頁面 (pendingTarget)，登入後就直接帶他去
                    if (pendingTarget) {
                        const target = pendingTarget;
                        pendingTarget = null; // 清除暫存
                        
                        if (target.view === 'thread_detail') {
                            loadThreadDetail(target.id);
                        } else if (target.view === 'plan_detail') {
                            jumpToPlan(target.id);
                        } else if (target.view) {
                            switchView(target.view);
                        }
                    } else {
                        // 否則預設回首頁
                        switchView('home');
                    }
                });
            } else {
                // 登入失敗 (例如密碼錯誤)
                showMsg('失敗', res.message, 'error');
            }
        })
        .catch(err => {
            toggleLoading(false);
            onFailure(err);
        });
};

window.handleRegister = function() {
    const data = {
        username: document.getElementById('reg-user').value.trim(),
        password: document.getElementById('reg-pass').value.trim(),
        nickname: document.getElementById('reg-nick').value.trim(),
        email: document.getElementById('reg-email').value.trim(),
        vCode: document.getElementById('reg-vcode').value.trim()
    };

    // 1. 基本欄位前端檢查
    if (!data.username || !data.password || !data.nickname) {
        return showMsg('資料不完整', '請填寫所有必填欄位', 'warning');
    }

    if (data.email && !data.vCode) {
        return showMsg('提示', '請完成信箱驗證', 'warning');
    }

    toggleLoading(true);

    // 2. 取代原本的 google.script.run.userRegister
    callGAS('userRegister', [data])
        .then(res => {
            toggleLoading(false);
            if (res.status === 'success') {
                // 註冊成功，關閉彈窗
                const authModalEl = document.getElementById('authModal');
                const bsModal = bootstrap.Modal.getInstance(authModalEl);
                if (bsModal) bsModal.hide();

                showMsg('註冊成功', '帳號已建立，請使用新帳號登入', 'success', () => {
                    // 自動切換到登入分頁
                    showAuthModal('login');
                });
            } else {
                // 註冊失敗處理
                if (res.message && res.message.includes("驗證碼")) {
                    // 呼叫你定義過的抖動動畫輔助函式
                    triggerInputError('reg-vcode');
                }
                showMsg('註冊失敗', res.message || '請檢查輸入內容', 'error');
            }
        })
        .catch(err => {
            toggleLoading(false);
            onFailure(err);
        });
};

/* ==========================================
   發送驗證碼 (sendVCode)
   ========================================== */
window.sendVCode = function(inputId) {
    const emailInput = document.getElementById(inputId);
    const email = emailInput ? emailInput.value.trim() : "";

    // 1. 前端基本格式檢查
    if (!email || !email.includes('@')) {
        return showMsg('格式錯誤', '請輸入正確的電子郵件', 'warning');
    }

    toggleLoading(true);

    // 2. 呼叫 GAS API 取代 google.script.run.sendVerificationCode
    callGAS('sendVerificationCode', [email])
        .then(res => {
            toggleLoading(false);
            
            if (res.status === 'success') {
                showMsg('發送成功', '驗證碼已寄至您的信箱，請於 10 分鐘內輸入。', 'success');

                // 3. 自動判斷是哪一個頁面的驗證碼輸入框
                // 註冊頁面 ID 為 reg-email，修改資料頁面 ID 為 edit-email
                const vId = (inputId === 'reg-email') ? 'reg-vcode' : 'edit-vcode';
                const vInput = document.getElementById(vId);

                if (vInput) {
                    // 清空前一次留下來的殘餘內容
                    vInput.value = ""; 
                    // 移除隱藏標籤並自動聚焦
                    vInput.classList.remove('hidden');
                    vInput.focus(); 
                }
            } else {
                // 後端回傳失敗 (例如：Email 已被註冊、發送頻率過快等)
                showMsg('發送失敗', res.message || '請稍後再試', 'error');
            }
        })
        .catch(err => {
            toggleLoading(false);
            onFailure(err);
        });
};

/* ==========================================
   小工具：觸發輸入框抖動並清空 (triggerInputError)
   ========================================== */
window.triggerInputError = function(elementId) {
    const el = document.getElementById(elementId);
    if (!el) return;

    // 1. 清空輸入內容
    el.value = ""; 
    
    // 2. 加入 CSS 動畫類別 (需配合你 style.css 裡的 .shake-error)
    el.classList.add('shake-error'); 
    
    // 3. 自動聚焦，讓使用者可以直接重新輸入
    el.focus();

    // 4. 0.5 秒後移除類別
    // 這是為了讓下次錯誤發生時，重新加入類別能再次觸發 CSS 動畫
    setTimeout(() => {
        el.classList.remove('shake-error');
    }, 500);
};
 
/* ==========================================
   提交個人資料設定 (submitProfileEdit)
   ========================================== */
window.submitProfileEdit = function() {
    // 1. 取得 UI 狀態與欄位值
    const isEmailEditing = !document.getElementById('email-edit-group').classList.contains('hidden');
    const emailInput = document.getElementById('edit-email');
    const vCodeInput = document.getElementById('edit-vcode');
    
    const emailVal = emailInput ? emailInput.value.trim() : "";
    const vCodeVal = vCodeInput ? vCodeInput.value.trim() : "";
    const nickVal = document.getElementById('edit-nick').value.trim();
    const passVal = document.getElementById('edit-pass').value; // 密碼通常不 trim，保留使用者原始輸入
    const notifyVal = document.getElementById('edit-notify').checked;

    // 2. 準備傳送給後端的資料物件
    const data = { 
        nickname: nickVal, 
        password: passVal, // 留空則 GAS 後端邏輯會跳過不更新密碼
        notify: notifyVal 
    };

    // 3. 關鍵修正：判斷是否真的有改信箱
    if (isEmailEditing) {
        // 如果使用者按了更改，但沒填東西，或者填的跟原本一模一樣 -> 視為不改
        if (emailVal === "" || emailVal === user.email) {
            // 不將 email 放入 data 物件，後端就不會動到這一欄
        } else {
            // 使用者填了新信箱且不為空，必須檢查驗證碼
            if (!vCodeVal) {
                return showMsg('提示', '請先完成電子郵件驗證碼驗證', 'warning');
            }
            data.email = emailVal;
            data.vCode = vCodeVal;
        }
    }

    toggleLoading(true);

    // 4. 呼叫 GAS API 取代 google.script.run
    callGAS('updateProfileSettings', [user.username, data])
        .then(res => {
            toggleLoading(false);
            
            if (res.status === 'success') {
                // 5. 更新本地緩存 user 物件
                user.nickname = nickVal;
                user.notify = notifyVal;
                
                // 只有在真的成功修改 email 時才更新本地資料
                if (data.email !== undefined) {
                    user.email = data.email;
                }
                
                // 寫回 localStorage 確保重新整理後資料還在
                localStorage.setItem('tb_user', JSON.stringify(user));

                // 6. 即時更新畫面 UI
                updateAuthUI(); 
                if(document.getElementById('profile-nickname')) {
                    document.getElementById('profile-nickname').textContent = nickVal;
                }
                if(document.getElementById('profile-email-label')) {
                    document.getElementById('profile-email-label').textContent = user.email || '';
                }

                // 7. 關閉彈窗並顯示成功訊息
                const modalEl = document.getElementById('profileEditModal');
                const bsModal = bootstrap.Modal.getInstance(modalEl);
                if (bsModal) bsModal.hide();
                
                showMsg('成功', '您的資料與設定已成功更新。', 'success');
            } else {
                // 8. 錯誤處理：如果是驗證碼錯誤，觸發輸入框抖動
                if (res.message && res.message.includes("驗證碼")) {
                    triggerInputError('edit-vcode');
                }
                showMsg('更新失敗', res.message || '請檢查您的輸入內容', 'error');
            }
        })
        .catch(err => {
            toggleLoading(false);
            onFailure(err);
        });
};
  
/* ==========================================
   登出邏輯 (logout)
   ========================================== */
window.logout = function() {
    // 1. 清除本地儲存的會員資料
    localStorage.removeItem('tb_user');
    
    // 2. 清除全域變數
    user = null;
    
    // 3. 更新導覽列 UI 顯示 (變回登入按鈕)
    updateAuthUI();
    
    // 4. 強制導回首頁
    switchView('home');
    
    // 5. 提示使用者已登出
    showMsg('提示', '您已成功登出。', 'info');
};

/* ==========================================
   身分驗證彈窗頁籤切換 (toggleAuthTab)
   ========================================== */
window.toggleAuthTab = function(tab) {
    const errorMsg = document.getElementById('auth-error-msg');
    if (errorMsg) errorMsg.classList.add('hidden');

    // 切換頁籤 (Login / Register) 的 Active 樣式
    document.getElementById('tab-login').classList.toggle('active', tab === 'login');
    document.getElementById('tab-register').classList.toggle('active', tab === 'register');

    // 切換對應的表單顯示
    document.getElementById('form-login').classList.toggle('hidden', tab !== 'login');
    document.getElementById('form-register').classList.toggle('hidden', tab !== 'register');
};

/* ==========================================
   顯示登入/註冊彈窗 (showAuthModal)
   ========================================== */
window.showAuthModal = function(tab) {
    // 先設定好要顯示哪一個頁籤
    toggleAuthTab(tab);
    
    // 取得或建立 Bootstrap Modal 實例並顯示
    const authModalEl = document.getElementById('authModal');
    const bsModal = bootstrap.Modal.getOrCreateInstance(authModalEl);
    bsModal.show();
};

/* ==========================================
   基礎小工具：HTML 逸出 (esc)
   ========================================== */
window.esc = function(s) {
    if (!s) return '';
    // 將特殊字元轉為 HTML 實體，防止 XSS 攻擊
    return String(s)
        .replace(/&/g, '&amp;') // 建議優先處理 &
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#039;');
};

/* ==========================================
   基礎小工具：日期格式化 (formatDate)
   ========================================== */
window.formatDate = function(ts) {
    if (!ts) return '';
    const d = new Date(ts);
    // 檢查日期是否有效
    if (isNaN(d.getTime())) return '';
    
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    
    return `${y}/${m}/${day}`;
};

/* ==========================================
   回覆功能 (replyTo) - 支援樓中樓標籤生成
   ========================================== */
window.replyTo = function(floor, subFloor) {
    // 1. 權限檢查
    if (!user) return forceAuth();
    
    // 2. 設定目前的目標主樓層
    replyToFloor = parseInt(floor); 

    const editor = tinymce.get('comment-input');
    const indicator = document.getElementById('reply-indicator');
    
    if (!editor || !indicator) return;

    // 3. 生成標籤文字 (例如 B1 或 B1-1)
    var tagText = "B" + floor;
    if (subFloor && subFloor > 0) {
        tagText += "-" + subFloor;
    }
    
    // 4. 生成 HTML 標籤 (移除背景色，保持簡潔的連結感)
    // 注意：data-floor 和 data-sub 是給 bindRefPreviews() 抓取預覽內容用的
    var tagHtml = '<span class="ref-link" contenteditable="false" ' +
                  'data-floor="' + floor + '" ' +
                  'data-sub="' + (subFloor || 0) + '" ' +
                  'style="color:#0d6efd; font-weight:bold; cursor:pointer; text-decoration:none; margin-right:2px;">' + 
                  tagText + '</span>&nbsp;';
    
    // 5. 更新 UI 指示器
    indicator.textContent = "正在回覆 " + tagText + " ...";
    indicator.classList.remove('hidden');
    
    // 6. 插入內容至 TinyMCE
    // 優化：使用 insertContent 而不是 setContent，這樣不會覆蓋掉使用者已經打好的字
    const currentContent = editor.getContent().trim();
    if (currentContent === '' || currentContent === '<p>&nbsp;</p>') {
        editor.setContent('<p>' + tagHtml + '</p>');
    } else {
        // 如果原本已經有內容，就在新的一行插入標籤
        editor.insertContent('<p>' + tagHtml + '</p>');
    }
    
    // 7. 聚焦並將游標移至最後
    editor.focus();
    editor.selection.select(editor.getBody(), true);
    editor.selection.collapse(false);
    
    // 8. 平滑滾動至留言區
    const commentSection = document.getElementById('comment-section');
    if (commentSection) {
        commentSection.scrollIntoView({ behavior: 'smooth' });
    }
};

/* ==========================================
   留言按讚功能 (toggleLike) - 樂觀更新版
   ========================================== */
window.toggleLike = function(commentId, btn) {
    // 1. 權限檢查
    if (!user) return forceAuth();

    const isLiked = btn.classList.contains('liked');
    const icon = btn.querySelector('i');
    const countSpan = btn.querySelector('.like-count');
    
    // 取得目前的按讚數 (如果是空的就當作 0)
    let currentCount = parseInt(countSpan.textContent) || 0;

    // --- 2. 樂觀更新：立即改變 UI 外觀 ---
    btn.classList.toggle('liked');
    
    // 切換 Bootstrap Icon 樣式
    icon.className = isLiked ? 'bi bi-heart' : 'bi bi-heart-fill';
    
    // 計算新數字
    let newCount = isLiked ? Math.max(0, currentCount - 1) : currentCount + 1;
    countSpan.textContent = newCount > 0 ? newCount : '';

    // 鎖定按鈕防止在請求完成前連續點擊
    btn.style.pointerEvents = 'none';

    // --- 3. 發送至後端 (取代 google.script.run) ---
    callGAS('toggleCommentLike', [commentId, user.username])
        .then(res => {
            btn.style.pointerEvents = 'auto'; // 解鎖按鈕
            
            if (res.status !== 'success') {
                // 失敗時彈回原狀：將 UI 切換回點擊前的樣子
                btn.classList.toggle('liked');
                icon.className = !isLiked ? 'bi bi-heart' : 'bi bi-heart-fill';
                countSpan.textContent = currentCount > 0 ? currentCount : '';
                
                showMsg('錯誤', res.message || '按讚失敗', 'error');
            } else {
                // 成功時可以選擇同步後端回傳的精確數字 (防止多人同時點擊的誤差)
                if (res.likeCount !== undefined) {
                    countSpan.textContent = res.likeCount > 0 ? res.likeCount : '';
                }
            }
        })
        .catch(err => {
            btn.style.pointerEvents = 'auto';
            // 網路連線錯誤時也彈回原狀
            btn.classList.toggle('liked');
            icon.className = !isLiked ? 'bi bi-heart' : 'bi bi-heart-fill';
            countSpan.textContent = currentCount > 0 ? currentCount : '';
            
            onFailure(err);
        });
};

/* ==========================================
   預覽小視窗邏輯 (bindRefPreviews)
   實現點擊 B1 顯示小視窗，支援視窗內再次點擊
   ========================================== */
window.bindRefPreviews = function() {
    const box = document.getElementById('ref-preview-box');
    const forumDetail = document.getElementById('forum-detail-view');
    
    // 如果畫面上沒有預覽盒子或不在討論串頁面，就跳出
    if (!box || !forumDetail) return;

    // --- 建立共用渲染函式 ---
    const showPreview = (target, e) => {
        const floor = target.getAttribute('data-floor');
        const sub = target.getAttribute('data-sub') || 0;
        const key = floor + '-' + sub;
        const targetComment = commentLookup[key];

        if (targetComment) {
            const floorTitle = "B" + floor + (parseInt(sub) > 0 ? "-" + sub : "");
            
            // 渲染視窗內容
            box.innerHTML = `
                <div style="border-bottom: 2px solid #198754; margin-bottom: 8px; padding-bottom: 4px; display: flex; justify-content: space-between; align-items: center;">
                    <strong style="color: #198754;">${floorTitle} ${esc(targetComment.authorName)}</strong>
                    <span style="cursor:pointer; font-size:1.2rem; line-height:1;" onclick="document.getElementById('ref-preview-box').style.display='none'">&times;</span>
                </div>
                <div class="preview-inner-content" style="max-height: 300px; overflow-y: auto;">
                    ${targetComment.content}
                </div>
            `;

            // 定位邏輯：只有從外部點擊 (e 存在) 時才重新計算位置
            if (e) {
                const rect = target.getBoundingClientRect();
                box.style.display = 'block';
                // 計算位置，考慮捲軸偏移量
                box.style.left = (rect.left + window.scrollX) + 'px';
                box.style.top = (rect.bottom + window.scrollY + 8) + 'px';
            }
            
            // 如果內容有圖片，調整寬度以防破版
            if (box.querySelector('img')) {
                box.style.width = '350px';
            } else {
                box.style.width = '280px'; // 預設寬度
            }
        }
    };

    // --- 監聽 1：討論區內的標籤點擊 ---
    // 先移除舊的監聽器（如果有的話），避免重複綁定
    forumDetail.onclick = null; 
    forumDetail.addEventListener('click', function(e) {
        const target = e.target.closest('.ref-link');
        if (target) {
            e.preventDefault();
            e.stopPropagation();
            showPreview(target, e);
        } else {
            // 點擊討論區其他地方，若不是點在小視窗內，就關閉小視窗
            if (!box.contains(e.target)) {
                box.style.display = 'none';
            }
        }
    });

    // --- 監聽 2：小視窗內部的點擊 (實現視窗內跳轉預覽) ---
    box.onclick = null;
    box.addEventListener('click', function(e) {
        const target = e.target.closest('.ref-link');
        if (target) {
            e.preventDefault();
            e.stopPropagation();
            // 在視窗內點擊標籤，不傳入 e，讓視窗維持原位僅更新內容
            showPreview(target, null); 
        }
        // 阻止事件冒泡，防止觸發 document 的關閉邏輯
        e.stopPropagation();
    });

    // --- 監聽 3：點擊全域任何地方關閉視窗 ---
    document.addEventListener('click', function(e) {
        if (box.style.display === 'block' && !box.contains(e.target)) {
            const isRefLink = e.target.closest('.ref-link');
            if (!isRefLink) box.style.display = 'none';
        }
    });
}; 

/* ==========================================
   討論串主文按讚 (toggleThreadLike)
   ========================================== */
window.toggleThreadLike = function(tid, btn) {
    // 1. 權限檢查
    if (!user) return forceAuth();

    const icon = btn.querySelector('i');
    const countSpan = btn.querySelector('.t-like-count');
    
    // 取得目前的按讚數
    let currentCount = parseInt(countSpan.textContent) || 0;
    const isLiking = icon.classList.contains('bi-heart');

    // --- 2. 樂觀更新：立即更新詳情頁 UI ---
    if (isLiking) {
        icon.className = 'bi bi-heart-fill';
        btn.classList.replace('btn-outline-danger', 'btn-danger');
        countSpan.textContent = currentCount + 1;
    } else {
        icon.className = 'bi bi-heart';
        btn.classList.replace('btn-danger', 'btn-outline-danger');
        countSpan.textContent = Math.max(0, currentCount - 1);
    }

    // 鎖定按鈕防止連點
    btn.style.pointerEvents = 'none';

    // --- 3. 發送至後端 (取代 google.script.run) ---
    callGAS('toggleThreadLike', [tid, user.username])
        .then(res => {
            btn.style.pointerEvents = 'auto'; // 解鎖按鈕
            
            if (res.status === 'success') {
                const sid = String(tid);
                
                // ★ 數據同步 1：更新本地討論列表快取 ★
                const threadInList = appData.threads.find(t => String(t.id) === sid);
                if (threadInList) {
                    threadInList.likes = res.likes;
                    threadInList.likeCount = res.likes.length;
                    
                    // 偷偷在背後重新渲染列表，確保返回時數據是新的
                    renderForumList(); 
                }

                // ★ 數據同步 2：更新目前詳情頁的全域快取 ★
                if (appData.currentThread && String(appData.currentThread.id) === sid) {
                    appData.currentThread.likes = res.likes;
                    appData.currentThread.likeCount = res.likes.length;
                }
                
                // 同步伺服器回傳的精確數字
                countSpan.textContent = res.likes.length;

            } else {
                // 失敗處理：跳出訊息並重載該篇詳情以恢復正確 UI
                showMsg('錯誤', res.message || '操作失敗', 'error');
                loadThreadDetail(tid); 
            }
        })
        .catch(err => {
            btn.style.pointerEvents = 'auto';
            // 網路連線錯誤時也重載恢復
            loadThreadDetail(tid);
            onFailure(err);
        });
};

/* ==========================================
   等級系統配置 (getLevelInfo)
   根據總積分回傳等級名稱、顏色及下一級門檻
   ========================================== */
window.getLevelInfo = function(points) {
    // 確保點數為數字型別，避免比對錯誤
    const p = parseInt(points) || 0;

    // 根據等級調整為由淺入深的綠色系
    // LV.5: 8001+
    if (p >= 8001) {
        return { 
            lv: 5, 
            name: '淨零造物主', 
            color: '#004b23', // 深沉森林綠
            next: 'MAX' 
        };
    }
    
    // LV.4: 3001 - 8000
    if (p >= 3001) {
        return { 
            lv: 4, 
            name: '環境守護神', 
            color: '#007f5f', // 科技翡翠綠
            next: 8001 
        };
    }
    
    // LV.3: 1001 - 3000
    if (p >= 1001) {
        return { 
            lv: 3, 
            name: '綠能導師', 
            color: '#2b9348', // 標準活力綠
            next: 3001 
        };
    }
    
    // LV.2: 201 - 1000
    if (p >= 201) {
        return { 
            lv: 2, 
            name: '減碳專員', 
            color: '#80b918', // 嫩芽草地綠
            next: 1001 
        };
    }
    
    // LV.1: 0 - 200
    return { 
        lv: 1, 
        name: '碳足跡者', 
        color: '#aacc00', // 檸檬淺綠色
        next: 201 
    };
};

/* ==========================================
   顯示積分明細彈窗 (showScoreDetails)
   ========================================== */
window.showScoreDetails = function() {
    if (!user) return;

    // 1. 初始化並顯示 Modal
    const modalEl = document.getElementById('scoreDetailModal');
    const modal = bootstrap.Modal.getOrCreateInstance(modalEl);
    modal.show();

    // 顯示載入中
    const contentArea = document.getElementById('score-detail-content');
    if (contentArea) contentArea.innerHTML = '<div class="text-center py-4"><div class="spinner-border text-success"></div></div>';

    // 2. 呼叫 GAS API 取代 google.script.run
    callGAS('getPointDetails', [user.username])
        .then(res => {
            if (res.error) {
                if (contentArea) contentArea.innerHTML = '<div class="alert alert-danger">讀取失敗</div>';
                return;
            }

            // 3. 定義計算項目與對應單價
            const rows = [
                { label: '發布簡報 (公開)', count: res.pptActive || 0, price: 50 },
                { label: '發布簡報 (典藏)', count: res.pptArchived || 0, price: 40 },
                { label: '發布教案 (公開)', count: res.planActive || 0, price: 40 },
                { label: '發布教案 (典藏)', count: res.planArchived || 0, price: 32 },
                { label: '討論區參與', count: res.threadCount || 0, price: 30 }
            ];

            // 4. 計算基礎總分與獎勵分 (互動獎勵 = 總分 - 基礎分總和)
            const baseTotal = rows.reduce((sum, r) => sum + (r.count * r.price), 0);
            const rewardScore = (res.totalScore || 0) - baseTotal;

            // 5. 渲染算式 HTML (使用等寬字體對齊)
            let html = `
                <style>
                    .score-grid {
                        display: grid;
                        grid-template-columns: 1fr auto auto auto auto auto;
                        gap: 4px 10px;
                        align-items: center;
                        font-family: "Courier New", Courier, monospace;
                        font-size: 0.95rem;
                    }
                    .score-grid div { padding: 4px 0; border-bottom: 1px dashed #eee; }
                    .text-right { text-align: right; }
                    .text-center { text-align: center; }
                </style>
                <div class="score-grid">
            `;

            rows.forEach(r => {
                html += `
                    <div class="text-secondary" style="border-bottom:none;">${r.label}</div>
                    <div class="text-right fw-bold">${r.count}</div>
                    <div class="text-center text-muted">×</div>
                    <div class="text-right">${r.price}</div>
                    <div class="text-center text-muted">=</div>
                    <div class="text-right fw-bold">${r.count * r.price}</div>
                `;
            });

            // 加入互動獎勵 (愛心、收藏等隱藏分數)
            html += `
                    <div class="text-secondary">互動獎勵 (愛心/收藏)</div>
                    <div></div><div></div><div></div>
                    <div class="text-center text-muted">+</div>
                    <div class="text-right fw-bold">${rewardScore}</div>
                </div>

                <div class="d-flex justify-content-between mt-3 pt-3 border-top fs-5 text-success">
                    <strong>目前總碳值</strong>
                    <strong style="font-family: 'Segoe UI', Arial, sans-serif;">${res.totalScore || 0} CP</strong>
                </div>
            `;

            if (contentArea) contentArea.innerHTML = html;
        })
        .catch(err => {
            if (contentArea) contentArea.innerHTML = '<div class="alert alert-danger">連線錯誤，請稍後再試</div>';
            onFailure(err);
        });
};