/* ---------------- API layer (talks to Google Apps Script Web App) ---------------- */
async function apiGet(action, params){
  const qs = new URLSearchParams({ action, ...(params||{}) }).toString();
  const res = await fetch(`${APPS_SCRIPT_URL}?${qs}`);
  return res.json();
}
async function apiPost(action, payload){
  const res = await fetch(`${APPS_SCRIPT_URL}?action=${action}`, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' }, // avoids CORS preflight
    body: JSON.stringify(payload || {})
  });
  return res.json();
}
function urlConfigured(){
  return APPS_SCRIPT_URL && !APPS_SCRIPT_URL.includes('PASTE_YOUR');
}

/* ---------------- app state ---------------- */
const CLASSES = ['大一','大二','大三','大四'];
const SECTIONS = ['甲','乙','丙','丁','戊','己'];

const ILLNESS_REASONS = ['發燒','昏眩','噁心嘔吐','頭痛','牙痛','胃痛','腹痛','腹瀉','經痛','氣喘','流鼻血','疹癢','眼疾','過敏','其它'];
const INJURY_REASONS = ['擦傷','裂割刺傷','夾壓傷','挫撞傷','扭傷','灼燙傷','叮咬傷','骨折','舊傷','肌肉拉傷','甲溝炎','起水泡','其它'];

const TREATMENT_OPTIONS = ['傷口護理','冰敷','熱敷','休息觀察','告知家長','通知導師','家長帶回','校方送醫','衛生教育','生理食鹽水沖洗','口罩給予','止血','擦藥','禁食','溫開水給予','補充糖水','補充水分','體溫監測','血氧監測','彈繃固定','氧氣給與','熱敷觀察','三角巾固定','頸圈固定','夾板固定','KED固定','長背板固定','保暖','CPR+AED','哈姆立克法','通知家長送醫','家長同意自行返家','家長同意自行就醫','家長未接聽到電話','電訪追蹤','救護車護送','教官或導師送醫','其它'];

// 匯出 Excel 用：身體不適 + 受傷的詳細原因合併去重（保留原順序，「其它」只出現一次）
const REASON_DETAIL_OPTIONS = Array.from(new Set([...ILLNESS_REASONS, ...INJURY_REASONS]));

const HEALTH_CHECK_ITEMS = [
  '血壓 mmHg', '血糖(AC)', 'T-CHOL mg/dl', 'T-G mg/dl', 'SGOT IU/L', 'SGPT IU/L',
  '尿酸 mg/dl', '高密度脂蛋白膽固醇_HDL', '胸部X光', '心電圖'
];
const EDU_GUIDANCE_OPTIONS = ['飲食控制', '運動管理', '體重管理', '定期追蹤', '生活作息'];

let state = {
  screen: 'home',
  loading: true,
  loadError: null,
  roster: [],
  records: [],
  loggedIn: false,
  sessionToken: null,
  toast: null,
  student: { grade:null, section:null, id:null, name:null, gender:null, reason:null, detail:null },
  loginErr: '',
  loginBusy: false,
  search: '',
  dashSearch: '',
  dashFrom: '',
  dashTo: '',
  pwdOld: '',
  pwdErr: '',
  pwdBusy: false,
  treatmentRecordId: null,
  treatmentSelected: [],
  treatmentOther: '',
  treatmentBusy: false,
  rosterFull: [],
  caseSearch: '',
  healthCheckStudentId: null,
  healthCheckStudent: null,
  healthCheckItems: {},
  healthCheckGuidance: [],
  healthCheckNote: '',
  healthCheckBusy: false,
  healthCheckLoading: false,
};

let refreshTimer = null;

function showToast(msg){
  state.toast = msg;
  render();
  setTimeout(()=>{ state.toast=null; render(); }, 2200);
}

/* ---------------- init ---------------- */
function todayStr(offsetDays){
  const d = new Date();
  d.setDate(d.getDate() + (offsetDays||0));
  return d.toISOString().slice(0,10);
}

async function init(){
  state.dashFrom = todayStr(-6);
  state.dashTo = todayStr(0);
  if(!urlConfigured()){
    state.loading = false;
    render();
    return;
  }
  await loadRoster();
  state.loading = false;
  render();
}

async function loadRoster(){
  try{
    const data = await apiGet('getRoster');
    if(data && data.ok){
      state.roster = data.roster || [];
      state.loadError = null;
    } else {
      state.loadError = '讀取名冊失敗，請確認 Apps Script 是否部署成功';
    }
  }catch(e){
    console.error(e);
    state.loadError = '無法連線到資料庫，請檢查網路連線或 Apps Script 網址設定';
  }
}

async function loadRecords(){
  try{
    const data = await apiGet('getRecords', { token: state.sessionToken });
    if(data && data.ok){
      state.records = data.records || [];
      state.loadError = null;
    } else if(data && data.error === 'unauthorized'){
      state.loggedIn = false;
      state.sessionToken = null;
      state.screen = 'nurse-login';
      state.loginErr = '登入已逾時，請重新登入';
    } else {
      state.loadError = '讀取報到紀錄失敗';
    }
  }catch(e){
    console.error(e);
    state.loadError = '無法連線到資料庫，請檢查網路連線';
  }
}

async function loadRosterFull(){
  try{
    const data = await apiGet('getRosterFull', { token: state.sessionToken });
    if(data && data.ok){
      state.rosterFull = data.roster || [];
      state.loadError = null;
    } else if(data && data.error === 'unauthorized'){
      state.loggedIn = false;
      state.sessionToken = null;
      state.screen = 'nurse-login';
      state.loginErr = '登入已逾時，請重新登入';
    } else {
      state.loadError = '讀取學生個案資料失敗';
    }
  }catch(e){
    console.error(e);
    state.loadError = '無法連線到資料庫，請檢查網路連線';
  }
}

async function loadHealthCheck(studentId){
  try{
    const data = await apiGet('getHealthCheck', { token: state.sessionToken, studentId: studentId });
    if(data && data.ok){
      state.healthCheckItems = data.items || {};
      state.healthCheckGuidance = data.guidance || [];
      state.healthCheckNote = data.note || '';
    } else if(data && data.error === 'unauthorized'){
      state.loggedIn = false;
      state.sessionToken = null;
      state.screen = 'nurse-login';
      state.loginErr = '登入已逾時，請重新登入';
    } else {
      showToast('讀取健檢資料失敗');
    }
  }catch(e){
    console.error(e);
    showToast('無法連線，請稍後再試');
  }
}

/* ---------------- icons ---------------- */
const ICONS = {
  cross: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"><path d="M12 6v12M6 12h12"/></svg>`,
  student: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M22 10 12 5 2 10l10 5 10-5Z"/><path d="M6 12v5c0 1.5 2.7 3 6 3s6-1.5 6-3v-5"/></svg>`,
  nurse: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M12 2v4M9 4h6"/><path d="M6 8h12v6a6 6 0 0 1-12 0V8Z"/><path d="M12 12v4M10 14h4"/></svg>`,
  bandage: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M12 2 2 12l10 10L22 12 12 2Z"/><path d="M9 9l6 6M9 15l6-6"/></svg>`,
  thermo: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M3 12h4l2-7 4 14 2-7h6"/></svg>`,
  check: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M20 6 9 17l-5-5"/></svg>`,
  upload: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M12 3v12M7 8l5-5 5 5"/><path d="M4 17v2a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-2"/></svg>`,
  logout: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4"/><path d="M16 17l5-5-5-5M21 12H9"/></svg>`,
  trash: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M3 6h18M8 6V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2m-9 0 1 14a2 2 0 0 0 2 2h4a2 2 0 0 0 2-2l1-14"/></svg>`,
  done: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M9 11l3 3L22 4"/><path d="M21 12v7a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11"/></svg>`,
  refresh: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M21 12a9 9 0 1 1-2.6-6.4M21 4v6h-6"/></svg>`,
  clipboard: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><rect x="6" y="4" width="12" height="17" rx="2"/><path d="M9 4V3a1 1 0 0 1 1-1h4a1 1 0 0 1 1 1v1M9 11h6M9 15h6"/></svg>`,
  download: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M12 3v12M7 10l5 5 5-5"/><path d="M4 19h16"/></svg>`,
};

function stepperSvg(step, total){
  total = total || 4;
  const w = 140 * total;
  const margin = 40;
  const gap = (w - margin*2) / (total-1);
  const cx = [];
  for(let i=0;i<total;i++) cx.push(margin + i*gap);

  let path = `M0 26 L20 26 L28 10 L36 42 L44 26 L${cx[0]-6} 26`;
  for(let i=0;i<total-1;i++){
    path += ` L${cx[i]+6} 26 L${cx[i]+ (cx[i+1]-cx[i])/2 - 8} 26 L${cx[i]+(cx[i+1]-cx[i])/2} 8 L${cx[i]+(cx[i+1]-cx[i])/2+8} 44 L${cx[i]+(cx[i+1]-cx[i])/2+16} 26 L${cx[i+1]-6} 26`;
  }
  path += ` L${w} 26`;
  let circles = cx.map((x,i)=>{
    const idx = i+1;
    const active = idx <= step;
    const fill = active ? 'var(--primary)' : 'var(--surface)';
    const stroke = active ? 'var(--primary)' : 'var(--border)';
    const textFill = active ? '#fff' : 'var(--muted)';
    return `<circle cx="${x}" cy="26" r="15" fill="${fill}" stroke="${stroke}" stroke-width="2"/>
      <text x="${x}" y="31" text-anchor="middle" font-size="13" font-weight="700" fill="${textFill}" font-family="Roboto Mono, monospace">${idx}</text>`;
  }).join('');
  return `<svg viewBox="0 0 ${w} 52" preserveAspectRatio="none">
    <path d="${path}" fill="none" stroke="var(--border)" stroke-width="2" stroke-linejoin="round" stroke-linecap="round"/>
    ${circles}
  </svg>`;
}

/* ---------------- render ---------------- */
function render(){
  const app = document.getElementById('app');

  if(!urlConfigured()){
    app.innerHTML = `
      <div class="app-header">
        <div class="brand">
          <div class="brand-mark">${ICONS.cross.replace('currentColor','#fff')}</div>
          <div>
            <div class="brand-text">健康中心報到系統</div>
            <div class="brand-sub">SCHOOL HEALTH CENTER</div>
          </div>
        </div>
      </div>
      <div class="card">
        <h1 class="title">尚未設定資料庫連線</h1>
        <p class="subtitle">請先完成 config.js 中的 APPS_SCRIPT_URL 設定，詳細步驟請參考專案的 README.md。</p>
      </div>`;
    return;
  }

  if(state.loading){
    app.innerHTML = `<div class="loading-wrap">讀取資料中…</div>`;
    return;
  }

  app.innerHTML = `
    <div class="app-header">
      <div class="brand">
        <div class="brand-mark">${ICONS.cross.replace('currentColor','#fff')}</div>
        <div>
          <div class="brand-text">健康中心報到系統</div>
          <div class="brand-sub">SCHOOL HEALTH CENTER</div>
        </div>
      </div>
      ${headerRight()}
    </div>
    ${state.loadError ? `<div class="setup-warning">${state.loadError}</div>` : ''}
    <div class="screen">${screenHtml()}</div>
    ${state.toast ? `<div class="toast">${state.toast}</div>` : ''}
  `;
  bindEvents();
}

function headerRight(){
  if(state.screen === 'home') return '';
  if(state.screen.startsWith('student')) return `<button class="pill-btn" data-act="go-home">回首頁</button>`;
  if(state.screen === 'nurse-login') return `<button class="pill-btn" data-act="go-home">回首頁</button>`;
  if(state.screen === 'nurse-change-password') return `<button class="pill-btn" data-act="back-dashboard">回儀表板</button>`;
  if(state.screen === 'nurse-record-treatment') return `<button class="pill-btn" data-act="back-dashboard">回儀表板</button>`;
  if(state.screen === 'nurse-case-management') return `<button class="pill-btn" data-act="back-dashboard">回儀表板</button>`;
  if(state.screen === 'nurse-health-check') return `<button class="pill-btn" data-act="back-case-management">回學生個案管理</button>`;
  if(state.screen === 'nurse-dashboard'){
    return `
      <div style="display:flex;gap:8px;flex-wrap:wrap;">
        <button class="pill-btn" data-act="go-case-management">學生個案管理</button>
        <button class="pill-btn" data-act="go-change-password">修改密碼</button>
        <button class="pill-btn" data-act="logout">${ICONS.logout} 登出</button>
      </div>`;
  }
  return '';
}

function screenHtml(){
  switch(state.screen){
    case 'home': return homeScreen();
    case 'student-1': return studentStep1();
    case 'student-2': return studentStep2();
    case 'student-3': return studentStep3();
    case 'student-4': return studentStep4();
    case 'student-5': return studentStep5();
    case 'student-done': return studentDone();
    case 'nurse-login': return nurseLogin();
    case 'nurse-change-password': return nurseChangePassword();
    case 'nurse-dashboard': return nurseDashboard();
    case 'nurse-record-treatment': return nurseRecordTreatment();
    case 'nurse-case-management': return nurseCaseManagement();
    case 'nurse-health-check': return nurseHealthCheck();
    default: return homeScreen();
  }
}

/* ---- HOME ---- */
function homeScreen(){
  return `
  <div class="home-hero">
    <h1 class="title">歡迎使用健康中心報到系統</h1>
    <p class="subtitle">請選擇您的身份以繼續</p>
  </div>
  <div class="role-grid">
    <div class="role-card student" data-act="start-student">
      <div class="icon-wrap">${ICONS.student}</div>
      <h3>我是學生</h3>
      <p>身體不適或受傷，前往報到</p>
    </div>
    <div class="role-card nurse" data-act="go-nurse-login">
      <div class="icon-wrap">${ICONS.nurse}</div>
      <h3>我是護理人員</h3>
      <p>登入查看與管理報到紀錄</p>
    </div>
  </div>`;
}

/* ---- STUDENT STEP 1: 班級 ---- */
function studentStep1(){
  const {grade, section} = state.student;
  return `
  ${stepperBlock(1)}
  <div class="card">
    <h1 class="title">選擇班級</h1>
    <p class="subtitle">請先選擇年級，再選擇班別</p>
    <div class="choice-grid" style="margin-bottom:18px;">
      ${CLASSES.map(g=>`<div class="choice-btn ${grade===g?'selected':''}" data-act="pick-grade" data-val="${g}">${g}</div>`).join('')}
    </div>
    ${grade ? `
      <p class="subtitle" style="margin-bottom:10px;">班別</p>
      <div class="choice-grid cols-6">
        ${SECTIONS.map(s=>`<div class="choice-btn ${section===s?'selected':''}" data-act="pick-section" data-val="${s}">${s}</div>`).join('')}
      </div>
    ` : ''}
    <div class="btn-row">
      <button class="btn btn-primary" data-act="to-step2" ${(grade&&section)?'':'disabled'} style="width:100%;">下一步</button>
    </div>
  </div>`;
}

/* ---- STUDENT STEP 2: 學號姓名 ---- */
function studentStep2(){
  const className = state.student.grade + state.student.section + '班';
  const list = state.roster.filter(r => r.class === className);
  const q = (state.search||'').trim();
  const filtered = q ? list.filter(r => r.id.includes(q) || r.name.includes(q)) : list;
  return `
  ${stepperBlock(2)}
  <div class="card">
    <h1 class="title">選擇學號與姓名</h1>
    <p class="subtitle">班級：${className}　共 ${list.length} 位學生</p>
    ${list.length ? `<input class="search-input" placeholder="輸入學號或姓名搜尋" id="search-box" value="${q}">` : ''}
    ${
      list.length === 0
      ? `<div class="empty-note">此班級尚未匯入學生名冊。<br>請聯繫護理人員以 Excel 匯入名冊後再試一次。</div>`
      : `<div class="student-list">
          ${filtered.map(r=>`
            <div class="student-row ${state.student.id===r.id?'selected':''}" data-act="pick-student" data-id="${r.id}" data-name="${r.name}">
              <span class="sname">${r.name}</span>
              <span class="sid">${r.id}</span>
            </div>`).join('') || `<div class="empty-note">找不到符合的學生</div>`}
        </div>`
    }
    <div class="btn-row">
      <button class="btn btn-ghost" data-act="back-step1">上一步</button>
      <button class="btn btn-primary" data-act="to-step3" ${state.student.id?'':'disabled'}>下一步</button>
    </div>
  </div>`;
}

/* ---- STUDENT STEP 3: 性別 ---- */
function studentStep3(){
  const g = state.student.gender;
  return `
  ${stepperBlock(3)}
  <div class="card">
    <h1 class="title">選擇性別</h1>
    <p class="subtitle">${state.student.name}（${state.student.id}）</p>
    <div class="choice-grid">
      <div class="choice-btn ${g==='生理男'?'selected':''}" data-act="pick-gender" data-val="生理男">生理男</div>
      <div class="choice-btn ${g==='生理女'?'selected':''}" data-act="pick-gender" data-val="生理女">生理女</div>
    </div>
    <div class="btn-row">
      <button class="btn btn-ghost" data-act="back-step2">上一步</button>
      <button class="btn btn-primary" data-act="to-step4" ${g?'':'disabled'}>下一步</button>
    </div>
  </div>`;
}

/* ---- STUDENT STEP 4: 傷病原因分類 ---- */
function studentStep4(){
  const r = state.student.reason;
  return `
  ${stepperBlock(4)}
  <div class="card">
    <h1 class="title">傷病原因</h1>
    <p class="subtitle">請選擇本次到健康中心的原因</p>
    <div class="reason-grid">
      <div class="reason-card illness ${r==='身體不適'?'selected':''}" data-act="pick-reason" data-val="身體不適">
        <div class="r-icon">${ICONS.thermo}</div>
        <h3>身體不適</h3>
      </div>
      <div class="reason-card injury ${r==='受傷'?'selected':''}" data-act="pick-reason" data-val="受傷">
        <div class="r-icon">${ICONS.bandage}</div>
        <h3>受傷</h3>
      </div>
    </div>
    <div class="btn-row">
      <button class="btn btn-ghost" data-act="back-step3">上一步</button>
      <button class="btn btn-primary" data-act="to-step5" ${r?'':'disabled'}>下一步</button>
    </div>
  </div>`;
}

/* ---- STUDENT STEP 5: 詳細原因 + 送出 ---- */
function studentStep5(){
  const s = state.student;
  const options = s.reason === '身體不適' ? ILLNESS_REASONS : INJURY_REASONS;
  const d = s.detail;
  return `
  ${stepperBlock(5)}
  <div class="card">
    <h1 class="title">${s.reason}詳細原因</h1>
    <p class="subtitle">請選擇最符合的項目</p>
    <div class="choice-grid cols-3">
      ${options.map(o=>`<div class="choice-btn ${d===o?'selected':''}" data-act="pick-detail" data-val="${o}">${o}</div>`).join('')}
    </div>
    <div class="summary-box" style="margin-top:22px;">
      <div class="summary-row"><span class="k">班級</span><span class="v">${s.grade}${s.section}班</span></div>
      <div class="summary-row"><span class="k">學號 / 姓名</span><span class="v">${s.id} ${s.name}</span></div>
      <div class="summary-row"><span class="k">性別</span><span class="v">${s.gender}</span></div>
      <div class="summary-row"><span class="k">原因</span><span class="v">${s.reason}${d ? '・'+d : ''}</span></div>
    </div>
    <div class="btn-row">
      <button class="btn btn-ghost" data-act="back-step4">上一步</button>
      <button class="btn btn-primary" id="submit-btn" data-act="submit-record" ${d?'':'disabled'}>提交</button>
    </div>
  </div>`;
}

function studentDone(){
  return `
  <div class="card done-wrap">
    <div class="done-check">${ICONS.check}</div>
    <h1 class="title">已送出報到</h1>
    <p class="subtitle">請直接前往健康中心，護理人員已收到您的資料。</p>
    <div class="btn-row">
      <button class="btn btn-primary" data-act="go-home" style="width:100%;">返回首頁</button>
    </div>
  </div>`;
}

function stepperBlock(step){
  const labels = ['班級','學號姓名','性別','原因','詳細原因'];
  return `
  <div class="stepper">
    ${stepperSvg(step, labels.length)}
    <div class="step-label-row">
      ${labels.map((l,i)=>`<div class="step-label ${i+1===step?'active':''}">${l}</div>`).join('')}
    </div>
  </div>`;
}

/* ---- NURSE LOGIN ---- */
function nurseLogin(){
  return `
  <div class="card" style="max-width:420px;margin:20px auto 0;">
    <h1 class="title">護理人員登入</h1>
    <p class="subtitle">請輸入帳號密碼以管理報到紀錄</p>
    <div class="field-group">
      <label>帳號</label>
      <input id="login-account" type="text" autocomplete="username">
    </div>
    <div class="field-group">
      <label>密碼</label>
      <input id="login-password" type="password" autocomplete="current-password">
    </div>
    ${state.loginErr ? `<div class="error-text">${state.loginErr}</div>` : ''}
    <div class="btn-row">
      <button class="btn btn-primary" id="login-btn" data-act="do-login" style="width:100%;" ${state.loginBusy?'disabled':''}>${state.loginBusy ? '登入中…' : '登入'}</button>
    </div>
    <div class="login-hint">預設示範帳號：nurse／密碼：1234（密碼已加密存放，建議登入後立即到「修改密碼」更換）</div>
  </div>`;
}

/* ---- NURSE CHANGE PASSWORD ---- */
function nurseChangePassword(){
  return `
  <div class="card" style="max-width:420px;margin:20px auto 0;">
    <h1 class="title">修改密碼</h1>
    <p class="subtitle">請輸入目前密碼與新密碼</p>
    <div class="field-group">
      <label>目前密碼</label>
      <input id="pwd-old" type="password">
    </div>
    <div class="field-group">
      <label>新密碼（至少 4 個字元）</label>
      <input id="pwd-new" type="password">
    </div>
    <div class="field-group">
      <label>再輸入一次新密碼</label>
      <input id="pwd-confirm" type="password">
    </div>
    ${state.pwdErr ? `<div class="error-text">${state.pwdErr}</div>` : ''}
    <div class="btn-row">
      <button class="btn btn-ghost" data-act="back-dashboard">取消</button>
      <button class="btn btn-primary" data-act="do-change-password" ${state.pwdBusy?'disabled':''}>${state.pwdBusy ? '處理中…' : '確認修改'}</button>
    </div>
  </div>`;
}

/* ---- NURSE RECORD TREATMENT ---- */
function nurseRecordTreatment(){
  const rec = state.records.find(r => r.recordId === state.treatmentRecordId);
  if(!rec){
    return `<div class="card"><div class="empty-note">找不到這筆紀錄，可能已被刪除。</div>
      <div class="btn-row"><button class="btn btn-primary" data-act="back-dashboard" style="width:100%;">回儀表板</button></div></div>`;
  }
  const sel = state.treatmentSelected;
  return `
  <div class="card">
    <h1 class="title">護理處置</h1>
    <p class="subtitle">${rec.name}（${rec.id}）・ ${rec.class} ・ ${rec.reason}${rec.detail ? '・'+rec.detail : ''}</p>
    ${rec.history ? `<div class="setup-warning" style="background:var(--injury-bg);color:var(--injury);">病史：${rec.history}</div>` : ''}
    <div class="choice-grid">
      ${TREATMENT_OPTIONS.map(o=>`<div class="choice-btn ${sel.includes(o)?'selected':''}" data-act="toggle-treatment" data-val="${o}">${o}</div>`).join('')}
    </div>
    ${sel.includes('其它') ? `
      <div class="field-group" style="margin-top:16px;">
        <label>其它（請說明）</label>
        <input id="treatment-other" type="text" value="${state.treatmentOther}" placeholder="請輸入處置內容">
      </div>
    ` : ''}
    <div class="btn-row">
      <button class="btn btn-ghost" data-act="back-dashboard">取消</button>
      <button class="btn btn-primary" data-act="save-treatment" ${state.treatmentBusy?'disabled':''}>${state.treatmentBusy ? '儲存中…' : '儲存'}</button>
    </div>
  </div>`;
}

/* ---- NURSE CASE MANAGEMENT (學生個案管理：名冊 + 病史) ---- */
function filterRosterFull(){
  const q = (state.caseSearch||'').trim();
  let list = [...state.rosterFull];
  if(q) list = list.filter(r => r.name.includes(q) || r.id.includes(q) || r.class.includes(q) || (r.history||'').includes(q));
  return list;
}

function nurseCaseManagement(){
  const q = (state.caseSearch||'').trim();
  const list = filterRosterFull();

  return `
  <div class="card" style="margin-bottom:18px;">
    <h1 class="title">學生名冊管理</h1>
    <p class="subtitle" style="margin-bottom:14px;">目前名冊共 ${state.rosterFull.length} 筆。匯入 Excel 需包含欄位：<b>學號</b>、<b>姓名</b>、<b>班級</b>（例如：大一甲班），可選填 <b>病史</b>。</p>
    <div class="dash-toolbar">
      <button class="btn btn-primary import-btn">${ICONS.upload} 匯入 Excel 名冊
        <input type="file" id="roster-file" accept=".xlsx,.xls,.csv">
      </button>
      ${state.rosterFull.length ? `<button class="btn btn-ghost" data-act="clear-roster">清空名冊</button>` : ''}
    </div>
  </div>

  <div class="card">
    <h1 class="title">學生個案搜尋</h1>
    <p class="subtitle" style="margin-bottom:14px;">可用姓名、學號、班級，或直接搜病史內容（例如「氣喘」）找出相關學生。</p>
    <div class="dash-toolbar">
      <input class="search-input" id="case-search" placeholder="搜尋姓名、學號、班級或病史" value="${q}">
      <button class="pill-btn" data-act="export-roster-excel">${ICONS.download} 匯出名單 Excel${q ? '（搜尋結果）' : ''}</button>
    </div>
    ${
      list.length === 0
      ? `<div class="empty-note">${state.rosterFull.length===0 ? '尚未匯入學生名冊' : '找不到符合的學生'}</div>`
      : `<div class="student-list" style="max-height:none;">
          ${list.map(r=>`
            <div class="student-row" style="cursor:default;align-items:flex-start;flex-direction:column;gap:4px;">
              <div style="display:flex;justify-content:space-between;align-items:center;width:100%;gap:8px;">
                <div>
                  <span class="sname">${r.name}</span>
                  <span class="sid" style="margin-left:8px;">${r.id}・${r.class}</span>
                </div>
                <button class="pill-btn" data-act="go-health-check" data-id="${r.id}">修改資料</button>
              </div>
              ${r.history ? `<div style="font-size:13px;color:var(--injury);">病史：${r.history}</div>` : `<div style="font-size:13px;color:var(--muted);">尚無病史紀錄</div>`}
              ${r.healthAbnormal && r.healthAbnormal.length ? `<div style="font-size:13px;color:var(--warn);font-weight:700;">⚠ 健檢異常：${r.healthAbnormal.join('、')}</div>` : ''}
            </div>`).join('')}
        </div>`
    }
  </div>`;
}

/* ---- NURSE HEALTH CHECK (學生健檢異常資料) ---- */
function nurseHealthCheck(){
  if(state.healthCheckLoading){
    return `<div class="loading-wrap">讀取資料中…</div>`;
  }
  const s = state.healthCheckStudent;
  if(!s){
    return `<div class="card"><div class="empty-note">找不到這位學生的資料。</div>
      <div class="btn-row"><button class="btn btn-primary" data-act="back-case-management" style="width:100%;">回學生個案管理</button></div></div>`;
  }
  return `
  <div class="card">
    <h1 class="title">健檢異常資料</h1>
    <p class="subtitle">${s.name}（${s.id}）・ ${s.class}</p>
    <p class="subtitle" style="margin-bottom:14px;">每個項目皆為選填，填入數值後可勾選是否異常。</p>
    ${HEALTH_CHECK_ITEMS.map(item=>{
      const entry = state.healthCheckItems[item] || {value:'', abnormal:false};
      return `
      <div class="field-group" style="display:flex;align-items:center;gap:10px;">
        <label style="flex:1;min-width:150px;margin-bottom:0;">${item}</label>
        <input type="text" class="hc-value" data-item="${item}" value="${entry.value}" placeholder="數值（選填）" style="flex:1.4;padding:10px 12px;border-radius:10px;border:1px solid var(--border);background:var(--surface-soft);">
        <div class="choice-btn hc-abnormal" data-act="toggle-hc-abnormal" data-item="${item}" style="flex:none;padding:10px 14px;font-size:13px;${entry.abnormal ? 'background:var(--injury);color:#fff;' : ''}">異常</div>
      </div>`;
    }).join('')}
    <div class="field-group" style="margin-top:22px;">
      <label>衛教指導</label>
      <div class="choice-grid">
        ${EDU_GUIDANCE_OPTIONS.map(o=>`<div class="choice-btn ${state.healthCheckGuidance.includes(o)?'selected':''}" data-act="toggle-hc-guidance" data-val="${o}">${o}</div>`).join('')}
      </div>
    </div>
    <div class="field-group" style="margin-top:18px;">
      <label>備註</label>
      <textarea id="hc-note" rows="3" style="width:100%;padding:12px 14px;border-radius:10px;border:1px solid var(--border);background:var(--surface-soft);font-family:var(--font-body);font-size:15px;">${state.healthCheckNote}</textarea>
    </div>
    <div class="btn-row">
      <button class="btn btn-ghost" data-act="back-case-management">取消</button>
      <button class="btn btn-primary" data-act="save-health-check" ${state.healthCheckBusy?'disabled':''}>${state.healthCheckBusy ? '儲存中…' : '儲存'}</button>
    </div>
  </div>`;
}

/* ---- NURSE DASHBOARD ---- */
function nurseDashboard(){
  const rangeRecords = filterByDateRange(state.records, state.dashFrom, state.dashTo);
  const total = rangeRecords.length;
  const illnessCount = rangeRecords.filter(r=>r.reason==='身體不適').length;
  const injuryCount = rangeRecords.filter(r=>r.reason==='受傷').length;

  const q = (state.dashSearch||'').trim();
  let list = [...rangeRecords].sort((a,b)=> b.ts - a.ts);
  if(q) list = list.filter(r => r.name.includes(q) || r.id.includes(q) || r.class.includes(q));

  return `
  <div class="card" style="margin-bottom:18px;">
    <h1 class="title" style="font-size:18px;">查看區間</h1>
    <div class="dash-toolbar">
      <div class="field-group" style="margin-bottom:0;flex:1;min-width:140px;">
        <label>起始日期</label>
        <input type="date" id="dash-from" value="${state.dashFrom}">
      </div>
      <div class="field-group" style="margin-bottom:0;flex:1;min-width:140px;">
        <label>結束日期</label>
        <input type="date" id="dash-to" value="${state.dashTo}">
      </div>
    </div>
    <div class="dash-toolbar" style="margin-top:4px;">
      <button class="pill-btn" data-act="range-preset" data-days="0">今天</button>
      <button class="pill-btn" data-act="range-preset" data-days="6">近 7 天</button>
      <button class="pill-btn" data-act="range-preset" data-days="29">近 30 天</button>
      <button class="pill-btn" data-act="range-all">全部紀錄</button>
    </div>
  </div>

  <div class="stat-row">
    <div class="stat-card"><div class="num">${total}</div><div class="lbl">區間總人次</div></div>
    <div class="stat-card"><div class="num" style="color:var(--illness)">${illnessCount}</div><div class="lbl">身體不適</div></div>
    <div class="stat-card"><div class="num" style="color:var(--injury)">${injuryCount}</div><div class="lbl">受傷</div></div>
  </div>

  <div class="card">
    <h1 class="title" style="font-size:18px;">報到紀錄（區間內）</h1>
    <div class="dash-toolbar">
      <input class="search-input" id="dash-search" placeholder="搜尋姓名、學號或班級" value="${q}">
      <button class="pill-btn" data-act="export-excel">${ICONS.download} 匯出 Excel</button>
    </div>
    ${
      list.length === 0
      ? `<div class="empty-note">此區間內沒有符合的報到紀錄</div>`
      : `<div class="record-list">
          ${list.map(r=>`
            <div class="record-card ${r.status==='done'?'done':''}">
              <span class="rc-badge ${r.reason==='身體不適'?'illness':'injury'}"></span>
              <div class="rc-main">
                <div class="rc-top">
                  <span class="rc-name">${r.name}</span>
                  <span class="rc-sid">${r.id}</span>
                  <span class="rc-class">${r.class}・${r.gender}</span>
                </div>
                <div class="rc-meta">${r.reason}${r.detail ? '・'+r.detail : ''} ・ ${formatTime(r.ts)} ${r.status==='done' ? '・ 已處理' : ''}</div>
                ${r.history ? `<div class="rc-meta" style="color:var(--injury);">病史：${r.history}</div>` : ''}
                ${r.treatment ? `<div class="rc-meta" style="color:var(--primary-dark);">處置：${r.treatment}</div>` : ''}
              </div>
              <div class="rc-actions">
                <div class="icon-btn" data-act="open-treatment" data-id="${r.recordId}" title="護理處置">${ICONS.clipboard}</div>
                <div class="icon-btn" data-act="toggle-status" data-id="${r.recordId}" title="標記已處理">${ICONS.done}</div>
                <div class="icon-btn danger" data-act="delete-record" data-id="${r.recordId}" title="刪除">${ICONS.trash}</div>
              </div>
            </div>`).join('')}
        </div>`
    }
  </div>`;
}

function filterByDateRange(records, from, to){
  if(!from && !to) return records;
  const fromTs = from ? new Date(from + 'T00:00:00').getTime() : -Infinity;
  const toTs = to ? new Date(to + 'T23:59:59').getTime() : Infinity;
  return records.filter(r => r.ts >= fromTs && r.ts <= toTs);
}

function formatTime(ts){
  const d = new Date(ts);
  const pad = n => String(n).padStart(2,'0');
  return `${d.getMonth()+1}/${d.getDate()} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
}

/* ---------------- events ---------------- */
function bindEvents(){
  document.querySelectorAll('[data-act]').forEach(el=>{
    el.addEventListener('click', onAct);
  });
  bindLiveSearch('search-box', v=>{ state.search = v; render(); preserveFocus('search-box'); });
  bindLiveSearch('dash-search', v=>{ state.dashSearch = v; render(); preserveFocus('dash-search'); });
  bindLiveSearch('case-search', v=>{ state.caseSearch = v; render(); preserveFocus('case-search'); });
  const dashFrom = document.getElementById('dash-from');
  if(dashFrom){
    dashFrom.addEventListener('change', e=>{ state.dashFrom = e.target.value; render(); });
  }
  const dashTo = document.getElementById('dash-to');
  if(dashTo){
    dashTo.addEventListener('change', e=>{ state.dashTo = e.target.value; render(); });
  }
  const treatmentOther = document.getElementById('treatment-other');
  if(treatmentOther){
    treatmentOther.addEventListener('input', e=>{ state.treatmentOther = e.target.value; });
  }
  document.querySelectorAll('.hc-value').forEach(el=>{
    el.addEventListener('input', e=>{
      const item = el.dataset.item;
      if(!state.healthCheckItems[item]) state.healthCheckItems[item] = {value:'', abnormal:false};
      state.healthCheckItems[item].value = e.target.value;
    });
  });
  const hcNote = document.getElementById('hc-note');
  if(hcNote){
    hcNote.addEventListener('input', e=>{ state.healthCheckNote = e.target.value; });
  }
  const rosterFile = document.getElementById('roster-file');
  if(rosterFile){
    rosterFile.addEventListener('change', handleRosterFile);
  }
  manageAutoRefresh();
}

// 中文（注音／拼音）等輸入法在選字前會經過「組字中」的狀態，
// 如果每個按鍵都立刻整頁重繪，選字過程會被打斷，導致打不出字、只能貼上。
// 這裡在組字期間（compositionstart ~ compositionend）先不重繪，
// 等使用者確定選字完成才真正更新畫面。
function bindLiveSearch(id, onChange){
  const el = document.getElementById(id);
  if(!el) return;
  let composing = false;
  el.addEventListener('compositionstart', ()=>{ composing = true; });
  el.addEventListener('compositionend', (e)=>{
    composing = false;
    onChange(e.target.value);
  });
  el.addEventListener('input', (e)=>{
    if(composing) return;
    onChange(e.target.value);
  });
}

function preserveFocus(id){
  const el = document.getElementById(id);
  // 只有在焦點真的跑掉時才需要重新 focus；重複呼叫 focus() 會讓瀏覽器
  // 每次都嘗試把輸入框捲動到可視範圍內，造成畫面一直跳動／晃動。
  if(el && document.activeElement !== el){
    try{ el.focus({preventScroll:true}); }catch(err){ el.focus(); }
    const v = el.value;
    el.value = v;
    el.setSelectionRange(v.length, v.length);
  }
}

function manageAutoRefresh(){
  if(refreshTimer){ clearInterval(refreshTimer); refreshTimer = null; }
  if(state.screen === 'nurse-dashboard'){
    refreshTimer = setInterval(async ()=>{
      await loadRecords();
      render();
    }, 20000); // 每 20 秒自動重新整理一次
  }
}

async function onAct(e){
  const el = e.currentTarget;
  const act = el.dataset.act;
  switch(act){
    case 'go-home':
      state.screen = 'home';
      state.student = { grade:null, section:null, id:null, name:null, gender:null, reason:null, detail:null };
      state.search = '';
      render(); break;

    case 'start-student':
      state.screen = 'student-1'; render(); break;

    case 'pick-grade':
      state.student.grade = el.dataset.val;
      state.student.section = null;
      render(); break;
    case 'pick-section':
      state.student.section = el.dataset.val;
      render(); break;
    case 'to-step2':
      state.screen = 'student-2'; state.search=''; render(); break;
    case 'back-step1':
      state.screen = 'student-1'; render(); break;

    case 'pick-student':
      state.student.id = el.dataset.id;
      state.student.name = el.dataset.name;
      render(); break;
    case 'to-step3':
      state.screen = 'student-3'; render(); break;
    case 'back-step2':
      state.screen = 'student-2'; render(); break;

    case 'pick-gender':
      state.student.gender = el.dataset.val; render(); break;
    case 'to-step4':
      state.screen = 'student-4'; render(); break;
    case 'back-step3':
      state.screen = 'student-3'; render(); break;

    case 'pick-reason':
      state.student.reason = el.dataset.val;
      state.student.detail = null;
      render(); break;
    case 'to-step5':
      state.screen = 'student-5'; render(); break;
    case 'back-step4':
      state.screen = 'student-4'; render(); break;

    case 'pick-detail':
      state.student.detail = el.dataset.val; render(); break;
    case 'submit-record':
      await submitRecord(); break;

    case 'go-nurse-login':
      state.screen = 'nurse-login'; state.loginErr=''; render(); break;
    case 'do-login':
      await doLogin(); break;
    case 'logout':
      state.loggedIn = false; state.sessionToken = null; state.records = []; state.screen = 'home'; render(); break;

    case 'go-change-password':
      state.pwdErr = ''; state.screen = 'nurse-change-password'; render(); break;
    case 'back-dashboard':
      state.screen = 'nurse-dashboard'; render(); break;
    case 'do-change-password':
      await doChangePassword(); break;

    case 'go-case-management':
      state.loading = true; render();
      await loadRosterFull();
      state.loading = false;
      state.screen = 'nurse-case-management';
      render();
      break;
    case 'export-roster-excel':
      await exportRosterToExcel(); break;

    case 'go-health-check': {
      const studentId = el.dataset.id;
      state.healthCheckStudentId = studentId;
      state.healthCheckStudent = state.rosterFull.find(r => r.id === studentId) || null;
      state.healthCheckItems = {};
      state.healthCheckGuidance = [];
      state.healthCheckNote = '';
      state.healthCheckLoading = true;
      state.screen = 'nurse-health-check';
      render();
      await loadHealthCheck(studentId);
      state.healthCheckLoading = false;
      render();
      break;
    }
    case 'back-case-management':
      state.screen = 'nurse-case-management'; render(); break;
    case 'toggle-hc-abnormal': {
      const item = el.dataset.item;
      if(!state.healthCheckItems[item]) state.healthCheckItems[item] = {value:'', abnormal:false};
      state.healthCheckItems[item].abnormal = !state.healthCheckItems[item].abnormal;
      render();
      break;
    }
    case 'toggle-hc-guidance': {
      const v = el.dataset.val;
      const idx = state.healthCheckGuidance.indexOf(v);
      if(idx === -1) state.healthCheckGuidance.push(v); else state.healthCheckGuidance.splice(idx,1);
      render();
      break;
    }
    case 'save-health-check':
      await saveHealthCheck(); break;

    case 'range-preset':
      state.dashFrom = todayStr(-Number(el.dataset.days));
      state.dashTo = todayStr(0);
      render(); break;
    case 'range-all':
      state.dashFrom = ''; state.dashTo = '';
      render(); break;

    case 'toggle-status':
      await toggleStatus(el.dataset.id); break;
    case 'delete-record':
      await deleteRecord(el.dataset.id); break;

    case 'open-treatment': {
      const rec = state.records.find(r => r.recordId === el.dataset.id);
      state.treatmentRecordId = el.dataset.id;
      const parsed = parseTreatmentString(rec ? rec.treatment : '');
      state.treatmentSelected = parsed.options;
      state.treatmentOther = parsed.other;
      state.screen = 'nurse-record-treatment';
      render();
      break;
    }
    case 'toggle-treatment': {
      const v = el.dataset.val;
      const idx = state.treatmentSelected.indexOf(v);
      if(idx === -1) state.treatmentSelected.push(v); else state.treatmentSelected.splice(idx,1);
      if(v === '其它' && idx !== -1) state.treatmentOther = '';
      render();
      break;
    }
    case 'save-treatment':
      await saveTreatment(); break;

    case 'export-excel':
      exportRecordsToExcel(); break;

    case 'clear-roster':
      if(confirm('確定要清空整份學生名冊嗎？此動作無法復原。')){
        await apiPost('clearRoster', { token: state.sessionToken });
        state.roster = [];
        state.rosterFull = [];
        render();
        showToast('名冊已清空');
      }
      break;
  }
}

async function submitRecord(){
  const btn = document.getElementById('submit-btn');
  if(btn){ btn.disabled = true; btn.textContent = '提交中…'; }
  const s = state.student;
  const rec = {
    recordId: 'r_' + Date.now() + '_' + Math.random().toString(36).slice(2,7),
    class: s.grade + s.section + '班',
    id: s.id, name: s.name, gender: s.gender, reason: s.reason, detail: s.detail,
    status: 'pending', ts: Date.now()
  };
  try{
    const res = await apiPost('addRecord', rec);
    if(res && res.ok){
      state.records = [...state.records, rec];
      state.screen = 'student-done';
      render();
    } else {
      showToast('送出失敗，請稍後再試一次');
      if(btn){ btn.disabled = false; btn.textContent = '提交'; }
    }
  }catch(err){
    console.error(err);
    showToast('無法連線，請確認網路連線後再試一次');
    if(btn){ btn.disabled = false; btn.textContent = '提交'; }
  }
}

async function doLogin(){
  const acc = document.getElementById('login-account').value.trim();
  const pwd = document.getElementById('login-password').value;
  state.loginBusy = true; render();
  try{
    const res = await apiPost('login', { account: acc, password: pwd });
    if(res && res.ok && res.success){
      state.loggedIn = true; state.loginErr=''; state.sessionToken = res.token;
      state.loading = true; render();
      await loadRecords();
      state.loading = false;
      state.screen = 'nurse-dashboard';
    } else {
      state.loginErr = (res && res.error) || '帳號或密碼錯誤，請再試一次';
    }
  }catch(err){
    console.error(err);
    state.loginErr = '無法連線到伺服器，請稍後再試';
  }
  state.loginBusy = false;
  render();
}

async function doChangePassword(){
  const oldPwd = document.getElementById('pwd-old').value;
  const newPwd = document.getElementById('pwd-new').value;
  const confirmPwd = document.getElementById('pwd-confirm').value;
  state.pwdErr = '';
  if(newPwd !== confirmPwd){
    state.pwdErr = '兩次輸入的新密碼不一致';
    render(); return;
  }
  state.pwdBusy = true; render();
  try{
    const res = await apiPost('changePassword', { token: state.sessionToken, oldPassword: oldPwd, newPassword: newPwd });
    if(res && res.ok && res.success){
      state.pwdBusy = false;
      state.screen = 'nurse-dashboard';
      render();
      showToast('密碼已更新');
    } else {
      state.pwdErr = (res && res.error) || '修改失敗，請再試一次';
      state.pwdBusy = false;
      render();
    }
  }catch(err){
    console.error(err);
    state.pwdErr = '無法連線到伺服器，請稍後再試';
    state.pwdBusy = false;
    render();
  }
}

async function toggleStatus(id){
  const rec = state.records.find(r=>r.recordId===id);
  if(!rec) return;
  const newStatus = rec.status === 'done' ? 'pending' : 'done';
  state.records = state.records.map(r => r.recordId===id ? {...r, status:newStatus} : r);
  render();
  try{
    await apiPost('updateRecordStatus', { token: state.sessionToken, recordId:id, status:newStatus });
  }catch(err){
    console.error(err);
    showToast('狀態更新失敗，請重新整理後再試');
  }
}

async function deleteRecord(id){
  if(!confirm('確定要刪除這筆報到紀錄嗎？')) return;
  state.records = state.records.filter(r=>r.recordId!==id);
  render();
  try{
    await apiPost('deleteRecord', { token: state.sessionToken, recordId:id });
  }catch(err){
    console.error(err);
    showToast('刪除失敗，請重新整理後再試');
  }
}

/* ---------------- 護理處置 ---------------- */
function parseTreatmentString(str){
  if(!str) return { options: [], other: '' };
  const parts = str.split('、').map(s=>s.trim()).filter(Boolean);
  const options = [];
  let other = '';
  parts.forEach(p=>{
    if(p.startsWith('其它：') || p.startsWith('其它:')){
      options.push('其它');
      other = p.replace(/^其它[：:]/, '');
    } else if(TREATMENT_OPTIONS.includes(p)){
      options.push(p);
    }
  });
  return { options, other };
}

function buildTreatmentString(options, other){
  return options.map(o => o === '其它' ? `其它：${other||''}` : o).join('、');
}

async function saveTreatment(){
  const treatmentStr = buildTreatmentString(state.treatmentSelected, state.treatmentOther);
  state.treatmentBusy = true; render();
  try{
    const res = await apiPost('updateTreatment', { token: state.sessionToken, recordId: state.treatmentRecordId, treatment: treatmentStr });
    if(res && res.ok){
      state.records = state.records.map(r => r.recordId===state.treatmentRecordId ? {...r, treatment: treatmentStr} : r);
      state.treatmentBusy = false;
      state.screen = 'nurse-dashboard';
      render();
      showToast('護理處置已儲存');
    } else {
      state.treatmentBusy = false;
      showToast('儲存失敗，請稍後再試一次');
      render();
    }
  }catch(err){
    console.error(err);
    state.treatmentBusy = false;
    showToast('無法連線，請稍後再試一次');
    render();
  }
}

async function saveHealthCheck(){
  state.healthCheckBusy = true; render();
  try{
    const res = await apiPost('updateHealthCheck', {
      token: state.sessionToken,
      studentId: state.healthCheckStudentId,
      items: state.healthCheckItems,
      guidance: state.healthCheckGuidance,
      note: state.healthCheckNote
    });
    if(res && res.ok){
      const abnormalList = HEALTH_CHECK_ITEMS.filter(item => state.healthCheckItems[item] && state.healthCheckItems[item].abnormal);
      state.rosterFull = state.rosterFull.map(r => r.id === state.healthCheckStudentId
        ? {...r, healthAbnormal: abnormalList, healthItems: state.healthCheckItems, healthGuidance: state.healthCheckGuidance, healthNote: state.healthCheckNote}
        : r);
      state.healthCheckBusy = false;
      state.screen = 'nurse-case-management';
      render();
      showToast('健檢資料已儲存');
    } else {
      state.healthCheckBusy = false;
      showToast('儲存失敗，請稍後再試一次');
      render();
    }
  }catch(err){
    console.error(err);
    state.healthCheckBusy = false;
    showToast('無法連線，請稍後再試一次');
    render();
  }
}

/* ---------------- 匯出名冊 Excel（依目前搜尋結果） ---------------- */
async function exportRosterToExcel(){
  const list = filterRosterFull();
  if(list.length === 0){
    showToast('目前沒有資料可以匯出');
    return;
  }
  if(typeof ExcelJS === 'undefined'){
    showToast('匯出功能載入失敗，請重新整理頁面後再試');
    return;
  }

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('學生名冊');

  const columns = [
    { header:'學號', key:'id', width:10 },
    { header:'姓名', key:'name', width:10 },
    { header:'班級', key:'class', width:12 },
    { header:'病史', key:'history', width:26 },
  ];
  HEALTH_CHECK_ITEMS.forEach(item=>{
    columns.push({ header:item+'（數值）', key:item+'_v', width:12 });
    columns.push({ header:item+'（異常）', key:item+'_a', width:10 });
  });
  EDU_GUIDANCE_OPTIONS.forEach(g=>{
    columns.push({ header:g, key:'g_'+g, width:9 });
  });
  columns.push({ header:'備註', key:'note', width:30 });
  ws.columns = columns;
  ws.getRow(1).font = { bold:true };
  ws.getRow(1).fill = { type:'pattern', pattern:'solid', fgColor:{argb:'FFDCEEEC'} };

  list.forEach(r=>{
    const rowData = { id:r.id, name:r.name, class:r.class, history:r.history || '', note:r.healthNote || '' };
    HEALTH_CHECK_ITEMS.forEach(item=>{
      const entry = (r.healthItems && r.healthItems[item]) || {value:'', abnormal:false};
      rowData[item+'_v'] = entry.value || '';
      rowData[item+'_a'] = entry.abnormal ? '異常' : '';
    });
    EDU_GUIDANCE_OPTIONS.forEach(g=>{
      rowData['g_'+g] = (r.healthGuidance||[]).includes(g) ? '✓' : '';
    });
    const row = ws.addRow(rowData);

    HEALTH_CHECK_ITEMS.forEach(item=>{
      const entry = (r.healthItems && r.healthItems[item]) || {value:'', abnormal:false};
      if(entry.abnormal){
        [row.getCell(item+'_v'), row.getCell(item+'_a')].forEach(cell=>{
          cell.fill = { type:'pattern', pattern:'solid', fgColor:{argb:'FFC1533F'} };
          cell.font = { color:{argb:'FFFFFFFF'}, bold:true };
        });
      }
    });
  });

  try{
    const buffer = await wb.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type:'application/octet-stream' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    const q = (state.caseSearch||'').trim();
    a.download = q ? `學生名冊_搜尋_${q}_${todayStr(0)}.xlsx` : `學生名冊_完整_${todayStr(0)}.xlsx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    showToast('已匯出 Excel');
  }catch(err){
    console.error(err);
    showToast('匯出失敗，請稍後再試一次');
  }
}

/* ---------------- 匯出 Excel ---------------- */
function exportRecordsToExcel(){
  const rangeRecords = filterByDateRange(state.records, state.dashFrom, state.dashTo);
  const q = (state.dashSearch||'').trim();
  let list = [...rangeRecords].sort((a,b)=> a.ts - b.ts);
  if(q) list = list.filter(r => r.name.includes(q) || r.id.includes(q) || r.class.includes(q));

  if(list.length === 0){
    showToast('目前區間內沒有資料可以匯出');
    return;
  }

  const rows = list.map(r => {
    const parsed = parseTreatmentString(r.treatment);
    const row = {
      '時間': formatTime(r.ts),
      '班級': r.class,
      '學號': r.id,
      '姓名': r.name,
      '性別': r.gender,
      '病史': r.history || '',
      '原因': r.reason,
      '狀態': r.status === 'done' ? '已處理' : '未處理',
    };
    REASON_DETAIL_OPTIONS.forEach(opt => {
      row[opt] = (r.detail === opt) ? 1 : '';
    });
    TREATMENT_OPTIONS.forEach(opt => {
      row[opt] = parsed.options.includes(opt) ? 1 : '';
    });
    row['其它內容'] = parsed.other || '';
    return row;
  });

  const ws = XLSX.utils.json_to_sheet(rows);
  const baseWidths = [14,10,10,10,8,20,10,8];
  const colWidths = [...baseWidths, ...new Array(REASON_DETAIL_OPTIONS.length).fill(6), ...new Array(TREATMENT_OPTIONS.length).fill(6), 30];
  ws['!cols'] = colWidths.map(w=>({wch:w}));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, '報到紀錄');
  const fname = `報到紀錄_${state.dashFrom||'全部'}_${state.dashTo||'全部'}.xlsx`;
  XLSX.writeFile(wb, fname);
  showToast('已匯出 Excel');
}

async function handleRosterFile(e){
  const file = e.target.files[0];
  if(!file) return;
  try{
    const data = await file.arrayBuffer();
    const wb = XLSX.read(data, {type:'array'});
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, {defval:''});
    const parsed = rows.map(row=>{
      const id = String(row['學號'] ?? row['id'] ?? '').trim();
      const name = String(row['姓名'] ?? row['name'] ?? '').trim();
      const cls = String(row['班級'] ?? row['class'] ?? '').trim();
      const historyRaw = row['病史'] ?? row['history'];
      const history = (historyRaw === undefined || historyRaw === null) ? undefined : String(historyRaw).trim();
      return {id, name, class:cls, history};
    }).filter(r=>r.id && r.name && r.class);

    if(parsed.length === 0){
      showToast('未找到有效資料，請確認欄位為「學號」「姓名」「班級」');
      return;
    }
    showToast('匯入中，請稍候…');
    const res = await apiPost('importRoster', { token: state.sessionToken, rows: parsed });
    if(res && res.ok){
      await loadRoster();
      await loadRosterFull();
      render();
      showToast(`已匯入 ${parsed.length} 筆學生資料`);
    } else {
      showToast('匯入失敗，請稍後再試一次');
    }
  }catch(err){
    console.error(err);
    showToast('匯入失敗，請確認檔案格式正確');
  }finally{
    e.target.value = '';
  }
}

init();
