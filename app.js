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
// Apps Script 偶爾會有短暫的連線波動，讀取類的請求自動重試一次，
// 減少單純因為網路小抖動就跳出「無法連線」的情況
async function apiGetWithRetry(action, params, retries){
  retries = (retries === undefined) ? 1 : retries;
  try{
    return await apiGet(action, params);
  }catch(e){
    if(retries > 0){
      await new Promise(resolve => setTimeout(resolve, 1200));
      return apiGetWithRetry(action, params, retries - 1);
    }
    throw e;
  }
}
function urlConfigured(){
  return APPS_SCRIPT_URL && !APPS_SCRIPT_URL.includes('PASTE_YOUR');
}

/* ---------------- 按需載入 Excel 相關函式庫（避免每個人一打開網站就先下載這兩包東西） ---------------- */
const scriptLoadPromises = {};
function loadScript(url){
  if(scriptLoadPromises[url]) return scriptLoadPromises[url];
  scriptLoadPromises[url] = new Promise((resolve, reject)=>{
    const s = document.createElement('script');
    s.src = url;
    s.onload = () => resolve();
    s.onerror = () => { delete scriptLoadPromises[url]; reject(new Error('無法載入：' + url)); };
    document.head.appendChild(s);
  });
  return scriptLoadPromises[url];
}
function ensureXLSX(){
  if(typeof XLSX !== 'undefined') return Promise.resolve();
  return loadScript('https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js');
}
function ensureExcelJS(){
  if(typeof ExcelJS !== 'undefined') return Promise.resolve();
  return loadScript('https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js');
}

/* ---------------- app state ---------------- */
const CLASSES = ['京劇學系','民俗技藝學系','戲曲音樂學系','歌仔戲學系','劇場藝術學系','客家戲學系','劇場藝術科'];
const SECTIONS = ['甲','乙','丙','丁','戊','己']; // 目前未使用，保留供未來擴充

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

// 家裡貓咪的大頭照，拿來當報到流程進度條的頭像 🐱
const STEP_CAT_AVATARS = [
  'data:image/jpeg;base64,/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAkGBwgHBgkIBwgKCgkLDRYPDQwMDRsUFRAWIB0iIiAdHx8kKDQsJCYxJx8fLT0tMTU3Ojo6Iys/RD84QzQ5Ojf/2wBDAQoKCg0MDRoPDxo3JR8lNzc3Nzc3Nzc3Nzc3Nzc3Nzc3Nzc3Nzc3Nzc3Nzc3Nzc3Nzc3Nzc3Nzc3Nzc3Nzc3Nzf/wAARCAB4AHgDASIAAhEBAxEB/8QAHAAAAQUBAQEAAAAAAAAAAAAABgACAwQFBwEI/8QAOxAAAgEDAgMEBgkDBAMAAAAAAQIDAAQRBRIGITETQVFhByIycYGRFCNCUmJyobGyFSTBNpLh8DPR4v/EABkBAAIDAQAAAAAAAAAAAAAAAAIDAAEEBf/EACURAAMAAgICAgICAwAAAAAAAAABAgMRITEEEhMyIlEFI0FCYf/aAAwDAQACEQMRAD8AArbUpLbIhRCc+0wqd+INQI5Sqv5UFZk9tcW4Tt4mTf7Oe+pLe0knPP1VHeaFNpcMc52+UPTWtQmuGVryTGcYBxRLbzM0a72JOO80zSNBsoLbdcqJJHOWJHStCSxtkXMLsvkTyqm2y1APRPt126X7yA1s2jEmq40cyXpuoHcuV2YK4U+41LbK8MzRyrtcd1Rk9Wuy/fLv02cAfZo54FH1J/IKDMb7WUeKGjTgU5hI/BUf1Bf2C0Cg/wBLYzwFqH5o/wCVGNB/pZ/0Jf8A5o/5UoOezm3AB26YvP7TfvRLuHPnXHotWv8AT2KWly0a7fZ7udV31bUpuT3twfIORRfG2DVpM7BcX1rbg9tcRJj7zCsq54o0iLOLtXPgnOuWiG4nbO2Rye85NWYNMu852BffRLGl2yvf/gaXPGdgoOyGZ/hilQqNHlf22x7hSq/WSts6ZJpDXsQimBR1PqiRcj4EVm3fDt/ZyoVjJUnkD3+410xNOCOfVyKt20cbBra5XcnVCRSt64Om4VLZxPUNQvEnaB43ixy2MCrE1PoWp3MlwbXapUr/AORgSUHj50VelXRrs2S36xpIsUoUSxrhljI6N8e+qGgaUllZKrEpcyKGYSLyI8jQ3bkbgwq3r/BsQadc2USyS/WooGMNnl3keNY+pgGSKdSNrEqD4+FaiX0iJJbvIwQZABHMcsU7RtJj1m87K4Y9mi7mK8vW/wCarDTq0g/OxTOJ0+NFK0+siZR3qf2ov4BbKEfgqpfaHBpMu633mNlwNxzzqXgB/XZfI1rqdSzgqlTTQcUIelf/AELf/mj/AJUXZoR9K3PgXUPzR/yFZx09nBbTS47tVmkLZ6EeNasGmQr7MSj4Va4esnnsFZcY3Gt6DTDj1iBUdFueTCSyUfZqZLQdy0RR6dGvUE1YS1RRyQfKq9iaBxLB2HJDSon7HypVPYmjo4j2sPLrXr26OvmOdQf1COC5FtdlQzHCuOhPh5H31dHtnwxyouGbH7TyUJ4UurSazuAGSVChBHWg+S5/tnij2m4VsKpwenLkT06Uf9mh6gFvOuV61LJbcSarZTW8piMm+KUYAGVBx8yaCop9D8GbHLftweTxmbE4wG5ZZQQM+Bor4MstsEl2QR2uMZ8uv60IWzvFEyY3BxzGOQNbMXEN7b2y29pHDAijAONxrR42Gpr2pGT+S8uMkfHD2FnEMayaazn2ovWHuod4FkU6hIFBC5bANYl3fXt3n6RdSuPu7sD5CtHgYNBqkokIAJLDBzyIp+ZJS2cvDvejotCnpQBPBGoAeKfyFbomfKynOCfZ8BQb6R+JdGk4fu9Nh1GCW8kZQIo23Hk3POOlc6bVdG5w5a2CXB0QOkRn8TfvREkHlWLwUM6PHj77fvRBPd2lkm67uYoh+JufyqPstnq24p4gHhWDe8a6ZBlbSKW6fxA2r8zWFccWa3qLmOwiWLP2YELt86JSwdoOZVihUvNIkajvY4pUDQ8J8QaswkvnaJT9q4fcf9tKr9Z/ZW3+jV0uyvLQb5r2aQHmEVyefvaj3SOK9PuCkEkogm5DZMdufcehoThbdbhsjHjWdq8EIszdTRyG3JI7RIyy594FaqwwltcAR5OTenyjrSyFnJz6uetc/wBdkE2t3j4HKTbn3ACs/h/jyewijhl02e7tYwAswYIQPPNDmq6zc3d/dTvqEdvFLKzrHCm+QAnoT0zQ4vwe2HnpXKSCWSSOJd0rqg8WOKyLniXTYnMcLvdSfct0Lfr0oYkms3mD/R5ruQdHupDj5CpUnvriQQWybd/JYrdME/LnTXmMqxGpc67fspKW0Fkh6NdSet/tFEno+Fxd6ok8c7tCkTvMwTakrNhVxn4n4Vgad6OtQkuYbq5kjhflIIppN7MO8culdO0Gyj0nTo7ZTl+RZwuM/Cud5Hmw5cp7NuHxaVJtBDgFAo5+eK+deK7X6HxPfwKvJLpyPcTn/NfQkLsScENisTUOHtIutUnvLmxjluJAA7Pk9AByFKw2huSWcY02fWZdtlpxuWj67IV8fEit+w4D1i9ftL50tlPUyNvf5V1W0tYLWIR20SRIOiooAqbZT/k/Qn1AvT+ANKtsNcmW6f8AGcL8hRHbafb2iBLWCOFR3IoFaO0V4VoXTfZetFTsqVWilKqIcMnXTYU2LLe3R8C5RPlXSvRprEF3pX9OEMcDW+QsaknKnnn9aD4OG7u7IMUGFPfiiDQuFb7TrpblJxGw6gDOaZl1c6Bx7mtmxxNwNY6rm4tZHsrnHJo/Yb8yf+qCZeBNXhkwtvFIB0khcbG+BwVPwrrIYiNQ/N25ZpjxhTtHtHqcVgfkZMXHZr+GL5OeaVwA0mJdRZ09YfURY3Y78nuozstH0rRo+1sLWOJs7WJ9se8nnWpLHLJZf2nqzg5Cvy3eQqKSWN7rZcIQzoAQe80q8uTJ2+A4iJ6Q9I1mmuHTG9QpjPcRjpUFzqFpb6PDe6hLFabzgmRsfADqTUt0URE7NiAjDJHhQlxxp2oXOtvNZxRXTWFsjW9q+CHLbsttPI4wOXfUxYlT5ZWS3K2E+nX1pfQdvp9yk0Y6leX6VoNholAGTnOaFuELL+nadA+pRhLl4cTIAMluo6dCKtm+uYruQW0ck9opGVJy6+JHiKnq8d/iy1/ZPJthfCltqDTLtLyMtG4cciCPDz86u7a2TXstmal6vRDtr3bUu2vNtGCR7aVSYpVZCJLcIMKox5U8RCrKpXuwVRCq6bV3YBI6ZqrNJ2O5jkn7Pme6r143ZQ7u8nFCs2sKdYkLQtJBaj2u7ee/4Vi8hbvRqw/UDdeg1BLiG/vkvLmWZ2CLFIyiNs4UcuYxz+VGNnqMsFu1lqzSTTRFTBclebfhY+I8akfUY2YyMSgc5whxnzpt9H9PgCMoSNF6Dljy99Ssu5U6CnE1W2zVW5EkY58mXAA65rEsbmZtZmE0JcrgfWLnl4U36Zd6WEt0sZyu0kzsoxy6c60NAsbi4s59Qv5Gd5iewSIE7V8T4mqnHTXA2nMrY2+fajgernqEXBx7hVa01KytbB2SGSVm5SllOCPD3VYhsZpLhjMZBk4+shP6ZP8AxTmshZI5aGWQeGRiijG+xV2uiThloJ7y6ntnGx1GUUYA9woixQ3w0udUlbs2i+q5r3HnyooIrRj4RmyfYjxXmKea8xmmADMUqfilVkJdte4GKdivcVRClqVu1xZyJHneBuXzI7qCrSeK0s5Xv4wJHcsVGOY8Of7UezXNvApMkqjHnXPNcjNxqQuPVKEscv7I99Zs2to04d6afQxhHeSRm1QRR5DEAHe3xPdWi1rIGQs7ddyqnN2/wB4k0y3uIyoRxHhR60gxjryAxWnE8d1Bt3YQc9oONw8z30ExsN20WTeYsCZ2XZt9bHsADuHj5k0PaXrVwibTZEW+49mVb1gueWR/itae1N1GI2cbc+tg9aXZwW4G+SNU6czWuWkuRW0edrFdlCIEYDqGTmv61HqE2BugkIPQlDgZ8PI+/NQ3F9KV22sBC97tgVWsbaUTmeVwM8m59fKhprfBWv2bHDKkXDvIp7RlxuYczRIRmsfSXElwCgwoXAFbWKuRddjNuK8xT68NECMxSpxFKrIVbzVobYlE+tk+6p6VhahrM02QHVAOqhjSpVgyZKbaN+LFKSZiLcT3FyEBLKTlsnuHXnVq7HJUADKw55HWlSpX+yHtbTIhaphFttoTuGO//v8AmnrbyJCzGVjluv8A3p/9UqVaF0ZWVGSRmLNNIEPLAPKmKO3jKbdqo/LB6ilSo0CyeAMuNrMAOWMZzWjGMqPE+PKlSpki6NPSWaO8RcYDfrRDSpUSF0eV5SpUQJ5SpUqhD//Z',
  'data:image/jpeg;base64,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',
  'data:image/jpeg;base64,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',
  'data:image/jpeg;base64,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',
];

let state = {
  screen: 'home',
  loading: true,
  loadError: null,
  roster: [],
  records: [],
  loggedIn: false,
  sessionToken: null,
  toast: null,
  student: { method:null, grade:null, id:null, name:null, class:null, gender:null, reason:null, detail:null },
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
  rosterImportBusy: false,
  rosterExportBusy: false,
  recordsExportBusy: false,
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
  render(); // 先畫出「讀取資料中…」，避免抓資料期間畫面完全空白
  await loadRoster();
  state.loading = false;
  render();
}

async function loadRoster(){
  try{
    const data = await apiGetWithRetry('getRoster');
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
    const data = await apiGetWithRetry('getRecords', { token: state.sessionToken });
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
    const data = await apiGetWithRetry('getRosterFull', { token: state.sessionToken });
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
  const r = 18;
  const gap = (w - margin*2) / (total-1);
  const cx = [];
  for(let i=0;i<total;i++) cx.push(margin + i*gap);

  let path = `M0 26 L20 26 L28 10 L36 42 L44 26 L${cx[0]-r-3} 26`;
  for(let i=0;i<total-1;i++){
    path += ` L${cx[i]+r+3} 26 L${cx[i]+ (cx[i+1]-cx[i])/2 - 8} 26 L${cx[i]+(cx[i+1]-cx[i])/2} 8 L${cx[i]+(cx[i+1]-cx[i])/2+8} 44 L${cx[i]+(cx[i+1]-cx[i])/2+16} 26 L${cx[i+1]-r-3} 26`;
  }
  path += ` L${w} 26`;
  let defs = cx.map((x,i)=>`<clipPath id="cat-clip-${i}"><circle cx="${x}" cy="26" r="${r-2}"/></clipPath>`).join('');
  let circles = cx.map((x,i)=>{
    const idx = i+1;
    const active = idx <= step;
    const stroke = active ? 'var(--primary)' : 'var(--border)';
    const avatar = STEP_CAT_AVATARS[i % STEP_CAT_AVATARS.length];
    return `
      <circle cx="${x}" cy="26" r="${r}" fill="var(--surface)" stroke="${stroke}" stroke-width="2"/>
      <image href="${avatar}" x="${x-(r-2)}" y="${26-(r-2)}" width="${(r-2)*2}" height="${(r-2)*2}" clip-path="url(#cat-clip-${i})" preserveAspectRatio="xMidYMid slice" style="${active ? '' : 'opacity:0.45;'}"/>
      <circle cx="${x}" cy="26" r="${r}" fill="none" stroke="${stroke}" stroke-width="2"/>`;
  }).join('');
  return `<svg viewBox="0 0 ${w} 52" preserveAspectRatio="none">
    <defs>${defs}</defs>
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
      <div class="brand" data-act="go-brand-home" style="cursor:pointer;">
        <div class="brand-mark">${ICONS.cross.replace('currentColor','#fff')}</div>
        <div>
          <div class="brand-text">健康中心報到系統</div>
          <div class="brand-sub">SCHOOL HEALTH CENTER</div>
        </div>
      </div>
      ${headerRight()}
    </div>
    ${state.loadError ? `<div class="setup-warning">${state.loadError}<div style="margin-top:10px;"><button class="btn btn-ghost" data-act="retry-load">重新嘗試</button></div></div>` : ''}
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
    case 'student-select': return studentSelect();
    case 'student-id': return studentIdSearch();
    case 'student-class': return studentClassBrowse();
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

/* ---- STUDENT SELECT: 選擇要用哪種方式找到自己 ---- */
function studentSelect(){
  return `
  ${stepperBlock(1)}
  <div class="card">
    <h1 class="title">選擇學生</h1>
    <p class="subtitle">可以直接輸入學號查詢，或選擇班級瀏覽名冊</p>
    <div class="role-grid">
      <div class="role-card student" data-act="go-id-search">
        <div class="icon-wrap">${ICONS.student}</div>
        <h3>輸入學號查詢</h3>
        <p>知道學號，直接查詢最快</p>
      </div>
      <div class="role-card nurse" data-act="go-class-browse">
        <div class="icon-wrap">${ICONS.nurse}</div>
        <h3>選擇班級瀏覽</h3>
        <p>從班級名單裡找自己</p>
      </div>
    </div>
    <div class="btn-row">
      <button class="btn btn-ghost" data-act="go-home" style="width:100%;">上一步</button>
    </div>
  </div>`;
}

/* ---- STUDENT ID SEARCH（獨立畫面） ---- */
function studentIdSearch(){
  const q = (state.search||'').trim();
  const matches = q ? state.roster.filter(r => r.id.includes(q) || r.name.includes(q)) : [];
  return `
  ${stepperBlock(1)}
  <div class="card">
    <h1 class="title">輸入學號查詢</h1>
    <p class="subtitle">輸入學號或姓名，找到自己後點選即可</p>
    <input class="search-input" id="search-box" placeholder="輸入學號或姓名查詢" value="${q}">
    ${
      !q
      ? `<div class="empty-note">請輸入學號或姓名開始查詢</div>`
      : matches.length === 0
        ? `<div class="empty-note">找不到符合的學生</div>`
        : `<div class="student-list">
            ${matches.slice(0,30).map(r=>`
              <div class="student-row ${state.student.id===r.id?'selected':''}" data-act="pick-student" data-id="${r.id}" data-name="${r.name}" data-class="${r.class}">
                <span class="sname">${r.name}</span>
                <span class="sid">${r.id}・${r.class}</span>
              </div>`).join('')}
          </div>`
    }
    <div class="btn-row">
      <button class="btn btn-ghost" data-act="back-select">上一步</button>
      <button class="btn btn-primary" data-act="to-step3" ${state.student.id?'':'disabled'}>下一步</button>
    </div>
  </div>`;
}

/* ---- STUDENT CLASS BROWSE（獨立畫面） ---- */
function studentClassBrowse(){
  const grade = state.student.grade;
  const q = (state.search||'').trim();
  let listBlock = '';
  if(grade){
    const list = state.roster.filter(r => r.class === grade);
    const filtered = q ? list.filter(r => r.id.includes(q) || r.name.includes(q)) : list;
    listBlock = `
      <p class="subtitle" style="margin:14px 0 10px;">${grade}　共 ${list.length} 位學生</p>
      ${list.length ? `<input class="search-input" id="search-box" placeholder="輸入學號或姓名搜尋" value="${q}">` : ''}
      ${
        list.length === 0
        ? `<div class="empty-note">此班級尚未匯入學生名冊。<br>請聯繫護理人員以 Excel 匯入名冊後再試一次。</div>`
        : `<div class="student-list">
            ${filtered.map(r=>`
              <div class="student-row ${state.student.id===r.id?'selected':''}" data-act="pick-student" data-id="${r.id}" data-name="${r.name}" data-class="${r.class}">
                <span class="sname">${r.name}</span>
                <span class="sid">${r.id}・${r.class}</span>
              </div>`).join('') || `<div class="empty-note">找不到符合的學生</div>`}
          </div>`
      }`;
  }
  return `
  ${stepperBlock(1)}
  <div class="card">
    <h1 class="title">選擇班級瀏覽</h1>
    <p class="subtitle">請先選擇班級</p>
    <div class="choice-grid">
      ${CLASSES.map(g=>`<div class="choice-btn ${grade===g?'selected':''}" data-act="pick-grade" data-val="${g}">${g}</div>`).join('')}
    </div>
    ${listBlock}
    <div class="btn-row">
      <button class="btn btn-ghost" data-act="back-select">上一步</button>
      <button class="btn btn-primary" data-act="to-step3" ${state.student.id?'':'disabled'}>下一步</button>
    </div>
  </div>`;
}

/* ---- STUDENT STEP 3: 性別 ---- */
function studentStep3(){
  const g = state.student.gender;
  return `
  ${stepperBlock(2)}
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
  ${stepperBlock(3)}
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
  ${stepperBlock(4)}
  <div class="card">
    <h1 class="title">${s.reason}詳細原因</h1>
    <p class="subtitle">請選擇最符合的項目</p>
    <div class="choice-grid cols-3">
      ${options.map(o=>`<div class="choice-btn ${d===o?'selected':''}" data-act="pick-detail" data-val="${o}">${o}</div>`).join('')}
    </div>
    <div class="summary-box" style="margin-top:22px;">
      <div class="summary-row"><span class="k">班級</span><span class="v">${s.class}</span></div>
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
  const labels = ['學生','性別','原因','詳細原因'];
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
    <p class="subtitle" style="margin-bottom:14px;">目前名冊共 ${state.rosterFull.length} 筆。匯入 Excel 需包含欄位：<b>學號</b>、<b>姓名</b>、<b>班級</b>（例如：京劇學系），可選填 <b>病史</b>。</p>
    <div class="dash-toolbar">
      <button class="btn btn-primary import-btn" ${state.rosterImportBusy?'disabled':''}>${state.rosterImportBusy ? '匯入中…' : `${ICONS.upload} 匯入 Excel 名冊`}
        <input type="file" id="roster-file" accept=".xlsx,.xls,.csv" ${state.rosterImportBusy?'disabled':''}>
      </button>
      ${state.rosterFull.length ? `<button class="btn btn-ghost" data-act="clear-roster" ${state.rosterImportBusy?'disabled':''}>清空名冊</button>` : ''}
    </div>
  </div>

  <div class="card">
    <h1 class="title">學生個案搜尋</h1>
    <p class="subtitle" style="margin-bottom:14px;">可用姓名、學號、班級，或直接搜病史內容（例如「氣喘」）找出相關學生。</p>
    <div class="dash-toolbar">
      <input class="search-input" id="case-search" placeholder="搜尋姓名、學號、班級或病史" value="${q}">
      <button class="pill-btn" data-act="export-roster-excel" ${state.rosterExportBusy?'disabled':''}>${state.rosterExportBusy ? '匯出中…' : `${ICONS.download} 匯出名單 Excel${q ? '（搜尋結果）' : ''}`}</button>
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
      <button class="pill-btn" data-act="export-excel" ${state.recordsExportBusy?'disabled':''}>${state.recordsExportBusy ? '匯出中…' : `${ICONS.download} 匯出 Excel`}</button>
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
    case 'go-brand-home':
      if(state.loggedIn){
        state.screen = 'nurse-dashboard';
      } else {
        state.screen = 'home';
        state.student = { method:null, grade:null, id:null, name:null, class:null, gender:null, reason:null, detail:null };
        state.search = '';
      }
      render(); break;

    case 'retry-load':
      state.loading = true; render();
      if(state.screen === 'nurse-case-management' || state.screen === 'nurse-health-check'){
        await loadRosterFull();
      } else if(state.screen === 'nurse-dashboard' || state.screen === 'nurse-record-treatment'){
        await loadRecords();
      } else {
        await loadRoster();
      }
      state.loading = false;
      render();
      break;

    case 'go-home':
      state.screen = 'home';
      state.student = { method:null, grade:null, id:null, name:null, class:null, gender:null, reason:null, detail:null };
      state.search = '';
      render(); break;

    case 'start-student':
      state.screen = 'student-select'; render(); break;

    case 'go-id-search':
      state.student.grade = null;
      state.student.id = null; state.student.name = null; state.student.class = null;
      state.search = '';
      state.student.method = 'id';
      state.screen = 'student-id';
      render(); break;

    case 'go-class-browse':
      state.student.grade = null;
      state.student.id = null; state.student.name = null; state.student.class = null;
      state.search = '';
      state.student.method = 'class';
      state.screen = 'student-class';
      render(); break;

    case 'back-select':
      state.screen = 'student-select'; render(); break;

    case 'pick-grade':
      state.student.grade = el.dataset.val;
      state.student.id = null; state.student.name = null; state.student.class = null;
      state.search = '';
      render(); break;

    case 'pick-student':
      state.student.id = el.dataset.id;
      state.student.name = el.dataset.name;
      state.student.class = el.dataset.class;
      render(); break;
    case 'to-step3':
      state.screen = 'student-3'; render(); break;
    case 'back-step2':
      state.screen = state.student.method === 'class' ? 'student-class' : 'student-id';
      render(); break;

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
      await exportRecordsToExcel(); break;

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
    class: s.class,
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
  state.rosterExportBusy = true; render();
  try{
    await ensureExcelJS();
  }catch(err){
    console.error(err);
    showToast('匯出功能載入失敗，請檢查網路連線後再試一次');
    state.rosterExportBusy = false; render();
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
  }finally{
    state.rosterExportBusy = false; render();
  }
}

/* ---------------- 匯出 Excel ---------------- */
async function exportRecordsToExcel(){
  const rangeRecords = filterByDateRange(state.records, state.dashFrom, state.dashTo);
  const q = (state.dashSearch||'').trim();
  let list = [...rangeRecords].sort((a,b)=> a.ts - b.ts);
  if(q) list = list.filter(r => r.name.includes(q) || r.id.includes(q) || r.class.includes(q));

  if(list.length === 0){
    showToast('目前區間內沒有資料可以匯出');
    return;
  }

  state.recordsExportBusy = true; render();
  try{
    await ensureXLSX();

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
  }catch(err){
    console.error(err);
    showToast('匯出失敗，請檢查網路連線後再試一次');
  }finally{
    state.recordsExportBusy = false; render();
  }
}

async function handleRosterFile(e){
  const file = e.target.files[0];
  if(!file) return;
  state.rosterImportBusy = true; render();
  try{
    await ensureXLSX();
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
      showToast(`已匯入 ${parsed.length} 筆學生資料`);
    } else {
      showToast('匯入失敗，請稍後再試一次');
    }
  }catch(err){
    console.error(err);
    showToast('匯入失敗，請確認檔案格式正確，或網路連線是否正常');
  }finally{
    state.rosterImportBusy = false; render();
    e.target.value = '';
  }
}

init();
