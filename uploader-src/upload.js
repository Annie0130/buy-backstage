// 飛鼠電商 自動上架腳本
// 用法：
//   node upload.js airtable          從收購紀錄（Airtable）讀「售價有填」的商品來上架
//   node upload.js airtable --list   只列出會上架哪些商品，不開瀏覽器（先看看用）
//   node upload.js 商品.xlsx          從 Excel 讀（舊方式，仍可用）
//   node upload.js shopee            蝦皮：保存狀況二手、NCC、15000 以上最低購買數量 3
//   node upload.js todo              待辦清單：哪些還沒上架／還是草稿／蝦皮還沒補
//
// 流程：開啟瀏覽器 → 你自己登入飛鼠後台 → 回終端機按 Enter →
//       腳本逐筆點「新增商品」、填表 → 依 mode 決定是否自動儲存

const { chromium } = require('playwright');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');
const readline = require('readline');

// 開啟 Chrome；同一份登入資料已經有 Chrome 在用時，給中文提示
// 錯誤截圖、欄位診斷這類檔案集中放在「診斷檔」資料夾，資料夾才不會亂
const DIAG_RE = /(錯誤截圖.*\.png|欄位診斷\.txt|欄位截圖\.png|圖片欄位.*|平台分類選項\.txt)$/;
function diagDir() { const d = path.join(__dirname, '診斷檔'); fs.mkdirSync(d, { recursive: true }); return d; }
function tidyDiagFiles() {
  for (const dir of [__dirname, path.dirname(__dirname)]) {
    try {
      for (const f of fs.readdirSync(dir)) if (DIAG_RE.test(f)) { try { fs.renameSync(path.join(dir, f), path.join(diagDir(), f)); } catch (e) {} }
    } catch (e) {}
  }
}
if (require.main === module) tidyDiagFiles();

// 舊的資料夾排法（程式都在最外層）→ 先整理資料夾，再從「系統」裡重新執行
if (require.main === module && fs.existsSync(path.join(__dirname, 'migrate.js'))) {
  const M = require('./migrate.js');
  if (M.needsMigration(__dirname) && M.migrate(__dirname)) {
    const r = require('child_process').spawnSync(process.execPath, [path.join(__dirname, M.SYS, 'upload.js'), ...process.argv.slice(2)], { stdio: 'inherit', cwd: path.join(__dirname, M.SYS) });
    process.exit(r.status || 0);
  }
}
// 舊版把登入資料放在程式資料夾的 chrome-profile：第一次自動搬到這台電腦自己的位置（不用重新登入）
function migrateProfile() {
  const old = path.join(__dirname, 'chrome-profile');
  if (fs.existsSync(CONFIG.profileDir) || !fs.existsSync(old)) return;
  const owner = fs.existsSync(path.join(old, '平台.txt')) ? fs.readFileSync(path.join(old, '平台.txt'), 'utf8').trim() : 'win32';
  if (owner !== process.platform) return; // 別台電腦的登入資料不能用
  try {
    console.log('（第一次：把登入資料搬到這台電腦，約 10～30 秒…）');
    fs.mkdirSync(path.dirname(CONFIG.profileDir), { recursive: true });
    fs.cpSync(old, CONFIG.profileDir, { recursive: true });
    fs.renameSync(old, old + '（舊，已搬走，可以刪）');
  } catch (e) { console.log('   登入資料沒搬成功，等一下請重新登入一次'); }
}
async function launchBrowser(opts) {
  migrateProfile();
  fs.mkdirSync(CONFIG.profileDir, { recursive: true });
  try {
    return await chromium.launchPersistentContext(CONFIG.profileDir, opts);
  } catch (e) {
    if (/has been closed|ProcessSingleton|user data directory is already in use/i.test(e.message)) {
      throw new Error('上一次腳本開的 Chrome 還開著（頂端寫「Chrome 目前受到自動測試軟體控制」），請把那個 Chrome 視窗關掉，再重新執行一次。');
    }
    throw e;
  }
}

// 有些網路的 IPv6 會讓連線卡住：一律先用 IPv4，並拉長連線等待時間
require('dns').setDefaultResultOrder('ipv4first');
try { require('net').setDefaultAutoSelectFamilyAttemptTimeout(2000); } catch (e) {}

// 網路時好時壞時自動重試（最多 5 次，每次多等一點）
async function fetchRetry(url, opts, tries = 5) {
  for (let i = 1; ; i++) {
    try {
      return await fetch(url, opts);
    } catch (e) {
      if (i >= tries) throw new Error(`連不到 ${new URL(url).host}（已重試 ${tries} 次）。請檢查網路，或改用手機熱點再試。`);
      console.log(`   …連線失敗，${i * 3} 秒後重試（第 ${i} 次）`);
      await new Promise((r) => setTimeout(r, i * 3000));
    }
  }
}

// ======================= 設定區（可自行修改） =======================
const CONFIG = {
  // 飛鼠後台網址（登入頁或首頁都可以）。留空的話就自己在瀏覽器輸入網址。
  startUrl: '',

  // review = 只填表，你檢查後自己按「儲存草稿 / 建立」，再回終端機按 Enter（建議先用這個）
  // draft  = 填完自動按「儲存草稿」
  // create = 填完自動按「建立」（會直接上架並同步到有開啟的平台）
  mode: 'review',

  // 一次最多處理幾筆（0 = 全部）。第一次測試建議設 1。
  limit: 0,

  // 先放的商品圖片（「商品拍攝中」佔位圖）。目前關閉 = ''（不上傳）
  // 想打開的話改成：path.join(__dirname, '拍攝中.jpg')
  placeholderImage: '',

  // 要寫進「商品屬性」表格的欄位。空的 = 不填（這些資訊商品描述裡都有了）
  // 想填的話改成：['公/水', '快門數', '盒單']
  attributeColumns: [],

  // 同步資訊裡要自動打開同步開關的平台（飛鼠官網、FB商店預設就是開的）
  syncPlatforms: ['蝦皮', 'Y拍'],

  // NCC / BSMI 字號對照表（同資料夾的 NCC字號.csv，用 Excel 或記事本編輯）
  nccFile: path.join(__dirname, 'NCC字號.csv'),

  // ---- 蝦皮同步（同步資訊區塊的蝦皮那一列）----
  shopee: {
    enable: true,          // 自動填蝦皮自訂價格
    markup: 1.18,          // 蝦皮抽成：售價 × 1.18
    roundTo: 100,          // 無條件進位到百位（7,080 → 7,100）
    highPrice: 15000,      // 售價 ≥ 這個金額：自訂價格留空（用原價），並提醒到蝦皮把最低購買數量改 3
    autoAfterListing: false, // 上架後要不要自動跑蝦皮補資料（false＝不跑，等你確認蝦皮都出現後自己按）
  },

  // 供貨商要寫進「備註 (管理人員使用)」嗎？（備註不會公開給買家）
  supplierToNote: true,

  // 用你電腦上的 Chrome；沒裝 Chrome 的話改成 undefined 會用 Playwright 內建瀏覽器
  browserChannel: 'chrome',

  // 登入狀態會存在這個資料夾，下次執行不用重新登入
  // 登入資料放在這台電腦自己的硬碟（Mac、Windows 各一份；共享資料夾裡放 Chrome 會很慢又容易壞）
  profileDir: path.join(require('os').homedir(), '.linscam-uploader', `chrome-profile-${process.platform}`),

  // 上架紀錄：成功的會記錄下來，重跑時自動跳過，不會重複上架
  logFile: path.join(__dirname, '上架紀錄.csv'),

  // ---- Airtable（收購紀錄）----
  airtable: {
    baseId: 'appAuc0SgpUMOFLsv', // 林相攝影收購紀錄
    tableId: 'tblnOp2xI0RPOAR3J',
    // 唯讀金鑰放在同資料夾的 airtable-token.txt（不要分享給別人）
    tokenFile: path.join(__dirname, 'airtable-token.txt'),
    // 只抓這天（含）之後的收購紀錄，留空 = 不限。例如 '2026-11-01'
    sinceDate: '',
    // true = 只抓在收購後台按過「🛒 標記上架飛鼠」的商品（送修的不要標記就不會上架）
    requireMark: true,
  },

  // Airtable 的「商品介紹」（型號規格介紹）要放哪裡：
  // 'description' = 商品描述內文單獨一行（【型號介紹】）、'feature1' = 商品特色 1、'none' = 不放
  introTo: 'description',
};
// ===================================================================

const EMPTY_MARKS = new Set(['', '—', '-', '－', '–', 'N/A', 'n/a']);

function cellText(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'object') {
    if (v.richText) return v.richText.map((t) => t.text).join('');
    if ('result' in v) return cellText(v.result); // 公式
    if (v.text) return String(v.text); // 超連結
    if (v instanceof Date) return v.toISOString().slice(0, 10);
  }
  if (typeof v === 'number') return Number.isInteger(v) ? String(v) : String(+v.toFixed(4));
  return String(v).trim();
}

function normHeader(h) {
  return h.replace(/[（(]\s*必填\s*[)）]/g, '').replace(/\s+/g, '').trim();
}

async function readRows(file) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(file);
  const ws = wb.worksheets[0];
  const headers = [];
  ws.getRow(1).eachCell({ includeEmpty: true }, (c, col) => {
    headers[col] = normHeader(cellText(c.value));
  });
  const rows = [];
  ws.eachRow((row, n) => {
    if (n === 1) return;
    const r = { _row: n };
    headers.forEach((h, col) => {
      if (h) r[h] = cellText(row.getCell(col).value);
    });
    if (r['商品名稱']) rows.push(r);
  });
  return rows;
}

// 用電腦預設的程式打開檔案／資料夾（Windows：start；Mac：open）
function openPath(p) {
  try {
    const { exec } = require('child_process');
    if (process.platform === 'win32') exec(`start "" "${p}"`);
    else if (process.platform === 'darwin') exec(`open "${p.replace(/"/g, '\\"')}"`);
  } catch (e) {}
}

// ---------- Airtable 讀取 ----------
// 只讀商品相關欄位，不讀客戶個資
const AT_FIELDS = ['器材一', '序號', '收購總金額', '售價', '貨品來源', '盒單', '配件清單',
  '新舊程度', '原廠保固', '功能備註', '快門數', '商品介紹', '交易日期', '品類', 'CCD／Y2K', '記憶卡類型', '選單語言', '飛鼠狀態', '進貨管道', '店保'];

async function readAirtable(formulaOverride) {
  const { baseId, tableId, tokenFile, sinceDate } = CONFIG.airtable;
  const token = (process.env.AIRTABLE_TOKEN || (fs.existsSync(tokenFile) ? fs.readFileSync(tokenFile, 'utf8') : '')).trim();
  if (!token) {
    console.log(`找不到 Airtable 金鑰。請把金鑰貼進 ${tokenFile}（見使用說明）`);
    process.exit(1);
  }
  const conds = ["{器材一}!=''"];
  if (formulaOverride) conds.push(formulaOverride);
  else {
    if (CONFIG.airtable.requireMark) conds.push('{上架飛鼠}');
    if (sinceDate) conds.push(`IS_AFTER({交易日期}, DATEADD('${sinceDate}', -1, 'days'))`);
  }
  const records = [];
  let offset = '';
  do {
    const qs = new URLSearchParams({ filterByFormula: `AND(${conds.join(',')})`, pageSize: '100' });
    AT_FIELDS.forEach((f) => qs.append('fields[]', f));
    qs.append('sort[0][field]', '交易日期');
    qs.append('sort[0][direction]', 'asc');
    if (offset) qs.set('offset', offset);
    const res = await fetchRetry(`https://api.airtable.com/v0/${baseId}/${tableId}?${qs}`, {
      headers: { Authorization: `Bearer ${token}` },
    });
    if (!res.ok) {
      console.log(`Airtable 讀取失敗（${res.status}）：${await res.text()}`);
      process.exit(1);
    }
    const data = await res.json();
    records.push(...data.records);
    offset = data.offset || '';
  } while (offset);
  return records.map((rec, i) => Object.assign(airtableToRow(rec.fields, `AT#${i + 1}`), { _id: rec.id, _fsStatus: String(rec.fields['飛鼠狀態'] || '') }));
}

// 配件清單整理：充電器、充電線 → 合併成一項「充電設備」（有原廠字樣就保留「原廠」）
// 想再合併其他配件，照格式加進 ACC_MERGE 就好
const ACC_MERGE = [
  { match: /充電(器|線)/, to: '充電設備' },
];
function normalizeAccessories(text) {
  if (!text || EMPTY_MARKS.has(text.trim())) return text;
  const items = text.split(/[、，,／\n]+/).map((x) => x.trim()).filter(Boolean);
  const out = [];
  const placed = new Map(); // 合併後名稱 → 在 out 的位置
  for (const it of items) {
    const rule = ACC_MERGE.find((r) => r.match.test(it));
    if (!rule) { out.push(it); continue; }
    const orig = /原廠/.test(it);
    if (placed.has(rule.to)) {
      const i = placed.get(rule.to);
      if (orig && !out[i].startsWith('原廠')) out[i] = '原廠' + rule.to;
    } else {
      placed.set(rule.to, out.length);
      out.push((orig ? '原廠' : '') + rule.to);
    }
  }
  return out.join('、');
}

// 簡單的 CSV 一行解析（支援雙引號包起來、裡面有逗號的欄位）
function parseCsvLine(line) {
  const out = []; let cur = '', q = false;
  for (let i = 0; i < line.length; i++) {
    const c = line[i];
    if (q) { if (c === '"' && line[i + 1] === '"') { cur += '"'; i++; } else if (c === '"') q = false; else cur += c; }
    else if (c === '"') q = true;
    else if (c === ',') { out.push(cur); cur = ''; }
    else cur += c;
  }
  out.push(cur);
  return out.map((x) => x.trim());
}
function readCsv(file) {
  if (!fs.existsSync(file)) return [];
  return fs.readFileSync(file, 'utf8').replace(/^\uFEFF/, '').split(/\r?\n/).slice(1).filter((l) => l.trim()).map(parseCsvLine);
}

// ---------- NCC / BSMI 字號對照表 ----------
// 檔案格式（第一行是標題）：型號,NCC字號,BSMI字號
// 型號只要寫商品名稱裡會出現的關鍵字，例如「A6600」就能對到「Sony A6600」；大小寫、空白不影響
let NCC_TABLE = null;
const normKey = (t) => String(t || '').toLowerCase().replace(/[\s\-_]/g, '');
function loadNccTable() {
  if (NCC_TABLE) return NCC_TABLE;
  NCC_TABLE = [];
  const f = CONFIG.nccFile;
  // 簡單的 CSV 解析（支援用雙引號包起來、裡面有逗號的欄位）
  const parse = (line) => {
    const out = []; let cur = '', q = false;
    for (let i = 0; i < line.length; i++) {
      const c = line[i];
      if (q) { if (c === '"' && line[i + 1] === '"') { cur += '"'; i++; } else if (c === '"') q = false; else cur += c; }
      else if (c === '"') q = true;
      else if (c === ',') { out.push(cur); cur = ''; }
      else cur += c;
    }
    out.push(cur);
    return out.map((x) => x.trim());
  };
  // 優先順序：你自己填的 NCC字號.csv ＞「NCC官方資料.txt」（官方公開資料完全相符）＞ 從官網舊商品頁抓來的（備註「官網有填」）
  const read = (file) => (fs.existsSync(file) ? fs.readFileSync(file, 'utf8').replace(/^\uFEFF/, '').split(/\r?\n/).slice(1).map(parse).filter((c) => c[0]) : []);
  const mine = read(f), official = read(path.join(__dirname, 'NCC官方資料.txt'));
  const crawled = mine.filter((c) => /官網有填/.test(c[3] || ''));
  const seen = new Set();
  for (const [model, ncc, bsmi, note = ''] of [...mine.filter((c) => !crawled.includes(c)), ...official, ...crawled]) {
    // 備註有「⚠️」或「同型號有多個」＝還沒確認，先不自動帶入（確認後把備註清掉就會開始用）
    if (/⚠️|同型號有多個/.test(note)) continue;
    const key = normKey(model);
    if (seen.has(key)) continue;
    seen.add(key);
    NCC_TABLE.push({ key, model, ncc, bsmi });
  }
  NCC_TABLE.sort((a, b) => b.key.length - a.key.length); // 長的關鍵字優先（A7M4 比 A7 優先）
  return NCC_TABLE;
}
function lookupNcc(name) {
  const n = normKey(name);
  // 型號後面不能緊接英數字（避免「Canon M6」對到「Canon M6II」、「Sony A1」對到「Sony A100」）
  return loadNccTable().find((e) => { const i = n.indexOf(e.key); return i >= 0 && !/[a-z0-9]/.test(n[i + e.key.length] || ''); }) || null;
}
// 對照表 NCC 欄填「無」＝這個型號沒有 Wi-Fi／藍牙等無線功能（你確認過的）
const NO_RADIO_TEXT = '本商品無 Wi-Fi／藍牙等無線傳輸功能，非電信管制射頻器材，無須 NCC 審驗';
const isNoRadio = (e) => !!e && /^(無|無無線功能|免)$/.test((e.ncc || '').trim());

// 文案的 BSMI／NCC 兩行：有字號才寫，查不到就整行不寫（不讓客人看到「請自行填入」）
// 公司貨（台灣代理商進口）機身標籤上是「該品牌代理商」的 BSMI 登錄碼；水貨／不影響不套用
const BRAND_BSMI = {
  Canon: 'R31219', Sony: 'R33021', Fujifilm: 'R33950', Nikon: 'R33697', Olympus: 'R33235', 'OM System': 'R33235',
  Panasonic: 'R31221', Leica: 'R3B375', Ricoh: 'R33B28', Pentax: 'R33B28', Epson: 'R33126',
};
function bsmiFor(name, origin) {
  const e = lookupNcc(name);
  if (e && e.bsmi) return e.bsmi;
  if (String(origin || '').trim() !== '公司貨') return '';
  const brand = guessBrand(name) || (/^(om\s*system|om-)/i.test(name) ? 'OM System' : '') || (/^lumix/i.test(name) ? 'Panasonic' : '');
  return BRAND_BSMI[brand] || '';
}
function certLines(name, ccd = false, origin = '') {
  const e = lookupNcc(name) || {};
  const lines = [];
  const bsmi = bsmiFor(name, origin);
  if (bsmi) lines.push(`【BSMI商檢字號】${bsmi}`);
  if (ccd) return lines.length ? lines.join('\n') + '\n' : ''; // CCD 相機只寫 BSMI，不寫 NCC
  if (isNoRadio(e)) lines.push(`NCC：${NO_RADIO_TEXT}`);
  else if (e.ncc) lines.push(`NCC：${e.ncc}`);
  return lines.length ? lines.join('\n') + '\n' : '';
}

// 判斷是否鏡頭（跟收購後台同一套規則）
// 名稱有「+」是機身＋鏡頭套組，算機身
const isLensName = (name) => !/\+/.test(name) && /\dmm|\dF\d|F\d\.\d/.test(name);

// 上架文案：照搬收購後台「📝 文案」的格式
function buildListingText(f) {
  const name = f['器材一'] || '';
  const serial = f['序號'] || '';
  const origin = f['貨品來源'] || '';
  const boxDoc = f['盒單'] || '';
  const acc = normalizeAccessories(f['配件清單'] || '—');
  const cond = f['新舊程度'] || '';
  const warranty = f['原廠保固'] || '';
  const fnNote = (f['功能備註'] || '').trim();
  const isLens = isLensName(name);

  const shutterRaw = f['快門數'] || '';
  const m = shutterRaw.match(/(\d+)次/) || shutterRaw.match(/^(\d+)$/);
  const shutter = m ? m[1] : '';

  const condMap = { 全新: '100%新', 近無使用痕跡: '90%新以上', 些許使用痕跡: '80%新', 使用痕跡多: '75%新' };
  const condPct = condMap[cond] || cond || '—';

  // 原廠在保 → 寫原廠保固；不然看「店保」（空白＝半年）
  const shopW = f['店保'] || '';
  const warrantyStr = warranty && warranty.startsWith('在保')
    ? `原廠保固${warranty.replace('在保', '').trim()}`
    : shopW.startsWith('7天') ? '商品售出 7 天保固'
    : shopW.startsWith('無') ? '本商品無店家保固'
    : '商品售出半年保固';

  // 不寫來源的情況：CCD／Y2K 相機、或收購時貨品來源選「不影響」；其他照寫公司貨／水貨
  const isRetro = !!(f['CCD／Y2K'] && (f['CCD／Y2K'].name || f['CCD／Y2K']));
  const originTxt = isRetro || origin === '不影響' ? '' : origin;
  const accDesc = `${originTxt}${boxDoc ? (originTxt ? ' ' : '') + boxDoc : ''}${originTxt || boxDoc ? '　' : ''}${isLens ? '鏡頭' : '機身'}：${acc}`;

  const hasNote = fnNote && fnNote !== '全數正常';
  let fnDesc;
  if (isLens) {
    const base = '鏡頭前後玉乾淨無刮傷、無入塵';
    fnDesc = (hasNote ? `${base}，${fnNote}` : base) + '　自動對焦功能正常';
  } else {
    fnDesc = hasNote ? `狀況經檢測功能正常　${fnNote}` : '狀況經檢測功能正常';
  }

  return `⚠️ 商品與現場同步販售，庫存依現場為準，所以請一定要詢問! 謝謝
☎️客服電話：02-29303697
🏠門市營業時間14-21點(二三公休)
台北市文山區羅斯福路六段5號
--------------------------------------
【商品貨號】${serial}
【名稱】${name}
${CONFIG.introTo === 'description' && (f['商品介紹'] || '').trim() ? `【型號介紹】${f['商品介紹'].trim()}\n` : ''}【品相描述】${fnDesc}${shutter && !isLens ? '　快門數' + shutter + '次' : ''}
【新舊程度】${condPct}
【保固狀態】${warrantyStr}
【配件說明】${accDesc}
${(f['記憶卡類型'] || '').trim() ? `【記憶卡】${f['記憶卡類型'].trim()}\n` : ''}${(f['選單語言'] || '').trim() ? `【選單語言】${f['選單語言'].trim()}\n` : ''}${certLines(name, /CCD/.test(String((f['CCD／Y2K'] && f['CCD／Y2K'].name) || f['CCD／Y2K'] || '')), f['貨品來源'])}--------------------------------------
【商品狀態分級說明】
● 100%新 = 全新未拆封新品
●   90%新 = 機況近全新、使用痕跡不明顯
●   85%新 = 僅些許使用小痕跡，不影響心情
●   80%新 = 外觀較重的使用痕跡，會特別標注位置`;
}

function airtableToRow(f, tag) {
  const str = (v) => (v === undefined || v === null ? '' : String(v).trim());
  const shutterRaw = str(f['快門數']);
  const sm = shutterRaw.match(/(\d+)次/) || shutterRaw.match(/^(\d+)$/);
  return {
    _row: tag,
    商品名稱: str(f['器材一']),
    分類: '',
    條碼: str(f['序號']),
    庫存量: '1',
    進貨價: str(f['收購總金額']),
    銷售價: str(f['售價']),
    品牌: '',
    供貨商: str(f['進貨管道']) || '現場客人',
    '公/水': str(f['貨品來源']),
    快門數: sm && !isLensName(str(f['器材一'])) ? sm[1] : '',
    盒單: str(f['盒單']),
    商品描述: buildListingText(f),
    商品介紹: str(f['商品介紹']),
    新舊程度: str(f['新舊程度']),
    Y2K: str(f['CCD／Y2K'] && f['CCD／Y2K'].name ? f['CCD／Y2K'].name : f['CCD／Y2K']), // 'CCD 相機'／'Y2K 復古'／''
    快門次數: (str(f['快門數'] || '').match(/(\d{2,})/) || [])[1] || '',
    交易日期: str(f['交易日期']),
    品類: str(f['品類'] && f['品類'].name ? f['品類'].name : f['品類']),
  };
}

// ---------- 商品名稱（SEO 標題）----------
// 規則：官網賣過同型號 → 沿用舊標題（拿掉舊的狀況詞、換新序號）；沒賣過 → 照店裡格式產生
//   (賣點標籤) 型號 特色… [快門少] [極新品] 二手相機／二手鏡頭 [林相攝影] #序號
//   快門 ≤ 1 萬次才加「快門少」；特色從「商品介紹」抓；特色太少才加「林相攝影」
const TITLE_FILE = path.join(__dirname, '商品標題.csv');
const TITLE_MAX = 70;
let TITLE_TABLE = null;
function lookupOldTitle(name) {
  if (!TITLE_TABLE) TITLE_TABLE = readCsv(TITLE_FILE).filter((c) => c[0] && c[1]).map((c) => ({ key: normKey(c[0]), title: c[1] }));
  const n = normKey(name);
  const title = (TITLE_TABLE.find((e) => e.key === n) || {}).title || '';
  // 安全檢查：舊標題裡一定要有這個型號的代號（例如 XT30III、S90、24-70），不然當作沒有舊標題，重新產生
  const flat = (x) => String(x).toUpperCase().replace(/[^A-Z0-9]/g, '');
  const code = (String(name).match(/[A-Za-z]*\d+[A-Za-z0-9]*/g) || [])[0];
  if (title && code && !flat(title).includes(flat(code))) return '';
  return title;
}
const BRAND_ZH = { Sony: '索尼', Canon: '佳能', Nikon: '尼康', Fujifilm: '富士', Olympus: '奧林巴斯', Panasonic: 'Lumix', Leica: '徠卡',
  Ricoh: '理光', Sigma: '適馬', Tamron: '騰龍', Pentax: '賓得士', Hasselblad: '哈蘇', Viltrox: '唯卓', Samyang: '森養' };
// 商品介紹裡的關鍵字 → 標題用的特色詞（照店裡常用的說法）
const FEATURE_RULES = [
  [/防手震|防震|IBIS/i, '防手震'], [/[4-8]\.?\d?K/i, (m) => `${m[0].toUpperCase()}錄影`], [/翻轉|可翻|上掀|自拍/, '翻轉螢幕'],
  [/全片幅/, '全片幅'], [/(\d{4})\s*萬畫素/, (m) => (Number(m[1]) >= 3000 ? '高畫素' : '')], [/眼部|主體偵測|追焦|相位對焦|對焦(快|迅速|靈敏)/, '對焦快速'],
  [/雙卡/, '雙卡槽'], [/防滴|防塵/, '防滴防塵'], [/連拍/, '高速連拍'], [/Vlog|錄影|影音/i, 'Vlog'], [/底片模擬|濾鏡|色彩/, '濾鏡好看'],
  [/輕巧|輕便|僅\s*\d+g|約\s*\d{3}g|小巧/, '輕巧好攜帶'], [/恆定光圈|F2\.8 恆定/, '恆定光圈'], [/大光圈|F1\.[0-8]/i, '大光圈'],
  [/散景/, '散景美'], [/微距/, '微距'], [/星空/, '星空'], [/人像/, '人像'], [/風景/, '風景'], [/運動|鳥|生態|飛機/, '拍運動生態'],
  [/閃燈/, '有閃燈'], [/觸控/, '觸控螢幕'], [/新手|入門/, '新手適合'],
];
const CONDITION_WORDS = /快門少|極新品|全新品|近全新|快門數?\s*\d+|外觀漂亮|新品/g;
// 拍攝場景詞（優先放進標題；每件最多 SCENE_MAX 個）
const SCENE_MAX = 3;
const SCENES = {
  macro: ['微距', '拍昆蟲', '牙科攝影', '商品攝影', '花卉', '珠寶'],
  fisheye: ['魚眼', '極限運動', '創意攝影'],
  tiltshift: ['移軸', '建築'],
  superTele: ['追星', '拍飛機', '打鳥', '生態', '運動賽事', '演唱會'],
  tele: ['人像', '運動', '演唱會', '舞台'],
  travel: ['旅遊', '一鏡到底', '出國必備'],
  ultraWide: ['風景', '星空', '建築', '室內', '旅遊'],
  wide: ['風景', '街拍', '旅遊', 'Vlog'],
  standardZoom: ['日常', '旅遊', '街拍', '一鏡走天下'],
  portrait: ['人像', '寫真', '婚紗', '散景美'],
  normal: ['人像', '街拍', '日常', '美食'],
  pancake: ['餅乾鏡', '輕巧', '隨身'],
  fast: ['夜拍', '散景美', '低光'],
  vlogCam: ['Vlog', '自拍', '網紅'],
  actionCam: ['運動', '潛水', '騎車'],
  teleCam: ['追星', '拍飛機', '演唱會'],
  ccd: ['CCD', '復古', 'Y2K'],
  y2k: ['復古', 'Y2K', '日系'],
};
function lensInfo(name) {
  const m = name.match(/(\d+(?:\.\d+)?)(?:\s*-\s*(\d+(?:\.\d+)?))?\s*mm/i);
  const a = name.match(/(?:^|\s)F\s?(\d+(?:\.\d+)?)/i) || name.match(/\d\/(\d\.\d)/);
  return { minF: m ? Number(m[1]) : 0, maxF: m ? Number(m[2] || m[1]) : 0, zoom: !!(m && m[2]), aperture: a ? Number(a[1]) : 0 };
}
function sceneWords(r) {
  const name = r['商品名稱'] || '';
  const text = `${r['商品介紹'] || ''} ${name}`;
  const t = gearType(r);
  const keys = [];
  if (t === 'lens') {
    const info = lensInfo(name);
    // 依片幅換算等效焦段（M43 ×2、APS-C ×1.5），判斷廣角／標準／望遠比較準
    const cat = r['品類'] || '';
    const crop = /M43/.test(cat) ? 2 : /APS-C/.test(cat) ? 1.5 : 1;
    const minF = info.minF * crop, maxF = info.maxF * crop, { zoom, aperture } = info;
    if (/macro|marco|微距/i.test(text)) keys.push('macro');
    else if (/fisheye|魚眼/i.test(text) || (minF === 8 && maxF === 15)) keys.push('fisheye');
    else if (/ts-?e|tilt|移軸/i.test(text)) keys.push('tiltshift');
    else if (maxF >= 300) keys.push('superTele');
    else if (zoom && maxF / minF >= 5) keys.push('travel');
    else if (maxF >= 135 && zoom) keys.push('tele');
    else if (minF && minF <= 18) keys.push('ultraWide'); // 等效 18mm 以下算超廣角
    else if (!zoom && minF >= 56 && minF <= 135) keys.push('portrait');
    else if (zoom && minF <= 28 && maxF <= 105) keys.push('standardZoom');
    else if (minF && minF <= 28) keys.push('wide');
    else if (!zoom && minF >= 30 && minF <= 55) keys.push('normal');
    if (/\b(22mm F2|24mm F2\.8|40mm F2\.8|20mm F1\.7|16mm F2\.8|27mm F2\.8|15mm F1\.7)/i.test(name) || /餅乾/.test(text)) keys.push('pancake');
    if (aperture && aperture <= 1.8) keys.push('fast');
  } else {
    // 收購同意書勾的 CCD／Y2K 最優先（CCD 才會寫「CCD」）
    if (/CCD/.test(r['Y2K'] || '')) keys.push('ccd');
    else if (/Y2K/.test(r['Y2K'] || '')) keys.push('y2k');
    if (/\bzv|g7\s?x|vlog|x-?s20/i.test(text)) keys.push('vlogCam');
    if (/action|gopro|insta360|osmo/i.test(name)) keys.push('actionCam');
    if (/p900|p1000|sx\d{2,3}|fz\d{2,4}|rx10|zs\d{2}/i.test(name)) keys.push('teleCam');
  }
  // 主要類型放 2 個，其他類型各放 1 個，總共最多 SCENE_MAX 個
  const out = [];
  const add = (w) => { if (w && !out.includes(w) && out.length < SCENE_MAX) out.push(w); };
  if (keys[0]) SCENES[keys[0]].slice(0, keys.length > 1 ? 2 : SCENE_MAX).forEach(add);
  keys.slice(1).forEach((k) => add(SCENES[k].find((w) => !out.includes(w))));
  if (keys[0]) SCENES[keys[0]].forEach(add); // 還有空位就補主要類型的詞
  return out;
}

function titleFeatures(r) {
  const text = `${r['商品介紹'] || ''} ${r['商品名稱'] || ''}`;
  const out = [...sceneWords(r)];
  for (const [re, word] of FEATURE_RULES) {
    const m = text.match(re);
    if (!m) continue;
    const w = typeof word === 'function' ? word(m) : word;
    if (w && !out.includes(w)) out.push(w);
  }
  return out;
}
function conditionWords(r) {
  const w = [];
  const shots = Number(r['快門次數']);
  if (gearType(r) !== 'lens' && shots > 0 && shots <= 10000) w.push('快門少');
  if (/全新|近無使用痕跡/.test(r['新舊程度'] || '')) w.push('極新品');
  return w;
}
// 組標題，太長時依序拿掉：①林相攝影 ②最後面的特色詞；「快門少／極新品」和型號一定保留
function fitTitle({ head, feats = [], cond = [], used = '', shop = false }, serial) {
  const tail = serial ? ` #${serial}` : '';
  const f = [...feats];
  let withShop = shop;
  const make = () => [head, ...f, ...cond, used, withShop ? '林相攝影' : ''].filter(Boolean).join(' ').replace(/\s+/g, ' ').trim() + tail;
  while (make().length > TITLE_MAX) {
    if (withShop && f.length) withShop = false; // 有特色詞時，先拿掉店名
    else if (f.length) f.pop();
    else if (withShop) withShop = false;
    else break;
  }
  return make();
}
function buildTitle(r) {
  const name = r['商品名稱'] || '';
  const serial = r['條碼'] || '';
  const t = gearType(r);
  const used = t === 'lens' ? '二手鏡頭' : '二手相機';
  const cond = conditionWords(r);
  const old = lookupOldTitle(name);
  if (old) {
    // 沿用舊標題：拿掉舊的 #序號 和舊的狀況詞，補上這台的狀況
    let base = old.replace(/\s*#\s*[\w\-]*\s*$/, '').replace(CONDITION_WORDS, '').replace(/\s+/g, ' ').trim();
    const hadShop = /林相攝影/.test(base);
    base = base.replace(/\s*林相攝影\s*/g, ' ').trim();
    const extra = cond.filter((w) => !base.includes(w));
    return fitTitle({ head: base, cond: extra, shop: hadShop && base.split(" ").length < 4 }, serial); // 舊標題本來就有特色，通常不用再加店名
  }
  const brand = BRAND_ZH[r['品牌']] || '';
  const cat = r['品類'] || '';
  const fmt = /全片幅/.test(cat) ? '全幅' : /APS-C/.test(cat) ? 'APSC' : /M43/.test(cat) ? 'M43' : '';
  const tag = {
    lens: `${brand}${fmt}${/\d+-\d+\s*mm/i.test(name) ? '變焦鏡' : '定焦鏡'}`,
    mirrorless: `${brand}${fmt}無反`, dslr: `${brand}${fmt}單眼`, compact: `${brand}隨身機`,
  }[t] || `${brand}攝影器材`;
  const feats = titleFeatures(r).filter((w) => !(w === '全片幅' && fmt === '全幅'));
  const kit = /\+/.test(name) && t !== 'lens' ? '含鏡頭' : '';
  // 真的想不出特色（少於 2 個）才加店名
  return fitTitle({ head: `(${tag}) ${name}${kit ? ' ' + kit : ''}`, feats: feats.slice(0, 4), cond, used, shop: feats.length < 2 }, serial);
}

// 蝦皮商品名稱：最多 25 個字（超過會被截斷），一定要有 #序號
const SHOPEE_TITLE_MAX = 60;
const SPEC_TOKENS = /\b(Di|II|III|IV|VXD|RXD|VC|USD|SP|DC|DN|DG|HSM|OS|OSS|IS|STM|USM|VR|ED|ASPH|LM|WR|R|PZ|G2|Art|Contemporary|Sports|ZA|GM|G|L|S|Pro|PRO|Macro|-A|A\d{3})\b/g;
function shortModelVariants(name) {
  const n = name.replace(/\s+/g, ' ').trim();
  const noSpec = n.replace(SPEC_TOKENS, ' ').replace(/\s+/g, ' ').trim();
  const noMount = noSpec.replace(/\s*for\s+\w+(\s+\w)?$/i, '').trim();
  const noAperture = noMount.replace(/\s+F\d+(\.\d+)?(-\d+(\.\d+)?)?(?=\s|$)/i, '').trim(); // 只拿掉「F2.8」這種光圈，不會動到「EF 100mm」
  const kitless = noAperture.split('+')[0].trim();
  // 再拿掉接環代號（FE、EF、RF、XF、Z…），例如「Sony FE 200-600mm」→「Sony 200-600mm」
  const noPrefix = kitless.replace(/\s+(FE|EF-?S|EF-?M|EF|RF-?S|RF|XF|XC|Z|DX|AF-?S|AF|E|M\.?Zuiko|DG|DC)(?=\s|\d)/gi, ' ').replace(/\s+/g, ' ').trim();
  return [...new Set([n, noSpec, noMount, noAperture, kitless, noPrefix])];
}
// 蝦皮標題：型號＋最重要的詞放最前面（手機大約只顯示前 20 字），#序號放最後，最多 60 字
function buildShopeeTitle(r) {
  const serial = r['條碼'] || '';
  const tail = serial ? ` #${serial}` : '';
  const t = gearType(r);
  const used = t === 'lens' ? '二手鏡頭' : '二手相機';
  const first = t === 'lens' ? [...sceneWords(r), ...conditionWords(r)] : [...conditionWords(r), ...sceneWords(r)];
  const words = [...new Set([...first, ...titleFeatures(r), used])];
  for (const model of shortModelVariants(r['商品名稱'] || '')) {
    if ((model + tail).length > SHOPEE_TITLE_MAX) continue;
    let out = model;
    for (const w of words) if (!out.includes(w) && (`${out} ${w}${tail}`).length <= SHOPEE_TITLE_MAX) out += ` ${w}`;
    return out + tail;
  }
  const room = SHOPEE_TITLE_MAX - tail.length;
  return (r['商品名稱'] || '').slice(0, Math.max(room, 0)).trim() + tail;
}

// ---------- 自動判斷品牌 / 分類（Excel 留空時才補） ----------
// 分類清單取自試算表「記得不能刪掉」分頁，必須跟飛鼠後台的分類名稱一致
const CATEGORIES = new Set([
  'Canon 相機', 'Canon 鏡頭', 'Fujifilm 相機', 'Fujifilm 鏡頭', 'Leica 相機', 'Leica 鏡頭',
  'Nikon 相機', 'Nikon 鏡頭', 'Olympus 相機', 'Olympus 鏡頭', 'Panasonic 相機', 'Panasonic 鏡頭',
  'Pentax 相機', 'Ricoh 相機', 'Sony 相機', 'Sony 鏡頭', 'Sigma 相機', 'Sigma 鏡頭', 'Hasselblad', '配件類',
]);
const BRAND_RULES = [
  ['Hasselblad', /hasselblad|哈蘇/i],
  ['Sigma', /sigma|適馬/i], // 放在前面：Sigma 的「for Sony」鏡頭品牌是 Sigma（分類依接環）
  ['Canon', /canon|佳能/i],
  ['Fujifilm', /fuji|富士|\bx100|\bx-?t\d|\bx-?s\d|\bx-?e\d|\bx-?h\d|\bx-?pro|\bgfx/i],
  ['Leica', /leica|徠卡|萊卡/i],
  ['Nikon', /nikon|尼康|nikkor/i],
  ['Olympus', /olympus|奧林巴斯|\bom-?\d|\bom system|zuiko/i],
  ['Panasonic', /panasonic|lumix|國際牌/i],
  ['Pentax', /pentax|賓得/i],
  ['Ricoh', /ricoh|理光|\bgr\s?i{1,3}x?\b/i],
  ['Sony', /sony|索尼|\bα\d|\bilce|\bzv-?e?\d|\ba7[a-z]*\d*\b|\ba6\d{3}\b/i],
];
const LENS_RE = /\d+(\.\d+)?\s*-\s*\d+(\.\d+)?\s*mm|\d+(\.\d+)?\s*mm|f\s*\/?\s*\d+(\.\d+)?\b|鏡頭|\blens\b|macro|nikkor|zuiko/i;

// 副廠鏡頭品牌（不在飛鼠分類清單裡）：品牌照填，分類留空讓你手動選
const THIRD_PARTY = [['Tamron', /tamron|騰龍/i], ['Viltrox', /viltrox|唯卓/i], ['Samyang', /samyang|森養/i],
  ['Tokina', /tokina/i], ['Laowa', /laowa|老蛙/i], ['TTArtisan', /ttartisan|銘匠/i], ['7Artisans', /7\s?artisans|七工匠/i],
  ['Voigtlander', /voigtl[aä]nder|福倫達/i], ['Zeiss', /zeiss|蔡司/i], ['Meike', /meike|美科/i], ['Yongnuo', /yongnuo|永諾/i]];
const THIRD_PARTY_NAMES = new Set(THIRD_PARTY.map(([b]) => b));

function guessBrand(name) {
  for (const [brand, re] of THIRD_PARTY) if (re.test(name)) return brand;
  for (const [brand, re] of BRAND_RULES) if (re.test(name)) return brand;
  return '';
}
// 副廠鏡頭依接環放進原廠的「鏡頭」分類，例如 Tamron … For fuji → Fujifilm 鏡頭
// M43、L 接環有多家共用，判斷不出放哪，留給你手動選
const MOUNT_RULES = [
  ['Fujifilm', /\bfor\s*(fuji(film)?|x\b|gfx)|\b(fuji(film)?|x|gfx)[\s-]*mount|富士(接環|卡口|用)?/i],
  ['Sony', /\bfor\s*(sony|e\b|fe\b)|\b(sony|e|fe)[\s-]*mount|索尼(接環|卡口|用)?/i],
  ['Canon', /\bfor\s*(canon|ef\b|rf\b|ef-?m\b)|\b(canon|ef|rf|ef-?m)[\s-]*mount|佳能(接環|卡口|用)?/i],
  ['Nikon', /\bfor\s*(nikon|z\b|f\b)|\b(nikon|z|f)[\s-]*mount|尼康(接環|卡口|用)?/i],
];
function guessMount(name) {
  for (const [brand, re] of MOUNT_RULES) if (re.test(name)) return brand;
  return '';
}

function guessCategory(name, brand) {
  if (!brand) return '';
  if (THIRD_PARTY_NAMES.has(brand)) {
    const mount = guessMount(name);
    return mount && LENS_RE.test(name) ? `${mount} 鏡頭` : '';
  }
  if (brand === 'Hasselblad') return 'Hasselblad';
  // Sigma 鏡頭也依接環放（for Sony → Sony 鏡頭）；沒寫接環才放 Sigma 鏡頭
  if (brand === 'Sigma' && LENS_RE.test(name)) {
    const mount = guessMount(name);
    return mount ? `${mount} 鏡頭` : 'Sigma 鏡頭';
  }
  const cat = `${brand} ${LENS_RE.test(name) ? '鏡頭' : '相機'}`;
  return CATEGORIES.has(cat) ? cat : '';
}
function autoFill(r) {
  const notes = [];
  if (!r['品牌']) {
    const b = guessBrand(r['商品名稱']);
    if (b) { r['品牌'] = b; notes.push(`品牌→${b}`); }
  }
  if (!r['分類']) {
    const c = guessCategory(r['商品名稱'], r['品牌']);
    if (c) { r['分類'] = c; notes.push(`分類→${c}`); }
  }
  return notes;
}

function validate(r) {
  const missing = [];
  if (!r['商品名稱']) missing.push('商品名稱');
  if (!r['分類']) missing.push('分類');
  if (!r['條碼']) missing.push('條碼/序號');
  if (r['銷售價'] === '' || r['銷售價'] === '0') missing.push('銷售價');
  return missing;
}

// ---------- 紀錄檔 ----------
function keyOf(r) {
  return `${r['商品名稱']}|${r['條碼'] || ''}`;
}
function loadDone() {
  const done = new Set();
  if (!fs.existsSync(CONFIG.logFile)) return done;
  for (const line of fs.readFileSync(CONFIG.logFile, 'utf8').split(/\r?\n/).slice(1)) {
    const cols = line.split('","').map((s) => s.replace(/^"|"$/g, ''));
    if (cols[3] && cols[3].startsWith('成功')) done.add(`${cols[1]}|${cols[2]}`);
  }
  return done;
}

// ---------- 回寫收購後台「飛鼠狀態」----------
// 需要金鑰有 data.records:write 權限；只寫「飛鼠狀態」這一欄。沒權限就跳過，不影響上架
const AT_PENDING = [];
let AT_WRITE_OFF = false;
function atToken() {
  const f = CONFIG.airtable.tokenFile;
  return (process.env.AIRTABLE_TOKEN || (fs.existsSync(f) ? fs.readFileSync(f, 'utf8') : '')).trim();
}
function setFeishuStatus(serial, text, id) {
  if (AT_WRITE_OFF || !serial) return;
  const { baseId, tableId } = CONFIG.airtable;
  const today = new Date().toLocaleDateString('zh-TW');
  const job = (async () => {
    const token = atToken();
    if (!token) return;
    const H = { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' };
    let rid = id;
    if (!rid) {
      const qs = new URLSearchParams({ filterByFormula: `{序號}="${String(serial).replace(/"/g, '\\"')}"`, maxRecords: '2' });
      qs.append('fields[]', '序號');
      const res = await fetchRetry(`https://api.airtable.com/v0/${baseId}/${tableId}?${qs}`, { headers: H }, 3);
      const recs = res.ok ? (await res.json()).records : [];
      if (recs.length !== 1) return;
      rid = recs[0].id;
    }
    const res = await fetchRetry(`https://api.airtable.com/v0/${baseId}/${tableId}/${rid}`, { method: 'PATCH', headers: H, body: JSON.stringify({ fields: { 飛鼠狀態: `${text} ${today}`, ...(text.includes('蝦皮已補') ? { 上架飛鼠: false } : {}) } }) }, 3); // 全部完成 → 自動取消上架標記，清單才不會越來越長
    if (res.status === 401 || res.status === 403) {
      if (!AT_WRITE_OFF) console.log('   ℹ️ Airtable 金鑰沒有「寫入」權限，收購後台的狀態不會更新（上架不受影響；要開請看使用說明「收購後台顯示上架狀態」）');
      AT_WRITE_OFF = true;
    }
  })().catch(() => {});
  AT_PENDING.push(job);
}
const flushAirtable = () => Promise.all(AT_PENDING);

function log(r, status, note = '') {
  if (!fs.existsSync(CONFIG.logFile)) {
    fs.writeFileSync(CONFIG.logFile, '﻿"時間","商品名稱","條碼","結果","備註"\n');
  }
  const esc = (s) => String(s).replace(/"/g, '""');
  const t = new Date().toLocaleString('zh-TW', { hour12: false });
  fs.appendFileSync(
    CONFIG.logFile,
    [t, r['商品名稱'], r['條碼'] || '', status, note].map((s) => `"${esc(s)}"`).join(',') + '\n'
  );
  if (/^成功（已(建立|發佈)）/.test(status)) setFeishuStatus(r['條碼'], '飛鼠已上架', r._id);
  else if (/^成功/.test(status) && status.includes('草稿')) setFeishuStatus(r['條碼'], '飛鼠草稿', r._id);
}

// ---------- 終端機互動 ----------
// 整個執行過程只開一個輸入介面（Windows 上反覆開關會讓程式直接結束）
let RL = null;
function ask(q, kind = '') {
  if (!RL) RL = readline.createInterface({ input: process.stdin, output: process.stdout });
  // 從控制台執行時：先送一個「我在問問題」的訊號，控制台就會跳出回覆按鈕
  if (process.env.PANEL) console.log('__ASK__' + JSON.stringify({ q: q.trim(), kind }));
  return new Promise((res) => RL.question(q, (a) => res(a.trim())));
}
function closeAsk() { if (RL) { RL.close(); RL = null; } }

// ---------- 依「欄位標題文字」找輸入框 ----------
// 飛鼠表單是「標題在上、輸入框在下」，所以找到標題文字後，取它後面第一個輸入框。
const TEXT_INPUT = "input[not(@type='hidden') and not(@type='checkbox') and not(@type='radio') and not(@type='file') and not(@readonly) and not(contains(@class,'tt-hint'))]";

function labelXPath(label, { prefix = false, all = false } = {}) {
  // 比對時忽略空白和 *（「SKU 貨號 *」「SKU貨號」都算）
  const t = "translate(normalize-space(translate(text(),'*','')),' \u00a0','')";
  const want = label.replace(/\s+/g, '');
  const cond = prefix ? `starts-with(${t}, '${want}')` : `${t}='${want}'`;
  const xp = `//*[not(self::script) and not(self::style) and not(self::option)][${cond}]`;
  return all ? xp : `(${xp})[1]`;
}

// 頁面上可能有藏起來的同名欄位（例如多重規格的範本），只挑看得見的那一個
async function visibleFieldAfter(page, label, tag = TEXT_INPUT, opts = {}) {
  // 標題有 <label for="xxx"> 時，直接用 id 找到真正的輸入框
  const labels = page.locator(`xpath=${labelXPath(label, { ...opts, all: true })}`);
  for (let k = 0, n = await labels.count(); k < n; k++) {
    const id = await labels.nth(k).getAttribute('for').catch(() => null);
    if (!id) continue;
    const el = page.locator(`[id="${id}"]`).first();
    const ok = await el.evaluate((e) => !e.readOnly && !e.disabled && /^(input|textarea)$/i.test(e.tagName)).catch(() => false);
    if (ok && (await el.isVisible().catch(() => false))) return el;
  }
  const cands = page.locator(`xpath=${labelXPath(label, { ...opts, all: true })}/following::${tag}[1]`);
  const n = await cands.count();
  for (let i = 0; i < n; i++) {
    const c = cands.nth(i);
    if (await c.isVisible().catch(() => false)) return c;
  }
  return n ? cands.first() : null;
}

function fieldAfter(page, label, tag = TEXT_INPUT, opts) {
  return page.locator(`xpath=${labelXPath(label, opts)}/following::${tag}[1]`);
}

// 這一筆填過的欄位（最後會再讀一次，確認值沒有被清掉）
let FILLED = [];
// 填完補送 change、再離開欄位，有些網頁要這樣才會真的記住值
const commitInput = (el) => el.evaluate((e) => { e.dispatchEvent(new Event('change', { bubbles: true })); e.blur && e.blur(); }).catch(() => {});
// 找「定價」自己的輸入框：從「定價」標題往上找最小的區塊，裡面只能有這一個輸入框、而且不能包含「售價」
async function listPriceInput(page) {
  const ok = await page.evaluate(() => {
    document.querySelectorAll('[data-list-price]').forEach((x) => x.removeAttribute('data-list-price'));
    const vis = (el) => { const r = el.getBoundingClientRect(); return r.width > 0 && r.height > 0; };
    const heads = [...document.querySelectorAll('label, span, div, th, td, p')].filter((el) => {
      const own = [...el.childNodes].filter((n) => n.nodeType === 3).map((n) => n.textContent).join('').replace(/\s/g, '');
      return /^定價\*?/.test(own) && vis(el);
    });
    for (const h of heads) {
      let box = h.parentElement;
      for (let up = 0; up < 5 && box; up++, box = box.parentElement) {
        const inputs = [...box.querySelectorAll('input')].filter((i) => vis(i) && !i.readOnly && !/hidden|checkbox|radio|file/.test(i.type));
        if (!inputs.length) continue;
        if (inputs.length > 1 || /售價|成本/.test(box.textContent)) break; // 太大了，包到別的欄位
        inputs[0].setAttribute('data-list-price', '1');
        return true;
      }
    }
    return false;
  }).catch(() => false);
  return ok ? page.locator('[data-list-price="1"]').first() : null;
}
async function ensureListPrice(page, price, problems) {
  if (!price) return;
  try {
    const el = await listPriceInput(page);
    if (!el) throw new Error('not found');
    const cur = String(await el.inputValue().catch(() => '')).replace(/[^\d]/g, '');
    if (cur === String(price).replace(/[^\d]/g, '')) { console.log(`   💲 定價：${price}`); return; }
    await el.click({ timeout: 3000 }).catch(() => {});
    await el.fill(String(price), { timeout: 4000 });
    await commitInput(el);
    const after = String(await el.inputValue().catch(() => '')).replace(/[^\d]/g, '');
    if (after !== String(price).replace(/[^\d]/g, '')) throw new Error('not kept');
    console.log(`   💲 定價補填：${price}（跟售價一樣）`);
  } catch (e) {
    problems.push(`「定價」沒填成功，請手動填：${price}`);
    await dumpFieldHtml(page, ['定價']).catch(() => {});
  }
}
async function fill(page, label, value, problems, opts = {}) {
  if (value === undefined || value === '') return;
  try {
    const el = await visibleFieldAfter(page, label, opts.tag || TEXT_INPUT, opts);
    if (!el) throw new Error('not found');
    await el.fill(String(value), { timeout: 4000 });
    await commitInput(el);
    FILLED.push({ label, el, value: String(value), opts });
  } catch (e) {
    problems.push(`「${label}」沒填成功，請手動填：${value}`);
  }
}
const normVal = (v) => String(v || '').replace(/[\s,，]/g, '');
// 全部填完後再讀一次：被清掉的自動重填一次，重填後還是不對才提醒手動填
const readVal = (el) => el.evaluate((e) => (e.isContentEditable ? e.innerText : e.value)).catch(() => null);
const sameVal = (now, want) => now !== null && (normVal(now) === normVal(want) || normVal(now).startsWith(normVal(want).slice(0, 30)));
async function verifyFilled(page, problems) {
  for (const f of FILLED) {
    if (sameVal(await readVal(f.el), f.value)) continue;
    // 重新找一次欄位再填（頁面重整後，原本的元素可能已經被換掉）
    try {
      const el = f.refind ? await f.refind() : await visibleFieldAfter(page, f.label, f.opts.tag || TEXT_INPUT, f.opts);
      if (el) { await el.fill(f.value, { timeout: 4000 }); await commitInput(el); await page.waitForTimeout(500); f.el = el; }
    } catch (e) {}
    if (sameVal(await readVal(f.el), f.value)) { console.log(`   🔁 「${f.label}」被清掉了，已自動重填`); continue; }
    problems.push(`「${f.label}」填完後被清掉了，請手動填：${f.value.length > 40 ? f.value.slice(0, 40) + '…' : f.value}`);
  }
}

// 飛鼠商品分類（可複選）：依品牌選「【 Sony 索尼 】相機、鏡頭」這類選項；隨身相機另外勾「所有口袋機」
// 副廠鏡頭依接環放進原廠品牌（Sigma … For Sony → Sony）
const FEISHU_BRAND_KEY = {
  Canon: 'Canon', Fujifilm: 'Fujifilm', Hasselblad: 'hasselblad', Leica: 'Leica', Nikon: 'Nikon',
  Olympus: 'Olympus', Panasonic: 'Lumix', Sony: 'Sony',
};
function feishuCategoryKeys(r) {
  const name = r['商品名稱'] || '';
  let brand = r['品牌'] || guessBrand(name);
  const t = gearType(r);
  if (t === 'lens' && !FEISHU_BRAND_KEY[brand]) brand = guessMount(name) || brand;
  const keys = [];
  // L 接環（Leica、Panasonic S 系列共用）的鏡頭：Leica 和 Lumix 兩個分類都勾
  if (t === 'lens' && /\bL[\s-]*mount\b|for\s*L\b|L\s*接環/i.test(name)) keys.push('Leica', 'Lumix');
  else if (FEISHU_BRAND_KEY[brand]) keys.push(FEISHU_BRAND_KEY[brand]);
  if (t === 'compact') keys.push('所有口袋機');
  if (r['Y2K']) keys.push('CCD', 'Y2K'); // 分類名稱只要包含 CCD 或 Y2K 就會勾（改名也沒關係）
  return keys;
}

async function setCategory(page, r, problems) {
  const keys = feishuCategoryKeys(r);
  if (!keys.length) { problems.push('「分類」判斷不出來，請手動選'); return; }
  try {
    const sel = page.locator(`xpath=${labelXPath('分類', { all: true })}/following::select[1]`);
    let picked = [];
    for (let k = 0, n = await sel.count(); k < n && !picked.length; k++) {
      picked = await sel.nth(k).evaluate((el, keys) => {
        const lower = keys.map((x) => x.toLowerCase());
        const hit = [...el.options].filter((o) => lower.some((key) => o.text.toLowerCase().includes(key)));
        if (!hit.length) return [];
        if (!el.multiple) [...el.options].forEach((o) => (o.selected = false));
        hit.forEach((o) => (o.selected = true));
        el.dispatchEvent(new Event('input', { bubbles: true }));
        el.dispatchEvent(new Event('change', { bubbles: true }));
        if (window.jQuery) {
          const $el = window.jQuery(el);
          $el.trigger('change');
          $el.trigger('chosen:updated');
        }
        return hit.map((o) => o.text.trim());
      }, keys).catch(() => []);
    }
    if (!picked.length) throw new Error('no option');
    console.log(`   🗂️ 飛鼠分類：${picked.join('、')}`);
  } catch (e) {
    problems.push(`「分類」沒選成功，請手動選（${keys.join('、')}）`);
  }
}

async function addAttributes(page, attrs, problems) {
  if (!attrs.length) return;
  const header = labelXPath('屬性名稱');
  try {
    for (let i = 0; i < attrs.length; i++) {
      const nameInput = page.locator(`xpath=${header}/following::${TEXT_INPUT}[${2 * i + 1}]`);
      const valueInput = page.locator(`xpath=${header}/following::${TEXT_INPUT}[${2 * i + 2}]`);
      if (i > 0) {
        // 按綠色「+」新增一列
        const plus = page.locator(
          `xpath=${header}/following::*[self::button or self::a or self::div][normalize-space(.)='+'][1]`
        );
        await plus.click({ timeout: 3000 });
        await page.waitForTimeout(300);
      }
      // 防呆：如果這格已經有值，代表定位錯了，停下來交給人工
      if ((await nameInput.inputValue({ timeout: 3000 })) !== '') throw new Error('row not empty');
      await nameInput.fill(attrs[i][0]);
      await valueInput.fill(attrs[i][1]);
    }
  } catch (e) {
    problems.push('「商品屬性」沒填完整，請手動補：' + attrs.map(([k, v]) => `${k}=${v}`).join('、'));
  }
}

function shopeePrice(price) {
  const p = Number(price);
  const { markup, roundTo, highPrice } = CONFIG.shopee;
  if (!p || p >= highPrice) return '';
  return String(Math.ceil((p * markup) / roundTo) * roundTo);
}

// 同步資訊：找到「蝦皮」那一列，打開同步開關、填自訂價格
// 把同步資訊各平台（飛鼠官網、蝦皮…）的分類下拉選單選項存成文字檔，方便設定自動選分類
let OPTIONS_DUMPED = false;
async function dumpSyncOptions(page) {
  if (OPTIONS_DUMPED) return;
  OPTIONS_DUMPED = true;
  try {
    const text = await page.evaluate(() => {
      const out = [];
      for (const name of ['飛鼠官網', 'FB商店', '蝦皮', 'Y拍']) {
        const badge = [...document.querySelectorAll('body *')].find(
          (e) => e.children.length === 0 && e.textContent.trim() === name && e.offsetParent !== null
        );
        if (!badge) continue;
        const others = ['飛鼠官網', 'FB商店', '蝦皮', 'Y拍'].filter((x) => x !== name);
        let row = badge;
        while (row && !row.querySelector('select') && row !== document.body) row = row.parentElement;
        // 往上找到的範圍如果包含別的平台，代表這一列自己沒有下拉選單
        if (!row || row === document.body) continue;
        const rowText = [...row.querySelectorAll('*')].filter((e) => e.children.length === 0).map((e) => e.textContent.trim());
        if (others.some((o) => rowText.includes(o))) continue;
        row.querySelectorAll('select').forEach((sel, i) => {
          const opts = [...sel.options].map((o) => o.text.trim()).filter(Boolean);
          out.push(`【${name}】第 ${i + 1} 個下拉選單（${opts.length} 個選項）`);
          opts.forEach((o) => out.push('  ' + o));
          out.push('');
        });
      }
      return out.join('\n');
    });
    if (text.trim()) {
      const f = path.join(diagDir(), '平台分類選項.txt');
      fs.writeFileSync(f, text, 'utf8');
      console.log(`   📝 已把平台分類的下拉選項存到：${f}`);
    }
  } catch (e) {}
}

// ---------- 同步資訊的平台分類（第 2 個下拉選單＝常用分類的完整路徑）----------
// 商品類型：lens 鏡頭、mirrorless 無反、dslr 單眼、compact 隨身、adapter 轉接環、other 其他
const COMPACT_RE = /\bgr\s?(i{1,3}|[1-4])x?\b|g7\s?x|g5\s?x|g1\s?x|g9\s?x|rx100|rx10|\brx1|zv-?1\b|zv1|x100|xf10|x\s?half|\btg-?\d|mju|\blx\d|\bfz\d|coolpix|coolplx|powershot|ixus|\bsx\d|\bs\d{2,3}\b|d-?lux|\bq[23]\b|\bp\d{3}\b|hx\d|tx\d|dsc|finepix|instax|contax|pocket|action|insta360|\bl10\b|zs\d/i;
const DSLR_RE = /\b\d{1,4}d\b|\b[567]d\d?\b|\bd\d{2,4}\b|\bk-?\d|\bf\d{3}\b|\ba99\b|\ba77\b|\ba68\b/i;
function gearType(r) {
  const cat = r['品類'] || '';
  const name = r['商品名稱'] || '';
  if (/轉接環|adapter|ftz/i.test(name)) return 'adapter';
  if (/鏡頭/.test(cat)) return 'lens';
  if (/無反/.test(cat)) return 'mirrorless';
  if (/單眼/.test(cat)) return 'dslr';
  if (/隨身/.test(cat)) return 'compact';
  if (cat === '其他') return 'other';
  // 沒有品類時用名稱猜
  if (/\+/.test(name)) { /* 機身+鏡頭套組：看機身 */ }
  const body = name.split('+')[0];
  if (!/\+/.test(name) && LENS_RE.test(name)) return 'lens';
  if (COMPACT_RE.test(body)) return 'compact';
  if (DSLR_RE.test(body)) return 'dslr';
  return 'mirrorless';
}
// Y拍品牌選項的寫法
const YAHOO_BRAND = {
  lens: { Canon: 'Canon', Sony: 'Sony', Fujifilm: 'Fujifilm', Nikon: 'Nikon', Olympus: 'Olympus', Panasonic: 'Panasonic', Leica: 'Leica', Sigma: 'Sigma' },
  dslr: { Canon: 'Canon', Sony: 'SONY', Fujifilm: 'FUJIFILM富士', Nikon: 'Nikon', Panasonic: 'Panasonic', Olympus: 'Olympus', Pentax: 'Pentax' },
  compact: { Sony: 'SONY', Canon: 'Canon', Nikon: 'Nikon', Panasonic: 'Panasonic', Leica: 'LEICA', Fujifilm: 'FUJIFILM富士', Ricoh: 'RICOH', Olympus: 'OLYMPUS' },
};
function platformCategory(platform, r) {
  const t = gearType(r);
  const brand = r['品牌'] || '';
  if (platform === '飛鼠官網') {
    return { lens: '相機和光學元件 > 相機和光學配件 > 相機/攝影機鏡頭 > 相機鏡頭', adapter: '相機和光學元件 > 攝影器材', other: '相機和光學元件 > 攝影器材' }[t]
      || '相機和光學元件 > 相機 > 數位相機';
  }
  if (platform === '蝦皮') {
    if (/CCD/.test(r['Y2K'] || '')) return '相機&空拍機 > 相機 > 全自動數位相機'; // CCD 一律放數位相機（傻瓜相機）
    return { lens: '相機&空拍機 > 鏡頭', mirrorless: '相機&空拍機 > 相機 > 無反光鏡可換鏡頭相機', dslr: '相機&空拍機 > 相機 > 數位單眼相機',
      compact: '相機&空拍機 > 相機 > 全自動數位相機', adapter: '相機&空拍機 > 鏡頭周邊配件 > 轉接環' }[t] || '';
  }
  if (platform === 'Y拍') {
    if (t === 'lens') return YAHOO_BRAND.lens[brand] ? `相機、攝影與周邊 > 鏡頭 > ${YAHOO_BRAND.lens[brand]}` : '';
    if (t === 'mirrorless' || t === 'dslr') {
      if (YAHOO_BRAND.dslr[brand]) return `相機、攝影與周邊 > 單眼相機 > ${YAHOO_BRAND.dslr[brand]}`;
    }
    if (t === 'mirrorless' || t === 'dslr' || t === 'compact') {
      return `相機、攝影與周邊 > 數位相機 > ${YAHOO_BRAND.compact[brand] || '其他品牌相機'}`;
    }
  }
  return '';
}
// 在這一列的第 n 個下拉選單，選文字完全一樣的選項（被美化套件藏起來的 select 也能選）
async function pickRowSelect(row, index, text) {
  const sel = row.locator('select').nth(index);
  if (!(await sel.count())) return false;
  return await sel.evaluate((el, want) => {
    const opt = [...el.options].find((o) => o.text.trim() === want);
    if (!opt) return false;
    el.value = opt.value;
    el.dispatchEvent(new Event('input', { bubbles: true }));
    el.dispatchEvent(new Event('change', { bubbles: true }));
    if (window.jQuery) window.jQuery(el).trigger('change');
    return true;
  }, text).catch(() => false);
}

// 同步資訊：找到某個平台那一列（例如「蝦皮」「Y拍」）
async function findSyncRow(page, name) {
  const badges = page.locator(`xpath=//*[normalize-space(text())='${name}']`);
  for (let k = 0, n = await badges.count(); k < n; k++) {
    const b = badges.nth(k);
    if (!(await b.isVisible().catch(() => false))) continue;
    const cand = b.locator("xpath=ancestor::*[self::tr or self::li or self::div][.//input[@type='checkbox'] or .//*[contains(@class,'switch') or contains(@class,'toggle')]][1]");
    if (await cand.count()) return cand.first();
  }
  return null;
}

// 打開這一列的同步開關，回傳 true / false / null（不確定）
async function turnSyncOn(page, row) {
  const cb = row.locator('input[type=checkbox]').first();
  const hasCb = (await cb.count()) > 0;
  const state = () => (hasCb ? cb.isChecked().catch(() => null) : Promise.resolve(null));
  const waitOn = async () => {
    for (let t = 0; t < 10; t++) { if ((await state()) === true) return true; await page.waitForTimeout(250); }
    return state();
  };
  if ((await state()) === true) return true;
  const sw = row.locator(".switch, .toggle, [class*='switch'], [class*='toggle'], label:has(input[type=checkbox])").first();
  for (let tries = 0; tries < 2; tries++) {
    if (await sw.count()) await sw.click({ timeout: 3000 }).catch(() => {});
    else if (hasCb) await cb.check({ force: true, timeout: 3000 }).catch(() => {});
    const on = await waitOn();
    if (on !== false) return on; // true，或沒有 checkbox 無法確認（null）
  }
  return false;
}

async function fillShopeeRow(page, r, problems) {
  await dumpSyncOptions(page);
  // 1) 打開各平台同步開關
  for (const name of CONFIG.syncPlatforms) {
    try {
      const row = await findSyncRow(page, name);
      if (!row) throw new Error('no row');
      if ((await turnSyncOn(page, row)) === false) throw new Error('still off');
    } catch (e) {
      problems.push(`「${name}同步」開關沒打開成功，請手動打開`);
    }
  }
  // 2) 各平台的平台分類（先選分類：選完那一列可能會重整，之後再填自訂標題／價格才不會被清掉）
  for (const name of ['飛鼠官網', ...CONFIG.syncPlatforms]) {
    const want = platformCategory(name, r);
    try {
      if (!want) throw new Error('no mapping');
      const row = await findSyncRow(page, name);
      if (!row) throw new Error('no row');
      await pickRowSelect(row, 0, '常用分類');
      await page.waitForTimeout(300);
      if (!(await pickRowSelect(row, 1, want))) throw new Error('option not found');
      console.log(`   🗂️ ${name}平台分類：${want}`);
    } catch (e) {
      problems.push(`「${name}平台分類」沒選成功，請手動選${want ? `：${want}` : ''}`);
    }
  }
  await page.waitForTimeout(800);
  // 蝦皮那一列的第 n 個文字框（第 1 個＝自訂價格、第 2 個＝自訂標題）；每次重新找，避免拿到舊的元素
  const shopeeBox = async (n) => { const row = await findSyncRow(page, '蝦皮'); return row ? row.locator(`xpath=.//${TEXT_INPUT}`).nth(n) : null; };
  const fillShopee = async (n, label, value) => {
    const el = await shopeeBox(n);
    if (!el) throw new Error('no row');
    await el.fill(value, { timeout: 3000 });
    await commitInput(el);
    FILLED.push({ label, el, value, opts: {}, refind: () => shopeeBox(n) });
  };
  // 3) 蝦皮自訂標題（這一列第二個文字框）：最多 60 字，含 #序號
  if (r['蝦皮標題']) {
    try {
      await fillShopee(1, '蝦皮自訂標題', r['蝦皮標題']);
      console.log(`   🏷️ 蝦皮標題（${r['蝦皮標題'].length} 字）：${r['蝦皮標題']}`);
    } catch (e) {
      problems.push(`「蝦皮自訂標題」沒填成功，請手動填：${r['蝦皮標題']}`);
    }
  }
  // 4) 蝦皮自訂價格（這一列第一個文字框）
  if (CONFIG.shopee.enable) {
    const sp = shopeePrice(r['銷售價']);
    try {
      if (sp) {
        await fillShopee(0, '蝦皮自訂價格', sp);
        console.log(`   💰 蝦皮自訂價格：${r['銷售價']} × ${CONFIG.shopee.markup} → ${sp}`);
      } else if (Number(r['銷售價']) >= CONFIG.shopee.highPrice) {
        // 15000 以上：蝦皮用原價 → 自訂價格要「清空」（以前門檻 20000 時可能被填過 ×1.18）
        const el = await shopeeBox(0);
        if (el && String(await el.inputValue().catch(() => '')).trim()) { await el.fill('', { timeout: 3000 }); await commitInput(el); }
        console.log(`   💰 蝦皮自訂價格：清空（售價 ${r['銷售價']} ≥ ${CONFIG.shopee.highPrice}，蝦皮用原價，最低購買數量之後改 3）`);
      }
    } catch (e) {
      problems.push(`「蝦皮自訂價格」沒填成功，請手動填：${sp}`);
    }
  }
  if (!lookupNcc(r['商品名稱']) && !isNoRadio(lookupNcc(r['商品名稱']))) problems.push(`NCC 字號對照表沒有「${r['商品名稱']}」，文案先不寫 NCC；查到後補進 NCC字號.csv，下次就會自動帶入`);
}

// 有欄位填不進去時，把那幾個欄位附近的網頁原始碼存成「欄位診斷.txt」，傳給 Claude 看就知道怎麼修
let DIAG_EXTRA = '';
let PENDING_PHOTO_RECORD = '';
async function dumpFieldHtml(page, labels) {
  if (!labels.length) return;
  try {
    const text = await page.evaluate((labels) => {
      const norm = (t) => (t || '').replace(/[\s\u00a0*＊]/g, '');
      const clean = (h) => h.replace(/<svg[\s\S]*?<\/svg>/g, '<svg/>').replace(/\s(style|d|class)="[^"]{60,}"/g, ' $1="…"');
      const out = [];
      for (const label of labels) {
        const want = norm(label);
        const hits = [...document.querySelectorAll('body *')].filter(
          (e) => [...e.childNodes].some((n) => n.nodeType === 3 && norm(n.textContent) === want)
        );
        out.push(`==== 「${label}」找到 ${hits.length} 個標題 ====`);
        hits.slice(0, 3).forEach((h, i) => {
          let box = h;
          for (let k = 0; k < 5 && box.parentElement && !box.querySelector('input,select,textarea,[contenteditable]'); k++) box = box.parentElement;
          out.push(`-- 第 ${i + 1} 個（${h.offsetParent ? '看得見' : '看不見'}）--`);
          out.push(clean(box.outerHTML).slice(0, 2500));
        });
        out.push('');
      }
      return out.join('\n');
    }, labels);
    const f = path.join(diagDir(), '欄位診斷.txt');
    fs.writeFileSync(f, (DIAG_EXTRA ? DIAG_EXTRA + '\n\n' : '') + text, 'utf8');
    DIAG_EXTRA = '';
    console.log(`   🔍 填不進去的欄位已存診斷檔：${f}（用記事本打開，全選複製貼給 Claude）`);
  } catch (e) {}
}

// 上傳商品圖片：找「商品圖片」後面的檔案上傳欄位（就算被拖曳框藏起來也能用），可一次傳多張
const asUpload = (file) => {
  const ext = path.extname(file).toLowerCase();
  const mimeType = ext === '.png' ? 'image/png' : ext === '.webp' ? 'image/webp' : 'image/jpeg';
  // 用檔案內容上傳（中文檔名直接用路徑上傳會失敗）
  return { name: `photo-${path.basename(file).replace(/[^\w.\-]/g, '_')}`, mimeType, buffer: fs.readFileSync(file) };
};
// 頁面上有好幾個上傳框（商品圖片、廣告圖…），只挑「商品圖片」那一個，絕不傳到廣告圖
const IMG_LABELS = /^(商品圖片|廣告|分享圖|規格圖|封面圖|SEO|主圖|圖片)/;
async function findImageInput(page) {
  const idx = await page.evaluate((src) => {
    const re = new RegExp(src);
    const inputs = [...document.querySelectorAll('input[type=file]')];
    // 找每個上傳欄位「前面最近的圖片標題」
    const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
    const heads = [];
    for (let n; (n = walker.nextNode());) {
      const t = n.textContent.replace(/[\s ]/g, '');
      if (t && t.length <= 20 && re.test(t)) heads.push({ node: n, text: t });
    }
    let best = -1;
    inputs.forEach((inp, i) => {
      let last = null;
      for (const h of heads) if (h.node.compareDocumentPosition(inp) & Node.DOCUMENT_POSITION_FOLLOWING) last = h;
      if (last && last.text.startsWith('商品圖片') && best < 0) best = i;
    });
    return best;
  }, IMG_LABELS.source).catch(() => -1);
  if (idx >= 0) return page.locator('input[type=file]').nth(idx);
  // 飛鼠的「商品圖片」是拖曳上傳框（Dropzone），真正的上傳欄位藏在頁面最底：
  // 可多選、只收圖片、沒有 name；廣告圖（adv-image）、尺寸表（size-chart-image）、規格圖都排除
  const dz = await page.evaluate(() => {
    const all = [...document.querySelectorAll('input[type=file]')];
    const bad = (i) => /adv|size|chart|spec|option|banner|share|seo/i.test(i.name + ' ' + i.id + ' ' + i.className);
    let k = all.findIndex((i) => i.classList.contains('dz-hidden-input') && !bad(i));
    if (k < 0) k = all.findIndex((i) => i.multiple && /image/.test(i.accept) && !i.name && !bad(i));
    return k;
  }).catch(() => -1);
  return dz >= 0 ? page.locator('input[type=file]').nth(dz) : null;
}
// 拖曳框上傳完會出現縮圖（.dz-preview）；等縮圖數量增加、而且都傳完
const dzCount = (page) => page.evaluate(() => ({
  all: document.querySelectorAll('.dz-preview').length,
  busy: document.querySelectorAll('.dz-preview.dz-processing:not(.dz-complete)').length,
  err: document.querySelectorAll('.dz-preview.dz-error').length,
})).catch(() => ({ all: 0, busy: 0, err: 0 }));
// 找不到隱藏的上傳欄位時：點「商品圖片」區塊裡的上傳按鈕，接住檔案選擇視窗
async function chooseViaClick(page, files) {
  const label = page.locator(`xpath=${labelXPath('商品圖片', { all: true })}`).first();
  if (!(await label.count())) return false;
  const box = label.locator(`xpath=following::*[self::button or self::a or self::label or contains(@class,'upload') or contains(@class,'Upload')][1]`).first();
  if (!(await box.count())) return false;
  const near = await box.evaluate((el) => {
    // 確認這個按鈕前面最近的標題還是「商品圖片」（不是廣告圖）
    let n = el; for (let k = 0; k < 6 && n; k++, n = n.parentElement) if (/廣告/.test(n.textContent) && !/商品圖片/.test(n.textContent)) return false;
    return true;
  }).catch(() => false);
  if (!near) return false;
  const [chooser] = await Promise.all([page.waitForEvent('filechooser', { timeout: 5000 }), box.click()]).catch(() => [null]);
  if (!chooser) return false;
  await chooser.setFiles(chooser.isMultiple() ? files.map(asUpload) : asUpload(files[0]));
  return chooser.isMultiple() ? files.length : 1;
}
// 商品圖片下面的「上傳後 AI 自動去背」要打勾（只動商品圖片這一個，不動廣告圖、規格圖的）
async function ensureAiCutout(page) {
  const r = await page.evaluate(() => {
    const labels = [...document.querySelectorAll('label')];
    const head = labels.find((l) => l.textContent.replace(/\s/g, '') === '商品圖片');
    if (!head) return 'no-head';
    const lab = labels.find((l) => (head.compareDocumentPosition(l) & Node.DOCUMENT_POSITION_FOLLOWING) && /^上傳後AI自動去背$/.test(l.textContent.replace(/\s/g, '')));
    if (!lab) return 'no-label';
    // 確認中間沒有隔著別的圖片標題（例如廣告圖片）
    const between = labels.filter((l) => (head.compareDocumentPosition(l) & Node.DOCUMENT_POSITION_FOLLOWING) && (l.compareDocumentPosition(lab) & Node.DOCUMENT_POSITION_FOLLOWING));
    if (between.some((l) => /圖片/.test(l.textContent))) return 'not-product';
    const box = lab.querySelector('input[type=checkbox]') || (lab.htmlFor && document.getElementById(lab.htmlFor));
    if (!box) return 'no-box';
    if (box.checked) return 'already';
    box.click();
    return box.checked ? 'checked' : 'failed';
  }).catch(() => 'error');
  if (r === 'checked') console.log('   ☑️ 已勾選「上傳後 AI 自動去背」');
  else if (r !== 'already') console.log(`   ⚠️ 找不到或勾不起來「上傳後 AI 自動去背」（${r}），請手動確認`);
}
// 飛鼠「商品圖片」區現在已經有幾張照片（在「商品圖片」和「廣告圖片」兩個標題中間的圖）；讀不到回傳 null
async function existingImageCount(page) {
  await page.waitForTimeout(1500); // 等舊照片的縮圖載入
  return page.evaluate(() => {
    const labels = [...document.querySelectorAll('label')];
    const flat = (el) => el.textContent.replace(/\s/g, '');
    const head = labels.find((l) => flat(l) === '商品圖片');
    if (!head) return null;
    const next = labels.find((l) => (head.compareDocumentPosition(l) & Node.DOCUMENT_POSITION_FOLLOWING) && /^(廣告圖片|尺寸表圖片|SPU編號)/.test(flat(l)));
    const inRange = (el) => (head.compareDocumentPosition(el) & Node.DOCUMENT_POSITION_FOLLOWING) && (!next || (el.compareDocumentPosition(next) & Node.DOCUMENT_POSITION_FOLLOWING));
    const previews = [...document.querySelectorAll('.dz-preview')].filter(inRange);
    if (previews.length) return previews.length;
    return [...document.querySelectorAll('img')].filter((im) => inRange(im) && (im.naturalWidth || im.width) >= 40).length;
  }).catch(() => null);
}
// 等商品圖片全部處理完（上傳、AI 去背）：沒有「處理中」的縮圖、畫面上沒有「去背中／處理中／上傳中」，而且連續幾秒沒有變化
async function waitImagesSettled(page, maxSec = 150) {
  let stable = 0, last = '';
  for (let t = 0; t < maxSec; t++) {
    const st = await page.evaluate(() => {
      const vis = (el) => { const r = el.getBoundingClientRect(); return r.width > 0 && r.height > 0; };
      const busy = [...document.querySelectorAll('.dz-preview.dz-processing:not(.dz-complete), [class*=loading], [class*=spinner], .progress-bar')].filter(vis).length;
      const txt = [...document.querySelectorAll('body *')].filter((e) => e.children.length === 0 && vis(e) && /去背中|處理中|上傳中|processing|uploading/i.test(e.textContent || '')).length;
      const imgs = [...document.querySelectorAll('.dz-preview img, .dz-image img')].map((i) => i.src.slice(-40)).join('|');
      return { busy: busy + txt, imgs };
    }).catch(() => ({ busy: 0, imgs: '' }));
    if (!st.busy && st.imgs === last) { if (++stable >= 4) return true; } else stable = 0;
    last = st.imgs;
    if (t === 15) console.log('   ⏳ 等飛鼠處理照片（AI 去背）…');
    await page.waitForTimeout(1000);
  }
  console.log('   ⚠️ 照片處理等了很久還沒好，先繼續');
  return false;
}
async function uploadImages(page, files, problems) {
  files = (files || []).filter(Boolean);
  if (!files.length) return;
  const missing = files.filter((f) => !fs.existsSync(f));
  if (missing.length) { problems.push(`找不到圖片檔 ${missing.map((f) => path.basename(f)).join('、')}，商品圖片請手動上傳`); return; }
  const waitFor = (list) => page.waitForTimeout(Math.min(2500 + list.reduce((a, f) => a + fs.statSync(f).size, 0) / 1e6 * 600, 20000));
  await ensureAiCutout(page);
  try {
    let left = files.slice();
    while (left.length) {
      const input = await findImageInput(page);
      const before = await dzCount(page);
      let n;
      if (input) {
        const multiple = await input.evaluate((e) => e.multiple).catch(() => false);
        n = multiple ? left.length : 1;
        await input.setInputFiles(multiple ? left.map(asUpload) : asUpload(left[0]), { timeout: 10000 });
      } else {
        n = await chooseViaClick(page, left);
        if (!n) throw new Error('找不到「商品圖片」的上傳欄位');
      }
      await waitFor(left.slice(0, n));
      // 有拖曳框縮圖的話，最多再等 90 秒讓它全部傳完
      for (let t = 0; t < 90; t++) {
        const c = await dzCount(page);
        if (c.all === before.all && t > 5) break;               // 沒有縮圖機制，就照時間等
        if (c.all >= before.all + n && !c.busy) { if (c.err > before.err) throw new Error('飛鼠回報圖片上傳錯誤'); break; }
        await page.waitForTimeout(1000);
      }
      left = left.slice(n);
    }
    console.log(`   🖼️ 已上傳商品圖片 ${files.length} 張`);
  } catch (e) {
    problems.push(`商品圖片沒上傳成功，請手動上傳（${files.length} 張）`);
    await dumpImageFields(page).catch(() => {});
  }
}
// 診斷用：把頁面上所有上傳欄位、圖片相關的文字寫進「圖片欄位診斷.txt」（不含個資，只有頁面結構）
async function dumpImageFields(page) {
  const info = await page.evaluate(() => {
    const out = [];
    const desc = (el) => el ? `<${el.tagName.toLowerCase()}${el.id ? ' id=' + el.id : ''}${el.className && typeof el.className === 'string' ? ' class="' + el.className.slice(0, 80) + '"' : ''}>` : '';
    const textsBefore = (el, n) => {
      const w = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT); const arr = [];
      for (let t; (t = w.nextNode());) {
        if (!(t.compareDocumentPosition(el) & Node.DOCUMENT_POSITION_FOLLOWING)) break;
        const x = t.textContent.replace(/\s+/g, ' ').trim(); if (x && !/^[{<]/.test(x)) arr.push(x.slice(0, 40));
      }
      return arr.slice(-n);
    };
    out.push('網址：' + location.pathname);
    document.querySelectorAll('input[type=file]').forEach((inp, i) => {
      const chain = []; let p = inp.parentElement; for (let k = 0; k < 5 && p; k++, p = p.parentElement) chain.push(desc(p));
      out.push(`\n[上傳欄位 ${i + 1}] name=${inp.name} accept=${inp.accept} multiple=${inp.multiple} 看得見=${!!inp.offsetParent}`);
      out.push('  外層：' + chain.join(' ← '));
      out.push('  前面的文字：' + textsBefore(inp, 6).join(' ｜ '));
    });
    const w = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
    const pics = []; for (let t; (t = w.nextNode());) { const x = t.textContent.replace(/\s+/g, ' ').trim(); if (/圖|照片|上傳/.test(x) && x.length < 40) pics.push(x + '　' + desc(t.parentElement)); }
    out.push('\n[頁面上有「圖／照片／上傳」的文字]'); out.push(...pics.map((x) => '  ' + x));
    return out.join('\n');
  });
  fs.writeFileSync(path.join(diagDir(), '圖片欄位診斷.txt'), '\uFEFF' + info.replace(/\n/g, '\r\n'), 'utf8');
  await page.screenshot({ path: path.join(diagDir(), '圖片欄位截圖.png'), fullPage: true }).catch(() => {});
  console.log('   📝 已存「系統\診斷檔\圖片欄位診斷.txt」和截圖，請把診斷檔內容貼給 Claude');
}
async function uploadImage(page, file, problems) { return uploadImages(page, file ? [file] : [], problems); }

async function fillProduct(page, r) {
  const problems = [];
  const barcode = r['條碼'] || '';
  FILLED = [];
  const price = r['銷售價'];

  // 圖片「最先」上傳，並等 AI 去背跑完：飛鼠處理完圖片會重整表單，先傳圖再填欄位，填好的才不會被清掉
  // 有整理好的實拍照（商品照片\序號）就傳實拍照，沒有才用佔位圖（佔位圖目前關閉）
  const photos = productPhotos(barcode);
  if (photos.length) {
    const before = problems.length;
    await uploadImages(page, photos, problems);
    if (problems.length === before) PENDING_PHOTO_RECORD = barcode; // 存檔成功後才記錄
    await waitImagesSettled(page);
  } else {
    await uploadImage(page, CONFIG.placeholderImage, problems);
    problems.push('沒有整理好的實拍照，商品圖片請手動上傳');
  }

  await fill(page, 'SPU 編號', barcode, problems);
  await fill(page, '商品名稱', r['標題'] || r['商品名稱'], problems);
  await fill(page, '品牌', r['品牌'], problems);
  await setCategory(page, r, problems);

  if (CONFIG.introTo === 'feature1' && r['商品介紹']) {
    await fill(page, '商品特色', r['商品介紹'], problems);
  }

  if (r['商品描述']) {
    await fill(page, '商品描述', r['商品描述'], problems, {
      tag: "*[self::textarea or @contenteditable='true']",
    });
  }

  const attrs = CONFIG.attributeColumns
    .filter((c) => r[c] !== undefined && !EMPTY_MARKS.has(r[c]))
    .map((c) => [c, r[c]]);
  await addAttributes(page, attrs, problems);

  // 商品規格（單一規格）
  await ensureListPrice(page, price, problems); // 定價＝售價（用自己的找法，避免填到售價那格）
  await fill(page, '售價', price, problems);
  await fill(page, '成本', r['進貨價'], problems);
  await fill(page, 'SKU 貨號', barcode, problems);
  await fill(page, '庫存數量', r['庫存量'] || '1', problems);

  if (CONFIG.supplierToNote && r['供貨商']) {
    await fill(page, '備註', `供貨商：${r['供貨商']}`, problems, { tag: 'textarea', prefix: true });
  }
  await fillShopeeRow(page, r, problems);
  // 填完再等一次圖片處理（AI 去背跑完飛鼠會重整表單），然後確認前面填的都還在
  await waitImagesSettled(page);
  await page.waitForTimeout(1500);
  await verifyFilled(page, problems);
  // 最後再確認一次「定價」＝「售價」（飛鼠現在定價是必填，空的會存不進去）
  await ensureListPrice(page, price, problems);
  const failed = problems.map((x) => (x.match(/^「(.+?)」(沒(填|選)成功|填完後被清掉了)/) || [])[1]).filter(Boolean);
  // 也記下「腳本實際填進去的是哪個元素」，比對網頁原始碼用
  const cleared = FILLED.filter((f) => failed.includes(f.label));
  if (cleared.length) {
    const info = [];
    for (const f of cleared) {
      info.push(`「${f.label}」腳本填的元素：` + (await f.el.evaluate((e) => e.outerHTML.slice(0, 400)).catch(() => '（讀不到）')));
    }
    DIAG_EXTRA = info.join('\n');
  }
  await dumpFieldHtml(page, failed);
  return problems;
}

// 記住「新增商品」頁的網址，下次直接打開，不用靠左側選單
const NEW_URL_FILE = path.join(__dirname, '飛鼠新增商品網址.txt');
const isNewProductPage = (page) =>
  page.getByText('建立新商品', { exact: true }).first().isVisible().catch(() => false);

// 頁面上目前看得到的提示／錯誤文字（用來比對按儲存之後「新出現」的訊息）
const SAVE_ERR_RE = /重複|已存在|已被使用|已經存在|不可重複|失敗|錯誤|必填|請輸入|請選擇|格式不正確/;
const SAVE_OK_RE = /成功|已儲存|已建立|已新增|建立完成|儲存完成|新增完成|更新完成|發佈完成|上架完成/;
async function visibleMessages(page) {
  return page.evaluate(() => {
    const sel = '.toast, .alert, .swal2-popup, .sweet-alert, .modal.show, .modal.in, .notification, .noty_body, .invalid-feedback, .help-block, .text-danger, .error, [role=alert]';
    const out = [];
    for (const el of document.querySelectorAll(sel)) {
      if (!el.offsetParent && getComputedStyle(el).position !== 'fixed') continue;
      const txt = (el.innerText || '').trim().replace(/\s+/g, ' ');
      if (txt) out.push(txt.slice(0, 120));
    }
    return out;
  }).catch(() => []);
}

// 按「儲存草稿／建立」後確認飛鼠真的存好了：出現新的錯誤訊息（例如編號重複）就丟出錯誤
async function confirmSaved(page, btn, before = [], stillOnForm = isNewProductPage) {
  const startUrl = page.url();
  const seen = new Set(before);
  for (let t = 0; t < 120; t++) { // 最多等 30 秒（有照片、AI 去背時存檔比較慢）
    await page.waitForTimeout(250);
    const fresh = (await visibleMessages(page)).filter((m) => !seen.has(m));
    const err = fresh.find((m) => SAVE_ERR_RE.test(m) && !SAVE_OK_RE.test(m));
    if (err) throw new Error(`飛鼠沒有存成功：${err}`);
    if (fresh.some((m) => SAVE_OK_RE.test(m))) return;
    if (page.url() !== startUrl) return;
    if (!(await stillOnForm(page))) return;
  }
  const msgs = (await visibleMessages(page)).filter((m) => !seen.has(m)).slice(0, 3);
  throw new Error(`按了「${btn}」30 秒後頁面沒有變化，不確定有沒有存成功，請到飛鼠確認這筆${msgs.length ? `（畫面上的訊息：${msgs.join('／')}）` : ''}`);
}

async function openNewProductPage(page) {
  // 1) 用記住的網址
  if (fs.existsSync(NEW_URL_FILE)) {
    const url = fs.readFileSync(NEW_URL_FILE, 'utf8').trim();
    if (url) {
      await page.goto(url, { waitUntil: 'domcontentloaded' }).catch(() => {});
      await page.getByText('建立新商品', { exact: true }).first().waitFor({ timeout: 15000 }).catch(() => {});
      if (await isNewProductPage(page)) { await page.waitForTimeout(800); return; }
    }
  }
  // 2) 點左側選單：商品管理 → 新增商品
  try {
    const link = page.getByText('新增商品', { exact: true }).first();
    if (!(await link.isVisible().catch(() => false))) {
      await page.getByText('商品管理', { exact: true }).first().click({ timeout: 8000 });
      await page.waitForTimeout(600);
    }
    await link.click({ timeout: 8000 });
    await page.getByText('建立新商品', { exact: true }).first().waitFor({ timeout: 15000 });
  } catch (e) {
    // 3) 都不行：請使用者自己點到新增商品頁
    console.log('   ⚠️ 找不到飛鼠左側選單（可能還沒登入、登入過期，或視窗太窄選單收起來了）');
    await ask('   請在瀏覽器登入飛鼠，自己點到「商品管理 → 新增商品」頁面，再回這裡按 Enter：');
    if (!(await isNewProductPage(page))) throw new Error('目前頁面不是「建立新商品」，這筆先跳過');
  }
  try { fs.writeFileSync(NEW_URL_FILE, page.url(), 'utf8'); } catch (e) {}
  await page.waitForTimeout(800);
}

async function main(srcOverride) {
  const src = srcOverride || process.argv[2];
  const listOnly = process.argv.includes('--list');
  const redo = process.argv.includes('--redo'); // 之前記成已上架的也重新處理
  // 指令後面可加：--draft（自動存草稿）、--create（自動建立）、--limit 1（只跑 1 筆）
  if (process.argv.includes('--draft')) CONFIG.mode = 'draft';
  if (process.argv.includes('--create')) CONFIG.mode = 'create';
  const li = process.argv.indexOf('--limit');
  if (li > 0 && process.argv[li + 1]) CONFIG.limit = parseInt(process.argv[li + 1], 10) || 0;
  let rows;
  if (src === 'airtable') {
    rows = await readAirtable();
    console.log(`從 Airtable 讀到 ${rows.length} 筆${CONFIG.airtable.requireMark ? '「已標記上架飛鼠」' : ''}的收購紀錄`);
  } else if (src && fs.existsSync(src)) {
    rows = await readRows(src);
  } else {
    console.log('用法： node upload.js airtable   或   node upload.js 商品.xlsx');
    process.exit(1);
  }

  const done = loadDone();
  const where = (r) => (typeof r._row === 'number' ? `第 ${r._row} 列` : `收購日 ${r['交易日期'] || '?'}`);
  let todo = [];
  for (const r of rows) {
    const guessed = autoFill(r);
    if (src === 'airtable') { r['標題'] = buildTitle(r); r['蝦皮標題'] = buildShopeeTitle(r); }
    if (guessed.length) console.log(`🤖 ${where(r)}「${r['商品名稱']}」自動判斷：${guessed.join('、')}（請在填表時確認）`);
    if (src === 'airtable' && !String(r['商品介紹'] || '').trim()) console.log(`⚠️ 「${r['商品名稱']}」收購紀錄沒有「商品介紹」，文案會少【型號介紹】。到收購後台打開這件的「🖨️ 印卡」，AI 會自動補上，再上架`);
    // review 模式下分類空白不略過，填表時讓你在瀏覽器手動選
    const missing = validate(r).filter((x) => !(x === '分類' && CONFIG.mode === 'review'));
    if (missing.length) {
      console.log(`⚠️  ${where(r)}「${r['商品名稱']}」缺必填欄位：${missing.join('、')}，略過`);
      if (!listOnly) log(r, '略過', '缺 ' + missing.join('、'));
    } else if (!redo && done.has(keyOf(r))) {
      console.log(`⏭️  ${where(r)}「${r['商品名稱']}」之前已上架，略過（要重做請加 --redo）`);
    } else if (!redo && /已上架/.test(r._fsStatus || '')) {
      console.log(`⏭️  ${where(r)}「${r['商品名稱']}」收購後台顯示飛鼠已上架，略過`);
    } else if (!redo && /草稿/.test(r._fsStatus || '')) {
      console.log(`⏭️  ${where(r)}「${r['商品名稱']}」飛鼠已經是草稿，不重新建立（補照片會另外處理）`);
    } else todo.push(r);
  }
  // 上架紀錄被刪掉時也不會重複上架：先查官網，已經有這個序號的就當作已上架
  if (src === 'airtable' && !redo && todo.length) {
    const live = [];
    const sf = await storefrontMany(todo.map((r) => r['條碼']));
    for (const r of todo) if (r['條碼'] && sf.get(r['條碼'])) live.push(r);
    live.forEach((r) => {
      console.log(`⏭️  「${r['商品名稱']}」序號 ${r['條碼']} 官網已經有了，當作已上架，略過`);
      if (!listOnly) log(r, '成功（已建立）', '官網已有，補記紀錄');
    });
    todo = todo.filter((r) => !live.includes(r));
  }
  if (CONFIG.onlyWithPhotos) {
    const noPhoto = todo.filter((r) => !productPhotos(r['條碼']).length);
    if (noPhoto.length) {
      console.log(`\n⏭️ 這 ${noPhoto.length} 件還沒有照片，這次先不處理（不會上架、也不會存草稿）：`);
      noPhoto.forEach((r) => console.log(`   ・${r['商品名稱']}　序號 ${r['條碼']}`));
    }
    todo = todo.filter((r) => productPhotos(r['條碼']).length);
  }
  if (CONFIG.limit > 0) todo = todo.slice(0, CONFIG.limit);
  console.log(`\n共 ${todo.length} 筆要處理，模式：${CONFIG.mode}\n`);
  if (!todo.length) return;
  if (listOnly) {
    todo.forEach((r, i) => console.log(`${i + 1}. ${r['標題'] || r['商品名稱']}\n   蝦皮：${r['蝦皮標題'] || '-'}\n   序號 ${r['條碼']}｜${r['分類']}｜售價 ${r['銷售價']}｜成本 ${r['進貨價'] || '-'}`));
    console.log('\n（--list 只列出，沒有開瀏覽器。拿掉 --list 就會開始上架）');
    return;
  }

  const ctx = await launchBrowser({
    headless: false,
    channel: CONFIG.browserChannel,
    viewport: null,
    args: ['--start-maximized'],
  });
  const page = ctx.pages()[0] || (await ctx.newPage());
  // 換到下一筆時若跳出「確定要離開此頁面？」，自動按確定
  page.on('dialog', (d) => d.accept().catch(() => {}));
  if (CONFIG.startUrl) await page.goto(CONFIG.startUrl);

  // 已經登入（記住的「新增商品」網址打得開）就直接開始，不用按 Enter
  let loggedIn = false;
  if (fs.existsSync(NEW_URL_FILE)) {
    const url = fs.readFileSync(NEW_URL_FILE, 'utf8').trim();
    if (url) {
      await page.goto(url, { waitUntil: 'domcontentloaded' }).catch(() => {});
      loggedIn = await page.getByText('建立新商品', { exact: true }).first().waitFor({ timeout: 15000 }).then(() => true, () => false);
    }
  }
  if (loggedIn) console.log('✅ 飛鼠已經是登入狀態，直接開始');
  else await ask('👉 請在瀏覽器登入飛鼠後台，看到左側選單後，回這裡按 Enter 開始...');

  let ok = 0, fail = 0;
  for (const [i, r] of todo.entries()) {
    console.log(`\n[${i + 1}/${todo.length}] ${r['商品名稱']}`);
    try {
      await openNewProductPage(page);
      const problems = await fillProduct(page, r);
      problems.forEach((p) => console.log('   • ' + p));

      // review：每筆都停下來；draft：不停（之後再進草稿補照片、分類）；create：有欄位沒填成功才停
      // 只有飛鼠必填的欄位（名稱、分類、售價、SPU、SKU、描述…）沒填好才算嚴重；
      // 平台分類、蝦皮自訂標題／價格沒選到不擋上架（飛鼠會用預設，最後的蝦皮價格檢查會再補）
      const isBad = (x) => /沒填成功|沒選成功|沒設定成功|判斷不出來|被清掉了/.test(x) && !/平台分類|蝦皮自訂|同步/.test(x);
      const serious = problems.some(isBad);
      const noPhotos = problems.some((x) => x.startsWith('沒有整理好的實拍照') || x.startsWith('商品圖片沒上傳成功'));
      // 直接建立模式：沒有照片、或有欄位沒填好 → 不上架，先存草稿（之後補好再建立）
      const asDraft = CONFIG.mode === 'create' && (serious || noPhotos);
      if (asDraft) {
        const bad = problems.filter(isBad).map((x) => x.replace(/，請.*$/, ''));
        console.log(`   ⚠️ ${noPhotos ? '沒有照片' : `有欄位沒填好（${bad.join('；')}）`}，這筆先存草稿、不上架。到飛鼠草稿補好後按「發佈」就好`);
      }
      const needHuman = CONFIG.mode === 'review';
      if (needHuman) {
        const a = await ask(
          '   請檢查頁面、補上缺的資料，自己按「儲存草稿」或「建立」。\n' +
            '   已存好按 Enter；沒存、要跳過輸入 s；不存直接結束輸入 q：'
        );
        if (a === 'q') { log(r, '略過', '沒存就結束'); break; }
        if (a === 's') { log(r, '略過', '手動跳過'); continue; }
        log(r, '成功（手動儲存）', problems.join('；'));
      } else {
        const btn = CONFIG.mode === 'create' && !asDraft ? '建立' : '儲存草稿';
        const btns = page.locator('button, a, [role=button], input[type=button], input[type=submit]')
          .filter({ hasText: new RegExp(`^[\\s+＋🖫💾]*${btn}\\s*$`) });
        let clicked = false;
        const before = await visibleMessages(page);
        for (let k = 0, n = await btns.count(); k < n && !clicked; k++) {
          if (await btns.nth(k).isVisible().catch(() => false)) { await btns.nth(k).click({ timeout: 5000 }); clicked = true; }
        }
        if (!clicked) throw new Error(`找不到「${btn}」按鈕，這筆已填好但還沒存，請手動按「${btn}」`);
        await confirmSaved(page, btn, before);
        if (PENDING_PHOTO_RECORD) { recordPhotoUpload(PENDING_PHOTO_RECORD); PENDING_PHOTO_RECORD = ''; }
        log(r, btn === '建立' ? '成功（已建立）' : asDraft ? '成功（草稿，需補）' : '成功（草稿）', problems.join('；'));
        console.log(`   ✅ 已按「${btn}」`);
      }
      ok++;
    } catch (e) {
      fail++;
      console.log('   ❌ 失敗：' + e.message.split('\n')[0]);
      log(r, '失敗', e.message.split('\n')[0]);
      const shot = path.join(diagDir(), `錯誤截圖-${i + 1}.png`);
      await page.screenshot({ path: shot, fullPage: true }).then(
        () => console.log(`   📸 已存錯誤截圖：${shot}`),
        () => {}
      );
      if (CONFIG.mode !== 'review') continue; // 不停模式：失敗的記下來，直接做下一筆
      const a = await ask('   按 Enter 繼續下一筆，輸入 q 結束：');
      if (a === 'q') break;
    }
  }

  console.log(`\n完成：成功 ${ok} 筆、失敗 ${fail} 筆。紀錄在 ${CONFIG.logFile}`);
  if (CONFIG.mode === 'review') await ask('按 Enter 關閉瀏覽器...');
  closeAsk();
  await ctx.close();
}


// =================== 照片：黑畫面分隔（node upload.js photos）===================
//   node upload.js photos --order   列出拍照順序（存成 拍照順序.txt）
//   node upload.js photos           把 商品照片\收件匣 的照片依黑畫面分組 → 預覽 → 確認後放進 商品照片\序號
const PHOTO_DIR = path.join(__dirname, '商品照片');
// 照片收件匣：整理過資料夾（程式在「系統」裡）→ 最外層的「照片放這裡」；舊排法 → 商品照片\收件匣
const ROOT_DIR = path.basename(__dirname) === '系統' ? path.dirname(__dirname) : __dirname;
const INBOX_DIR = ROOT_DIR !== __dirname ? path.join(ROOT_DIR, '照片放這裡') : path.join(PHOTO_DIR, '收件匣');
const INBOX_NAME = ROOT_DIR !== __dirname ? '「照片放這裡」資料夾' : '商品照片\\收件匣';
const DONE_DIR = path.join(PHOTO_DIR, '已處理');
const IMG_RE = /\.(jpe?g|png)$/i;
const BLACK_MEAN = 60; // 平均亮度低於這個（0～255）…（手機／LINE 傳的照片壓縮過，黑畫面不會是全黑）
const BLACK_P90 = 95;  // …而且 90% 的像素都比這個暗（容許邊角漏一點光、雜訊），就當作黑畫面
// 每件商品大約拍幾張（只用來檢查有沒有漏拍；分組一律看黑畫面）
const PHOTO_COUNT = { camera: [6, 99], lens: [4, 99] }; // [最少, 最多]：多拍沒關係，只有太少才提醒
const expectedPhotos = (r) => (r && gearType(r) === 'lens' ? PHOTO_COUNT.lens : PHOTO_COUNT.camera);
const countText = ([a, b]) => (b >= 99 ? `至少 ${a} 張` : a === b ? `${a} 張` : `${a}～${b} 張`);
const PHOTO_MAX_SIDE = 2000; // 上傳前縮小到最長邊 2000px（飛鼠建議 1600x1600，太大傳很慢）
const safeName = (sn) => String(sn).replace(/[\\/:*?"<>|]/g, '_');
// 照片上傳紀錄：記下每個序號「上傳過哪一批照片」，避免同一批重複上傳
const PHOTO_LOG = path.join(__dirname, '照片上傳紀錄.csv');
function photoSignature(serial) {
  return productPhotos(serial).map((f) => `${path.basename(f)}:${fs.statSync(f).size}`).join('|');
}
function photoUploadedSigs() {
  const m = new Map();
  for (const c of readCsv(PHOTO_LOG)) m.set(c[1], c[2]);
  return m;
}
function recordPhotoUpload(serial) {
  if (!fs.existsSync(PHOTO_LOG)) fs.writeFileSync(PHOTO_LOG, '\uFEFF"時間","序號","照片"\n');
  const t = new Date().toLocaleString('zh-TW', { hour12: false });
  fs.appendFileSync(PHOTO_LOG, [t, serial, photoSignature(serial)].map((x) => `"${String(x).replace(/"/g, '""')}"`).join(',') + '\n');
}
const photosPending = (serial) => productPhotos(serial).length && photoUploadedSigs().get(String(serial)) !== photoSignature(serial);

function productPhotos(serial) {
  if (!serial) return [];
  const dir = path.join(PHOTO_DIR, safeName(serial));
  if (!fs.existsSync(dir)) return [];
  return fs.readdirSync(dir).filter((f) => IMG_RE.test(f)).sort((a, b) => a.localeCompare(b, 'en', { numeric: true })).map((f) => path.join(dir, f));
}
// 讀照片的拍攝時間（EXIF DateTimeOriginal），讀不到就回傳 null
function exifTime(buf) {
  try {
    if (buf[0] !== 0xff || buf[1] !== 0xd8) return null;
    let o = 2;
    while (o + 4 < buf.length) {
      if (buf[o] !== 0xff) return null;
      const marker = buf[o + 1];
      const len = buf.readUInt16BE(o + 2);
      if (marker === 0xe1 && buf.toString('latin1', o + 4, o + 10) === 'Exif\0\0') {
        const t = o + 10;
        const le = buf.toString('latin1', t, t + 2) === 'II';
        const r16 = (p) => (le ? buf.readUInt16LE(p) : buf.readUInt16BE(p));
        const r32 = (p) => (le ? buf.readUInt32LE(p) : buf.readUInt32BE(p));
        const find = (off, tag) => {
          const n = r16(t + off);
          for (let i = 0; i < n; i++) { const e = t + off + 2 + i * 12; if (r16(e) === tag) return e; }
          return null;
        };
        const ascii = (e) => { const cnt = r32(e + 4); const vo = t + r32(e + 8); return buf.toString('latin1', vo, vo + cnt - 1); };
        const ifd0 = r32(t + 4);
        const exifPtr = find(ifd0, 0x8769);
        let str = null;
        if (exifPtr !== null) { const e = find(r32(exifPtr + 8), 0x9003); if (e !== null) str = ascii(e); }
        if (!str) { const e = find(ifd0, 0x0132); if (e !== null) str = ascii(e); }
        const m = str && str.match(/(\d{4}):(\d\d):(\d\d) (\d\d):(\d\d):(\d\d)/);
        return m ? new Date(+m[1], m[2] - 1, +m[3], +m[4], +m[5], +m[6]).getTime() : null;
      }
      if (marker === 0xda) return null;
      o += 2 + len;
    }
  } catch (e) {}
  return null;
}
// 這次要拍照的商品：收購後台有標記、還沒上架的，加上「存成草稿、還沒有照片」的；順序跟上架時一樣
function lastFeishuNotes() {
  const m = new Map();
  for (const c of readCsv(CONFIG.logFile)) m.set(`${c[1]}|${c[2]}`, `${c[3] || ''} ${c[4] || ''}`);
  return m;
}
function lastFeishuStatus() {
  const m = new Map();
  for (const c of readCsv(CONFIG.logFile)) m.set(`${c[1]}|${c[2]}`, c[3] || '');
  return m;
}
// 補照片清單：已經上架的商品要補拍照時，把序號一行一個打在這個檔案裡（照拍照順序）
const REFILL_FILE = path.join(__dirname, '補照片清單.txt');
let LAST_REFILL = []; // 這次分組用到的補照片清單序號
function ensureRefillFile() {
  if (fs.existsSync(REFILL_FILE)) return;
  fs.writeFileSync(REFILL_FILE, '\uFEFF# 補照片清單：已經上架、要補拍照片的商品，照拍照順序一行打一個序號\r\n# 例如：\r\n# 70039\r\n# 008362\r\n# （# 開頭的行不會被讀取；清單是空的就照收購後台的標記）\r\n', 'utf8');
}
function refillSerials() {
  if (!fs.existsSync(REFILL_FILE)) return [];
  return fs.readFileSync(REFILL_FILE, 'utf8').replace(/^\uFEFF/, '').split(/\r?\n/).map((l) => l.trim()).filter((l) => l && !l.startsWith('#'));
}
async function rowsBySerials(serials) {
  const esc = (x) => String(x).replace(/"/g, '\\"');
  const found = await readAirtable(`OR(${serials.map((sn) => `{序號}="${esc(sn)}"`).join(',')})`);
  return serials.map((sn) => found.find((r) => r['條碼'] === sn) || { 商品名稱: `（Airtable 找不到序號 ${sn}）`, 條碼: sn, _missing: true });
}

let HAVE_PHOTOS = []; // 已經有分好照片、不用再拍的（拍照順序會另外列出來）
async function photoOrderRows() {
  // 先排「補照片清單」裡的已上架商品，再排收購後台標記、還沒上架的新商品
  const list = refillSerials();
  const refill = list.length ? await rowsBySerials(list) : [];
  refill.forEach((r) => (r._refill = true));
  const rows = await readAirtable();
  const st = lastFeishuStatus();
  HAVE_PHOTOS = rows.filter((r) => r['條碼'] && !list.includes(r['條碼']) && productPhotos(r['條碼']).length && !/已上架/.test(r._fsStatus || ''));
  const cand = rows.filter((r) => !list.includes(r['條碼'])).filter((r) => {
    if (!r['條碼'] || productPhotos(r['條碼']).length) return false;  // 已經有分好的照片就不用再拍
    if (/已上架/.test(r._fsStatus || '')) return false;               // 收購後台顯示已上架
    const s = st.get(keyOf(r)) || '';
    return !s.startsWith('成功') || s.includes('草稿');                // 還沒上架、或是草稿
  });
  // 再到官網確認：已經上架的（例如在飛鼠手動按了發佈）就不列，並補記紀錄
  const out = [];
  const sf = await storefrontMany(cand.map((r) => r['條碼']));
  for (const r of cand) {
    if (sf.get(r['條碼'])) { log(r, '成功（已建立）', '官網已有，補記紀錄'); setFeishuStatus(r['條碼'], '飛鼠已上架', r._id); continue; }
    r._isDraft = (st.get(keyOf(r)) || '').includes('草稿') || /草稿/.test(r._fsStatus || '');
    out.push(r);
  }
  await flushAirtable();
  return refill.concat(out);
}
async function photosOrderMain() {
  fs.mkdirSync(INBOX_DIR, { recursive: true });
  ensureRefillFile();
  const rows = await photoOrderRows();
  if (rows.some((r) => r._refill)) console.log(`（補照片清單 ${rows.filter((r) => r._refill).length} 件排在最前面，後面是新商品）\n`);
  rows.filter((r) => r._missing).forEach((r) => console.log(`⚠️ 補照片清單裡的序號 ${r['條碼']} 在收購紀錄找不到，請檢查有沒有打錯`));
  const lines = [`拍照順序　${new Date().toLocaleString('zh-TW', { hour12: false })}`, '',
    `照這個順序拍：相機 ${countText(PHOTO_COUNT.camera)}、鏡頭 ${countText(PHOTO_COUNT.lens)}。`,
    '每拍完一件，一定要蓋上鏡頭蓋拍 1 張「全黑」當分隔！', ''];
  rows.forEach((r, i) => { autoFill(r); lines.push(`${i + 1}. ${r['商品名稱']}　序號 ${r['條碼']}　（${gearType(r) === 'lens' ? '鏡頭' : '相機'} ${countText(expectedPhotos(r))}）${r._isDraft ? '　【飛鼠已是草稿，等補照片】' : ''}${r._refill ? '　【補照片】' : ''}`, '   → 拍完蓋鏡頭蓋拍 1 張全黑'); });
  if (HAVE_PHOTOS.length) {
    lines.push('', `📁 這 ${HAVE_PHOTOS.length} 件已經有分好的照片，不用拍（按「一鍵上架」會直接用這些照片上架）：`);
    HAVE_PHOTOS.forEach((r) => lines.push(`   ・${r['商品名稱']}　序號 ${r['條碼']}（${productPhotos(r['條碼']).length} 張）`));
    lines.push('   （想重拍的話，把「系統\\商品照片\\序號」那個資料夾刪掉，它就會回到拍照順序裡）', '');
  }
  if (!rows.length) {
    lines.push('（目前沒有要拍的商品）');
    lines.push('・新商品：請先在收購後台填售價、按「🛒 標記上架飛鼠」');
    lines.push('・已經有照片的商品不會再列出來（照片在 商品照片\\序號 資料夾）');
    lines.push('・要幫已上架的商品補照片：把序號一行一個打進「補照片清單.txt」，再雙擊一次');
  }
  lines.push('', `拍完把 SD 卡裡的照片全部放進：${INBOX_DIR}`, '再雙擊「★一鍵上架.bat」就會分組並上架');
  const text = lines.join('\n');
  console.log(text);
  fs.writeFileSync(path.join(__dirname, '拍照順序.txt'), '\uFEFF' + text.replace(/\n/g, '\r\n'), 'utf8');
}

async function photosMain() {
  fs.mkdirSync(INBOX_DIR, { recursive: true });
  const rows = await photoOrderRows();
  const all = fs.readdirSync(INBOX_DIR).filter((f) => !f.startsWith('.'));
  const files = all.filter((f) => IMG_RE.test(f));
  const skipped = all.length - files.length;
  if (!files.length) { console.log(`收件匣是空的（${INBOX_DIR}），這次沒有新照片要分組`); return null; }
  console.log(`收件匣有 ${files.length} 張照片${skipped ? `（另有 ${skipped} 個不是 JPG／PNG 的檔案，例如 RAW，會略過）` : ''}，商品 ${rows.length} 件，分析中…`);

  // 依拍攝時間排序（同一秒的照檔名排）
  const items = files.map((f) => {
    const full = path.join(INBOX_DIR, f);
    const buf = fs.readFileSync(full, { start: 0, end: 256 * 1024 });
    return { file: f, src: full, exif: exifTime(buf) };
  });
  // 排順序：每張都有拍攝時間（EXIF）→ 照拍攝時間；只要有一張沒有（例如手機傳過來、軟體匯出的）→ 照檔名的號碼
  // （檔案的「修改時間」複製後會亂掉，不能用）
  const byName = (a, b) => a.file.localeCompare(b.file, 'en', { numeric: true });
  const allExif = items.every((it) => it.exif);
  items.sort(allExif ? (a, b) => a.exif - b.exif || byName(a, b) : byName);
  console.log(allExif ? '（照拍攝時間排順序）' : '（有照片沒有拍攝時間，改照檔名號碼排順序）');

  // 用瀏覽器量亮度、做縮圖和縮小版（不用另外裝套件）
  const browser = await chromium.launch({ channel: CONFIG.browserChannel, headless: true }).catch(() => chromium.launch({ headless: true }));
  const page = await browser.newPage();
  await page.route('https://photo.local/**', (route) => {
    const name = decodeURIComponent(route.request().url().split('/').pop());
    route.fulfill({ body: fs.readFileSync(path.join(INBOX_DIR, name)), contentType: 'image/jpeg' });
  });
  await page.goto('about:blank');
  for (const [k, it] of items.entries()) {
    const res = await page.evaluate(async ({ url, max }) => {
      const blob = await (await fetch(url)).blob();
      const bmp = await createImageBitmap(blob);
      const c = document.createElement('canvas');
      c.width = 64; c.height = 64;
      const x = c.getContext('2d');
      x.drawImage(bmp, 0, 0, 64, 64);
      const d = x.getImageData(0, 0, 64, 64).data;
      const lum = [];
      for (let i = 0; i < d.length; i += 4) lum.push(0.299 * d[i] + 0.587 * d[i + 1] + 0.114 * d[i + 2]);
      const mean = lum.reduce((a, b) => a + b, 0) / lum.length;
      const p90 = lum.sort((a, b) => a - b)[Math.floor(lum.length * 0.9)];
      const scale = (m) => Math.min(1, m / Math.max(bmp.width, bmp.height));
      const t = document.createElement('canvas'); const ts = scale(220);
      t.width = Math.round(bmp.width * ts); t.height = Math.round(bmp.height * ts);
      t.getContext('2d').drawImage(bmp, 0, 0, t.width, t.height);
      const f = document.createElement('canvas'); const fs_ = scale(max);
      f.width = Math.round(bmp.width * fs_); f.height = Math.round(bmp.height * fs_);
      f.getContext('2d').drawImage(bmp, 0, 0, f.width, f.height);
      return { mean, p90, thumb: t.toDataURL('image/jpeg', 0.7), full: f.toDataURL('image/jpeg', 0.9) };
    }, { url: 'https://photo.local/' + encodeURIComponent(it.file), max: PHOTO_MAX_SIDE }).catch((e) => ({ error: e.message }));
    Object.assign(it, res);
    if ((k + 1) % 10 === 0 || k === items.length - 1) console.log(`   …${k + 1}/${items.length}`);
  }
  await browser.close();
  // 判斷黑畫面：夠暗（固定門檻），或是「比這批其他照片暗很多」（手機／LINE 傳的黑畫面常常是灰黑、有雜訊）
  const ok = items.filter((it) => !it.error);
  const med = (arr) => { const a = arr.slice().sort((x, y) => x - y); return a.length ? a[Math.floor(a.length / 2)] : 0; };
  const medMean = med(ok.map((it) => it.mean)), medP90 = med(ok.map((it) => it.p90));
  for (const it of ok) {
    it.black = (it.mean < BLACK_MEAN && it.p90 < BLACK_P90) || (it.mean < 100 && it.mean < medMean * 0.45 && it.p90 < medP90 * 0.6);
  }
  const bad = items.filter((it) => it.error);

  const blacks = items.filter((it) => it.black).length;
  if (!blacks && rows.length > 1) {
    console.log('⚠️ 照片裡沒有找到黑畫面，沒辦法分組。請每拍完一件商品，蓋上鏡頭蓋拍 1 張全黑，再跑一次。');
    console.log('   （收件匣的照片都沒有動）');
    return false;
  }
  // 依黑畫面分組（連續好幾張黑的算一個分隔；最前面的黑畫面忽略）
  const groups = [];
  let cur = [];
  for (const it of items) {
    if (it.error) continue;
    if (it.black) { if (cur.length) { groups.push(cur); cur = []; } continue; }
    cur.push(it);
  }
  if (cur.length) groups.push(cur);
  const isOdd = (g, i) => { const [a, b] = expectedPhotos(rows[i]); return g.length < a || g.length > b; };
  const oddGroups = groups.map((g, i) => (isOdd(g, i) ? i : -1)).filter((i) => i >= 0);

  // 預覽頁
  const warns = [];
  if (groups.length !== rows.length) warns.push(`⚠️ 分出 ${groups.length} 組照片，但商品有 ${rows.length} 件，數量不一樣！可能漏拍或多拍了黑畫面，請檢查。`);
  if (oddGroups.length) warns.push(`⚠️ 第 ${oddGroups.map((i) => i + 1).join('、')} 組張數不對（相機應該 ${countText(PHOTO_COUNT.camera)}、鏡頭 ${countText(PHOTO_COUNT.lens)}），可能漏拍、多拍，或順序跟清單不一樣，請檢查。`);
  const warn = warns.join('<br>');
  const esc = (t) => String(t).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  // 預覽頁（可以手動調整：每張照片選「放到第幾組／不要」、設為主圖；在控制台按「分組正確」會照調整後的結果）
  const nG = Math.max(groups.length, rows.length);
  const model = {
    rows: Array.from({ length: nG }, (_, i) => rows[i] ? { name: rows[i]['商品名稱'], sn: rows[i]['條碼'], expect: expectedPhotos(rows[i]) } : null),
    groups: Array.from({ length: nG }, (_, i) => (groups[i] || []).map((it) => it.file)),
    trash: items.filter((it) => it.black).map((it) => it.file),
    thumbs: Object.fromEntries(items.filter((it) => !it.error).map((it) => [it.file, it.thumb])),
    info: Object.fromEntries(items.filter((it) => !it.error).map((it) => [it.file, `亮度 ${Math.round(it.mean)}`])),
  };
  const html = `<!doctype html><meta charset="utf-8"><title>照片分組預覽</title>
<style>body{font-family:system-ui,"Microsoft JhengHei",sans-serif;margin:20px;background:#f6f6f4;color:#222}h1{font-size:22px}h2{font-size:17px;margin:22px 0 8px}small{color:#777;font-weight:400}
.warn{background:#fff3cd;border:1px solid #f0c36d;padding:12px 14px;border-radius:8px;font-weight:700}.tip{background:#e7f5ff;border:1px solid #a5d8ff;padding:10px 14px;border-radius:8px}.g{display:flex;flex-wrap:wrap;gap:8px;min-height:40px}
figure{margin:0;background:#fff;padding:6px;border-radius:6px;border:1px solid #ddd;width:232px}figure.main{border:3px solid #d9480f}img{display:block;max-width:220px;max-height:160px}figcaption{font-size:11px;color:#888;margin-top:4px;word-break:break-all}
select,button{font:inherit;font-size:12px;margin-top:4px}figure.sel{outline:4px solid #1f6feb;outline-offset:-2px}figure{cursor:grab}figure img{cursor:pointer}.ctl{display:flex;gap:4px;align-items:center;flex-wrap:wrap}.bar{position:sticky;top:0;z-index:5;background:#fffbe6;border:1px solid #f5d98a;border-radius:8px;padding:8px 12px;margin:8px 0;font-size:14px}button{cursor:pointer;border:1px solid #ccc;background:#fff;border-radius:6px;padding:2px 8px}.bad{color:#c00;font-weight:700}.trash figure{opacity:.6}</style>
<h1>照片分組預覽</h1><p>共 ${items.length} 張：商品照 ${items.length - blacks - bad.length} 張、黑畫面 ${blacks} 張${bad.length ? `、讀不到 ${bad.length} 張` : ''}。確認每一組都是對的商品後，${process.env.PANEL ? '按下面的「✅ 分組正確，開始上架」' : '回黑色視窗輸入 y'}。</p>
${warn ? `<p class="warn">${warn}</p>` : ''}
<p class="tip">🖐️ 分錯了可以直接調整：每張照片下面選「放到哪一組」，或選「🗑 不要」（黑畫面、多拍的）；按「⭐ 設為主圖」會變成第一張。${process.env.PANEL ? '' : '（調整功能要在「控制台」裡用）'}</p>
<div id="root"></div>
<script>
const M = ${JSON.stringify(model).replace(/</g, '\\u003c')};
const esc = (t) => String(t).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const SEL = new Set();
let DRAG = null;
const listOf = (to) => (to === 'trash' ? M.trash : M.groups[+to]);
function locate(f){ for (let i=0;i<M.groups.length;i++){ const k=M.groups[i].indexOf(f); if(k>=0) return [String(i),k]; } const k=M.trash.indexOf(f); return k>=0?['trash',k]:[null,-1]; }
function take(f){ const r=locate(f); if(r[0]!==null) listOf(r[0]).splice(r[1],1); }
function moveMany(files, to, before){ files=files.filter((f)=>f!==before); files.forEach(take); const L=listOf(to); let at = before ? L.indexOf(before) : -1; if (at<0) at=L.length; L.splice(at,0,...files); render(); }
function move(f, to){ moveMany([f], to); }
function nudge(f, d){ const r=locate(f); const L=listOf(r[0]); const j=r[1]+d; if(j<0||j>=L.length) return; const t=L[r[1]]; L[r[1]]=L[j]; L[j]=t; render(); }
function frontSel(){ const fs=[...SEL]; const groups={}; fs.forEach((f)=>{ const r=locate(f); (groups[r[0]]=groups[r[0]]||[]).push(f); }); Object.keys(groups).forEach((g)=>{ const L=listOf(g); groups[g].forEach((f)=>L.splice(L.indexOf(f),1)); L.unshift(...groups[g]); }); SEL.clear(); render(); }
function toggle(f){ if(SEL.has(f)) SEL.delete(f); else SEL.add(f); render(); }
function selectGroup(g){ const L=listOf(g); const all=L.length&&L.every((f)=>SEL.has(f)); L.forEach((f)=>{ if(all) SEL.delete(f); else SEL.add(f); }); render(); }
function applySel(to){ if(!SEL.size||to==='') return; moveMany([...SEL], to); SEL.clear(); render(); }
function dropOn(to, before){ if(!DRAG) return; const fs=SEL.has(DRAG)?[...SEL]:[DRAG]; moveMany(fs, to, before); SEL.clear(); DRAG=null; render(); }
function opts(cur, head){ return (head?'<option value="">'+head+'</option>':'') + M.groups.map((_,i)=>'<option value="'+i+'" '+(cur===String(i)?'selected':'')+'>'+(head?'':'放到：')+(i+1)+'. '+esc(M.rows[i]?M.rows[i].name:'（沒有對應商品）')+'</option>').join('') + '<option value="trash" '+(cur==='trash'?'selected':'')+'>🗑 不要（黑畫面／多拍）</option>'; }
function fig(f, k, g){
  const sel=SEL.has(f), main=k===0&&g!=='trash';
  return '<figure draggable="true" data-f="'+esc(f)+'" data-g="'+g+'" class="'+(main?'main ':'')+(sel?'sel':'')+'" ondragstart="DRAG=this.dataset.f" ondragover="event.preventDefault()" ondrop="event.preventDefault();event.stopPropagation();dropOn(this.dataset.g,this.dataset.f)">'
    + '<img src="'+M.thumbs[f]+'" data-f="'+esc(f)+'" onclick="toggle(this.dataset.f)" title="點一下選取">'
    + '<figcaption>'+(sel?'☑️ ':'')+(g!=='trash'?'第 '+(k+1)+' 張'+(main?'（主圖）':'')+'　':'')+esc(f)+'　'+esc(M.info[f]||'')+'</figcaption>'
    + '<div class="ctl"><button data-f="'+esc(f)+'" onclick="nudge(this.dataset.f,-1)" title="往前一張">◀</button><button data-f="'+esc(f)+'" onclick="nudge(this.dataset.f,1)" title="往後一張">▶</button>'
    + '<select data-f="'+esc(f)+'" onchange="move(this.dataset.f,this.value)">'+opts(String(g))+'</select></div></figure>';
}
function render(){
  let h='<div class="bar">'+(SEL.size?'已選 <b>'+SEL.size+'</b> 張：<select onchange="applySel(this.value)">'+opts('', '選取的照片放到…')+'</select> <button onclick="frontSel()">⭐ 排到最前面（當主圖）</button> <button onclick="SEL.clear();render()">取消選取</button>':'👆 點照片可以多選　✋ 拖曳照片可以排順序、換組　◀ ▶ 往前／往後一張')+'</div>';
  M.groups.forEach((g,i)=>{ const r=M.rows[i]; const a=r?r.expect[0]:0, b=r?r.expect[1]:0; const odd=r&&(g.length<a||g.length>b);
    h+='<section><h2>'+(i+1)+'. '+(r?esc(r.name)+'　<small>序號 '+esc(r.sn)+'</small>':'<span class="bad">（沒有對應的商品，這組不會使用）</span>')+'　<small class="'+(odd?'bad':'')+'">'+g.length+' 張'+(odd?'（應該'+(b>=99?'至少 '+a:(a===b?'是 '+a:'是 '+a+'～'+b))+' 張）':'')+'</small> <button onclick="selectGroup(\\''+i+'\\')">全選這組</button></h2>'
      +'<div class="g" data-g="'+i+'" ondragover="event.preventDefault()" ondrop="event.preventDefault();dropOn(this.dataset.g)">'+(g.length?g.map((f,k)=>fig(f,k,String(i))).join(''):'<span class="bad">（沒有照片，可以把照片拖到這裡）</span>')+'</div></section>'; });
  h+='<section class="trash"><h2>🗑 不使用（黑畫面、多拍的）<small>'+M.trash.length+' 張</small> <button onclick="selectGroup(\\'trash\\')">全選這組</button></h2><div class="g" data-g="trash" ondragover="event.preventDefault()" ondrop="event.preventDefault();dropOn(this.dataset.g)">'+M.trash.map((f,k)=>fig(f,k,'trash')).join('')+'</div></section>';
  document.getElementById('root').innerHTML=h;
}
window.getGrouping = () => ({ groups: M.groups });
render();
</script>`;
  const preview = path.join(__dirname, '照片分組預覽.html');
  fs.writeFileSync(preview, html, 'utf8');
  console.log(`\n分成 ${groups.length} 組（黑畫面 ${blacks} 張），商品 ${rows.length} 件`);
  if (warn) console.log(warns.join('\n'));
  console.log(`預覽：${preview}`);
  if (!process.env.PANEL) openPath(preview); // 控制台會自己在網頁裡顯示預覽

  const a = await ask(warn
    ? '\n⚠️ 上面有警告！確定還是要套用嗎？輸入 yes 套用，其他鍵取消：'
    : '\n預覽正確嗎？輸入 y 套用（照片放進各商品資料夾），其他鍵取消：', warn ? 'preview-yes' : 'preview-y');
  if (warn ? a.toLowerCase() !== 'yes' : a.toLowerCase() !== 'y') { console.log('已取消，收件匣的照片都沒有動。可以調整後再跑一次。'); return false; }

  // 在控制台手動調整過分組：照調整後的結果
  const editFile = path.join(__dirname, '照片分組修改.json');
  if (fs.existsSync(editFile)) {
    try {
      const ed = JSON.parse(fs.readFileSync(editFile, 'utf8'));
      if (fs.statSync(editFile).mtimeMs >= fs.statSync(preview).mtimeMs && Array.isArray(ed.groups)) {
        const byFile = Object.fromEntries(items.map((it) => [it.file, it]));
        groups.length = 0;
        ed.groups.forEach((g) => groups.push((g || []).map((f) => byFile[f]).filter((it) => it && !it.error)));
        console.log('   🖐️ 照你在預覽頁調整後的分組：' + groups.map((g, i) => `${i + 1}組 ${g.length} 張`).join('、'));
      }
    } catch (e) { console.log('   ⚠️ 讀不到調整後的分組，照原本的分組'); }
    try { fs.unlinkSync(editFile); } catch (e) {}
  }
  const stamp = new Date().toISOString().replace(/[-:T]/g, '').slice(0, 12);
  const backup = path.join(DONE_DIR, stamp);
  fs.mkdirSync(backup, { recursive: true });
  groups.forEach((g, i) => {
    if (!g.length) return;
    const r = rows[i];
    if (!r) return;
    const dir = path.join(PHOTO_DIR, safeName(r['條碼']));
    if (fs.existsSync(dir) && fs.readdirSync(dir).length) fs.renameSync(dir, path.join(backup, `舊照片_${safeName(r['條碼'])}`)); // 舊的先移到已處理
    fs.mkdirSync(dir, { recursive: true });
    g.forEach((it, k) => fs.writeFileSync(path.join(dir, `${String(k + 1).padStart(2, '0')}.jpg`), Buffer.from(it.full.split(',')[1], 'base64')));
    console.log(`   📁 ${r['商品名稱']} → 商品照片\\${safeName(r['條碼'])}（${g.length} 張）`);
  });
  for (const it of items) fs.renameSync(it.src, path.join(backup, it.file)); // 原始照片移到已處理（不會刪掉）
  console.log(`\n完成！原始照片已移到 ${backup}`);
  LAST_REFILL = rows.filter((r) => r._refill).map((r) => r['條碼']);
  if (rows.some((r) => r._refill) && fs.existsSync(REFILL_FILE)) {
    fs.renameSync(REFILL_FILE, path.join(backup, '補照片清單（已用）.txt'));
    ensureRefillFile();
    console.log('「補照片清單.txt」已用完，已封存到 已處理，下次重新打序號就好。');
    if (process.argv[2] !== 'all') console.log('接著雙擊「補照片.bat」把照片傳上飛鼠。');
  } else console.log('上架時會自動上傳這些照片。');
  return true;
}

// =================== 草稿補照片並發佈（node upload.js drafts）===================
// 之前存成草稿、現在 商品照片\序號 有照片的商品：打開草稿 → 上傳照片 → 按「發佈」
const isEditPage = (page) => page.getByText('商品編輯', { exact: true }).first().isVisible().catch(() => false);
async function draftTargets() {
  const st = lastFeishuStatus();
  const rows = await readAirtable();
  const cand = rows.filter((r) => r['條碼'] && (st.get(keyOf(r)) || '').includes('草稿') && photosPending(r['條碼']));
  // 紀錄寫草稿、但其實已經上架了（例如在飛鼠手動按了發佈）→ 不再處理，紀錄改成已上架
  const out = [];
  for (const r of cand) {
    if (/已上架/.test(r._fsStatus || '') || (await isOnStorefront(r['條碼']))) {
      console.log(`⏭️  ${r['商品名稱']}　序號 ${r['條碼']} 已經上架了（不是草稿），略過`);
      log(r, '成功（已建立）', '官網已有，補記紀錄');
      recordPhotoUpload(r['條碼']);
    } else out.push(r);
  }
  return out;
}
async function openDraftEditor(page, serial, { useSearch = false } = {}) {
  // 商品管理 → 商品列表 → 「草稿」分頁
  // 直接用網址打開商品列表（在編輯頁時左側選單會收起來，點選單不可靠）
  const origin = (() => { try { return new URL(page.url()).origin; } catch (e) { return ''; } })();
  let opened = false;
  if (/^https?:/.test(origin)) {
    await page.goto(origin + '/admin/product', { waitUntil: 'domcontentloaded' }).catch(() => {});
    opened = await page.locator('table, input[placeholder*="搜尋"]').first().waitFor({ timeout: 10000 }).then(() => true, () => false);
  }
  if (!opened) {
    const listLink = page.getByText('商品列表', { exact: true }).first();
    if (!(await listLink.isVisible().catch(() => false))) {
      await page.getByText('商品管理', { exact: true }).first().click({ timeout: 8000 }).catch(() => {});
      await page.waitForTimeout(800);
    }
    await listLink.click({ timeout: 10000 }).catch(() => { throw new Error('找不到左側選單的「商品列表」，請確認已登入飛鼠'); });
  }
  await page.waitForTimeout(1500);
  if (useSearch) {
    // 已上架的商品很多，用上方的搜尋框找序號
    const box = page.getByPlaceholder(/搜尋/).first();
    if (await box.isVisible().catch(() => false)) {
      await box.fill(String(serial));
      await box.press('Enter');
      await page.waitForTimeout(2500);
    }
  } else {
    const tab = page.getByText(/^\s*草稿\s*[（(]\s*\d+\s*[)）]\s*$/).first();
    if (await tab.isVisible().catch(() => false)) { await tab.click(); await page.waitForTimeout(1500); }
  }
  // 找 SPU 完全一樣的那一列，點商品名稱打開編輯頁
  // 只認表格裡「文字剛好等於序號」的那一格，連結也只在同一列（tr）裡找
  const cell = page.locator('xpath=//tr//*[normalize-space(text())=' + JSON.stringify(String(serial)) + ']').first();
  if (!(await cell.waitFor({ timeout: 10000 }).then(() => true, () => false))) throw new Error(`商品列表${useSearch ? '搜尋' : '的草稿裡'}找不到 SPU ${serial}`);
  const sameRows = await page.locator('xpath=//tr[.//*[normalize-space(text())=' + JSON.stringify(String(serial)) + ']]').count();
  if (sameRows > 1) throw new Error(`商品列表有 ${sameRows} 筆都是 SPU ${serial}，請到飛鼠手動處理`);
  const row = cell.locator('xpath=ancestor::tr[1]');
  const link = row.locator('a').first();
  const popup = page.context().waitForEvent('page', { timeout: 5000 }).catch(() => null);
  if (await link.count()) await link.click({ timeout: 8000 });
  else await row.click({ timeout: 8000 });
  const p = (await popup) || page;
  if (!(await p.getByText('商品編輯', { exact: true }).first().waitFor({ timeout: 15000 }).then(() => true, () => false))) {
    throw new Error('沒有打開商品編輯頁');
  }
  await p.waitForTimeout(1000);
  // 安全檢查：編輯頁的 SPU 編號一定要跟序號一樣，不一樣就不動（避免照片傳到別的商品）
  const spuBox = await visibleFieldAfter(p, 'SPU 編號').catch(() => null);
  const spu = spuBox ? String(await spuBox.inputValue().catch(() => '')).trim() : '';
  if (spu !== String(serial)) {
    if (p !== page) await p.close().catch(() => {});
    throw new Error(`打開的商品 SPU 是「${spu || '讀不到'}」，不是 ${serial}，已停止，什麼都沒改`);
  }
  return p;
}
async function draftsMain(opts = {}) {
  const targets = await draftTargets();
  if (!targets.length) { console.log('沒有要補照片的草稿（草稿的照片要先放在 商品照片\\序號，可以用「0_照片分組」整理）'); return 0; }
  console.log(`要補照片並發佈的草稿 ${targets.length} 件：`);
  targets.forEach((r, i) => console.log(`  ${i + 1}. ${r['商品名稱']}　序號 ${r['條碼']}（${productPhotos(r['條碼']).length} 張）`));
  const ctx = await launchBrowser({ headless: false, channel: CONFIG.browserChannel, viewport: null, args: ['--start-maximized'] });
  const page = ctx.pages()[0] || (await ctx.newPage());
  page.on('dialog', (d) => d.accept().catch(() => {}));
  let loggedIn = false;
  if (fs.existsSync(NEW_URL_FILE)) {
    const url = fs.readFileSync(NEW_URL_FILE, 'utf8').trim();
    if (url) {
      await page.goto(url, { waitUntil: 'domcontentloaded' }).catch(() => {});
      loggedIn = await page.getByText('商品管理', { exact: true }).first().waitFor({ timeout: 15000 }).then(() => true, () => false);
    }
  } else if (CONFIG.startUrl) await page.goto(CONFIG.startUrl).catch(() => {});
  if (loggedIn) console.log('✅ 飛鼠已經是登入狀態，直接開始');
  else await ask('👉 請在瀏覽器登入飛鼠後台，看到左側選單後，回這裡按 Enter 開始...');

  let ok = 0;
  for (const [i, r] of targets.entries()) {
    console.log(`\n[${i + 1}/${targets.length}] ${r['商品名稱']}　序號 ${r['條碼']}`);
    let p = page;
    try {
      p = await openDraftEditor(page, r['條碼']);
      const problems = [];
      const have = await existingImageCount(p);
      if (have) console.log(`   ⏭️ 飛鼠上已經有 ${have} 張商品圖片，不重複上傳，直接發佈`);
      else {
        await uploadImages(p, productPhotos(r['條碼']), problems);
        if (problems.length) throw new Error(problems.join('；'));
      }
      recordPhotoUpload(r['條碼']); // 傳完馬上記下來，就算後面發佈沒確認到，下次也不會重傳
      const btns = p.locator('button, a, [role=button], input[type=button], input[type=submit]').filter({ hasText: /^[\s✓✔]*發佈\s*$/ });
      let clicked = false;
      const before = await visibleMessages(p);
      for (let k = 0, n = await btns.count(); k < n && !clicked; k++) {
        if (await btns.nth(k).isVisible().catch(() => false)) { await btns.nth(k).click({ timeout: 5000 }); clicked = true; }
      }
      if (!clicked) throw new Error('照片已上傳，但找不到「發佈」按鈕，請手動按');
      await confirmSaved(p, '發佈', before, isEditPage);
      recordPhotoUpload(r['條碼']);
      log(r, '成功（已發佈）', '草稿補照片後發佈');
      console.log('   ✅ 已補照片並發佈');
      ok++;
    } catch (e) {
      const msg = e.message.split('\n')[0];
      console.log('   ❌ 失敗：' + msg);
      await p.screenshot({ path: path.join(diagDir(), `草稿錯誤截圖-${i + 1}.png`), fullPage: true }).catch(() => {});
    } finally {
      if (p !== page) await p.close().catch(() => {});
    }
  }
  console.log(`\n草稿補照片：成功 ${ok} 件、失敗 ${targets.length - ok} 件`);
  await ctx.close(); // 做完自動關掉瀏覽器（失敗的有存截圖）
  return ok;
}

// =================== 單獨補照片（node upload.js addphotos [序號…]）===================
// 已經在飛鼠（草稿或已上架）、商品照片\序號 有「還沒傳過的照片」的：搜尋 → 打開編輯 → 上傳 → 按「發佈／更新／儲存」
async function addPhotosMain(opts = {}) {
  const args = opts.keepQuiet ? [] : process.argv.slice(3).filter((a) => !a.startsWith('--'));
  const st = lastFeishuStatus();
  const rows = await readAirtable().catch(() => []);
  let targets;
  if (args.length) {
    targets = args.map((sn) => rows.find((r) => r['條碼'] === sn) || { 商品名稱: '', 條碼: sn }).filter((r) => productPhotos(r['條碼']).length);
  } else {
    // 商品照片 底下每個序號資料夾，只要有「還沒傳過」的照片就處理；
    // 收購後台有標記、但還沒上架過的不在這裡處理（那些會在上架時一起傳）
    const notes = lastFeishuNotes();
    const notYetListed = new Set(rows.filter((r) => !(st.get(keyOf(r)) || '').startsWith('成功')
      && !/草稿|已上架/.test(r._fsStatus || '')                       // 收購後台顯示已在飛鼠
      && !/重複|已存在|已被使用/.test(notes.get(keyOf(r)) || '')).map((r) => r['條碼'])); // 建立時說編號重複 = 飛鼠上已經有（多半是草稿）
    const dirs = fs.existsSync(PHOTO_DIR) ? fs.readdirSync(PHOTO_DIR, { withFileTypes: true }).filter((d) => d.isDirectory() && !['收件匣', '已處理'].includes(d.name)).map((d) => d.name) : [];
    let pending = dirs.filter((sn) => photosPending(sn) && !notYetListed.has(sn));
    if (opts.keepQuiet) pending = pending.filter((sn) => LAST_REFILL.includes(sn)); // 一鍵上架：只補「補照片清單」裡的
    targets = pending.length ? await rowsBySerials(pending) : [];
  }
  if (!targets.length) {
    if (opts.keepQuiet) { console.log('沒有要補照片的已上架商品。'); return; }
    console.log('沒有要補照片的商品。');
    console.log('（只會處理：已經在飛鼠上、而且 商品照片\\序號 裡有「還沒傳過」的照片；也可以指定序號：node upload.js addphotos 70039）');
    return;
  }
  console.log(`要補照片的商品 ${targets.length} 件：`);
  targets.forEach((r, i) => console.log(`  ${i + 1}. ${r['商品名稱'] || ''}　序號 ${r['條碼']}（${productPhotos(r['條碼']).length} 張）`));
  const ctx = await launchBrowser({ headless: false, channel: CONFIG.browserChannel, viewport: null, args: ['--start-maximized'] });
  const page = ctx.pages()[0] || (await ctx.newPage());
  page.on('dialog', (d) => d.accept().catch(() => {}));
  let loggedIn = false;
  const url = fs.existsSync(NEW_URL_FILE) ? fs.readFileSync(NEW_URL_FILE, 'utf8').trim() : CONFIG.startUrl;
  if (url) {
    await page.goto(url, { waitUntil: 'domcontentloaded' }).catch(() => {});
    loggedIn = await page.getByText('商品管理', { exact: true }).first().waitFor({ timeout: 15000 }).then(() => true, () => false);
  }
  if (loggedIn) console.log('✅ 飛鼠已經是登入狀態，直接開始');
  else await ask('👉 請在瀏覽器登入飛鼠後台，看到左側選單後，回這裡按 Enter 開始...');

  let ok = 0, published = 0;
  for (const [i, r] of targets.entries()) {
    console.log(`\n[${i + 1}/${targets.length}] ${r['商品名稱'] || ''}　序號 ${r['條碼']}`);
    let p = page;
    try {
      p = await openDraftEditor(page, r['條碼'], { useSearch: true });
      const problems = [];
      const have = await existingImageCount(p);
      if (have && !process.argv.includes('--force')) {
        recordPhotoUpload(r['條碼']);
        console.log(`   ⏭️ 飛鼠上已經有 ${have} 張商品圖片，不重複上傳（要換照片：先到飛鼠刪掉舊照片，再補一次）`);
        ok++;
        continue;
      }
      await uploadImages(p, productPhotos(r['條碼']), problems);
      if (problems.length) throw new Error(problems.join('；'));
      recordPhotoUpload(r['條碼']); // 傳完馬上記下來，就算後面儲存沒確認到，下次也不會重傳
      // 不按「儲存草稿」（已上架的會變回草稿）；依序找 發佈 → 更新 → 儲存
      let clicked = '';
      const before = await visibleMessages(p);
      for (const word of ['發佈', '更新', '儲存']) {
        const btns = p.locator('button, a, [role=button], input[type=button], input[type=submit]').filter({ hasText: new RegExp(`^[\\s✓✔💾🖫]*${word}\\s*$`) });
        for (let k = 0, n = await btns.count(); k < n && !clicked; k++) {
          if (await btns.nth(k).isVisible().catch(() => false)) { await btns.nth(k).click({ timeout: 5000 }); clicked = word; }
        }
        if (clicked) break;
      }
      if (!clicked) throw new Error('照片已上傳，但找不到「發佈／更新／儲存」按鈕，請手動按');
      await confirmSaved(p, clicked, before, isEditPage);
      recordPhotoUpload(r['條碼']);
      if ((st.get(keyOf(r)) || '').includes('草稿') && clicked === '發佈') { log(r, '成功（已發佈）', '補照片後發佈'); published++; }
      console.log(`   ✅ 已補照片並按「${clicked}」`);
      ok++;
    } catch (e) {
      console.log('   ❌ 失敗：' + e.message.split('\n')[0]);
      await p.screenshot({ path: path.join(diagDir(), `補照片錯誤截圖-${i + 1}.png`), fullPage: true }).catch(() => {});
    } finally {
      if (p !== page) await p.close().catch(() => {});
    }
  }
  console.log(`\n補照片：成功 ${ok} 件、失敗 ${targets.length - ok} 件`);
  await ctx.close();
  if (!opts.keepQuiet && published) { if (CONFIG.shopee.autoAfterListing) await waitThenShopee(); else shopeeLaterNote(); }
  return published;
}

// =================== 一鍵上架（node upload.js all）===================
// 照片分組 → 有照片的直接「建立」上架（沒照片或有問題的存草稿）→ 等同步 → 蝦皮補資料
// 飛鼠建立後，蝦皮狀態常常是橘色勾勾、沒有真的傳到蝦皮：打開商品編輯頁再按一次「發佈」，就會推到蝦皮
async function clickPublish(p) {
  const before = await visibleMessages(p);
  // 找畫面上看得到的按鈕：草稿要按「發佈」；已上架的按「發佈／更新」；只有在沒有「儲存草稿」的頁面才會按「儲存」
  const pick = await p.evaluate(() => {
    const vis = (el) => { const r = el.getBoundingClientRect(); const st = getComputedStyle(el); return r.width > 0 && r.height > 0 && st.visibility !== 'hidden' && st.display !== 'none'; };
    const els = [...document.querySelectorAll('button, a, [role=button], input[type=button], input[type=submit]')].filter(vis);
    const txt = (el) => (el.value || el.textContent || '').replace(/[\s✓✔💾🖫]/g, '');
    const texts = [...new Set(els.map(txt).filter((t) => t && t.length <= 8))];
    const hasDraftBtn = texts.includes('儲存草稿');
    const order = hasDraftBtn ? ['發佈'] : ['發佈', '更新', '儲存'];
    for (const w of order) {
      const el = els.find((e) => txt(e) === w);
      if (el) { document.querySelectorAll('[data-pub-pick]').forEach((x) => x.removeAttribute('data-pub-pick')); el.setAttribute('data-pub-pick', '1'); return { word: w, texts }; }
    }
    return { word: '', texts };
  }).catch(() => ({ word: '', texts: [] }));
  if (!pick.word) { console.log('   （畫面上的按鈕：' + pick.texts.join('、') + '）'); return ''; }
  await p.locator('[data-pub-pick="1"]').first().click({ timeout: 5000 });
  await confirmSaved(p, pick.word, before, isEditPage).catch(() => {}); // 編輯頁存完常常沒有變化，不算失敗
  return pick.word;
}
async function republishMain(serials) {
  serials = (serials || process.argv.slice(3).filter((a) => !a.startsWith('--'))).filter(Boolean);
  if (!serials.length) { console.log('沒有要重新發佈的商品。'); return 0; }
  console.log(`要重新發佈（推到蝦皮）的商品 ${serials.length} 件：${serials.join('、')}`);
  const priceOf = Object.fromEntries((await rowsBySerials(serials).catch(() => [])).map((r) => [r['條碼'], { price: r['銷售價'] }]));
  const ctx = await launchBrowser({ headless: false, channel: CONFIG.browserChannel, viewport: null, args: ['--start-maximized'] });
  const page = ctx.pages()[0] || (await ctx.newPage());
  page.on('dialog', (d) => d.accept().catch(() => {}));
  const url = fs.existsSync(NEW_URL_FILE) ? fs.readFileSync(NEW_URL_FILE, 'utf8').trim() : CONFIG.startUrl;
  let loggedIn = false;
  if (url) {
    await page.goto(url, { waitUntil: 'domcontentloaded' }).catch(() => {});
    loggedIn = await page.getByText('商品管理', { exact: true }).first().waitFor({ timeout: 15000 }).then(() => true, () => false);
  }
  if (!loggedIn) await ask('👉 請在瀏覽器登入飛鼠後台，看到左側選單後，回這裡按 Enter 開始...');
  let ok = 0;
  for (const [i, sn] of serials.entries()) {
    console.log(`\n[${i + 1}/${serials.length}] 序號 ${sn}`);
    let p = page;
    try {
      p = await openDraftEditor(page, sn, { useSearch: true });
      const fixMsg = await fixFeishuShopeePrice(p, (priceOf[sn] || {}).price).catch(() => '');
      if (fixMsg) console.log('   💰 ' + fixMsg);
      const w = await clickPublish(p);
      if (!w) throw new Error('找不到「發佈」按鈕，請手動點進去按發佈');
      // 確認真的上架了（官網查得到）；還是草稿的話通常是有必填欄位沒填，飛鼠不讓發佈
      let live = false;
      for (let t = 0; t < 4 && !live; t++) { await new Promise((r) => setTimeout(r, 5000)); live = !!(await isOnStorefront(sn, { fresh: true })); }
      if (!live) throw new Error(`按了「${w}」，但官網還看不到，可能還是草稿（有必填欄位沒填？）。請點進去看飛鼠的紅字提示，補好再按發佈`);
      console.log(`   ✅ 已重新按「${w}」，官網已上架，2～5 分鐘內會推到蝦皮`);
      ok++;
    } catch (e) {
      console.log('   ❌ 失敗：' + e.message.split('\n')[0]);
    } finally {
      if (p !== page) await p.close().catch(() => {});
    }
  }
  await ctx.close();
  console.log(`\n重新發佈：成功 ${ok} 件、失敗 ${serials.length - ok} 件`);
  return ok;
}
function shopeeLaterNote() {
  console.log('\n===== 飛鼠上架完成 =====');
  console.log('蝦皮這次不會自動跑。請等 5～10 分鐘，到飛鼠／蝦皮確認商品都已經推到蝦皮了，');
  console.log('再到控制台按「🦐 蝦皮補資料＋價格檢查」（或雙擊「蝦皮價格檢查.bat」）。');
}
// 等飛鼠同步到蝦皮，再自動補蝦皮資料（終端機不要關）
async function waitThenShopee() {
  const wait = 150;
  console.log(`\n===== 等 ${wait / 60} 分半讓飛鼠同步到蝦皮（視窗不要關）… =====`);
  for (let t = wait; t > 0; t -= 30) { console.log(`   還有 ${t} 秒`); await new Promise((r) => setTimeout(r, Math.min(30, t) * 1000)); }
  console.log('\n===== 蝦皮補資料＋價格檢查（保存狀況、NCC／BSMI、價格、最低購買數量）=====');
  if (!process.argv.includes('--auto')) process.argv.push('--auto');
  await shopeeMain();
}
async function hasReadyPhotos() {
  const st = lastFeishuStatus();
  return (await readAirtable()).some((r) => r['條碼'] && !(st.get(keyOf(r)) || '').startsWith('成功') && productPhotos(r['條碼']).length);
}
async function allMain() {
  console.log('===== ① 照片分組 =====');
  const applied = await photosMain();
  if (applied === false) { console.log('照片分組取消了，一鍵上架停止。'); return; }
  if (applied === null) {
    // 收件匣沒新照片：如果有商品已經分好照片（商品照片\序號），就直接上架那些
    const st0 = lastFeishuStatus();
    const ready = (await readAirtable()).filter((r) => r['條碼'] && !(st0.get(keyOf(r)) || '').startsWith('成功') && productPhotos(r['條碼']).length);
    if (ready.length) {
      console.log(`\n收件匣沒有新照片，但有 ${ready.length} 件已經分好照片，直接上架：`);
      ready.forEach((r, i) => console.log(`  ${i + 1}. ${r['商品名稱']}　序號 ${r['條碼']}（${productPhotos(r['條碼']).length} 張）`));
    }
  }
  if (applied === null && !(await hasReadyPhotos())) {
    // 收件匣沒照片、也沒有分好照片的商品 → 這次只告訴你拍照順序，什麼都不上架
    console.log('\n📷 收件匣還沒有照片，先告訴你拍照順序（什麼都還沒上架）：\n');
    await photosOrderMain();
    openPath(path.join(__dirname, '拍照順序.txt'));
    console.log('\n👉 照上面的順序拍（每件拍完蓋鏡頭蓋拍 1 張全黑），');
    console.log(`   照片放進 ${INBOX_NAME} 後，再按一次「一鍵上架」就會開始上架。`);
    return;
  }
  console.log('\n===== ② 飛鼠上架（有照片的直接建立）=====');
  CONFIG.mode = 'create';
  CONFIG.closeWhenDone = true;
  CONFIG.onlyWithPhotos = true; // 一鍵上架只處理有照片的商品，沒照片的完全不動（之後拍好再一鍵上架）
  const logBefore = readCsv(CONFIG.logFile).length;
  console.log('\n===== ②-1 之前存成草稿的：補照片並發佈 =====');
  await draftsMain({ keepQuiet: true });
  console.log('\n===== ②-2 新商品：直接建立 =====');
  await main('airtable');
  console.log('\n===== ②-3 已上架的商品：補照片（補照片清單）=====');
  await addPhotosMain({ keepQuiet: true });
  // 這次新上架的（不含「官網已有，補記紀錄」）
  const created = [...new Set(readCsv(CONFIG.logFile).slice(logBefore)
    .filter((c) => /^成功（已(建立|發佈)）/.test(c[3] || '') && !/官網已有/.test(c[4] || '')).map((c) => c[2]).filter(Boolean))];
  if (!created.length) { console.log('\n這次沒有新上架的商品，蝦皮補資料就不用跑了。'); return; }
  console.log('\n===== ③ 重新發佈一次，讓飛鼠真的推到蝦皮 =====');
  await new Promise((r) => setTimeout(r, 20000)); // 先等飛鼠存好
  await republishMain(created);
  if (!CONFIG.shopee.autoAfterListing) { shopeeLaterNote(); return; }
  await waitThenShopee();
  console.log('\n===== 全部完成 =====');
  console.log('存成草稿的商品（沒照片或有欄位要補），雙擊「待辦清單.bat」可以看到。');
}

// =================== 待辦清單（node upload.js todo）===================
// 列出：① 已標記但還沒上架　② 存成草稿但還沒按建立　③ 已建立但蝦皮還沒補資料
// 官網查得到這個序號＝已上架。結果會記 30 分鐘（有上架的）／5 分鐘（還沒上架的），重複查不用再等
const SF_CACHE_FILE = path.join(__dirname, '官網查詢快取.json');
let SF_CACHE = null;
function sfCache() {
  if (!SF_CACHE) { try { SF_CACHE = JSON.parse(fs.readFileSync(SF_CACHE_FILE, 'utf8')); } catch (e) { SF_CACHE = {}; } }
  return SF_CACHE;
}
async function isOnStorefront(serial, { fresh = false } = {}) {
  const c = sfCache()[serial];
  if (!fresh && c && Date.now() - c.t < (c.live ? 30 : 5) * 60 * 1000) return c.live;
  try {
    const res = await fetchRetry(`https://linscam.fslol.com/product/${encodeURIComponent(serial)}/x`, { headers: { 'User-Agent': 'Mozilla/5.0' } }, 2);
    const live = res.status === 200 && (await res.text()).includes(`【${serial}】`);
    sfCache()[serial] = { live, t: Date.now() };
    try { fs.writeFileSync(SF_CACHE_FILE, JSON.stringify(SF_CACHE)); } catch (e) {}
    return live;
  } catch (e) { return null; } // 查不到（網路問題）
}
// 一次查很多件：最多 6 件同時查
async function storefrontMany(serials) {
  const out = new Map(), queue = [...new Set(serials.filter(Boolean))];
  await Promise.all(Array.from({ length: Math.min(6, queue.length) }, async () => {
    while (queue.length) { const sn = queue.shift(); out.set(sn, await isOnStorefront(sn)); }
  }));
  return out;
}

// =================== 控制台用：每件商品的狀態（node upload.js status）===================
async function statusMain() {
  const rows = await readAirtable();
  const last = (file, keyFn) => { const m = new Map(); for (const c of readCsv(file)) m.set(keyFn(c), { status: c[3] || '', note: c[4] || '', time: c[0] }); return m; };
  const feishu = last(CONFIG.logFile, (c) => `${c[1]}|${c[2]}`);
  const shopee = last(path.join(__dirname, '蝦皮補資料紀錄.csv'), (c) => c[2]);
  const refill = refillSerials();
  const out = [];
  await Promise.all(rows.map(async (r) => {
    autoFill(r);
    const sn = r['條碼'];
    const f = feishu.get(keyOf(r)), sp = shopee.get(sn);
    const live = !sn ? false : /已上架/.test(r._fsStatus || '') ? true : await isOnStorefront(sn); // 收購後台已經是「已上架」就不用再查官網
    let stage, hint;
    const miss = validate(r);
    if (!sn) { stage = 'error'; hint = '沒有序號'; }
    else if (live) {
      if (/蝦皮已補/.test(r._fsStatus || '')) { stage = 'done'; hint = '全部完成'; }
      else if (sp && sp.status.includes('🚨')) { stage = 'danger'; hint = sp.note.slice(0, 80); }
      else { stage = 'shopee'; hint = sp ? `蝦皮上次${sp.status}：${sp.note.slice(0, 50)}` : '蝦皮資料還沒確認（保存狀況、NCC、價格、最低購買數量）→ 按「💰 蝦皮價格檢查」'; }
    } else if (f && f.status.startsWith('成功') || /草稿/.test(r._fsStatus || '')) { stage = 'draft'; hint = '飛鼠草稿（官網還看不到）'; }
    else if (miss.length) { stage = 'error'; hint = '缺 ' + miss.join('、'); }
    else { stage = 'new'; hint = f ? `上次${f.status}：${f.note.slice(0, 50)}` : '還沒上架'; }
    // 收購後台狀態沒跟上的，順便補寫回去
    if (live && sn) {
      // 只補寫「飛鼠已上架」；「蝦皮已補」只有蝦皮真的存成功時才會寫
      if (!/已上架/.test(r._fsStatus || '')) setFeishuStatus(sn, '飛鼠已上架', r._id);
    }
    // NCC／BSMI 對照表裡沒有的相機（鏡頭、配件不用）：CCD 只要 BSMI，其他相機要 NCC
    const gt = gearType(r), isCam = ['mirrorless', 'dslr', 'compact'].includes(gt);
    const ce = lookupNcc(r['商品名稱']), isCcd = /CCD/.test(r['Y2K'] || '');
    // 缺哪些字號（相機才要；CCD 只要 BSMI）
    const needNcc = isCam && !isCcd && !(ce && (ce.ncc || isNoRadio(ce)));
    const needBsmi = isCam && !bsmiFor(r['商品名稱'], r['公/水']);
    const certMissing = [needNcc ? 'NCC' : '', needBsmi ? 'BSMI' : ''].filter(Boolean).join('、');
    if (!String(r['商品介紹'] || '').trim() && stage !== 'done') hint += '　⚠️ 缺商品介紹（到收購後台按「印卡」讓 AI 補）';
    out.push({ certMissing, model: String(r['商品名稱'] || '').split('+')[0].trim(), name: r['商品名稱'], serial: sn, price: r['銷售價'], type: gearType(r) === 'lens' ? '鏡頭' : '相機',
      photos: sn ? productPhotos(sn).length : 0, pending: sn ? photosPending(sn) : false, stage, hint, refill: refill.includes(sn), date: r['交易日期'] || '' });
  }));
  out.sort((a, b) => String(a.date).localeCompare(String(b.date)));
  await flushAirtable();
  console.log('__STATUS__' + JSON.stringify({ writeOff: AT_WRITE_OFF, items: out, refill, inbox: fs.existsSync(INBOX_DIR) ? fs.readdirSync(INBOX_DIR).filter((x) => IMG_RE.test(x)).length : 0 }));
}
// 控制台用：蝦皮已經手動改好 → 記成完成（node upload.js markshopee 序號…）
async function markShopeeMain() {
  const serials = process.argv.slice(3).filter((a) => !a.startsWith('--'));
  const rows = await readAirtable().catch(() => []);
  for (const sn of serials) {
    const r = rows.find((x) => x['條碼'] === sn) || {};
    shopeeLog(r['商品名稱'] || '', sn, '成功（手動確認）', '在控制台按「已手動補好」');
    console.log(`✅ ${r['商品名稱'] || ''} 序號 ${sn}：蝦皮記成已補好`);
  }
}
async function todoMain() {
  const rows = await readAirtable();
  const lastStatus = (file, keyFn) => {
    const m = new Map();
    for (const c of readCsv(file)) m.set(keyFn(c), { status: c[3] || '', note: c[4] || '', time: c[0] });
    return m;
  };
  const feishu = lastStatus(CONFIG.logFile, (c) => `${c[1]}|${c[2]}`);
  const shopee = lastStatus(path.join(__dirname, '蝦皮補資料紀錄.csv'), (c) => c[2]);

  const notListed = [], draft = [], needShopee = [], done = [];
  console.log(`檢查 ${rows.length} 筆已標記的商品（會順便到飛鼠官網確認有沒有建立）…\n`);
  for (const r of rows) {
    autoFill(r);
    const tag = `${r['商品名稱']}｜序號 ${r['條碼'] || '（沒有序號）'}${r['銷售價'] ? `｜售價 ${r['銷售價']}` : ''}`;
    const f = feishu.get(keyOf(r));
    if (!f || !f.status.startsWith('成功')) {
      const miss = validate(r);
      const why = miss.length ? `缺 ${miss.join('、')}` : f ? `上次${f.status}：${f.note.slice(0, 40)}` : '還沒跑上架';
      notListed.push(`${tag}　→ ${why}`);
      continue;
    }
    const live = await isOnStorefront(r['條碼']);
    if (live === false) { draft.push(`${tag}　→ 飛鼠官網還看不到，可能還是草稿`); continue; }
    const sp = shopee.get(r['條碼']);
    if (!sp || !sp.status.startsWith('成功')) {
      needShopee.push(`${tag}　→ ${sp ? `上次${sp.status}：${sp.note.slice(0, 40)}` : '還沒補蝦皮資料'}${live === null ? '（官網連不上，沒確認）' : ''}`);
      continue;
    }
    done.push(tag);
  }
  const section = (title, list, hint) => [
    `■ ${title}（${list.length} 筆）${list.length ? '　' + hint : ''}`,
    ...(list.length ? list.map((x, i) => `   ${i + 1}. ${x}`) : ['   （沒有）']),
    '',
  ];
  const lines = [
    `待辦清單　${new Date().toLocaleString('zh-TW', { hour12: false })}`,
    '',
    ...section('① 已標記但還沒上架', notListed, '→ 補好資料後雙擊「其他功能\\1_飛鼠上架（存草稿）.bat」'),
    ...section('② 存成草稿但還沒按建立', draft, '→ 到飛鼠補照片、按「建立」'),
    ...section('③ 已建立但蝦皮還沒補資料', needShopee, '→ 雙擊「其他功能\\2_蝦皮補資料.bat」'),
    `■ 全部完成（${done.length} 筆）`,
    '',
    '※ 做完的商品可以到收購後台按「取消上架標記」，清單會比較乾淨',
  ];
  const text = lines.join('\n');
  console.log(text);
  const f = path.join(__dirname, '待辦清單.txt');
  fs.writeFileSync(f, '\uFEFF' + text.replace(/\n/g, '\r\n'), 'utf8');
  console.log(`\n（也存了一份在 ${f}）`);
}

// =================== 蝦皮「最低購買數量」改 3（node upload.js shopee）===================
// 售價 15000 以上、飛鼠已上架成功的商品，到蝦皮賣家中心把最低購買數量改成 3
//   node upload.js shopee              自動找商品
//   node upload.js shopee 70039        只處理指定序號
//   node upload.js shopee --auto       改好自動按「更新」
const SHOPEE = {
  listUrl: 'https://seller.shopee.tw/portal/product/list/all',
  minQty: '3',
  logFile: path.join(__dirname, '蝦皮補資料紀錄.csv'),
  condition: '二手',   // 商品保存狀況
};

function loadDoneSerials() {
  const done = new Set();
  if (!fs.existsSync(SHOPEE.logFile)) return done;
  for (const line of fs.readFileSync(SHOPEE.logFile, 'utf8').split(/\r?\n/).slice(1)) {
    const c = line.split('","').map((s) => s.replace(/^"|"$/g, ''));
    if (c[3] && c[3].startsWith('成功')) done.add(c[2]);
  }
  return done;
}
function shopeeLog(name, serial, status, note = '') {
  if (/^成功/.test(status)) setFeishuStatus(serial, '飛鼠已上架＋蝦皮已補');
  if (!fs.existsSync(SHOPEE.logFile)) fs.writeFileSync(SHOPEE.logFile, '﻿"時間","商品名稱","序號","結果","備註"\n');
  const t = new Date().toLocaleString('zh-TW', { hour12: false });
  const esc = (s) => String(s).replace(/"/g, '""');
  fs.appendFileSync(SHOPEE.logFile, [t, name, serial, status, note].map((s) => `"${esc(s)}"`).join(',') + '\n');
}

// 飛鼠上架紀錄裡成功過的「商品名稱|序號」
function loadFeishuDone() {
  const done = new Set();
  if (!fs.existsSync(CONFIG.logFile)) return done;
  for (const line of fs.readFileSync(CONFIG.logFile, 'utf8').split(/\r?\n/).slice(1)) {
    const c = line.split('","').map((s) => s.replace(/^"|"$/g, ''));
    if (c[3] && c[3].startsWith('成功')) done.add(`${c[1]}|${c[2]}`);
  }
  return done;
}

async function findTargets() {
  const args = process.argv.slice(3).filter((a) => !a.startsWith('--'));
  const rows = await readAirtable();
  const mk = (r) => ({ name: r['商品名稱'], serial: r['條碼'], price: Number(r['銷售價']) || 0, ccd: /CCD/.test(r['Y2K'] || ''), origin: r['公/水'] || '', cam: ['mirrorless', 'dslr', 'compact'].includes(gearType(r)), fsStatus: r._fsStatus || '' });
  if (args.length) {
    // 指定序號：盡量從 Airtable 補上名稱、售價
    return args.map((sn) => { const r = rows.find((x) => x['條碼'] === sn); return r ? mk(r) : { name: '', serial: sn, price: 0 }; });
  }
  const feishuDone = loadFeishuDone();
  // 飛鼠已上架：上架紀錄有成功、收購後台顯示已上架、或官網查得到（在飛鼠手動上架的也算）
  const need = rows.filter((r) => r['條碼'] && !feishuDone.has(keyOf(r)) && !/已上架/.test(r._fsStatus || ''));
  const sf = await storefrontMany(need.map((r) => r['條碼']));
  return rows.filter((r) => r['條碼'] && (feishuDone.has(keyOf(r)) || /已上架/.test(r._fsStatus || '') || sf.get(r['條碼']))).map(mk);
}

async function openEditBySerial(ctx, page, serial) {
  await page.goto(SHOPEE.listUrl, { waitUntil: 'domcontentloaded' });
  const search = page.getByPlaceholder(/主商品貨號/).first();
  await search.waitFor({ timeout: 20000 }).catch(() => {
    throw new Error('找不到蝦皮「我的商品」的搜尋框，請確認已登入蝦皮賣家中心');
  });
  await search.fill(serial);
  await search.press('Enter');
  const apply = page.getByText('套用', { exact: true }).first();
  if (await apply.isVisible().catch(() => false)) await apply.click();
  await page.waitForTimeout(3000);

  // 找到「主商品貨號: 序號」完全相同的那一筆
  const esc = serial.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const tag = page.getByText(new RegExp(`主商品貨號[:：]\\s*${esc}\\s*$`)).first();
  if (!(await tag.isVisible().catch(() => false))) {
    throw new Error(`蝦皮找不到主商品貨號 ${serial}（飛鼠可能還在同步，等 2 分鐘再跑一次）`);
  }
  // 1) 從同一列讀「商品 ID」，直接打開編輯頁（不用點「編輯」，避免被彈出視窗擋住）
  const rowText = await tag
    .locator("xpath=ancestor::*[contains(normalize-space(.),'商品 ID') or contains(normalize-space(.),'商品ID')][1]")
    .first().innerText().catch(() => '');
  const idm = rowText.match(/商品\s*ID[:：]\s*(\d{5,})/);
  if (idm) {
    await page.goto(`https://seller.shopee.tw/portal/product/${idm[1]}`, { waitUntil: 'domcontentloaded' });
    console.log(`   🔗 直接打開蝦皮商品 ${idm[1]} 的編輯頁`);
    return page;
  }
  // 2) 讀不到商品 ID：關掉彈出視窗，點這一列看得見的「編輯」
  await page.keyboard.press('Escape').catch(() => {});
  const edits = tag
    .locator("xpath=ancestor::*[.//*[normalize-space(text())='編輯']][1]")
    .getByText('編輯', { exact: true });
  let edit = null;
  for (let k = 0, n = await edits.count(); k < n; k++) {
    if (await edits.nth(k).isVisible().catch(() => false)) { edit = edits.nth(k); break; }
  }
  if (!edit) throw new Error('找到商品了，但找不到「編輯」按鈕');
  const popup = ctx.waitForEvent('page', { timeout: 5000 }).catch(() => null);
  await edit.click({ timeout: 5000 }).catch(() => edit.click({ timeout: 5000, force: true }));
  const newPage = await popup;
  const p = newPage || page;
  await p.waitForLoadState('domcontentloaded');
  return p;
}

// 飛鼠編輯頁：把同步資訊「蝦皮」那列的自訂價格改成符合規則（15000 以下 ×1.18；以上清空＝用原價）
async function fixFeishuShopeePrice(page, price) {
  const n = Number(price);
  if (!n) return '';
  const want = shopeePrice(n); // 15000 以上是 ''
  const row = await findSyncRow(page, '蝦皮');
  if (!row) return '';
  const el = row.locator(`xpath=.//${TEXT_INPUT}`).nth(0);
  if (!(await el.count())) return '';
  const cur = String(await el.inputValue().catch(() => '')).replace(/[^\d]/g, '');
  if (cur === want) return '';
  await el.fill(want, { timeout: 3000 });
  await commitInput(el);
  return `蝦皮自訂價格 ${cur || '空白'} → ${want || '清空（用原價）'}`;
}
// 蝦皮：最低購買數量改成 want（15000 以上 3、以下 1），改完再讀一次確認
async function setMinQty(p, want = SHOPEE.minQty) {
  const tab = p.getByText('銷售資訊', { exact: true }).first();
  if (await tab.isVisible().catch(() => false)) { await tab.click().catch(() => {}); await p.waitForTimeout(800); }
  let box = null;
  const labs = p.locator(`xpath=//*[not(self::script) and not(self::style)][starts-with(normalize-space(text()),'最低購買數量')]`);
  for (let k = 0, n = await labs.count(); k < n && !box; k++) {
    const lab = labs.nth(k);
    if (!(await lab.isVisible().catch(() => false))) continue;
    for (let up = 1; up <= 5 && !box; up++) {
      const b = lab.locator(`xpath=ancestor::*[${up}]`);
      const inp = b.locator('input:not([type=hidden]):not([type=checkbox]):not([type=radio])');
      if ((await inp.count()) === 1) box = inp.first();
    }
  }
  if (!box) box = await visibleFieldAfter(p, '最低購買數量');
  if (!box) throw new Error('找不到「最低購買數量」欄位');
  const before = String(await box.inputValue().catch(() => '')).trim();
  if (before === String(want)) return before;
  await box.click({ timeout: 3000 }).catch(() => {});
  await box.press(process.platform === 'darwin' ? 'Meta+A' : 'Control+A').catch(() => {});
  await box.type(String(want), { delay: 40 });
  await box.press('Tab').catch(() => {});
  await p.waitForTimeout(400);
  const after = String(await box.inputValue().catch(() => '')).trim();
  if (after !== String(want)) throw new Error(`最低購買數量填了 ${want} 但欄位顯示 ${after}`);
  return before;
}

// 蝦皮價格檢查：應該是多少 → 15000 以下 售價×1.18 進位到百；15000 以上 原價
function expectedShopeePrice(price) {
  const n = Number(price);
  if (!n) return 0;
  return n >= CONFIG.shopee.highPrice ? n : Number(shopeePrice(n));
}
async function checkShopeePrice(p, price) {
  const want = expectedShopeePrice(price);
  if (!want) return { ok: false, msg: '收購紀錄沒有售價，蝦皮價格無法檢查' };
  const tab = p.getByText('銷售資訊', { exact: true }).first();
  if (await tab.isVisible().catch(() => false)) { await tab.click().catch(() => {}); await p.waitForTimeout(800); }
  const box = await visibleFieldAfter(p, '價格');
  if (!box) return { ok: false, msg: `找不到蝦皮「價格」欄位，請手動確認價格是 ${want}` };
  const cur = Number(String(await box.inputValue().catch(() => '')).replace(/[^\d.]/g, '')) || 0;
  if (cur === want) return { ok: true, cur, want, changed: false };
  await box.fill(String(want), { timeout: 5000 });
  await box.press('Tab').catch(() => {});
  await p.waitForTimeout(400);
  const now = Number(String(await box.inputValue().catch(() => '')).replace(/[^\d.]/g, '')) || 0;
  if (now !== want) return { ok: false, msg: `蝦皮價格是 ${cur}，應該是 ${want}，改不過去，請馬上手動改！` };
  return { ok: true, cur, want, changed: true };
}
// 蝦皮按「更新」後確認真的存好了：出現成功訊息、或頁面跳回商品列表 → ok；欄位錯誤 → 回傳錯誤；都沒有 → 不確定
async function saveShopeeVerified(p) {
  const startUrl = p.url();
  if (!(await save(p))) return { ok: false, msg: '找不到「更新」按鈕' };
  for (let t = 0; t < 40; t++) {
    const st = await p.evaluate(() => {
      const vis = (el) => { const r = el.getBoundingClientRect(); return r.width > 0 && r.height > 0; };
      const errs = [...document.querySelectorAll('[class*="error-msg"], [class*="error-message"], [class*="form-item__error"], [class*="--error"] [class*="message"], .text-danger')]
        .filter(vis).map((e) => (e.innerText || '').trim()).filter(Boolean);
      const body = document.body ? document.body.innerText : '';
      return { errs: errs.slice(0, 3), ok: /更新成功|儲存成功|已成功更新|商品已更新|已更新商品|成功更新/.test(body) };
    }).catch(() => ({ errs: [], ok: false }));
    if (st.ok) return { ok: true };
    if (p.url() !== startUrl && !/\/portal\/product\/\d+/.test(p.url())) return { ok: true }; // 跳回商品列表
    if (st.errs.length && t >= 2) return { ok: false, msg: '蝦皮說：' + st.errs.join('／') };
    await new Promise((r) => setTimeout(r, 500));
  }
  return { ok: null, msg: '按了「更新」20 秒沒看到成功訊息，請到蝦皮確認這筆' };
}
// 沒看到成功訊息時：重新打開商品編輯頁，把價格、最低購買數量讀回來，跟要改的一樣就算成功
async function verifyShopeeByReload(p, editUrl, price) {
  try {
    await p.goto(editUrl, { waitUntil: 'domcontentloaded' });
    await p.waitForTimeout(4000);
    const tab = p.getByText('銷售資訊', { exact: true }).first();
    if (await tab.isVisible().catch(() => false)) { await tab.click().catch(() => {}); await p.waitForTimeout(800); }
    const num = async (label) => { const b = await visibleFieldAfter(p, label).catch(() => null); return b ? String(await b.inputValue().catch(() => '')).replace(/[^\d]/g, '') : ''; };
    const gotPrice = await num('價格'), gotQty = await num('最低購買數量');
    const wantPrice = String(expectedShopeePrice(price) || ''), wantQty = Number(price) >= CONFIG.shopee.highPrice ? String(SHOPEE.minQty) : '1';
    const okPrice = !wantPrice || gotPrice === wantPrice, okQty = !price || gotQty === wantQty;
    if (okPrice && okQty) return { ok: true, msg: `重新打開確認：價格 ${gotPrice}、最低購買數量 ${gotQty}` };
    return { ok: false, msg: `重新打開後：價格 ${gotPrice || '?'}（應該 ${wantPrice}）、最低購買數量 ${gotQty || '?'}（應該 ${wantQty}），沒有存成功` };
  } catch (e) {
    return { ok: null, msg: '按了「更新」沒看到成功訊息，重新打開也讀不到，請到蝦皮確認這筆' };
  }
}
async function save(p) {
  const btns = p.locator('button').filter({ hasText: /^\s*(更新|儲存並上架|儲存)\s*$/ });
  for (let k = 0, n = await btns.count(); k < n; k++) {
    if (await btns.nth(k).isVisible().catch(() => false)) {
      await btns.nth(k).click({ timeout: 5000 });
      await p.waitForTimeout(3000);
      return true;
    }
  }
  return false;
}

// 找「標題文字包含 keyword」且看得見的那一塊（往上找到包含輸入元件的範圍）
async function shopeeFieldBox(p, keyword) {
  const labels = p.locator(`xpath=//*[not(self::script) and not(self::style)][contains(normalize-space(text()),'${keyword}')]`);
  for (let k = 0, n = await labels.count(); k < n; k++) {
    const lab = labels.nth(k);
    if (!(await lab.isVisible().catch(() => false))) continue;
    for (let up = 1; up <= 5; up++) {
      const box = lab.locator(`xpath=ancestor::*[${up}]`);
      const has = await box.evaluate((e) => !!e.querySelector('input,select,textarea,[role=combobox],[role=radio],[class*=select],[class*=radio]')).catch(() => false);
      if (has) return box;
    }
  }
  return null;
}

// 有欄位填不進去時，把附近原始碼存成「蝦皮欄位診斷.txt」
async function dumpShopeeFields(p, keywords) {
  try {
    const text = await p.evaluate((keywords) => {
      const out = [];
      for (const kw of keywords) {
        const hits = [...document.querySelectorAll('body *')].filter((e) => [...e.childNodes].some((n) => n.nodeType === 3 && n.textContent.includes(kw)));
        out.push(`==== 「${kw}」找到 ${hits.length} 個 ====`);
        hits.slice(0, 3).forEach((h, i) => {
          let box = h;
          for (let k = 0; k < 5 && box.parentElement && !box.querySelector('input,select,textarea,[role=combobox]'); k++) box = box.parentElement;
          out.push(`-- 第 ${i + 1} 個（${h.offsetParent ? '看得見' : '看不見'}）--`);
          out.push(box.outerHTML.replace(/<svg[\s\S]*?<\/svg>/g, '<svg/>').slice(0, 2500));
        });
        out.push('');
      }
      return out.join('\n');
    }, keywords);
    const f = path.join(diagDir(), '蝦皮欄位診斷.txt');
    fs.writeFileSync(f, text, 'utf8');
    console.log(`   🔍 已存診斷檔：${f}（用記事本打開，全選複製貼給 Claude）`);
  } catch (e) {}
}

// 商品保存狀況 → 二手（可能是單選按鈕，也可能是下拉選單）
async function setShopeeCondition(p) {
  for (const tab of ['其他', '屬性']) {
    const t = p.getByText(tab, { exact: true }).first();
    if (await t.isVisible().catch(() => false)) { await t.click().catch(() => {}); await p.waitForTimeout(500); }
  }
  const box = await shopeeFieldBox(p, '保存狀況');
  if (!box) return false;
  const want = SHOPEE.condition;
  const direct = box.getByText(want, { exact: true }).first();
  if (await direct.isVisible().catch(() => false)) {
    await direct.click();
  } else {
    const trigger = box.locator("[role=combobox], [class*=select], input:not([type=hidden])").first();
    await trigger.click({ timeout: 4000 });
    await p.waitForTimeout(500);
    const opts = p.getByText(want, { exact: true });
    let clicked = false;
    for (let k = (await opts.count()) - 1; k >= 0 && !clicked; k--) {
      if (await opts.nth(k).isVisible().catch(() => false)) { await opts.nth(k).click(); clicked = true; }
    }
    if (!clicked) return false;
  }
  await p.waitForTimeout(400);
  return (await box.innerText().catch(() => '')).includes(want);
}

// NCC 字號 → 從 NCC字號.csv 對照表帶入
async function setShopeeField(p, label, value) {
  // 先找「標題剛好是 label」的欄位（避免抓到上面的違規說明文字），找不到才用包含
  let box = null;
  const exact = p.locator(`xpath=//*[not(self::script) and not(self::style)][normalize-space(text())='${label}' or (starts-with(normalize-space(text()),'${label}') and string-length(normalize-space(text())) <= ${label.length + 8})]`);
  for (let k = 0, n = await exact.count(); k < n && !box; k++) {
    const lab = exact.nth(k);
    if (!(await lab.isVisible().catch(() => false))) continue;
    for (let up = 1; up <= 5 && !box; up++) {
      const b = lab.locator(`xpath=ancestor::*[${up}]`);
      if (await b.evaluate((e) => !!e.querySelector('input:not([type=hidden]),textarea')).catch(() => false)) box = b;
    }
  }
  if (!box) box = await shopeeFieldBox(p, label);
  if (!box) return false;
  const input = box.locator(`xpath=.//${TEXT_INPUT}`).first();
  if (!(await input.count())) return false;
  await input.fill(value, { timeout: 4000 });
  return true;
}
async function setShopeeNcc(p, ncc) {
  return setShopeeField(p, 'NCC', ncc);
}
async function setShopeeNccOld(p, ncc) {
  const box = await shopeeFieldBox(p, 'NCC');
  if (!box) return false;
  const input = box.locator(`xpath=.//${TEXT_INPUT}`).first();
  if (!(await input.count())) return false;
  await input.fill(ncc, { timeout: 4000 });
  return true;
}

// =================== 蝦皮違規：批次補 BSMI（node upload.js bsmifix）===================
// 「BSMI補填.csv」：序號,商品,BSMI字號,備註,結果 —— BSMI 字號請照機身／盒子標籤填真的號碼，空白的會略過
const BSMI_FILE = path.join(__dirname, 'BSMI補填.csv');
const BSMI_SEED = [
  ['059285', 'Sony TX100V'], ['21539', 'Fujifilm Finepix F80EXR'], ['001449', 'Lumix LX5'], ['101300', 'Ricoh GR 一代'],
  ['402466', 'Canon S90'], ['02744', 'Fujifilm X100S'], ['000841', 'Fujifilm FinePix F700'], ['001348', 'Lumix ZS3'],
  ['m65074', 'Fujifilm X20'], ['70664', 'Sony DSC-T5'],
];
const csvLine = (a) => a.map((x) => `"${String(x ?? '').replace(/"/g, '""')}"`).join(',');
function ensureBsmiFile() {
  if (fs.existsSync(BSMI_FILE)) return false;
  const rows = BSMI_SEED.map(([sn, name]) => {
    const e = lookupNcc(name);
    const hint = '請看機身底部或盒子上的 BSMI 標籤（⊕ 旁邊 D 或 R 開頭的號碼）；沒有標籤就留空白、不要上蝦皮';
    return [sn, name, e && e.bsmi ? e.bsmi : '', e && e.bsmi ? '官網同型號有填過，請對照機身標籤確認' : hint, ''];
  });
  fs.writeFileSync(BSMI_FILE, '﻿' + [csvLine(['序號', '商品', 'BSMI字號', '備註', '結果']), ...rows.map(csvLine)].join('\r\n') + '\r\n', 'utf8');
  return true;
}
// 把確認過的 BSMI 記進 NCC字號.csv（同型號下次上架會自動帶入）
function rememberBsmi(model, bsmi) {
  try {
    const f = CONFIG.nccFile;
    const lines = fs.existsSync(f) ? fs.readFileSync(f, 'utf8').replace(/^﻿/, '').split(/\r?\n/).filter(Boolean) : ['型號,NCC字號,BSMI字號,備註'];
    const key = normKey(model);
    let hit = false;
    const out = lines.map((l, i) => {
      if (i === 0) return l;
      const c = l.split(',');
      if (normKey(c[0] || '') === key) { hit = true; if (!c[2]) c[2] = bsmi; return c.join(','); }
      return l;
    });
    if (!hit) out.push(`${model},,${bsmi},從 BSMI補填 加入`);
    fs.writeFileSync(f, '﻿' + out.join('\r\n') + '\r\n', 'utf8');
    NCC_TABLE = null;
  } catch (e) {}
}
// BSMI 空白的：查收購紀錄，公司貨就照品牌填代理商碼（或你填過的／官方相符的）；水貨、品牌不在名單的留空白
async function autoFillBsmi() {
  const rows = readCsv(BSMI_FILE).filter((c) => c[0]);
  const blank = rows.filter((c) => !(c[2] || '').trim() && !/^成功/.test(c[4] || ''));
  if (!blank.length) return;
  const recs = await rowsBySerials(blank.map((c) => c[0]));
  const filled = [], skipped = [];
  const note = new Map();
  recs.forEach((r, i) => {
    const c = blank[i];
    if (r._missing) { skipped.push(`${c[1]}（${c[0]}）：收購紀錄找不到這個序號`); note.set(c[0], ['', '收購紀錄找不到序號，請看機身標籤']); return; }
    const name = r['商品名稱'] || c[1], origin = (r['公/水'] || '').trim();
    const code = bsmiFor(name, origin);
    if (code) {
      const why = lookupNcc(name) && lookupNcc(name).bsmi ? '你填過的／官方資料相符' : `公司貨，${guessBrand(name) || '品牌'}代理商碼（自動）`;
      filled.push(`${c[1]}（${c[0]}）→ ${code}　${why}`); note.set(c[0], [code, why]);
    } else {
      const why = origin !== '公司貨' ? `貨品來源是「${origin || '空白'}」，不是公司貨，不套代理商碼，請看機身標籤` : '品牌不在代理商名單，請看機身標籤';
      skipped.push(`${c[1]}（${c[0]}）：${why}`); note.set(c[0], ['', why]);
    }
  });
  const out = readCsv(BSMI_FILE).map((c) => (note.has(c[0]) && !(c[2] || '').trim() ? [c[0], c[1], note.get(c[0])[0], note.get(c[0])[1], c[4] || ''] : c));
  fs.writeFileSync(BSMI_FILE, '\uFEFF' + [csvLine(['序號', '商品', 'BSMI字號', '備註', '結果']), ...out.map(csvLine)].join('\r\n') + '\r\n', 'utf8');
  if (filled.length) { console.log(`依品牌自動填好 BSMI ${filled.length} 件：`); filled.forEach((x) => console.log('  ✅ ' + x)); }
  if (skipped.length) { console.log(`沒辦法自動填、會略過的 ${skipped.length} 件（請看機身標籤自己填進 BSMI補填.csv）：`); skipped.forEach((x) => console.log('  ⚠️ ' + x)); }
  console.log('');
}
async function bsmiFixMain() {
  if (ensureBsmiFile()) console.log('已建立「BSMI補填.csv」。');
  await autoFillBsmi();
  const rows = readCsv(BSMI_FILE).filter((c) => c[0]);
  const todo = rows.filter((c) => (c[2] || '').trim() && !/^成功/.test(c[4] || ''));
  const bad = todo.filter((c) => !/^[A-Za-z]\d{3,}[A-Za-z0-9]*$|^[A-Za-z]{1,3}\d{4,}$/.test(c[2].trim()));
  if (bad.length) bad.forEach((c) => console.log(`⚠️ ${c[1]}（${c[0]}）的 BSMI「${c[2]}」格式看起來不對，請再確認`));
  if (!todo.length) { console.log('「BSMI補填.csv」裡沒有要補的（BSMI字號都是空的，或都補過了）。'); openPath(BSMI_FILE); return; }
  console.log(`要補 BSMI 的蝦皮商品 ${todo.length} 件：`);
  todo.forEach((c, i) => console.log(`  ${i + 1}. ${c[1]}　序號 ${c[0]}　BSMI ${c[2].trim()}`));
  const ctx = await launchBrowser({ headless: false, channel: CONFIG.browserChannel, viewport: null, args: ['--start-maximized'] });
  const page = ctx.pages()[0] || (await ctx.newPage());
  page.on('dialog', (d) => d.accept().catch(() => {}));
  await page.goto(SHOPEE.listUrl).catch(() => {});
  const shopeeIn = await page.getByPlaceholder(/主商品貨號/).first().waitFor({ timeout: 15000 }).then(() => true, () => false);
  if (!shopeeIn) await ask('👉 請在瀏覽器登入蝦皮賣家中心，看到「我的商品」後，回這裡按 Enter 開始...');
  const result = new Map();
  for (const [i, c] of todo.entries()) {
    const [sn, name] = c, bsmi = c[2].trim();
    console.log(`\n[${i + 1}/${todo.length}] ${name}　序號 ${sn}`);
    let p = page;
    try {
      p = await openEditBySerial(ctx, page, sn);
      await p.waitForTimeout(1500);
      if (!(await setShopeeField(p, 'BSMI', bsmi).catch(() => false))) throw new Error('找不到蝦皮的「BSMI」欄位，請手動填');
      { const sv = await saveShopeeVerified(p); if (sv.ok !== true) throw new Error(sv.msg); }
      console.log(`   ✅ BSMI 已填 ${bsmi} 並儲存（蝦皮會重新審核）`);
      result.set(sn, '成功 ' + new Date().toLocaleDateString('zh-TW'));
      rememberBsmi(name, bsmi);
    } catch (e) {
      console.log('   ❌ 失敗：' + e.message.split('\n')[0]);
      result.set(sn, '失敗：' + e.message.split('\n')[0].slice(0, 60));
      await dumpShopeeFields(p, ['BSMI']).catch(() => {});
    } finally {
      if (p !== page) await p.close().catch(() => {});
    }
  }
  await ctx.close();
  // 結果寫回 BSMI補填.csv
  const out = readCsv(BSMI_FILE).map((c) => (result.has(c[0]) ? [c[0], c[1], c[2], c[3], result.get(c[0])] : c));
  fs.writeFileSync(BSMI_FILE, '\uFEFF' + [csvLine(['序號', '商品', 'BSMI字號', '備註', '結果']), ...out.map(csvLine)].join('\r\n') + '\r\n', 'utf8');
  const ok = [...result.values()].filter((x) => x.startsWith('成功')).length;
  console.log(`\n完成：成功 ${ok} 件、失敗 ${result.size - ok} 件（結果寫在 BSMI補填.csv 的「結果」欄）`);
}

// =================== 批次改蝦皮分類（node upload.js catfix [序號…]）===================
// 在飛鼠打開商品 → 同步資訊的蝦皮那列分類改成「數位相機（傻瓜相機）」→ 發佈，讓飛鼠推到蝦皮
const SHOPEE_COMPACT_CAT = '相機&空拍機 > 相機 > 全自動數位相機';
async function catFixMain() {
  let serials = process.argv.slice(3).filter((a) => !a.startsWith('--'));
  if (!serials.length) { ensureBsmiFile(); serials = readCsv(BSMI_FILE).map((c) => c[0]).filter(Boolean); }
  if (!serials.length) { console.log('沒有要改分類的商品（BSMI補填.csv 是空的）。'); return; }
  console.log(`要把蝦皮分類改成「${SHOPEE_COMPACT_CAT}」的商品 ${serials.length} 件：${serials.join('、')}`);
  const ctx = await launchBrowser({ headless: false, channel: CONFIG.browserChannel, viewport: null, args: ['--start-maximized'] });
  const page = ctx.pages()[0] || (await ctx.newPage());
  page.on('dialog', (d) => d.accept().catch(() => {}));
  const url = fs.existsSync(NEW_URL_FILE) ? fs.readFileSync(NEW_URL_FILE, 'utf8').trim() : CONFIG.startUrl;
  let loggedIn = false;
  if (url) {
    await page.goto(url, { waitUntil: 'domcontentloaded' }).catch(() => {});
    loggedIn = await page.getByText('商品管理', { exact: true }).first().waitFor({ timeout: 15000 }).then(() => true, () => false);
  }
  if (!loggedIn) await ask('👉 請在瀏覽器登入飛鼠後台，看到左側選單後，回這裡按 Enter 開始...');
  let ok = 0;
  for (const [i, sn] of serials.entries()) {
    console.log(`\n[${i + 1}/${serials.length}] 序號 ${sn}`);
    let p = page;
    try {
      p = await openDraftEditor(page, sn, { useSearch: true });
      const row = await findSyncRow(p, '蝦皮');
      if (!row) throw new Error('找不到同步資訊的「蝦皮」那一列');
      await pickRowSelect(row, 0, '常用分類');
      await p.waitForTimeout(300);
      if (!(await pickRowSelect(row, 1, SHOPEE_COMPACT_CAT))) throw new Error(`蝦皮分類裡找不到「${SHOPEE_COMPACT_CAT}」，請手動選`);
      const w = await clickPublish(p);
      if (!w) throw new Error('分類改好了，但找不到「發佈」按鈕，請手動按');
      console.log(`   ✅ 蝦皮分類已改，按了「${w}」`);
      ok++;
    } catch (e) {
      console.log('   ❌ 失敗：' + e.message.split('\n')[0]);
    } finally {
      if (p !== page) await p.close().catch(() => {});
    }
  }
  await ctx.close();
  console.log(`\n改分類：成功 ${ok} 件、失敗 ${serials.length - ok} 件。等 5 分鐘讓飛鼠推到蝦皮，再按「🧾 批次補 BSMI」`);
}

const DANGER = [];
async function shopeeMain() {
  const auto = process.argv.includes('--auto');
  const recheck = process.argv.includes('--check'); // 全部重新檢查（補過的也再看一次價格）
  const done = loadDoneSerials();
  // 補過的才略過：電腦裡的紀錄寫成功、而且收購後台也是「蝦皮已補」（舊紀錄可能誤記成功）
  let targets = (await findTargets()).filter((t) => recheck || !(done.has(t.serial) && /蝦皮已補/.test(t.fsStatus || '')));

  if (!targets.length) {
    console.log('沒有需要補蝦皮資料的商品。');
    console.log('（只會處理：收購後台標記上架、而且飛鼠已上架成功的商品；補過的不會重複補）');
    return;
  }
  console.log(`要補蝦皮資料的商品（${targets.length} 筆）：保存狀況選「${SHOPEE.condition}」、填 NCC；售價 ≥ ${CONFIG.shopee.highPrice} 的最低購買數量改 ${SHOPEE.minQty}`);
  targets.forEach((t, i) => console.log(`  ${i + 1}. ${t.name || ''} 序號 ${t.serial}${t.price ? `　售價 ${t.price}` : ''}`));
  console.log('');

  const ctx = await launchBrowser({
    headless: false,
    channel: CONFIG.browserChannel,
    viewport: null,
    args: ['--start-maximized'],
  });
  const page = ctx.pages()[0] || (await ctx.newPage());
  page.on('dialog', (d) => d.accept().catch(() => {}));
  await page.goto(SHOPEE.listUrl).catch(() => {});
  // 已經登入（看得到商品搜尋框）就直接開始，不用按 Enter
  const shopeeIn = await page.getByPlaceholder(/主商品貨號/).first().waitFor({ timeout: 15000 }).then(() => true, () => false);
  if (shopeeIn) console.log('✅ 蝦皮已經是登入狀態，直接開始');
  else await ask('👉 請在瀏覽器登入蝦皮賣家中心，看到「我的商品」後，回這裡按 Enter 開始...');

  let ok = 0, fail = 0;
  for (const [i, t] of targets.entries()) {
    console.log(`\n[${i + 1}/${targets.length}] ${t.name || ''} 序號 ${t.serial}`);
    let p = page;
    try {
      p = await openEditBySerial(ctx, page, t.serial);
      const notes = [], failed = [];
      await p.waitForTimeout(1500);
      // 1) 保存狀況
      if (await setShopeeCondition(p).catch(() => false)) console.log(`   ✏️ 保存狀況：${SHOPEE.condition}`);
      else { failed.push('保存狀況'); notes.push(`保存狀況沒選成功，請手動選「${SHOPEE.condition}」`); }
      // 2) NCC／BSMI
      const nccRow = t.name ? lookupNcc(t.name) : null;
      const bsmi = t.cam && t.name ? bsmiFor(t.name, t.origin) : '';
      if (bsmi) {
        if (await setShopeeField(p, 'BSMI', bsmi).catch(() => false)) console.log(`   ✏️ BSMI：${bsmi}${nccRow && nccRow.bsmi ? '' : '（公司貨，品牌代理商碼）'}`);
        else console.log(`   ℹ️ 蝦皮頁面沒有 BSMI 欄位，BSMI（${bsmi}）已寫在商品內文，不用另外填`);
      } else if (t.cam) notes.push(`「${t.name || t.serial}」沒有 BSMI（水貨或對照表沒有），蝦皮 BSMI 請手動填`);
      if (t.ccd) {
        // CCD 相機：只填 BSMI，不填 NCC
      } else if (isNoRadio(nccRow)) {
        notes.push('對照表標示「無無線功能」，蝦皮 NCC 欄位不填（文案已寫明）');
      } else if (nccRow && nccRow.ncc) {
        if (await setShopeeNcc(p, nccRow.ncc).catch(() => false)) console.log(`   ✏️ NCC：${nccRow.ncc}`);
        else { failed.push('NCC'); notes.push(`NCC 沒填成功，請手動填：${nccRow.ncc}`); }
      } else {
        notes.push(`NCC字號.csv 沒有「${t.name || t.serial}」，NCC 請手動填`);
      }
      // 3) 15000 以上：最低購買數量 3（改不成功一定要提醒，不然會被用錯的方式下標）
      if (t.price) {
        const wantQty = t.price >= CONFIG.shopee.highPrice ? String(SHOPEE.minQty) : '1';
        const before = await setMinQty(p, wantQty).catch((e) => { notes.push('🚨 ' + e.message); return null; });
        if (before === null) { failed.push('最低購買數量'); notes.push(`🚨 最低購買數量沒改成 ${wantQty}，請馬上到蝦皮手動改！`); }
        else if (before === wantQty) console.log(`   ✏️ 最低購買數量：${wantQty}（本來就對）`);
        else console.log(`   ✏️ 最低購買數量：${before || '?'} → ${wantQty}`);
      } else notes.push('收購紀錄沒有售價，最低購買數量和價格沒有檢查');
      // 4) 價格檢查：不對就改成正確價格
      const pr = await checkShopeePrice(p, t.price).catch((e) => ({ ok: false, msg: '價格檢查失敗：' + e.message }));
      if (!pr.ok) { failed.push('價格'); notes.push('🚨 ' + pr.msg); }
      else if (pr.changed) console.log(`   💰 蝦皮價格不對：${pr.cur} → 改成 ${pr.want}`);
      else console.log(`   💰 蝦皮價格正確：${pr.want}`);
      notes.forEach((x) => console.log('   • ' + x));
      if (failed.length) await dumpShopeeFields(p, failed.filter((x) => ['保存狀況', 'NCC', 'BSMI'].includes(x)));
      const danger = failed.some((x) => x === '價格' || x === '最低購買數量');
      if (auto && failed.length && !danger) {
        // 保存狀況／NCC／BSMI 有沒填到的：其他改好的（價格、最低購買數量…）照樣存，沒填到的記下來請你手動補
        const editUrl = p.url();
        let sv = await saveShopeeVerified(p);
        if (sv.ok === null) sv = await verifyShopeeByReload(p, editUrl, t.price);
        const miss = failed.filter((x) => ['保存狀況', 'NCC', 'BSMI'].includes(x)).join('、');
        if (sv.ok === false) throw new Error(sv.msg);
        console.log(`   ⚠️ 其他欄位已存好，但「${miss}」沒填進去，請到蝦皮手動補（已記在紀錄檔）`);
        shopeeLog(t.name, t.serial, '需手動補', `沒填到：${miss}；` + notes.join('；'));
        continue;
      }
      if (danger) {
        // 價格／最低購買數量有問題：先把改得動的存起來，並大聲提醒
        DANGER.push(`${t.name || ''} 序號 ${t.serial}：${notes.filter((x) => x.startsWith('🚨')).join('；')}`);
        if (auto) {
          await save(p).catch(() => false);
          shopeeLog(t.name, t.serial, '🚨需馬上處理', notes.join('；'));
          continue;
        }
      }
      if (auto) {
        const editUrl = p.url();
        let sv = await saveShopeeVerified(p);
        if (sv.ok === null) { sv = await verifyShopeeByReload(p, editUrl, t.price); if (sv.ok) console.log('   🔎 ' + sv.msg); }
        if (sv.ok === false) throw new Error(sv.msg);
        if (sv.ok === null) { console.log('   ⚠️ ' + sv.msg); shopeeLog(t.name, t.serial, '需確認', sv.msg); continue; }
        console.log('   ✅ 已儲存（蝦皮確認成功）');
        shopeeLog(t.name, t.serial, '成功（自動儲存）');
      } else {
        const a = await ask('   請確認後在蝦皮按「更新」儲存。完成按 Enter，跳過輸入 s，結束輸入 q：');
        if (a === 'q') break;
        if (a === 's') { shopeeLog(t.name, t.serial, '略過'); continue; }
        shopeeLog(t.name, t.serial, '成功（手動儲存）');
      }
      ok++;
    } catch (e) {
      fail++;
      const msg = e.message.split('\n')[0];
      console.log('   ❌ 失敗：' + msg);
      shopeeLog(t.name, t.serial, '失敗', msg);
      const shot = path.join(diagDir(), `蝦皮錯誤截圖-${i + 1}.png`);
      await p.screenshot({ path: shot, fullPage: true }).then(() => console.log(`   📸 已存錯誤截圖：${shot}`), () => {});
      if (auto) continue; // 不停模式：失敗的記下來，直接做下一筆
      const a = await ask('   按 Enter 繼續下一筆，輸入 q 結束：');
      if (a === 'q') break;
    } finally {
      if (p !== page) await p.close().catch(() => {});
    }
  }
  console.log(`\n完成：成功 ${ok} 筆、失敗 ${fail} 筆。紀錄在 ${SHOPEE.logFile}`);
  if (DANGER.length) {
    console.log('\n🚨🚨🚨 下面這些蝦皮價格／最低購買數量有問題，請「馬上」到蝦皮手動改，不然可能被用錯的價格下標：');
    DANGER.forEach((x) => console.log('   🚨 ' + x));
    console.log('   （不確定的話，可以先在蝦皮把這幾件「下架」，改好再上架）');
  }
  if (!auto) await ask('按 Enter 關閉瀏覽器...');
  closeAsk();
  await ctx.close();
}


// 兩台電腦（例如 Mac 和 Windows 共用資料夾）不能同時跑，不然紀錄會打架
const LOCK_FILE = path.join(__dirname, '使用中.lock');
const HOST = require('os').hostname();
const MACHINE = process.platform === 'darwin' ? 'Mac' : process.platform === 'win32' ? 'Windows 電腦' : '另一台電腦';
function takeLock() {
  try {
    if (fs.existsSync(LOCK_FILE)) {
      const l = JSON.parse(fs.readFileSync(LOCK_FILE, 'utf8'));
      const fresh = Date.now() - l.time < 3 * 60 * 1000; // 3 分鐘內有更新 = 還在跑
      let alive = true;
      if (l.host === HOST) { try { process.kill(l.pid, 0); } catch (e) { alive = false; } } // 同一台電腦：那個程式已經關掉就不算
      if (fresh && alive && !(l.host === HOST && l.pid === process.pid)) {
        const who = l.host === HOST ? '這台電腦的另一個視窗' : `${l.machine}（${l.host}）`;
        console.log(`\n⛔ ${who}正在跑「${l.what}」，請等它跑完再按。`);
        console.log('   （如果確定沒有在跑，等 3 分鐘後再試，或刪掉資料夾裡的「使用中.lock」）');
        process.exit(1);
      }
    }
    const write = () => fs.writeFileSync(LOCK_FILE, JSON.stringify({ host: HOST, machine: MACHINE, pid: process.pid, what: ({ all: '一鍵上架', photos: process.argv.includes('--order') ? '拍照順序' : '照片分組', addphotos: '補照片', shopee: '蝦皮補資料', airtable: '飛鼠上架', drafts: '草稿補照片', markshopee: '記錄蝦皮完成' })[process.argv[2]] || process.argv.slice(2).join(' '), time: Date.now() }));
    write();
    setInterval(() => { try { write(); } catch (e) {} }, 60 * 1000).unref();
    const release = () => { try { const l = JSON.parse(fs.readFileSync(LOCK_FILE, 'utf8')); if (l.host === HOST && l.pid === process.pid) fs.unlinkSync(LOCK_FILE); } catch (e) {} };
    process.on('exit', release);
    process.on('SIGINT', () => { release(); process.exit(130); });
    process.on('SIGTERM', () => { release(); process.exit(143); });
  } catch (e) {}
}
const READ_ONLY_CMDS = ['status', 'todo'];
// 有新版時提醒（最多等 3 秒，查不到就算了）
async function checkUpdate() {
  try {
    const local = fs.existsSync(path.join(__dirname, '版本.txt')) ? fs.readFileSync(path.join(__dirname, '版本.txt'), 'utf8').trim() : '';
    const one = (u) => Promise.race([fetch(u + '?t=' + Date.now()).then((r) => (r.ok ? r.json() : null)), new Promise((r) => setTimeout(() => r(null), 3000))]).catch(() => null);
    const ms = (await Promise.all(['https://raw.githubusercontent.com/Annie0130/buy-backstage/main/tools/feishu-uploader/manifest.json',
      'https://precious-otter-3873db.netlify.app/tools/feishu-uploader/manifest.json'].map(one))).filter(Boolean);
    if (!ms.length) return;
    const m = ms.sort((a, b) => String(b.version).localeCompare(String(a.version)))[0];
    if (m.version && m.version !== local) console.log(`\n🔔 上架程式有新版（${m.version}）！跑完後雙擊「更新程式.bat」就會自動更新${m.note ? '：' + m.note : ''}\n`);
  } catch (e) {}
}
if (require.main === module && !READ_ONLY_CMDS.includes(process.argv[2]) && !process.argv.includes('--list')) takeLock();
if (require.main === module) (process.argv[2] === 'status' ? Promise.resolve() : checkUpdate()).then(() => (process.argv[2] === 'status' ? statusMain() : process.argv[2] === 'markshopee' ? markShopeeMain() : process.argv[2] === 'republish' ? republishMain() : process.argv[2] === 'bsmifix' ? bsmiFixMain() : process.argv[2] === 'catfix' ? catFixMain() : process.argv[2] === 'shopee' ? shopeeMain() : process.argv[2] === 'todo' ? todoMain()
  : process.argv[2] === 'photos' ? (process.argv.includes('--order') ? photosOrderMain() : photosMain())
  : process.argv[2] === 'all' ? allMain() : process.argv[2] === 'drafts' ? draftsMain()
  : process.argv[2] === 'addphotos' ? addPhotosMain() : main())).then(() => flushAirtable()).then(() => closeAsk()).catch(async (e) => { await flushAirtable(); closeAsk(); console.error('❌ ' + (e.message || e)); process.exit(1); });
module.exports = { autoFillBsmi, addPhotosMain, photosPending, draftsMain, allMain, exifTime, productPhotos, photosMain, photosOrderMain, sceneWords, buildShopeeTitle, buildTitle, todoMain, isOnStorefront, confirmSaved, visibleMessages, shopeeMain, feishuCategoryKeys, platformCategory, gearType, readAirtable, visibleFieldAfter, ask, closeAsk, keyOf, TEXT_INPUT, fetchRetry, lookupNcc, dumpFieldHtml, normalizeAccessories, readRows, fillProduct, validate, autoFill, airtableToRow, buildListingText, CONFIG };
