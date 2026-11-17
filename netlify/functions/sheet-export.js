// 把勾選的收購紀錄寫進 Google 試算表「匯入時只能有他」。
// 以前瀏覽器直接呼叫公開的 Apps Script，任何人拿到網址都能寫。現在改由這個函式代送：
// - 要帶正確的後台密碼
// - 若 Netlify 有設定 SHEET_SECRET，會多送一個 secret 參數（試算表那邊要檢查，見檔案最下面）
// - 儲存格若以 = + - @ 開頭，先加上 '，避免被試算表當成公式
//
// 試算表程式（buy-form 的 apps-script/Code.gs，實際執行的是試算表裡的那份）要對應改成：
// 1. doGet / doPost 進來時，若「指令碼屬性」有 SHEET_SECRET，就比對參數 secret，不對就拒絕。
//    屬性還沒填時先不要拒絕，否則收購同意書網站（也用同一個網址）會突然寫不進去。
// 2. 寫入每一格前，用跟 sheetCell 一樣的規則：開頭是 = + - @ 就前面加 '。
//    純數字條碼維持原本的 asText（前面加 '，開頭的 0 才不會不見）。已經有 ' 的不要再加一次。
// 改完要在 Apps Script 按「部署 → 管理部署 → 編輯 → 版本：新版本 → 部署」。
import { adminAuthResult, adminAuthResponse } from '../lib/auth.js';
import { sheetCell } from '../lib/sheet.js';

const SHEET_URL = process.env.SHEET_URL
  || 'https://script.google.com/macros/s/AKfycbx-afMzrAEbPKTpP7Kc4F0EWCUYXH87sfPGUUYXCAGipSW_TVJ7xZFZ_3hLW-l0sd0G/exec';

const MAX_ROWS = 20;

function rowQuery(row) {
  const cells = row && row.cells;
  if (!cells || typeof cells !== 'object' || Array.isArray(cells)) return null;
  const params = new URLSearchParams();
  for (const [key, value] of Object.entries(cells)) {
    if (key === 'secret' || key === 'airtableId') continue;
    if (value === undefined || value === null) continue;
    params.append(key, sheetCell(value));
  }
  if (!params.get('商品名稱')) return null;
  if (row.airtableId && /^rec[A-Za-z0-9]{14}$/.test(String(row.airtableId))) {
    params.append('airtableId', String(row.airtableId));
  }
  if (process.env.SHEET_SECRET) params.append('secret', process.env.SHEET_SECRET);
  return params;
}

async function writeRow(params) {
  const res = await fetch(`${SHEET_URL}?${params.toString()}`, { method: 'GET', redirect: 'follow' });
  const text = await res.text();
  if (!res.ok) return false;
  try {
    const data = JSON.parse(text);
    if (data && data.error) return false;
    if (data && data.success) return true;
  } catch {
    // 舊版程式有時只回純文字
  }
  return text.trim() === 'OK' || text.includes('"success":true') || text.includes('"success": true');
}

export default async (req) => {
  if (!process.env.ADMIN_PASSWORD) {
    return Response.json({ error: '伺服器未設定 ADMIN_PASSWORD' }, { status: 500 });
  }
  const auth = await adminAuthResult(req);
  if (auth !== 'ok') return adminAuthResponse(auth);
  if (req.method !== 'POST') return Response.json({ error: 'Method not allowed' }, { status: 405 });

  let body;
  try { body = await req.json(); } catch { return Response.json({ error: '資料格式錯誤' }, { status: 400 }); }
  const rows = Array.isArray(body && body.rows) ? body.rows : null;
  if (!rows || !rows.length || rows.length > MAX_ROWS) {
    return Response.json({ error: `一次要 1～${MAX_ROWS} 筆` }, { status: 400 });
  }

  let success = 0;
  let fail = 0;
  for (const row of rows) {
    const params = rowQuery(row);
    if (!params) { fail += 1; continue; }
    try {
      if (await writeRow(params)) success += 1;
      else fail += 1;
    } catch {
      fail += 1;
    }
  }
  return Response.json({ success, fail });
};
