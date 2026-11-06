// 收購後台的 Airtable 讀寫代理：金鑰與後台密碼都只存在 Netlify 環境變數裡。
// GET  ?offset=xxx      讀取一頁紀錄（每頁 100 筆）
// PATCH ?id=recXXXX     更新一筆紀錄的欄位
// POST                  單筆新增商品（維修回來、面交、寄賣等不是從合約網站來的）
// 不提供刪除；每個請求都要帶正確的後台密碼（x-admin-password）。
import { adminPasswordOk } from '../lib/auth.js';

const BASE_ID = 'appAuc0SgpUMOFLsv';
const TABLE_ID = 'tblnOp2xI0RPOAR3J';
const AIRTABLE_URL = `https://api.airtable.com/v0/${BASE_ID}/${TABLE_ID}`;

export default async (req) => {
  const token = process.env.AIRTABLE_TOKEN;
  const adminPassword = process.env.ADMIN_PASSWORD;
  if (!token || !adminPassword) {
    return Response.json({ error: '伺服器未設定 AIRTABLE_TOKEN 或 ADMIN_PASSWORD' }, { status: 500 });
  }
  if (!adminPasswordOk(req)) {
    return Response.json({ error: '密碼錯誤' }, { status: 401 });
  }

  const params = new URL(req.url).searchParams;
  const auth = { Authorization: `Bearer ${token}` };

  if (req.method === 'GET') {
    const query = new URLSearchParams({ pageSize: '100' });
    if (params.get('offset')) query.set('offset', params.get('offset'));
    const res = await fetch(`${AIRTABLE_URL}?${query}`, { headers: auth });
    return Response.json(await res.json(), { status: res.status });
  }

  if (req.method === 'PATCH') {
    const id = params.get('id') || '';
    if (!/^rec[A-Za-z0-9]{14}$/.test(id)) {
      return Response.json({ error: '紀錄 ID 錯誤' }, { status: 400 });
    }
    let body;
    try {
      body = await req.json();
    } catch {
      return Response.json({ error: '資料格式錯誤' }, { status: 400 });
    }
    if (!body || typeof body.fields !== 'object' || Array.isArray(body.fields)) {
      return Response.json({ error: '缺少 fields' }, { status: 400 });
    }
    const res = await fetch(`${AIRTABLE_URL}/${id}`, {
      method: 'PATCH',
      headers: { ...auth, 'Content-Type': 'application/json' },
      body: JSON.stringify({ fields: body.fields }),
    });
    return Response.json(await res.json(), { status: res.status });
  }

  if (req.method === 'POST') {
    let body;
    try { body = await req.json(); } catch { return Response.json({ error: '資料格式錯誤' }, { status: 400 }); }
    const f = body && body.fields;
    if (!f || typeof f !== 'object' || Array.isArray(f)) return Response.json({ error: '缺少 fields' }, { status: 400 });
    if (!String(f['器材一'] || '').trim()) return Response.json({ error: '請填器材名稱' }, { status: 400 });
    // 單筆新增只收商品相關欄位（不收客人個資）
    const ALLOWED = ['器材一', '序號', '交易日期', '收購總金額', '售價', '貨品來源', '新舊程度', '盒單', '配件清單', '快門數',
      '功能備註', '原廠保固', '店保', 'CCD／Y2K', '品類', '商品介紹', '上架飛鼠', '進貨管道', '經手人員'];
    const fields = Object.fromEntries(Object.entries(f).filter(([k, v]) => ALLOWED.includes(k) && v !== '' && v !== null && v !== undefined));
    const res = await fetch(AIRTABLE_URL, {
      method: 'POST',
      headers: { ...auth, 'Content-Type': 'application/json' },
      body: JSON.stringify({ fields, typecast: true }),
    });
    return Response.json(await res.json(), { status: res.status });
  }

  return Response.json({ error: 'Method not allowed' }, { status: 405 });
};
