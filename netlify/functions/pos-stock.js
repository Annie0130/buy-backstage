// POS 庫存專區的 Airtable 讀寫代理（跟收購紀錄分開的「POS庫存」表）。
// GET   ?offset=xxx   讀取一頁（每頁 100 筆）
// PATCH ?id=recXXXX   更新一筆
// POST  {records:[{id?, fields}]}  批次新增／更新（一次最多 10 筆；有 id 的更新、沒 id 的新增）
// 不提供刪除（賣掉的改成「已售出」）；每個請求都要帶正確的後台密碼（x-admin-password）。
import { adminPasswordOk } from '../lib/auth.js';

const BASE_ID = 'appAuc0SgpUMOFLsv';
const TABLE_ID = 'tbl3FSkUsoTqeTt4j';
const AIRTABLE_URL = `https://api.airtable.com/v0/${BASE_ID}/${TABLE_ID}`;

const ALLOWED = ['名稱', '條碼', '狀態', 'POS分類', '品牌', '售價', '庫存量', '貨品來源', '盒單', '原廠盒', '保卡', 'POS備註',
  '新舊程度', '狀況說明', '選單語言', '配件', '快門數', '店保', '原廠保固', '原廠保固到期', '原廠保固備註', '商品介紹', '品類',
  'CCD／Y2K', '最後匯入', '印卡日期', '寄賣'];
const clean = (f) => Object.fromEntries(Object.entries(f || {}).filter(([k]) => ALLOWED.includes(k))
  .map(([k, v]) => [k, v === '' ? null : v]));
const isId = (id) => /^rec[A-Za-z0-9]{14}$/.test(id || '');

export default async (req) => {
  const token = process.env.AIRTABLE_TOKEN;
  if (!token || !process.env.ADMIN_PASSWORD) {
    return Response.json({ error: '伺服器未設定 AIRTABLE_TOKEN 或 ADMIN_PASSWORD' }, { status: 500 });
  }
  if (!adminPasswordOk(req)) return Response.json({ error: '密碼錯誤' }, { status: 401 });

  const params = new URL(req.url).searchParams;
  const H = { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' };
  const send = async (url, method, body) => {
    const res = await fetch(url, { method, headers: H, body: body ? JSON.stringify(body) : undefined });
    return Response.json(await res.json(), { status: res.status });
  };

  if (req.method === 'GET') {
    const q = new URLSearchParams({ pageSize: '100' });
    if (params.get('offset')) q.set('offset', params.get('offset'));
    return send(`${AIRTABLE_URL}?${q}`, 'GET');
  }

  let body;
  try { body = await req.json(); } catch { return Response.json({ error: '資料格式錯誤' }, { status: 400 }); }

  if (req.method === 'PATCH') {
    const id = params.get('id');
    if (!isId(id)) return Response.json({ error: '紀錄 ID 錯誤' }, { status: 400 });
    return send(`${AIRTABLE_URL}/${id}`, 'PATCH', { fields: clean(body && body.fields), typecast: true });
  }

  if (req.method === 'POST') {
    const recs = Array.isArray(body && body.records) ? body.records : [];
    if (!recs.length || recs.length > 10) return Response.json({ error: '一次要 1～10 筆' }, { status: 400 });
    const upd = recs.filter((r) => r.id), add = recs.filter((r) => !r.id);
    if (upd.some((r) => !isId(r.id))) return Response.json({ error: '紀錄 ID 錯誤' }, { status: 400 });
    const out = [];
    if (upd.length) {
      const res = await fetch(AIRTABLE_URL, { method: 'PATCH', headers: H, body: JSON.stringify({ records: upd.map((r) => ({ id: r.id, fields: clean(r.fields) })), typecast: true }) });
      const data = await res.json();
      if (!res.ok) return Response.json(data, { status: res.status });
      out.push(...data.records);
    }
    if (add.length) {
      const res = await fetch(AIRTABLE_URL, { method: 'POST', headers: H, body: JSON.stringify({ records: add.map((r) => ({ fields: clean(r.fields) })), typecast: true }) });
      const data = await res.json();
      if (!res.ok) return Response.json(data, { status: res.status });
      out.push(...data.records);
    }
    return Response.json({ records: out });
  }

  return Response.json({ error: 'Method not allowed' }, { status: 405 });
};
