// 印卡區（Airtable「印卡清單」表）的讀寫代理：排隊中的商品卡、印卡紀錄。
// GET   ?view=queue        排隊中的（全部）
// GET   ?view=history      最近 60 天印過的（每頁 100 筆，用 offset 翻頁）
// POST  {records:[{id?, fields}]}  批次新增／更新（一次最多 10 筆；有 id 的更新、沒 id 的新增）
// 不提供刪除（移出排隊＝狀態改「已移除」）；每個請求都要帶正確的後台密碼（x-admin-password）。
import { adminPasswordOk } from '../lib/auth.js';

const BASE_ID = 'appAuc0SgpUMOFLsv';
const TABLE_ID = 'tblKNAod1qDKxsdpu';
const AIRTABLE_URL = `https://api.airtable.com/v0/${BASE_ID}/${TABLE_ID}`;
const ALLOWED = ['名稱', '來源', '紀錄ID', '條碼', '狀態', '加入時間', '印卡時間', '批次'];
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

  if (req.method === 'GET') {
    const view = params.get('view') === 'history' ? 'history' : 'queue';
    const q = new URLSearchParams({ pageSize: '100' });
    q.set('filterByFormula', view === 'queue'
      ? "{狀態}='排隊中'"
      : "AND({狀態}='已印', IS_AFTER({印卡時間}, DATEADD(NOW(), -60, 'days')))");
    q.append('sort[0][field]', view === 'queue' ? '加入時間' : '印卡時間');
    q.append('sort[0][direction]', view === 'queue' ? 'asc' : 'desc');
    if (params.get('offset')) q.set('offset', params.get('offset'));
    const res = await fetch(`${AIRTABLE_URL}?${q}`, { headers: H });
    return Response.json(await res.json(), { status: res.status });
  }

  if (req.method === 'POST') {
    let body;
    try { body = await req.json(); } catch { return Response.json({ error: '資料格式錯誤' }, { status: 400 }); }
    const recs = Array.isArray(body && body.records) ? body.records : [];
    if (!recs.length || recs.length > 10) return Response.json({ error: '一次要 1～10 筆' }, { status: 400 });
    const upd = recs.filter((r) => r.id), add = recs.filter((r) => !r.id);
    if (upd.some((r) => !isId(r.id))) return Response.json({ error: '紀錄 ID 錯誤' }, { status: 400 });
    const out = [];
    for (const [method, list] of [['PATCH', upd], ['POST', add]]) {
      if (!list.length) continue;
      const res = await fetch(AIRTABLE_URL, { method, headers: H, body: JSON.stringify({
        records: list.map((r) => (r.id ? { id: r.id, fields: clean(r.fields) } : { fields: clean(r.fields) })), typecast: true }) });
      const data = await res.json();
      if (!res.ok) return Response.json(data, { status: res.status });
      out.push(...data.records);
    }
    return Response.json({ records: out });
  }

  return Response.json({ error: 'Method not allowed' }, { status: 405 });
};
