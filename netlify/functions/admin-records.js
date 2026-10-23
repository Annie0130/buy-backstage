// 收購後台的 Airtable 讀寫代理：金鑰與後台密碼都只存在 Netlify 環境變數裡。
// GET  ?offset=xxx      讀取一頁紀錄（每頁 100 筆）
// PATCH ?id=recXXXX     更新一筆紀錄的欄位
// 不提供刪除；每個請求都要帶正確的後台密碼（x-admin-password）。
import { timingSafeEqual } from 'node:crypto';

const BASE_ID = 'appAuc0SgpUMOFLsv';
const TABLE_ID = 'tblnOp2xI0RPOAR3J';
const AIRTABLE_URL = `https://api.airtable.com/v0/${BASE_ID}/${TABLE_ID}`;

function passwordOk(given, expected) {
  const a = Buffer.from(given || '');
  const b = Buffer.from(expected);
  return a.length === b.length && timingSafeEqual(a, b);
}

export default async (req) => {
  const token = process.env.AIRTABLE_TOKEN;
  const adminPassword = process.env.ADMIN_PASSWORD;
  if (!token || !adminPassword) {
    return Response.json({ error: '伺服器未設定 AIRTABLE_TOKEN 或 ADMIN_PASSWORD' }, { status: 500 });
  }
  if (!passwordOk(req.headers.get('x-admin-password'), adminPassword)) {
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

  return Response.json({ error: 'Method not allowed' }, { status: 405 });
};
