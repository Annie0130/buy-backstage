// 「重寄」收購確認信。瀏覽器只打到這裡，並帶後台密碼。
// 通過後才用環境變數 EMAIL_API_SECRET（表頭 x-email-secret）轉送到寄信服務。
// 只轉固定範本的交易欄位，不接受自訂主旨或整封 HTML。
import { adminAuthResult, adminAuthResponse } from '../lib/auth.js';

const EMAIL_API = 'https://linscam-email-api-lpvi.vercel.app/api/send-email';

function equipmentList(value) {
  const list = Array.isArray(value) ? value : [];
  return list.map((item) => String(item ?? '').trim()).filter(Boolean).slice(0, 3);
}

export default async (req) => {
  if (req.method !== 'POST') {
    return Response.json({ error: '只接受 POST' }, { status: 405 });
  }
  if (!process.env.ADMIN_PASSWORD) {
    return Response.json({ error: '伺服器未設定 ADMIN_PASSWORD' }, { status: 500 });
  }
  const authResult = await adminAuthResult(req);
  if (authResult !== 'ok') return adminAuthResponse(authResult);

  const secret = process.env.EMAIL_API_SECRET;
  if (!secret) {
    return Response.json({ error: '未設定 EMAIL_API_SECRET' }, { status: 500 });
  }

  let body;
  try {
    body = await req.json();
  } catch {
    return Response.json({ error: '資料格式錯誤' }, { status: 400 });
  }
  if (!body || typeof body !== 'object' || Array.isArray(body)) {
    return Response.json({ error: '資料格式錯誤' }, { status: 400 });
  }

  const forward = {
    template: 'buyout-receipt',
    to: body.to,
    customerName: body.customerName,
    txDate: body.txDate,
    amount: body.amount,
    payment: body.payment,
    equipment: equipmentList(body.equipment),
  };

  const res = await fetch(EMAIL_API, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-email-secret': secret,
    },
    body: JSON.stringify(forward),
  });
  const data = await res.json().catch(() => ({}));
  return Response.json(data, { status: res.status });
};
