// 後台小程式共用的密碼檢查：密碼存在 Netlify 環境變數 ADMIN_PASSWORD
import { timingSafeEqual } from 'node:crypto';

export function adminPasswordOk(req) {
  const expected = process.env.ADMIN_PASSWORD;
  if (!expected) return false;
  const a = Buffer.from(req.headers.get('x-admin-password') || '');
  const b = Buffer.from(expected);
  return a.length === b.length && timingSafeEqual(a, b);
}
