// 後台小程式共用的密碼檢查：密碼存在 Netlify 環境變數 ADMIN_PASSWORD。
// 猜錯會先等一秒；同一個網路在 10 分鐘內猜錯 5 次，要再等 10 分鐘。
// （Netlify 函式各自記在自己的記憶體，換一台機器或冷啟動後次數會重算，但仍會每次猜錯都慢一秒。）
import { timingSafeEqual } from 'node:crypto';

const LIMITS = {
  failDelayMs: 1000,
  maxFails: 5,
  windowMs: 10 * 60 * 1000,
  lockMs: 10 * 60 * 1000,
};

const attempts = new Map();

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function clientKey(req) {
  const ip = req.headers.get('x-nf-client-connection-ip')
    || (req.headers.get('x-forwarded-for') || '').split(',')[0];
  return (ip || 'unknown').trim() || 'unknown';
}

function passwordsMatch(given, expected) {
  const a = Buffer.from(given);
  const b = Buffer.from(expected);
  return a.length === b.length && timingSafeEqual(a, b);
}

// 測試用：縮短等待，並清掉累計次數。正式環境不會呼叫。
export function _testAuthLimits(overrides = {}) {
  Object.assign(LIMITS, overrides);
  attempts.clear();
}

/**
 * @returns {Promise<'ok'|'bad'|'locked'>}
 * locked：這個網路還在鎖定期，正確密碼也不放行。
 */
export async function adminAuthResult(req) {
  const expected = process.env.ADMIN_PASSWORD;
  if (!expected) return 'bad';

  const key = clientKey(req);
  const now = Date.now();
  const rec = attempts.get(key);
  if (rec && rec.lockedUntil > now) {
    await sleep(LIMITS.failDelayMs);
    return 'locked';
  }
  if (rec && rec.lockedUntil && rec.lockedUntil <= now) {
    rec.lockedUntil = 0;
    rec.fails = [];
  }

  const given = req.headers.get('x-admin-password') || '';
  if (passwordsMatch(given, expected)) {
    attempts.delete(key);
    return 'ok';
  }

  const next = rec || { fails: [], lockedUntil: 0 };
  next.fails = (next.fails || []).filter((t) => now - t < LIMITS.windowMs);
  next.fails.push(now);
  if (next.fails.length >= LIMITS.maxFails) {
    next.lockedUntil = now + LIMITS.lockMs;
    next.fails = [];
  }
  attempts.set(key, next);
  await sleep(LIMITS.failDelayMs);
  return next.lockedUntil > now ? 'locked' : 'bad';
}

export function adminAuthResponse(result) {
  if (result === 'locked') {
    return Response.json(
      { error: '密碼嘗試次數過多，請稍後再試' },
      { status: 429, headers: { 'retry-after': String(Math.ceil(LIMITS.lockMs / 1000)) } },
    );
  }
  return Response.json({ error: '密碼錯誤' }, { status: 401 });
}
