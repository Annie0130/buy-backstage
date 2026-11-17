import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { describe, test } from 'node:test';
import { adminAuthResult, _testAuthLimits } from '../netlify/lib/auth.js';
import { sheetCell } from '../netlify/lib/sheet.js';

process.env.ADMIN_PASSWORD = 'shop-test-password';
process.env.EMAIL_API_SECRET = 'email-secret-at-least-16';
_testAuthLimits({ failDelayMs: 0, maxFails: 5, windowMs: 10 * 60 * 1000, lockMs: 10 * 60 * 1000 });

function req(password, ip) {
  return new Request('https://example.test/', {
    headers: {
      'x-admin-password': password,
      'x-nf-client-connection-ip': ip,
    },
  });
}

describe('security', { concurrency: 1 }, () => {
test('sheet cells that look like formulas are forced to text', () => {
  assert.equal(sheetCell('=1+1'), "'=1+1");
  assert.equal(sheetCell('+886'), "'+886");
  assert.equal(sheetCell('-10'), "'-10");
  assert.equal(sheetCell('@SUM(A1)'), "'@SUM(A1)");
  assert.equal(sheetCell('Sony A7C'), 'Sony A7C');
  assert.equal(sheetCell('008362'), '008362');
  assert.equal(sheetCell(''), '');
  assert.equal(sheetCell(12000), '12000');
});

test('wrong password is rejected and locks after repeated guesses', async () => {
  const ip = '203.0.113.20';
  assert.equal(await adminAuthResult(req(process.env.ADMIN_PASSWORD, ip)), 'ok');
  for (let i = 0; i < 4; i++) {
    assert.equal(await adminAuthResult(req('nope', ip)), 'bad');
  }
  assert.equal(await adminAuthResult(req('nope', ip)), 'locked');
  assert.equal(await adminAuthResult(req(process.env.ADMIN_PASSWORD, ip)), 'locked');
});

test('a correct password clears earlier mistakes from another network', async () => {
  const ip = '203.0.113.21';
  assert.equal(await adminAuthResult(req('nope', ip)), 'bad');
  assert.equal(await adminAuthResult(req(process.env.ADMIN_PASSWORD, ip)), 'ok');
  assert.equal(await adminAuthResult(req('nope', ip)), 'bad');
});

test('send-receipt refuses a bad password before calling the email service', async () => {
  let called = false;
  const original = globalThis.fetch;
  globalThis.fetch = async () => { called = true; throw new Error('should not send'); };
  try {
    const { default: handler } = await import('../netlify/functions/send-receipt.js');
    const res = await handler(new Request('https://example.test/send-receipt', {
      method: 'POST',
      headers: {
        'content-type': 'application/json',
        'x-admin-password': 'wrong',
        'x-nf-client-connection-ip': '203.0.113.30',
      },
      body: JSON.stringify({ to: 'guest@example.com' }),
    }));
    assert.equal(res.status, 401);
    assert.equal(called, false);
  } finally {
    globalThis.fetch = original;
  }
});

test('send-receipt forwards only the receipt fields and the email secret', async () => {
  const seen = [];
  const original = globalThis.fetch;
  globalThis.fetch = async (url, opts) => {
    seen.push({ url, opts, body: JSON.parse(opts.body) });
    return Response.json({ success: true });
  };
  try {
    const { default: handler } = await import('../netlify/functions/send-receipt.js');
    const res = await handler(new Request('https://example.test/send-receipt', {
      method: 'POST',
      headers: {
        'content-type': 'application/json',
        'x-admin-password': process.env.ADMIN_PASSWORD,
        'x-nf-client-connection-ip': '203.0.113.31',
      },
      body: JSON.stringify({
        to: 'guest@example.com',
        customerName: '王小明',
        txDate: '2026-11-17',
        amount: 12000,
        payment: '現金',
        equipment: [' Sony A7C ', '', 'FE 35mm'],
        html: '<script>alert(1)</script>',
        subject: '自訂主旨',
        cc: 'attacker@example.com',
      }),
    }));
    assert.equal(res.status, 200);
    assert.equal(seen.length, 1);
    assert.equal(seen[0].url, 'https://linscam-email-api-lpvi.vercel.app/api/send-email');
    assert.equal(seen[0].opts.headers['x-email-secret'], process.env.EMAIL_API_SECRET);
    assert.deepEqual(seen[0].body, {
      template: 'buyout-receipt',
      to: 'guest@example.com',
      customerName: '王小明',
      txDate: '2026-11-17',
      amount: 12000,
      payment: '現金',
      equipment: ['Sony A7C', 'FE 35mm'],
    });
  } finally {
    globalThis.fetch = original;
  }
});

test('the public page does not call the email or sheet URLs itself', () => {
  const html = readFileSync(new URL('../public/index.html', import.meta.url), 'utf8');
  const toml = readFileSync(new URL('../netlify.toml', import.meta.url), 'utf8');
  assert.match(toml, /publish\s*=\s*"public"/);
  assert.doesNotMatch(html, /linscam-email-api/);
  assert.doesNotMatch(html, /script\.google\.com/);
  assert.doesNotMatch(html, /onclick="showCustomerHistory\(/);
  assert.doesNotMatch(html, /onclick="openCreditModal\('\$\{/);
  assert.match(html, /onclick="openCreditModal\(\)"/);
  assert.match(html, /data-customer-history/);
  assert.match(html, /function esc\(t\)/);
  assert.match(html, /\/\.netlify\/functions\/send-receipt/);
  assert.match(html, /\/\.netlify\/functions\/sheet-export/);
});
});
