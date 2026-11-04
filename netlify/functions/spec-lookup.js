// 商品卡「✨ 查規格」與自動填入：請 Claude 判斷品類並寫一段介紹。
// Anthropic API Key 存在 Netlify 環境變數 ANTHROPIC_API_KEY，要帶正確的後台密碼才能用。
import Anthropic from '@anthropic-ai/sdk';
import { adminPasswordOk } from '../lib/auth.js';
import { describeGear } from '../lib/describe.js';

export default async (req) => {
  if (req.method !== 'POST') {
    return Response.json({ error: 'Method not allowed' }, { status: 405 });
  }
  if (!adminPasswordOk(req)) {
    return Response.json({ error: '密碼錯誤' }, { status: 401 });
  }
  if (!process.env.ANTHROPIC_API_KEY) {
    return Response.json({ error: '伺服器未設定 ANTHROPIC_API_KEY' }, { status: 500 });
  }

  let model = '';
  let notes;
  try {
    const body = await req.json();
    model = String(body.model || '').trim();
    // 匯入舊庫存時會帶備註，請 AI 一併拆出配件、快門、保固、瑕疵
    if (body.notes !== undefined) notes = String(body.notes || '').slice(0, 500);
  } catch {
    return Response.json({ error: '資料格式錯誤' }, { status: 400 });
  }
  if (!model || model.length > 100) {
    return Response.json({ error: '請輸入型號（100 字以內）' }, { status: 400 });
  }

  try {
    const result = await describeGear(model, notes);
    if (!result) {
      return Response.json({ error: '這個型號查不到，請手動填寫' }, { status: 422 });
    }
    const { intro, ...rest } = result;
    if (!intro && notes === undefined) {
      return Response.json({ error: '查無此型號，請手動填寫', ...rest }, { status: 422 });
    }
    return Response.json({ text: intro, ...rest });
  } catch (err) {
    if (err instanceof Anthropic.AuthenticationError) {
      return Response.json({ error: 'Anthropic API Key 錯誤，請檢查 Netlify 的 ANTHROPIC_API_KEY' }, { status: 500 });
    }
    if (err instanceof Anthropic.RateLimitError) {
      return Response.json({ error: '查詢太頻繁或額度不足，請稍後再試' }, { status: 429 });
    }
    if (err instanceof Anthropic.APIError) {
      const detail = err.error?.error?.message || err.message || '';
      console.error('Anthropic API 錯誤', err.status, detail);
      if (/credit balance/i.test(detail)) {
        return Response.json({ error: 'AI 額度用完了，請到 console.anthropic.com 的 Billing 儲值' }, { status: 402 });
      }
      return Response.json({ error: `AI 服務錯誤（${err.status}）：${detail.slice(0, 200)}` }, { status: 502 });
    }
    throw err;
  }
};
