// 商品卡「✨ 查規格」：請 Claude 寫一段器材介紹。
// Anthropic API Key 存在 Netlify 環境變數 ANTHROPIC_API_KEY，要帶正確的後台密碼才能用。
import Anthropic from '@anthropic-ai/sdk';
import { adminPasswordOk } from '../lib/auth.js';

const SYSTEM = `你是二手相機店「林相攝影」的商品文案助手。使用者會給你一個相機或鏡頭型號，請用繁體中文寫一段放在商品卡上的介紹。

要求：
- 一段話，60 到 80 字，不要條列、不要標題、不要換行
- 先寫發售年份與最重要的規格（相機：畫素、感光元件尺寸；鏡頭：焦段、光圈、適用片幅），再寫一兩個主要特色和適合拍什麼
- 例：2021 年發售，2420 萬畫素全片幅背照式 CMOS，693 點相位對焦，機身僅 509g，高感光出色，輕巧全片幅首選。
- 只寫你確定的規格；不確定的數字寧可不寫，不要猜
- 認不出這個型號時，只回覆「查無此型號，請手動填寫」`;

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
  try {
    model = String((await req.json()).model || '').trim();
  } catch {
    return Response.json({ error: '資料格式錯誤' }, { status: 400 });
  }
  if (!model || model.length > 100) {
    return Response.json({ error: '請輸入型號（100 字以內）' }, { status: 400 });
  }

  const client = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });
  try {
    const msg = await client.beta.messages.create({
      model: 'claude-opus-5-5',
      max_tokens: 2048,
      betas: ['server-side-fallback-2026-07-01'],
      fallbacks: 'default',
      output_config: { effort: 'low' },
      system: SYSTEM,
      messages: [{ role: 'user', content: `型號：${model}` }],
    });
    if (msg.stop_reason === 'refusal') {
      return Response.json({ error: '這個型號查不到，請手動填寫' }, { status: 422 });
    }
    const text = msg.content.filter((b) => b.type === 'text').map((b) => b.text).join('').trim();
    return Response.json({ text: text || '查無結果，請手動填寫' });
  } catch (err) {
    if (err instanceof Anthropic.AuthenticationError) {
      return Response.json({ error: 'Anthropic API Key 錯誤，請檢查 Netlify 的 ANTHROPIC_API_KEY' }, { status: 500 });
    }
    if (err instanceof Anthropic.RateLimitError) {
      return Response.json({ error: '查詢太頻繁或額度不足，請稍後再試' }, { status: 429 });
    }
    if (err instanceof Anthropic.APIError) {
      return Response.json({ error: `AI 服務錯誤（${err.status}）` }, { status: 502 });
    }
    throw err;
  }
};
