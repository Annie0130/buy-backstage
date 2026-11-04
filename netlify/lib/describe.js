// 請 Claude 依型號判斷品類（片幅）並寫一段商品介紹。
// 合約網站（buy-form）和收購後台（buy-backstage）各有一份相同的檔案，修改時兩邊一起改。
import Anthropic from '@anthropic-ai/sdk';

// 要跟 Airtable「品類」欄位的選項一致
export const CATEGORIES = ['全片幅無反相機', 'APS-C 無反相機', 'M43 無反相機', '全片幅單眼相機', 'APS-C 單眼相機',
  '全片幅定焦鏡頭', '全片幅變焦鏡頭', 'APS-C 定焦鏡頭', 'APS-C 變焦鏡頭', 'M43 鏡頭', '隨身相機', '其他'];

const SYSTEM = `你是二手相機店「林相攝影」的商品資料助手。使用者會給你一個相機或鏡頭型號，請回傳品類與商品介紹。

品類（category）：從選項中選一個。鏡頭依「設計給哪種片幅」判斷（例如 Canon RF-S、Sony E 非 FE、Fujifilm XF/XC、Nikon DX 是 APS-C；RF、FE、Z 非 DX、EF 非 EF-S 是全片幅）。固定鏡頭的小型相機、運動相機、手持雲台相機選「隨身相機」。不確定就選「無法判斷」。

商品介紹（intro）：繁體中文一段話，60 到 80 字，不要條列、不要換行。先寫發售年份與最重要的規格（相機：畫素、感光元件尺寸；鏡頭：焦段、光圈、適用片幅），再寫一兩個主要特色和適合拍什麼。例：2021 年發售，2420 萬畫素全片幅背照式 CMOS，693 點相位對焦，機身僅 509g，高感光出色，輕巧全片幅首選。
只寫你確定的規格，不確定的數字寧可不寫，不要猜。認不出這個型號時，intro 回空字串。`;

// 回傳 {category, intro}；category 判斷不出來時是空字串。被拒答時回傳 null。
export async function describeGear(name) {
  const client = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });
  const msg = await client.beta.messages.create({
    model: 'claude-opus-5-5',
    max_tokens: 2048,
    betas: ['server-side-fallback-2026-07-01'],
    fallbacks: 'default',
    output_config: {
      effort: 'low',
      format: {
        type: 'json_schema',
        schema: {
          type: 'object',
          properties: {
            category: { type: 'string', enum: [...CATEGORIES, '無法判斷'] },
            intro: { type: 'string' },
          },
          required: ['category', 'intro'],
          additionalProperties: false,
        },
      },
    },
    system: SYSTEM,
    messages: [{ role: 'user', content: `型號：${name}` }],
  });
  if (msg.stop_reason === 'refusal') return null;
  const out = JSON.parse(msg.content.filter((b) => b.type === 'text').map((b) => b.text).join(''));
  return {
    category: CATEGORIES.includes(out.category) ? out.category : '',
    intro: (out.intro || '').trim(),
  };
}
