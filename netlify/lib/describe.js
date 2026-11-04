// 請 Claude 依型號判斷品類（片幅）並寫一段商品介紹。
// 合約網站（buy-form）和收購後台（buy-backstage）各有一份相同的檔案，修改時兩邊一起改。
import Anthropic from '@anthropic-ai/sdk';

// 要跟 Airtable「品類」欄位的選項一致
export const CATEGORIES = ['全片幅無反相機', 'APS-C 無反相機', 'M43 無反相機', '全片幅單眼相機', 'APS-C 單眼相機',
  '全片幅定焦鏡頭', '全片幅變焦鏡頭', 'APS-C 定焦鏡頭', 'APS-C 變焦鏡頭', 'M43 鏡頭', '隨身相機', '其他'];

const SYSTEM = `你是二手相機店「林相攝影」的商品資料助手。使用者會給你一個相機或鏡頭型號，請回傳品類與商品介紹。

品類（category）：從選項中選一個。鏡頭依「設計給哪種片幅」判斷（例如 Canon RF-S、Sony E 非 FE、Fujifilm XF/XC、Nikon DX 是 APS-C；RF、FE、Z 非 DX、EF 非 EF-S 是全片幅；副廠鏡頭看代號：Sigma 的 DG 是全片幅、DC 是 APS-C，Tamron 的 Di 與 Di III 是全片幅、Di II 與 Di III-A 是 APS-C，Tokina 的 FF 是全片幅、DX 是 APS-C）。固定鏡頭的小型相機、運動相機、手持雲台相機選「隨身相機」。不確定就選「無法判斷」。

商品介紹（intro）：繁體中文一段話，60 到 80 字，不要條列、不要換行。先寫發售年份與最重要的規格（相機：畫素、對焦、防震、重量等；鏡頭：對焦馬達、鏡片結構、重量等），再寫一兩個主要特色和適合拍什麼。商品卡上方已經顯示品牌、型號、品類（片幅與相機／鏡頭類型）和貨品來源，介紹裡不要再重複這些。例：2021 年發售，2420 萬畫素背照式 CMOS，693 點相位對焦，機身僅 509g，高感光表現出色，旅行街拍輕巧好帶。
只寫你確定的規格，不確定的數字寧可不寫，不要猜。認不出這個型號時，intro 回空字串。`;

// 匯入舊庫存時才用：把備註欄（配件、快門、保固、瑕疵混在一起）拆成卡片欄位
const NOTES_GUIDE = `

另外會附上一段店內備註，裡面可能混著隨附配件、快門數、原廠保固到期日、外觀與功能狀況。請拆成：
- shutter：快門數，只寫數字（「快9000次內」「9000內」寫 9000；「17XXX」這種不精確的寫 17000）；沒有就空字串
- accessories：隨附配件清單，例如 ["電池","背帶","充電器","遮光罩"]；「無遮光罩」這種寫明沒有的不要列；沒有就空陣列
- warranty_until：原廠保固到期日，格式 YYYY-MM-DD；沒寫就空字串
- condition：外觀新舊，從選項中選。備註沒提到外觀問題就選「近無使用痕跡」；提到外觀痕跡、刮傷、掉漆等小瑕疵選「些許使用痕跡」；寫明痕跡多或明顯就選「使用痕跡多」
- remarks：功能或外觀瑕疵，照備註原意改寫成給客人看的一句話（例：「前玉輕微發霉，不影響成像」）；沒有瑕疵就空字串`;

const NOTES_SCHEMA = {
  shutter: { type: 'string' },
  accessories: { type: 'array', items: { type: 'string' } },
  warranty_until: { type: 'string' },
  condition: { type: 'string', enum: ['全新', '近無使用痕跡', '些許使用痕跡', '使用痕跡多', ''] },
  remarks: { type: 'string' },
};

// 回傳 {category, intro}；有給 notes 時另外回傳 {shutter, accessories, warranty_until, condition, remarks}。
// category 判斷不出來時是空字串。被拒答時回傳 null。
export async function describeGear(name, notes) {
  const withNotes = notes !== undefined;
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
            ...(withNotes ? NOTES_SCHEMA : {}),
          },
          required: ['category', 'intro', ...(withNotes ? Object.keys(NOTES_SCHEMA) : [])],
          additionalProperties: false,
        },
      },
    },
    system: withNotes ? SYSTEM + NOTES_GUIDE : SYSTEM,
    messages: [{ role: 'user', content: withNotes ? `型號：${name}\n店內備註：${notes || '（無）'}` : `型號：${name}` }],
  });
  if (msg.stop_reason === 'refusal') return null;
  const out = JSON.parse(msg.content.filter((b) => b.type === 'text').map((b) => b.text).join(''));
  const result = {
    category: CATEGORIES.includes(out.category) ? out.category : '',
    intro: (out.intro || '').trim(),
  };
  if (withNotes) {
    Object.assign(result, {
      shutter: String(out.shutter || '').replace(/\D/g, ''),
      accessories: (out.accessories || []).map((a) => String(a).trim()).filter(Boolean),
      warranty_until: /^\d{4}-\d{2}-\d{2}$/.test(out.warranty_until || '') ? out.warranty_until : '',
      condition: out.condition || '',
      remarks: (out.remarks || '').trim(),
    });
  }
  return result;
}
