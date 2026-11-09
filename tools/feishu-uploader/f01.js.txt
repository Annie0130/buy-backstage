// 資料夾整理（只會做一次）：最外層只留「★控制台」「★一鍵上架」「更新程式」「照片放這裡」，其他全部搬進「系統」
// 舊版的程式放在最外層；新版從 系統 資料夾執行。這支會被 update.js／panel.js／upload.js 自動呼叫。
const fs = require('fs');
const path = require('path');

const SYS = '系統';
const INBOX = '照片放這裡';

// 這個程式所在的資料夾是「舊的最外層」嗎？（還沒有 系統 資料夾、而且自己不在 系統 裡）
function needsMigration(dir) {
  return path.basename(dir) !== SYS && !fs.existsSync(path.join(dir, SYS)) && fs.existsSync(path.join(dir, 'upload.js'));
}

const LAUNCHERS = {
  '★控制台.bat': '@echo off\r\nchcp 65001 >nul\r\npushd "%~dp0系統"\r\ntitle 上架控制台（這個視窗不要關）\r\necho ===== 上架控制台 =====\r\necho 瀏覽器會自動打開 http://localhost:3939\r\necho 這個黑色視窗不要關，關掉控制台就停了\r\necho.\r\nnode panel.js\r\npause\r\n',
  '★一鍵上架.bat': '@echo off\r\nchcp 65001 >nul\r\npushd "%~dp0系統"\r\ntitle 一鍵上架\r\necho ===== 一鍵上架（沒照片時：顯示拍照順序／有照片時：上架）=====\r\necho.\r\nnode upload.js all\r\necho.\r\necho ===== 跑完了，可以關掉這個視窗 =====\r\npause\r\n',
  '更新程式.bat': '@echo off\r\nchcp 65001 >nul\r\npushd "%~dp0系統"\r\ntitle 更新程式\r\necho ===== 更新上架程式到最新版 =====\r\necho （不會動金鑰、紀錄、照片）\r\necho.\r\nnode update.js\r\necho.\r\necho ===== 跑完了，可以關掉這個視窗 =====\r\npause\r\n',
  '更新程式.command': '#!/bin/bash\n# Mac 用：雙擊執行\ncd "$(dirname "$0")/系統"\necho "===== 更新上架程式到最新版 ====="\necho\nnode update.js\necho\nread -p "跑完了，按 Enter 關閉…"\n',
  '★控制台.command': '#!/bin/bash\n# Mac 用：雙擊執行\ncd "$(dirname "$0")/系統"\necho "===== 上架控制台（這個視窗不要關）====="\necho\nnode panel.js\necho\nread -p "跑完了，按 Enter 關閉…"\n',
};
function writeLaunchers(root) {
  for (const [name, body] of Object.entries(LAUNCHERS)) {
    const p = path.join(root, name);
    fs.writeFileSync(p, body, 'utf8');
    if (name.endsWith('.command')) { try { fs.chmodSync(p, 0o755); } catch (e) {} }
  }
}

function migrate(root) {
  if (!needsMigration(root)) return false;
  console.log('\n📁 整理資料夾：最外層只留「★控制台」「★一鍵上架」「更新程式」「照片放這裡」，其他搬進「系統」…');
  const sys = path.join(root, SYS);
  const inbox = path.join(root, INBOX);
  fs.mkdirSync(sys, { recursive: true });
  fs.mkdirSync(inbox, { recursive: true });
  const failed = [];
  for (const name of fs.readdirSync(root)) {
    if (name === SYS || name === INBOX) continue;
    try { fs.renameSync(path.join(root, name), path.join(sys, name)); } catch (e) { failed.push(name); }
  }
  // 舊的收件匣（系統\商品照片\收件匣）裡還沒處理的照片，搬到 照片放這裡
  const oldInbox = path.join(sys, '商品照片', '收件匣');
  if (fs.existsSync(oldInbox)) {
    for (const f of fs.readdirSync(oldInbox)) {
      try { fs.renameSync(path.join(oldInbox, f), path.join(inbox, f)); } catch (e) {}
    }
  }
  // 系統 裡的舊啟動檔用不到了（最外層會放新的），移到 舊版備份
  const bk = path.join(sys, '舊版備份', '整理資料夾前的按鈕');
  for (const name of Object.keys(LAUNCHERS)) {
    const p = path.join(sys, name);
    if (fs.existsSync(p)) { try { fs.mkdirSync(bk, { recursive: true }); fs.renameSync(p, path.join(bk, name)); } catch (e) {} }
  }
  writeLaunchers(root);
  if (failed.length) console.log(`   ⚠️ 這些沒搬成功（可能正在使用中），請關掉所有黑色視窗後，自己拖進「系統」：${failed.join('、')}`);
  console.log('   ✅ 整理好了！以後雙擊最外層的「★控制台」就好');
  return true;
}

module.exports = { migrate, needsMigration, writeLaunchers, SYS, INBOX };
if (require.main === module) migrate(process.argv[2] || __dirname);
