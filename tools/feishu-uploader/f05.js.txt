// 已併入 upload.js：這支等同於「node upload.js shopee」
const U = require('./upload.js');
process.argv.splice(2, 0, 'shopee');
U.shopeeMain().then(() => U.closeAsk()).catch((e) => { U.closeAsk(); console.error('❌ ' + (e.message || e)); process.exit(1); });
