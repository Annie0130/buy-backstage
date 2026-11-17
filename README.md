# 林相攝影收購後台

收購後台網站（Netlify `precious-otter`）和店裡電腦的飛鼠上架程式。給接手的人先看 `uploader-src/交接說明.md`；老闆看 `uploader-src/使用說明.txt`。

## Airtable 分成兩個 base

收購紀錄那個 base 超過免費版 1,000 筆，POS 庫存和印卡清單已搬到新 base。收購紀錄留在舊的。

| 資料 | Base | 表 | 程式 |
|---|---|---|---|
| 收購紀錄 | `appAuc0SgpUMOFLsv`（林相攝影收購紀錄） | `tblnOp2xI0RPOAR3J` | `netlify/functions/admin-records.js`、上架程式的 `baseId` |
| POS 庫存 | `appbohLgY8zi2592Z`（林相攝影POS庫存） | `tblIljPcfonMR2pia` | `netlify/functions/pos-stock.js`、上架程式的 `posBaseId` |
| 印卡清單 | 同一個新 base | `tblcfYzEVp6d0quU2` | `netlify/functions/print-queue.js` |

網站只用一支 `AIRTABLE_TOKEN`。POS 和印卡的 ID 已寫在程式裡，Netlify 不用改環境變數。要換表再設 `AIRTABLE_POS_BASE_ID`、`AIRTABLE_POS_TABLE_ID`、`AIRTABLE_PRINT_TABLE_ID`。

上架程式的 `airtable-token.txt` 和網站那支權杖，都要能讀寫這兩個 base（`data.records:read` 和 `data.records:write`）。少開新 base 的話，POS 庫存、印卡、以及 POS 的飛鼠狀態都寫不回去。
