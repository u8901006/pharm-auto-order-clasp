# 叫藥自動化（省略規格版）— Apps Script + clasp
## 使用
```bash
npm i -g @google/clasp
clasp login
# 若已有 Script ID：編輯 .clasp.json 填入 scriptId，然後
clasp push
# 或新建專案：
# clasp create --title "Pharm Auto Order (no spec)" --type standalone
# clasp push
把 src/Code.gs 的 SPREADSHEET_ID 設為你的試算表（若非容器綁定）。
分頁需有：

【結果】：含廠商與藥品自由文字

【常見量對照】：商品｜常見叫藥數量（可選：規格｜廠商）
執行 generateVendorOrderLines() 後會產生【訂單文字】。
