# gmail-fund-tracker

## 語言規則
所有回覆、審查意見、建議與說明，一律使用**繁體中文**。

## 專案目標
從 Gmail 抓取元大投信基金收益分配通知書（PDF 附件，密碼為身分證字號），解析後彙整成 Excel。

## 架構
- `main.py` — 主程式（Gmail 認證、抓附件、解析 PDF、輸出 Excel）
- `credentials.json` — Google OAuth 憑證（使用者自行放入，不進 git）
- `token.json` — OAuth token 快取（自動產生，不進 git）
- `.env` — 身分證字號與搜尋條件（不進 git）

## 使用方式
1. 複製 `.env.example` 為 `.env`，填入身分證字號
2. 將 `credentials.json` 放入專案根目錄
3. `pip install -r requirements.txt`
4. `python main.py`
5. 首次執行會開啟瀏覽器進行 Gmail 授權

## PDF 欄位萃取
元大投信 PDF 格式若有變動，請調整 `main.py` 內 `_extract_fields()` 的 regex。
