# Spec: gmail-fund-tracker

## Objective
從 Gmail 抓取元大投信基金收益分配通知書 PDF 附件（密碼為身分證字號），
解析欄位後彙整輸出為 Excel。
使用者：持有元大投信基金的個人投資者。
成功標準：每次執行後 Excel 正確新增當期收益分配資料，無重複。

## Tech Stack
Python 3.10+、google-auth-oauthlib、google-api-python-client、
pymupdf、openpyxl、python-dotenv、tenacity

## Commands
```
安裝：pip install -r requirements.txt
執行：python main.py
```

## Project Structure
```
main.py             → 主程式（Gmail 認證、抓附件、解析、輸出）
credentials.json    → Google OAuth 憑證（不進 git）
token.json          → OAuth token 快取（自動產生，不進 git）
.env                → 身分證字號與搜尋條件（不進 git）
基金收益分配彙整.xlsx → 輸出結果
```

## Code Style
```python
def _extract_fields(pdf_bytes: bytes, password: str) -> dict:
    ...
```
命名：snake_case，繁體中文變數名稱可用於 Excel 欄位定義。

## Testing Strategy
目前無自動測試。手動驗證：首次執行 Gmail OAuth 授權流程、PDF 解析欄位正確性。
PDF 格式若異動需調整 `_extract_fields()` 內的 regex。

## Boundaries
- Always: credentials.json / token.json / .env 不進 git
- Ask first: 新增支援其他投信公司格式
- Never: 明文儲存身分證字號於程式碼內

## Success Criteria
- OAuth 授權一次後 token.json 可快取，後續免瀏覽器
- 每封郵件的 PDF 正確解密並解析欄位
- Excel 逐次 append，不覆蓋歷史資料
