# ESG 永續報告書下載系統

以 Streamlit 包裝原有的臺灣證券交易所 ESG 永續報告書查詢與 PDF 下載流程。

## 本機執行

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Excel 格式

工作表名稱：`公司清單`。第一列須包含 `證券代碼`、`公司名稱`。第一版每次最多 5 家公司。

## 部署

1. 將 `app.py`、`backend.py`、`requirements.txt`、`.gitignore` 上傳到 GitHub repository 根目錄。
2. 到 https://share.streamlit.io/ 登入 GitHub 並建立應用程式。
3. 選擇 repository、branch，main file path 設為 `app.py`。
4. 部署後使用提供的網址測試一家公司是否能成功下載 PDF。

## 注意

- 不保證 TWSE API 在所有雲端環境都可存取；如遇 403、逾時或其他 API 變更，需進一步排查。
- 公開測試版限制每次 5 家公司、單份 PDF 60 MB，避免過度消耗免費主機資源。
- 使用前請遵守資料來源網站的服務條款與合理請求頻率。
- 查詢年度的實際可用資料由 API 決定。
