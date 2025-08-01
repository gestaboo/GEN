
# Flask Word 表單產生器

## ✅ 功能
- 手機或電腦瀏覽器開啟 HTML 表單
- 使用者填寫欄位後，套用 Word 模板並下載 `.docx`

## 🚀 安裝方式
```bash
pip install -r requirements.txt
python app.py
```

## 🌐 預設網址
```
http://127.0.0.1:5000
```

## ☁️ 免費部署推薦：Railway.app
1. 到 https://railway.app 註冊帳號
2. 新增專案 → 選擇「Deploy from GitHub repo」
3. 上傳本專案（或初始化 Git 後 push）
4. 設定執行指令：
   ```
   python app.py
   ```

## 📁 目錄說明
- `templates/form_template.docx`：Word 表單模板
- `templates/form.html`：網頁填寫表單
- `app.py`：Flask 主程式
- `requirements.txt`：依賴套件清單
