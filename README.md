# Gmail 郵件內容下載工具

一個自動化工具，用於從 Gmail 下載附件與郵件內文，支援多種郵件類型的篩選與整理。

## ✨ 主要功能

- **📥 自動下載**: 從 Gmail 自動下載附件與郵件內文
- **🔍 智能篩選**: 支援主旨關鍵字、時間範圍、檔案類型篩選
- **📁 自動整理**: 依主旨自動建立資料夾並分類檔案
- **📅 日期分類**: 支援按日期自動創建子資料夾（`all/YYYY-MM-DD/`）
- **🔐 安全認證**: 自動處理 Gmail API 認證與 token 更新
- **📊 完整日誌**: 詳細的執行記錄與錯誤追蹤

## 🚀 快速開始

### 1. 克隆專案
```bash
git clone git@github.com:kuonumber/gmail_parser.git
cd gmail_parser
```

### 2. 安裝依賴
```bash
pip install -r requirements.txt
```

### 3. 設定 Gmail API
1. 前往 [Google Cloud Console](https://console.cloud.google.com/) 建立專案
2. 啟用 Gmail API
3. 建立 OAuth 2.0 用戶端 ID（桌面應用程式）
4. 下載認證檔，重新命名為 `gmail_cred.json` 放在專案根目錄

### 4. 建立組態檔
建立 `.env` 檔案：
```ini
# 基本設定
GMAIL_DOWNLOAD_DIR=./downloads
GMAIL_SUBJECTS=發票,通知,報告
GMAIL_FILE_TYPES=pdf,xls,xlsx,csv
GMAIL_DOWNLOAD_CONTENT=true

# 時間範圍（今天）
GMAIL_DATE_RANGE=today

# 主旨對應資料夾（留空表示存入 all 資料夾）
GMAIL_SUBJECT_FOLDER_MAPPING=
```

### 5. 執行下載
```bash
python gmail_downloader.py
```

## 📖 使用說明

### 下載郵件與附件
```bash
python gmail_downloader.py
```
- 依 `.env` 組態執行
- 在 `downloads/` 內產出對應主旨資料夾
- 若啟用 `GMAIL_DOWNLOAD_CONTENT=true`，同時生成 `*_content.txt`
- 支援按日期自動分類：`all/YYYY-MM-DD/` 格式

## 🏗️ 專案結構

```
gmail_parser/
├── gmail_downloader.py          # 主要下載工具
├── requirements.txt             # Python 依賴套件
├── gmail_cred.json              # Gmail API 認證檔
├── token.pickle                 # 認證 Token
├── downloads/                   # 下載輸出目錄
└── .env                        # 組態檔案
```

## ⚙️ 組態選項

| 變數 | 說明 | 預設值 |
|------|------|--------|
| `GMAIL_DOWNLOAD_DIR` | 下載目標目錄 | `./downloads` |
| `GMAIL_SUBJECTS` | 郵件主旨關鍵字（逗號分隔） | `發票,通知,報告` |
| `GMAIL_FILE_TYPES` | 要下載的檔案類型 | `pdf,xls,xlsx,csv` |
| `GMAIL_DOWNLOAD_CONTENT` | 是否下載郵件內文 | `true` |
| `GMAIL_DATE_RANGE` | 搜尋時間範圍 | `30d` |
| `GMAIL_SUBJECT_FOLDER_MAPPING` | 主旨到資料夾的對應 | 無 |
| `GMAIL_DATE_RANGE` | 搜尋時間範圍 | `today` |

### 時間範圍格式
- 預設選項：`today`（今天）, `yesterday`（昨天）, `week`（週）, `month`（月）, `year`（年）, `all`（全部）
- 相對範圍：`7d`, `14d`, `30d`（最近 N 天）
- 絕對日期：`yyyy/mm/dd`（需搭配 `GMAIL_START_DATE` 和 `GMAIL_END_DATE`）

## 🔐 認證流程

1. **首次執行**: 開啟瀏覽器授權 Gmail 存取權限
2. **Token 快取**: 認證成功後儲存於 `token.pickle`
3. **自動更新**: Token 過期時自動重新認證

## 📝 日誌記錄

- 日誌檔案：`app.log`
- 記錄內容：執行過程、錯誤訊息、下載狀態
- 支援日誌輪替（按日期分割）

## ⚠️ 注意事項

- `gmail_cred.json` 與 `token.pickle` 為敏感資訊，請勿上傳至 Git
- `downloads/` 目錄已加入 `.gitignore`，不會被版本控制
- 建議使用 `.env` 檔案管理組態，避免將私密資訊寫入程式碼

## 🐛 常見問題

### 認證錯誤
- 確認 `gmail_cred.json` 存在且格式正確
- 確認已啟用 Gmail API
- 檢查網路連線狀態

### 沒有下載到檔案
- 檢查主旨關鍵字是否正確
- 確認時間範圍設定
- 檢查附件副檔名是否在 `GMAIL_FILE_TYPES` 中

### 權限錯誤
- 確認目錄可寫入
- 檢查磁碟空間是否充足
- 確認檔案權限設定

## 🤝 貢獻指南

歡迎提交 Issue 或 Pull Request 來改善專案！

## 📄 授權

本專案僅供個人使用，請遵守相關服務條款與隱私政策。

## 📞 支援

如有問題或建議，請在 GitHub 上建立 Issue。

---

**重要提醒**: 請妥善保護你的憑證與組態檔，不要將任何私密資訊提交到版本控制系統。
