# Gmail 電子發票解析器

一個自動化工具，用於從Gmail下載財政部電子發票整合服務平台的附件檔案。

## 🚀 功能特色

- **自動化下載**: 自動從Gmail下載電子發票附件
- **支援多種格式**: 支援PDF、Excel (.xls, .xlsx)、CSV等檔案格式
- **重複檢查**: 避免重複下載已處理的郵件
- **日誌記錄**: 完整的執行日誌記錄，支援日誌輪轉
- **認證管理**: 自動處理Gmail API認證和token更新

## 📋 系統需求

- Python 3.7+
- macOS/Linux/Windows
- Gmail帳戶
- 網路連線

## 🛠️ 安裝步驟

### 1. 克隆專案
```bash
git clone git@github.com:kuonumber/gmail_parser.git
cd gmail_parser
```

### 2. 安裝依賴
```bash
pip install -r requirements.txt
```

### 3. 設定Gmail API認證

#### 3.1 建立Google Cloud專案
1. 前往 [Google Cloud Console](https://console.cloud.google.com/)
2. 建立新專案或選擇現有專案
3. 啟用 Gmail API

#### 3.2 建立認證檔案
1. 在「API和服務」→「認證」中建立OAuth 2.0用戶端ID
2. 選擇「桌面應用程式」
3. 下載JSON認證檔案
4. 將檔案重新命名為 `gmail_cred.json` 並放在專案根目錄

## 🔧 使用方法

### 基本使用
```bash
python email_downloader.py
```

### 自訂處理數量
程式會提示輸入要處理的郵件數量，預設為10筆。

### 下載位置
- 電子發票附件會下載到 `einvoice/` 目錄
- 檔案命名格式: `{郵件ID}_{原始檔名}`

## 📁 專案結構

```
gmail_parser/
├── email_downloader.py      # 主要程式檔案
├── requirements.txt          # Python依賴套件
├── gmail_cred.json          # Gmail API認證檔案 (需自行設定)
├── token.pickle             # 認證token快取 (自動產生)
├── einvoice/                # 下載的電子發票檔案目錄
├── app.log*                 # 應用程式日誌檔案
├── already_parsed_mails.txt # 已處理郵件記錄
└── README.md               # 專案說明文件
```

## 🔐 認證說明

- **首次執行**: 會開啟瀏覽器要求授權Gmail存取權限
- **Token快取**: 認證成功後會儲存到 `token.pickle` 檔案
- **自動更新**: 當token過期時會自動重新認證

## 📊 支援的檔案格式

- PDF (.pdf)
- Excel (.xls, .xlsx)
- CSV (.csv)

## 📝 日誌記錄

- 日誌檔案: `app.log`
- 日誌輪轉: 每日自動備份，最多保留30天
- 記錄內容: 執行過程、錯誤訊息、下載狀態

## ⚠️ 注意事項

1. **認證檔案**: `gmail_cred.json` 包含敏感資訊，請勿上傳到公開儲存庫
2. **API配額**: Gmail API有使用配額限制，請注意使用頻率
3. **檔案大小**: 大量附件下載可能耗時較長
4. **網路連線**: 需要穩定的網路連線

## 🐛 故障排除

### 常見問題

1. **認證錯誤**
   - 檢查 `gmail_cred.json` 檔案是否存在且格式正確
   - 確認Gmail API已啟用

2. **權限錯誤**
   - 確認程式有寫入當前目錄的權限
   - 檢查磁碟空間是否充足

3. **網路連線問題**
   - 檢查網路連線狀態
   - 確認防火牆設定

## 📄 授權

本專案僅供個人使用，請遵守相關服務條款和隱私政策。

## 🤝 貢獻

歡迎提交Issue和Pull Request來改善這個專案。

## 📞 支援

如有問題或建議，請在GitHub上建立Issue。

---

**注意**: 使用本工具時請遵守Gmail的使用條款和相關法規。
