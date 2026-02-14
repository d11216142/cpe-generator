# CPE 編號產生器 (CPE Generator)

一個基於 Flask 的 Web 應用程式，用於管理和生成 Common Platform Enumeration (CPE) 編號。

## 功能特色

### 資料庫連線管理
1. 支援新增本地資料庫連線設定（僅限 localhost，基於安全考量）
2. 可儲存多個資料庫連線配置
3. 支援 Windows 驗證與 SQL Server 驗證兩種方式
4. 提供連線測試功能，確保連線設定正確
5. 可隨時切換、編輯或刪除已儲存的連線
6. 連線資訊安全儲存，密碼在介面上會被遮蔽

### 功能一：CPE 編號擷取與解析
1. 透過 CPE 字典自動抓取 CPE 編號
2. 分析該 CPE 編號是否真實，如果否，將其刪除並繼續下一筆
3. 將抓取到的 CPE 編號拆分成「供應商」、「產品名稱」、「版本號」，如果其他欄位有內容，將其整理列在其他欄位並敘述
4. 將拆分的資料列在畫面上，並模擬其在電腦上應該的大小(MB)、安裝日期、安裝位置(預設 C 槽)
5. 讓使用者可下載成 CSV 檔

### 功能二：隨機產生 CPE 編號
1. 依照常見的公司名稱，隨機生成供應商、產品名稱、版本號、大小(MB)、安裝日期、安裝位置(預設 C 槽)
2. 將其製作成 CPE 編號
3. 將製作號的 CPE 編號拿去 CPE 字典查找，如果不正確，將其改成最接近的結果

## 安裝說明

### 前置需求
- Python 3.7 或更高版本
- pip (Python 套件管理工具)

### 安裝步驟

1. 複製此專案到本地端
```bash
git clone https://github.com/d11216142/cpe-generator.git
cd cpe-generator
```

2. 安裝所需套件
```bash
pip install -r requirements.txt
```

## 使用方法

### 啟動應用程式

```bash
python app.py
```

應用程式將在 `http://localhost:5000` 啟動

**注意**: 預設情況下，應用程式在生產模式下運行（debug=False）。如果需要開發模式，請設定環境變數：
```bash
export FLASK_DEBUG=true
python app.py
```

⚠️ **安全提醒**: 切勿在生產環境中啟用 debug 模式，這可能會讓攻擊者透過除錯器執行任意程式碼。

### 使用介面

1. **資料庫連線管理**
   - 在「資料庫連線管理」區塊設定資料庫連線
   - 輸入連線名稱（例如：Production DB, Test DB）
   - 輸入伺服器位址（僅限本地：localhost, 127.0.0.1, localhost\SQLEXPRESS 等）
   - 輸入資料庫名稱
   - 選擇驗證方式：
     - 勾選「使用 Windows 驗證」：使用當前 Windows 帳號
     - 不勾選：需輸入 SQL Server 使用者名稱和密碼
   - 點擊「測試連線」確認設定正確
   - 點擊「儲存連線」保存配置
   - 在「已儲存的連線」列表中可以：
     - 編輯：修改連線設定
     - 使用：切換到該連線
     - 刪除：移除該連線配置

2. **自動抓取 CPE 編號**
   - 在「功能一」區塊輸入要抓取的數量 (預設 10，可調整為 1-100)
   - 點擊「自動抓取 CPE」按鈕
   - 系統會自動從 CPE 字典中隨機抓取指定數量的 CPE 編號
   - 每個 CPE 會被解析成供應商、產品名稱、版本號等欄位
   - 如果其他欄位有資訊，會顯示在「其他欄位」並說明內容
   - 系統會自動模擬大小(MB)、安裝日期、安裝位置(預設 C 槽)
   - 結果會顯示在下方的表格中，包含統計資訊

2. **產生隨機 CPE**
   - 在「功能二」區塊設定要產生的數量 (1-50)
   - 點擊「產生隨機 CPE」按鈕
   - 系統會根據常見公司名稱隨機產生 CPE 資料

3. **匯出 CSV**
   - 在解析或產生 CPE 資料後
   - 點擊「下載 CSV 檔案」按鈕
   - CSV 檔案會自動下載到您的電腦

## 資料庫配置說明

### 連線資訊儲存
- 資料庫連線配置儲存在 `db_connections.json` 檔案中
- 此檔案已加入 `.gitignore`，不會被提交到版本控制系統
- 連線資訊包含：連線名稱、伺服器位址、資料庫名稱、驗證方式、使用者名稱、密碼

### 安全建議
⚠️ **重要安全提醒**：
- 不要將包含實際密碼的 `db_connections.json` 檔案分享給他人
- 在生產環境中，建議使用 Windows 驗證而非 SQL Server 驗證
- 確保資料庫伺服器已設定適當的防火牆規則
- 定期更新資料庫密碼並限制存取權限

### 資料庫需求
- 支援 Microsoft SQL Server
- 需要安裝 ODBC Driver 17 for SQL Server 或更高版本
- pyodbc 套件（已包含在 requirements.txt）

## API 端點說明

### 資料庫連線管理 API
- `GET /api/db-connections` - 取得所有已儲存的連線
- `POST /api/db-connections` - 新增資料庫連線
- `PUT /api/db-connections/<name>` - 更新現有連線
- `DELETE /api/db-connections/<name>` - 刪除連線
- `POST /api/db-connections/test` - 測試連線
- `POST /api/db-connections/set-current` - 設定當前使用的連線

### CPE 相關 API
- `POST /api/auto-fetch-cpe` - 自動抓取 CPE 編號
- `POST /api/generate-random` - 產生隨機 CPE
- `POST /api/export-csv` - 匯出 CSV
- `POST /api/export-xlsx` - 匯出 XLSX
- `POST /api/export-json` - 匯出 JSON

## 技術架構

- **後端**: Flask (Python)
- **前端**: HTML5, CSS3, JavaScript (原生)
- **資料格式**: CPE 2.3 URI 格式
- **匯出格式**: CSV (UTF-8 with BOM)

## CPE 格式說明

CPE (Common Platform Enumeration) 是一種標準化的方法，用於描述和識別資訊科技系統、軟體和套件。

CPE 2.3 URI 格式範例:
```
cpe:2.3:a:microsoft:windows:10:*:*:*:*:*:*:*
```

格式說明:
- `cpe`: CPE 前綴
- `2.3`: CPE 版本
- `a`: 部分 (a=應用程式, o=作業系統, h=硬體)
- `microsoft`: 供應商
- `windows`: 產品名稱
- `10`: 版本號
- 其他欄位可包含更新、版本、語言等資訊

## 資料欄位說明

應用程式會顯示以下欄位:
- **CPE 編號**: 完整的 CPE URI 字串
- **供應商**: 軟體或產品的製造商
- **產品名稱**: 產品的名稱
- **版本號**: 產品的版本
- **其他欄位**: 包含更新、版本、語言等額外資訊
- **大小 (MB)**: 模擬的安裝大小
- **安裝日期**: 模擬的安裝日期
- **安裝位置**: 模擬的安裝路徑 (預設 C 槽)

## 授權

此專案為開源專案。

## 常見問題與疑難排解

### 安裝相關問題

#### Q1: 安裝 pyodbc 時出現錯誤
**問題描述**: 執行 `pip install -r requirements.txt` 時，pyodbc 安裝失敗。

**解決方案**:
1. **Windows 系統**:
   - 確保已安裝 Microsoft Visual C++ 14.0 或更高版本
   - 下載並安裝 [Microsoft C++ Build Tools](https://visualstudio.microsoft.com/visual-cpp-build-tools/)
   - 或安裝完整的 Visual Studio

2. **Linux 系統**:
   ```bash
   # Ubuntu/Debian
   sudo apt-get update
   sudo apt-get install unixodbc unixodbc-dev
   
   # CentOS/RHEL
   sudo yum install unixODBC unixODBC-devel
   ```

3. **macOS 系統**:
   ```bash
   brew install unixodbc
   ```

#### Q2: 缺少 ODBC Driver for SQL Server
**問題描述**: 連線資料庫時出現「找不到 ODBC Driver」錯誤。

**解決方案**:
- **Windows**: 從 Microsoft 官網下載並安裝 [ODBC Driver 17 for SQL Server](https://docs.microsoft.com/sql/connect/odbc/download-odbc-driver-for-sql-server)
- **Linux**: 參考 [Microsoft 官方文件](https://docs.microsoft.com/sql/connect/odbc/linux-mac/installing-the-microsoft-odbc-driver-for-sql-server)安裝對應版本的 ODBC Driver
- **macOS**: 使用 Homebrew 安裝:
  ```bash
  brew tap microsoft/mssql-release https://github.com/Microsoft/homebrew-mssql-release
  brew update
  brew install msodbcsql17
  ```

### 資料庫連線問題

#### Q3: 無法連線到資料庫
**可能原因與解決方案**:

1. **伺服器位址錯誤**:
   - 確認 SQL Server 服務已啟動
   - 本地連線請使用: `localhost` 或 `127.0.0.1`
   - SQL Server Express: `localhost\SQLEXPRESS`
   - LocalDB: `(localdb)\mssqllocaldb`

2. **Windows 驗證失敗**:
   - 確保勾選「使用 Windows 驗證」選項
   - 確認當前 Windows 帳號有資料庫存取權限
   - 在 SQL Server 中授予適當的權限

3. **SQL Server 驗證失敗**:
   - 取消勾選「使用 Windows 驗證」
   - 確認使用者名稱和密碼正確
   - 確保 SQL Server 已啟用混合驗證模式
   - 檢查使用者帳號是否已啟用且未過期

4. **防火牆設定**:
   - 確保 Windows 防火牆允許 SQL Server 的連線 (預設 port 1433)
   - 檢查 SQL Server 是否已啟用 TCP/IP 協定

#### Q4: 顯示「僅支援本地資料庫連線」錯誤
**問題描述**: 嘗試連線到遠端資料庫時被拒絕。

**解決方案**:
- 此為安全性設計，應用程式僅允許連線到本地資料庫
- 允許的伺服器位址: `localhost`, `127.0.0.1`, `::1`, `(local)`, `.`, `localhost\SQLEXPRESS`, `(localdb)\mssqllocaldb`
- 如需連線遠端資料庫，請在 `db_config.py` 中修改 `is_localhost()` 函式（不建議）

#### Q5: 資料庫連線超時
**解決方案**:
- 增加連線超時時間：在 `db_config.py` 中修改 `CONNECTION_TIMEOUT` 常數（預設 10 秒）
- 檢查網路連線狀態
- 確認 SQL Server 服務正常運行

### 應用程式執行問題

#### Q6: Port 5000 已被占用
**問題描述**: 啟動應用程式時顯示「Address already in use」。

**解決方案**:
1. **更改應用程式 Port**:
   - 在 `app.py` 的最後一行，將 `port=5000` 改為其他 port（例如 `port=5001`）
   
2. **釋放被占用的 Port**:
   ```bash
   # Windows
   netstat -ano | findstr :5000
   taskkill /PID <PID> /F
   
   # Linux/macOS
   lsof -i :5000
   kill -9 <PID>
   ```

#### Q7: Python 版本不相容
**問題描述**: 某些功能無法正常運作或出現語法錯誤。

**解決方案**:
- 確保使用 Python 3.7 或更高版本
- 檢查 Python 版本：`python --version` 或 `python3 --version`
- 如有多個 Python 版本，使用虛擬環境避免衝突:
  ```bash
  python -m venv venv
  # Windows
  venv\Scripts\activate
  # Linux/macOS
  source venv/bin/activate
  ```

#### Q8: 瀏覽器無法開啟應用程式
**解決方案**:
- 確認應用程式已成功啟動（終端機顯示「Running on http://127.0.0.1:5000」）
- 檢查防火牆是否封鎖本地連線
- 嘗試使用不同瀏覽器（推薦：Chrome, Firefox, Edge）
- 清除瀏覽器快取和 Cookie
- 使用無痕模式測試

### 資料處理問題

#### Q9: CSV 檔案中文顯示亂碼
**問題描述**: 下載的 CSV 檔案在 Excel 中開啟時中文顯示為亂碼。

**解決方案**:
- 應用程式已使用 UTF-8 with BOM 編碼，應該可以正常顯示中文
- 如仍有問題，請在 Excel 中使用「資料」>「從文字/CSV」匯入，並選擇 UTF-8 編碼
- 或使用 Google Sheets、LibreOffice 等其他軟體開啟

#### Q10: CPE 解析失敗
**問題描述**: 某些 CPE 編號無法正確解析。

**解決方案**:
- 確認 CPE 格式符合 CPE 2.3 URI 標準
- CPE 必須以 `cpe:2.3:` 開頭
- 檢查 CPE 編號中是否有特殊字元需要跳脫
- 查看應用程式 log 了解詳細錯誤訊息

#### Q11: 隨機產生的 CPE 不夠真實
**說明**:
- 隨機產生功能主要用於測試和演示
- 產生的 CPE 會嘗試與 CPE 字典中的相近項目對比
- 如需真實的 CPE 資料，建議使用「自動抓取 CPE」功能

### 匯出功能問題

#### Q12: 無法下載匯出的檔案
**解決方案**:
- 確認瀏覽器未封鎖彈出視窗或下載
- 檢查瀏覽器下載設定
- 確認磁碟空間充足
- 嘗試使用不同的匯出格式（CSV, XLSX, JSON）

#### Q13: XLSX 檔案無法開啟
**解決方案**:
- 確保已安裝 Microsoft Excel 或相容軟體（LibreOffice, WPS Office）
- 檢查檔案是否完整下載（檔案大小不為 0）
- 嘗試重新匯出

### 效能相關問題

#### Q14: 大量資料處理緩慢
**建議**:
- 單次抓取建議不超過 100 筆 CPE（已設定上限）
- 隨機產生建議不超過 50 筆（已設定上限）
- 如需處理大量資料，建議分批處理
- 考慮使用資料庫儲存功能，避免重複抓取

#### Q15: 記憶體使用過高
**解決方案**:
- 減少單次處理的資料量
- 處理完畢後重新整理頁面釋放記憶體
- 關閉不必要的瀏覽器分頁
- 考慮增加系統記憶體

### 安全性問題

#### Q16: 擔心資料庫密碼安全性
**安全建議**:
- 優先使用 Windows 驗證，避免使用密碼
- 資料庫連線資訊儲存在 `db_connections.json`，此檔案已加入 `.gitignore`
- 不要將 `db_connections.json` 分享給他人
- 在介面上顯示的密碼會被遮蔽為 `***`
- 建議定期更新資料庫密碼
- 在生產環境中，考慮使用環境變數或加密金鑰管理系統

#### Q17: Debug 模式安全性
**重要警告**:
- 預設情況下，應用程式在生產模式運行（`debug=False`）
- **切勿在生產環境或公開網路中啟用 debug 模式**
- Debug 模式可能讓攻擊者透過除錯器執行任意程式碼
- 僅在本地開發環境中使用 debug 模式

### 開發與測試問題

#### Q18: 想要修改 CPE 字典內容
**說明**:
- CPE 字典定義在 `app.py` 的 `CPE_DICTIONARY` 列表中
- 可以新增、修改或刪除 CPE 項目
- 確保 CPE 格式符合 CPE 2.3 URI 標準
- 修改後重新啟動應用程式

#### Q19: 想要調整資料驗證規則
**說明**:
- 資料驗證邏輯主要在 `validate_cpe_with_nvd()` 函式
- 可以根據需求調整驗證條件
- 修改後建議進行充分測試

#### Q20: 如何貢獻程式碼
**步驟**:
1. Fork 此專案
2. 建立新的功能分支
3. 進行修改並測試
4. 提交 Pull Request
5. 等待審核與合併

## 聯絡方式

如遇到其他問題或需要協助，歡迎：
- 在 GitHub 上提交 Issue
- 提供詳細的錯誤訊息和環境資訊（作業系統、Python 版本、錯誤 log）

## 貢獻

歡迎提交 Issue 或 Pull Request 來改進此專案。