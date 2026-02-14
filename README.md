# CPE 編號產生器 (CPE Generator)

一個基於 Flask 的 Web 應用程式，用於管理和生成 Common Platform Enumeration (CPE) 編號。

## 功能特色

### 資料庫連線管理
1. 支援新增本地或遠端的資料庫連線設定
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
   - 輸入伺服器位址（本地：localhost，遠端：IP 位址或主機名稱）
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

## 貢獻

歡迎提交 Issue 或 Pull Request 來改進此專案。