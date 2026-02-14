# db_config.py - 資料庫設定檔
import json
import os

try:
    import pyodbc
    PYODBC_AVAILABLE = True
except ImportError:
    PYODBC_AVAILABLE = False
    print("Warning: pyodbc is not installed. Database functionality will be disabled.")
    print("To enable database features, install pyodbc: pip install pyodbc")

# 預設 SQL Server 連線設定
# 注意：此應用程式僅支援本地資料庫連線（localhost）以確保安全性
# 在生產環境中，請使用環境變數或配置檔案來儲存敏感資訊
# 請勿將實際的資料庫憑證提交到版本控制系統
DEFAULT_DB_CONFIG = {
    'server': 'localhost',  # 僅支援 localhost 連線
    'database': 'YourDatabaseName',  # 改成您的資料庫名稱
    'username': 'your_username',  # 如果使用 SQL Server 驗證
    'password': 'your_password',  # 如果使用 SQL Server 驗證
    'trusted_connection': True  # 如果使用 Windows 驗證，設為 True
}

# 允許的本地主機名稱
ALLOWED_LOCALHOST_NAMES = [
    'localhost',
    '127.0.0.1',
    '::1',
    '(local)',
    '.',
    'localhost\\sqlexpress',
    '(localdb)\\mssqllocaldb'
]

# 用於儲存動態配置的檔案
DB_CONFIG_FILE = 'db_connections.json'

# 目前使用的資料庫配置
current_db_config = None

def load_db_connections():
    """從檔案載入已儲存的資料庫連線設定"""
    if os.path.exists(DB_CONFIG_FILE):
        try:
            with open(DB_CONFIG_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            print(f"載入資料庫連線設定失敗: {e}")
            return {}
    return {}

def save_db_connections(connections):
    """儲存資料庫連線設定到檔案"""
    try:
        with open(DB_CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(connections, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        print(f"儲存資料庫連線設定失敗: {e}")
        return False

def set_current_db_config(config):
    """設定目前使用的資料庫配置"""
    global current_db_config
    current_db_config = config

def get_current_db_config():
    """取得目前使用的資料庫配置"""
    global current_db_config
    return current_db_config if current_db_config else DEFAULT_DB_CONFIG

def is_localhost(server):
    """
    檢查伺服器位址是否為本地主機
    
    Args:
        server: 伺服器位址字串
    
    Returns:
        bool: 如果是本地主機返回 True，否則返回 False
    """
    if not server:
        return False
    
    # 將伺服器名稱轉換為小寫進行比較
    server_lower = server.lower().strip()
    
    # 檢查是否在允許的本地主機名稱列表中
    for allowed in ALLOWED_LOCALHOST_NAMES:
        if server_lower == allowed.lower() or server_lower.startswith(allowed.lower() + '\\'):
            return True
    
    return False

def get_error_suggestion(error_message):
    """
    根據錯誤訊息提供修改建議
    
    Args:
        error_message: 錯誤訊息字串
    
    Returns:
        str: 建議訊息
    """
    error_lower = str(error_message).lower()
    
    suggestions = []
    
    # ODBC Driver 相關錯誤
    if 'driver' in error_lower or 'odbc' in error_lower:
        suggestions.append("🔧 ODBC 驅動程式問題:")
        suggestions.append("   • 請確認已安裝 ODBC Driver 17 for SQL Server 或更高版本")
        suggestions.append("   • 下載連結: https://docs.microsoft.com/sql/connect/odbc/download-odbc-driver-for-sql-server")
        suggestions.append("   • 安裝後請重新啟動應用程式")
    
    # 登入失敗相關錯誤
    if 'login' in error_lower or 'authentication' in error_lower or '18456' in error_lower:
        suggestions.append("🔐 驗證失敗:")
        suggestions.append("   • 請確認使用者名稱和密碼是否正確")
        suggestions.append("   • 如果使用 Windows 驗證，請確認已勾選「使用 Windows 驗證」")
        suggestions.append("   • 如果使用 SQL Server 驗證，請確認：")
        suggestions.append("     - SQL Server 已啟用混合模式驗證")
        suggestions.append("     - 使用者帳號已在 SQL Server 中建立")
        suggestions.append("     - 使用者帳號擁有存取該資料庫的權限")
    
    # 資料庫不存在相關錯誤
    if 'database' in error_lower and ('not' in error_lower or 'cannot' in error_lower or '4060' in error_lower):
        suggestions.append("📁 資料庫不存在:")
        suggestions.append("   • 請確認資料庫名稱是否正確（區分大小寫）")
        suggestions.append("   • 請確認資料庫已在 SQL Server 中建立")
        suggestions.append("   • 使用 SQL Server Management Studio 或指令確認資料庫是否存在：")
        suggestions.append("     SELECT name FROM sys.databases;")
    
    # 連線逾時相關錯誤
    if 'timeout' in error_lower or 'timed out' in error_lower:
        suggestions.append("⏱️ 連線逾時:")
        suggestions.append("   • 請確認 SQL Server 服務是否已啟動")
        suggestions.append("   • 檢查方法：")
        suggestions.append("     - 開啟「服務」(services.msc)")
        suggestions.append("     - 尋找 SQL Server 相關服務")
        suggestions.append("     - 確認狀態為「執行中」")
    
    # 無法連線到伺服器相關錯誤
    if 'server' in error_lower and ('not' in error_lower or 'cannot' in error_lower or 'unable' in error_lower):
        suggestions.append("🖥️ 無法連線到伺服器:")
        suggestions.append("   • 請確認 SQL Server 服務是否已啟動")
        suggestions.append("   • 請確認 SQL Server Browser 服務是否已啟動（如果使用具名執行個體）")
        suggestions.append("   • 請確認 SQL Server 已設定為接受 TCP/IP 連線：")
        suggestions.append("     - 開啟 SQL Server Configuration Manager")
        suggestions.append("     - 在「SQL Server 網路組態」中啟用 TCP/IP 通訊協定")
        suggestions.append("   • 如果使用具名執行個體（如 localhost\\SQLEXPRESS），請確認：")
        suggestions.append("     - 執行個體名稱是否正確")
        suggestions.append("     - SQL Server Browser 服務已啟動")
    
    # 權限相關錯誤
    if 'permission' in error_lower or 'denied' in error_lower or 'access' in error_lower:
        suggestions.append("🚫 權限不足:")
        suggestions.append("   • 請確認使用者帳號擁有存取該資料庫的權限")
        suggestions.append("   • 使用管理員帳號執行以下 SQL 指令來授予權限：")
        suggestions.append("     USE [YourDatabase];")
        suggestions.append("     CREATE USER [YourUser] FOR LOGIN [YourUser];")
        suggestions.append("     ALTER ROLE db_datareader ADD MEMBER [YourUser];")
        suggestions.append("     ALTER ROLE db_datawriter ADD MEMBER [YourUser];")
    
    # 如果沒有匹配到特定錯誤，提供一般建議
    if not suggestions:
        suggestions.append("💡 一般建議:")
        suggestions.append("   • 請確認 SQL Server 服務是否已啟動")
        suggestions.append("   • 請確認資料庫名稱、使用者名稱和密碼是否正確")
        suggestions.append("   • 請確認已安裝 ODBC Driver 17 for SQL Server")
        suggestions.append("   • 嘗試使用 SQL Server Management Studio 連線以確認設定")
    
    return '\n'.join(suggestions)

def get_db_connection(config=None):
    """
    建立並返回資料庫連線
    此函式僅支援連線到本地資料庫（localhost）
    
    Args:
        config: 資料庫配置字典，如果為 None 則使用當前配置
    
    Returns:
        tuple: (connection, error_message)
            - connection: pyodbc.Connection 或 None
            - error_message: 錯誤訊息字串（如果成功則為 None）
    """
    if not PYODBC_AVAILABLE:
        error_msg = "pyodbc 模組未安裝，無法連線到資料庫。"
        suggestion = get_error_suggestion(error_msg)
        return None, f"{error_msg}\n\n{suggestion}"
    
    # 使用提供的配置或當前配置
    db_config = config if config else get_current_db_config()
    
    # 驗證伺服器位址是否為本地主機
    server = db_config.get('server', '')
    if not is_localhost(server):
        error_msg = f"❌ 安全限制：此應用程式僅支援連線到本地資料庫\n"
        error_msg += f"您嘗試連線的伺服器: {server}\n\n"
        error_msg += "🔒 允許的本地伺服器位址:\n"
        for allowed in ALLOWED_LOCALHOST_NAMES:
            error_msg += f"   • {allowed}\n"
        error_msg += "\n💡 建議:\n"
        error_msg += "   • 請將伺服器位址改為 'localhost' 或 '127.0.0.1'\n"
        error_msg += "   • 如果您需要連線到遠端資料庫，請聯絡系統管理員\n"
        return None, error_msg
    
    try:
        if db_config.get('trusted_connection', False):
            # Windows 驗證
            connection_string = (
                f"DRIVER={{ODBC Driver 17 for SQL Server}};"
                f"SERVER={db_config['server']};"
                f"DATABASE={db_config['database']};"
                f"Trusted_Connection=yes;"
            )
        else:
            # SQL Server 驗證
            connection_string = (
                f"DRIVER={{ODBC Driver 17 for SQL Server}};"
                f"SERVER={db_config['server']};"
                f"DATABASE={db_config['database']};"
                f"UID={db_config['username']};"
                f"PWD={db_config['password']};"
            )
        
        conn = pyodbc.connect(connection_string, timeout=10)
        return conn, None
    except pyodbc.Error as e:
        error_msg = f"❌ 資料庫連線失敗\n\n"
        error_msg += f"錯誤訊息: {str(e)}\n\n"
        suggestion = get_error_suggestion(str(e))
        error_msg += f"{suggestion}"
        return None, error_msg
    except Exception as e:
        error_msg = f"❌ 發生未預期的錯誤\n\n"
        error_msg += f"錯誤訊息: {str(e)}\n\n"
        suggestion = get_error_suggestion(str(e))
        error_msg += f"{suggestion}"
        return None, error_msg

def test_db_connection(config):
    """
    測試資料庫連線是否成功
    
    Args:
        config: 資料庫配置字典
    
    Returns:
        tuple: (success: bool, message: str)
    """
    if not PYODBC_AVAILABLE:
        error_msg = "pyodbc 模組未安裝"
        suggestion = get_error_suggestion(error_msg)
        return False, f"{error_msg}\n\n{suggestion}"
    
    # 驗證伺服器位址是否為本地主機
    server = config.get('server', '')
    if not is_localhost(server):
        error_msg = f"❌ 安全限制：此應用程式僅支援連線到本地資料庫\n"
        error_msg += f"您嘗試連線的伺服器: {server}\n\n"
        error_msg += "🔒 允許的本地伺服器位址:\n"
        for allowed in ALLOWED_LOCALHOST_NAMES:
            error_msg += f"   • {allowed}\n"
        error_msg += "\n💡 建議:\n"
        error_msg += "   • 請將伺服器位址改為 'localhost' 或 '127.0.0.1'\n"
        error_msg += "   • 如果您使用 SQL Server Express，可以嘗試 'localhost\\SQLEXPRESS'\n"
        error_msg += "   • 如果您需要連線到遠端資料庫，請聯絡系統管理員\n"
        return False, error_msg
    
    try:
        conn, error_msg = get_db_connection(config)
        if conn:
            conn.close()
            return True, "✅ 連線成功！資料庫連線正常運作。"
        else:
            return False, error_msg
    except Exception as e:
        error_msg = f"❌ 連線測試失敗\n\n"
        error_msg += f"錯誤訊息: {str(e)}\n\n"
        suggestion = get_error_suggestion(str(e))
        error_msg += f"{suggestion}"
        return False, error_msg

def save_cpe_to_database(cpe_data):
    """
    將 CPE 資料儲存到資料庫
    
    Args:
        cpe_data: 單筆 CPE 資料字典
    
    Returns:
        tuple: (success: bool, message: str)
    """
    conn, error_msg = get_db_connection()
    if not conn:
        return False, error_msg if error_msg else "無法連線到資料庫"
    
    try:
        cursor = conn.cursor()
        
        insert_query = """
        INSERT INTO cpe_records 
        (vendor, product_name, version, other_fields, size_mb, install_date, install_path)
        VALUES (?, ?, ?, ?, ?, ?, ?)
        """
        
        cursor.execute(insert_query, (
            cpe_data.get('vendor', ''),
            cpe_data.get('product', ''),
            cpe_data.get('version', ''),
            cpe_data.get('other_fields', ''),
            cpe_data.get('size_mb'),
            cpe_data.get('install_date'),
            cpe_data.get('install_path', 'C:\\')
        ))
        
        conn.commit()
        cursor.close()
        conn.close()
        return True, "✅ 資料已成功儲存到資料庫"
    except pyodbc.Error as e:
        error_msg = f"❌ 儲存到資料庫失敗\n\n錯誤訊息: {str(e)}\n\n"
        
        # 檢查是否為資料表不存在的錯誤
        if 'invalid object name' in str(e).lower() or 'cpe_records' in str(e).lower():
            error_msg += "💡 建議:\n"
            error_msg += "   • 資料表 'cpe_records' 不存在\n"
            error_msg += "   • 請使用以下 SQL 指令建立資料表：\n\n"
            error_msg += "   CREATE TABLE cpe_records (\n"
            error_msg += "       id INT PRIMARY KEY IDENTITY(1,1),\n"
            error_msg += "       vendor NVARCHAR(255),\n"
            error_msg += "       product_name NVARCHAR(255),\n"
            error_msg += "       version NVARCHAR(100),\n"
            error_msg += "       other_fields NVARCHAR(MAX),\n"
            error_msg += "       size_mb DECIMAL(10,2),\n"
            error_msg += "       install_date DATE,\n"
            error_msg += "       install_path NVARCHAR(500),\n"
            error_msg += "       created_at DATETIME DEFAULT GETDATE()\n"
            error_msg += "   );\n"
        else:
            error_msg += get_error_suggestion(str(e))
        
        if conn:
            conn.close()
        return False, error_msg
    except Exception as e:
        error_msg = f"❌ 發生未預期的錯誤\n\n錯誤訊息: {str(e)}\n\n"
        error_msg += get_error_suggestion(str(e))
        if conn:
            conn.close()
        return False, error_msg

def save_multiple_cpe_to_database(cpe_list):
    """
    批次將多筆 CPE 資料儲存到資料庫
    
    Args:
        cpe_list: CPE 資料列表
    
    Returns:
        dict: {'success': 成功筆數, 'failed': 失敗筆數, 'message': 訊息}
    """
    conn, error_msg = get_db_connection()
    if not conn:
        return {
            'success': 0, 
            'failed': len(cpe_list),
            'message': error_msg if error_msg else "無法連線到資料庫"
        }
    
    success_count = 0
    failed_count = 0
    
    try:
        cursor = conn.cursor()
        
        insert_query = """
        INSERT INTO cpe_records 
        (vendor, product_name, version, other_fields, size_mb, install_date, install_path)
        VALUES (?, ?, ?, ?, ?, ?, ?)
        """
        
        # Prepare data for batch insertion
        data_to_insert = [
            (
                cpe_data.get('vendor', ''),
                cpe_data.get('product', ''),
                cpe_data.get('version', ''),
                cpe_data.get('other_fields', ''),
                cpe_data.get('size_mb'),
                cpe_data.get('install_date'),
                cpe_data.get('install_path', 'C:\\')
            )
            for cpe_data in cpe_list
        ]
        
        # Use executemany for better performance
        cursor.executemany(insert_query, data_to_insert)
        success_count = len(cpe_list)
        
        conn.commit()
        cursor.close()
        conn.close()
        
        return {
            'success': success_count, 
            'failed': failed_count,
            'message': f"✅ 成功儲存 {success_count} 筆資料到資料庫"
        }
    except pyodbc.Error as e:
        error_msg = f"❌ 批次儲存失敗\n\n錯誤訊息: {str(e)}\n\n"
        
        # 檢查是否為資料表不存在的錯誤
        if 'invalid object name' in str(e).lower() or 'cpe_records' in str(e).lower():
            error_msg += "💡 建議:\n"
            error_msg += "   • 資料表 'cpe_records' 不存在\n"
            error_msg += "   • 請使用以下 SQL 指令建立資料表：\n\n"
            error_msg += "   CREATE TABLE cpe_records (\n"
            error_msg += "       id INT PRIMARY KEY IDENTITY(1,1),\n"
            error_msg += "       vendor NVARCHAR(255),\n"
            error_msg += "       product_name NVARCHAR(255),\n"
            error_msg += "       version NVARCHAR(100),\n"
            error_msg += "       other_fields NVARCHAR(MAX),\n"
            error_msg += "       size_mb DECIMAL(10,2),\n"
            error_msg += "       install_date DATE,\n"
            error_msg += "       install_path NVARCHAR(500),\n"
            error_msg += "       created_at DATETIME DEFAULT GETDATE()\n"
            error_msg += "   );\n"
        else:
            error_msg += get_error_suggestion(str(e))
        
        if conn:
            conn.close()
        return {
            'success': 0, 
            'failed': len(cpe_list),
            'message': error_msg
        }
    except Exception as e:
        error_msg = f"❌ 發生未預期的錯誤\n\n錯誤訊息: {str(e)}\n\n"
        error_msg += get_error_suggestion(str(e))
        if conn:
            conn.close()
        return {
            'success': 0, 
            'failed': len(cpe_list),
            'message': error_msg
        }
