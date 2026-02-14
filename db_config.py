# db_config.py - 資料庫設定檔
import pyodbc

# SQL Server 連線設定
DB_CONFIG = {
    'server': 'localhost',  # 改成您的 SQL Server 位址
    'database': 'YourDatabaseName',  # 改成您的資料庫名稱
    'username': 'your_username',  # 如果使用 SQL Server 驗證
    'password': 'your_password',  # 如果使用 SQL Server 驗證
    'trusted_connection': True  # 如果使用 Windows 驗證，設為 True
}

def get_db_connection():
    """建立並返回資料庫連線"""
    try:
        if DB_CONFIG['trusted_connection']:
            # Windows 驗證
            connection_string = (
                f"DRIVER={{ODBC Driver 17 for SQL Server}};"
                f"SERVER={DB_CONFIG['server']};"
                f"DATABASE={DB_CONFIG['database']};"
                f"Trusted_Connection=yes;"
            )
        else:
            # SQL Server 驗證
            connection_string = (
                f"DRIVER={{ODBC Driver 17 for SQL Server}};"
                f"SERVER={DB_CONFIG['server']};"
                f"DATABASE={DB_CONFIG['database']};"
                f"UID={DB_CONFIG['username']};"
                f"PWD={DB_CONFIG['password']};"
            )
        
        conn = pyodbc.connect(connection_string)
        return conn
    except Exception as e:
        print(f"資料庫連線失敗: {e}")
        return None

def save_cpe_to_database(cpe_data):
    """
    將 CPE 資料儲存到資料庫
    
    Args:
        cpe_data: 單筆 CPE 資料字典
    
    Returns:
        bool: 成功返回 True，失敗返回 False
    """
    conn = get_db_connection()
    if not conn:
        return False
    
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
        return True
    except Exception as e:
        print(f"儲存到資料庫失敗: {e}")
        if conn:
            conn.close()
        return False

def save_multiple_cpe_to_database(cpe_list):
    """
    批次將多筆 CPE 資料儲存到資料庫
    
    Args:
        cpe_list: CPE 資料列表
    
    Returns:
        dict: {'success': 成功筆數, 'failed': 失敗筆數}
    """
    conn = get_db_connection()
    if not conn:
        return {'success': 0, 'failed': len(cpe_list)}
    
    success_count = 0
    failed_count = 0
    
    try:
        cursor = conn.cursor()
        
        insert_query = """
        INSERT INTO cpe_records 
        (vendor, product_name, version, other_fields, size_mb, install_date, install_path)
        VALUES (?, ?, ?, ?, ?, ?, ?)
        """
        
        for cpe_data in cpe_list:
            try:
                cursor.execute(insert_query, (
                    cpe_data.get('vendor', ''),
                    cpe_data.get('product', ''),
                    cpe_data.get('version', ''),
                    cpe_data.get('other_fields', ''),
                    cpe_data.get('size_mb'),
                    cpe_data.get('install_date'),
                    cpe_data.get('install_path', 'C:\\')
                ))
                success_count += 1
            except Exception as e:
                print(f"儲存單筆資料失敗: {e}")
                failed_count += 1
        
        conn.commit()
        cursor.close()
        conn.close()
        
        return {'success': success_count, 'failed': failed_count}
    except Exception as e:
        print(f"批次儲存失敗: {e}")
        if conn:
            conn.close()
        return {'success': success_count, 'failed': len(cpe_list) - success_count}
