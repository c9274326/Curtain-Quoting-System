from services.excel_io import ExcelManager
import json

def create_default_files():
    """建立預設檔案"""
    # 載入設定
    with open('data/config.json', 'r', encoding='utf-8') as f:
        config = json.load(f)
    
    # 建立Excel管理器
    excel_manager = ExcelManager(config)
    
    # 建立價格表範本
    excel_manager.create_price_template('data/price_table.xlsx')
    print("已建立預設價格表: data/price_table.xlsx")

if __name__ == "__main__":
    import os
    
    # 建立資料夾
    os.makedirs("data", exist_ok=True)
    os.makedirs("assets", exist_ok=True)
    
    create_default_files()
