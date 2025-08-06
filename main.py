import tkinter as tk
from gui.view_main import CurtainPricingApp
import json, os

def load_config():
    config_path = "data/config.json"
    if os.path.exists(config_path):
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            pass
    # 建立預設設定
    default = {
        "company_name": "窗簾專家",
        "company_phone": "02-1234-5678",
        "company_address": "台北市中正區xx路xx號",
        "price_table_path": "data/price_table.xlsx",
        "tax_rate": 0.05
    }
    os.makedirs("data", exist_ok=True)
    with open(config_path, 'w', encoding='utf-8') as f:
        json.dump(default, f, ensure_ascii=False, indent=2)
    return default

def main():
    config = load_config()
    root = tk.Tk()
    # root.title is now set in CurtainPricingApp
    try:
        root.iconbitmap("assets/logo.ico")
    except:
        pass
    app = CurtainPricingApp(root, config)
    root.mainloop()

if __name__ == "__main__":
    main()
