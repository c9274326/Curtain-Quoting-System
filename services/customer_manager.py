# services/customer_manager.py

import json
import os
import uuid

class CustomerManager:
    """管理客戶資料，支援客戶與案子的 CRUD 操作"""
    
    def __init__(self, data_path="data/customers.json"):
        self.data_path = data_path
        os.makedirs(os.path.dirname(self.data_path), exist_ok=True)
        if not os.path.exists(self.data_path):
            with open(self.data_path, 'w', encoding='utf-8') as f:
                json.dump([], f, ensure_ascii=False, indent=2)
        self.customers = self.load()
    
    def load(self):
        """從 JSON 檔案載入客戶資料"""
        try:
            with open(self.data_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                return data if isinstance(data, list) else []
        except:
            return []
    
    def save(self):
        """將客戶資料保存到 JSON 檔案"""
        with open(self.data_path, 'w', encoding='utf-8') as f:
            json.dump(self.customers, f, ensure_ascii=False, indent=2)
    
    def add(self, name, phone, address, template_path):
        """新增客戶"""
        cust = {
            "id": str(uuid.uuid4()),
            "name": name,
            "phone": phone,
            "address": address,
            "template_path": template_path,
            "projects": []  # 初始化空案子列表
        }
        self.customers.append(cust)
        self.save()
        return cust
    
    def update(self, cust_id, **kwargs):
        """更新客戶資訊"""
        for cust in self.customers:
            if cust["id"] == cust_id:
                cust.update(kwargs)
                self.save()
                return cust
        return None
    
    def delete(self, cust_id):
        """刪除客戶"""
        self.customers = [c for c in self.customers if c["id"] != cust_id]
        self.save()
    
    def get_all(self):
        """取得所有客戶"""
        return self.customers
    
    def get_by_id(self, cust_id):
        """根據 ID 取得客戶"""
        for cust in self.customers:
            if cust["id"] == cust_id:
                return cust
        return None
    
    def add_project(self, cust_id, project_name):
        """為客戶新增案子"""
        cust = self.get_by_id(cust_id)
        if not cust:
            return None
        
        project = {
            "id": str(uuid.uuid4()),
            "name": project_name
        }
        
        # 確保 projects 列表存在
        if "projects" not in cust:
            cust["projects"] = []
        
        cust["projects"].append(project)
        self.save()
        return project
    
    def update_project(self, cust_id, project_id, new_name):
        """更新客戶案子名稱"""
        cust = self.get_by_id(cust_id)
        if not cust:
            return None
        
        for project in cust["projects"]:
            if project["id"] == project_id:
                project["name"] = new_name
                self.save()
                return project
        return None
    
    def delete_project(self, cust_id, project_id):
        """刪除客戶的案子"""
        cust = self.get_by_id(cust_id)
        if not cust:
            return False
        
        # 過濾掉要刪除的案子
        cust["projects"] = [p for p in cust["projects"] if p["id"] != project_id]
        self.save()
        return True
    
    def get_projects(self, cust_id):
        """取得客戶的所有案子"""
        cust = self.get_by_id(cust_id)
        if not cust:
            return []
        return cust.get("projects", [])
