import json, os
from dataclasses import is_dataclass, asdict
from services.pricing import ItemGroup, SewingItem, SubItem

class CustomEncoder(json.JSONEncoder):
    def default(self, o):
        if is_dataclass(o):
            return asdict(o)
        return super().default(o)

class HistoryManager:
    """管理每位客戶報價歷史，存為 data/history_<cust_id>.json"""
    def __init__(self, base_dir="data"):
        self.base_dir = base_dir
        self.data_path = os.path.join(base_dir, "history.json") # A dummy path, not actually used for history files
        os.makedirs(self.base_dir, exist_ok=True)

    def _path(self, cust_id):
        return os.path.join(self.base_dir, f"history_{cust_id}.json")

    def load_all_projects(self, cust_id):
        path = self._path(cust_id)
        if not os.path.exists(path):
            return {}
        with open(path, 'r', encoding='utf-8') as f:
            try:
                data = json.load(f)
                if isinstance(data, list):  # Legacy format
                    return {"未分類": data}
                return data
            except json.JSONDecodeError:
                return {}

    def load_project_items(self, cust_id, project_id):
        all_data = self.load_all_projects(cust_id)
        project_data = all_data.get(project_id, [])
        
        loaded_groups = []
        for group_data in project_data:
            sewing_data = group_data.get('sewing_item')
            if not sewing_data: continue
            
            sewing_item = SewingItem(**sewing_data)
            sub_items = [SubItem(**sub) for sub in group_data.get('sub_items', [])]
            
            loaded_groups.append(ItemGroup(
                item_number=group_data['item_number'],
                sewing_item=sewing_item,
                sub_items=sub_items
            ))
        return loaded_groups

    def save_project_items(self, cust_id, project_id, records):
        all_data = self.load_all_projects(cust_id)
        all_data[project_id] = records
        path = self._path(cust_id)
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(all_data, f, ensure_ascii=False, indent=2, cls=CustomEncoder)
