import json, os
from dataclasses import is_dataclass, asdict
from services.pricing import ItemGroup, SewingItem, SubItem

class CustomEncoder(json.JSONEncoder):
    def default(self, o):
        if is_dataclass(o):
            return asdict(o)
        return super().default(o)

class HistoryManager:
    """管理每個案子的報價歷史，存為 data/history_<project_id>.json"""
    def __init__(self, base_dir="data"):
        self.base_dir = base_dir
        os.makedirs(self.base_dir, exist_ok=True)

    def _get_path_for_project(self, project_id):
        return os.path.join(self.base_dir, f"history_{project_id}.json")

    def load_project_items(self, project_id):
        path = self._get_path_for_project(project_id)
        if not os.path.exists(path):
            return []
        with open(path, 'r', encoding='utf-8') as f:
            try:
                project_data = json.load(f)
            except json.JSONDecodeError:
                return []
        
        loaded_groups = []
        for group_data in project_data:
            sewing_data = group_data.get('sewing_item')
            if not sewing_data: continue
            
            sewing_item = SewingItem(**sewing_data)
            
            sub_items = []
            for sub_data in group_data.get('sub_items', []):
                # Handle old data that might not have an id
                if 'id' not in sub_data:
                    sub_data['id'] = f"sub-{uuid.uuid4().hex[:6]}"
                sub_items.append(SubItem(**sub_data))

            loaded_groups.append(ItemGroup(
                item_number=group_data['item_number'],
                sewing_item=sewing_item,
                sub_items=sub_items
            ))
        return loaded_groups

    def save_project_items(self, project_id, records):
        path = self._get_path_for_project(project_id)
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(records, f, ensure_ascii=False, indent=2, cls=CustomEncoder)

    def delete_project_file(self, project_id):
        """根據案子 ID 刪除對應的歷史紀錄檔案"""
        path = self._get_path_for_project(project_id)
        if os.path.exists(path):
            try:
                os.remove(path)
            except OSError as e:
                print(f"Error deleting file {path}: {e}")
