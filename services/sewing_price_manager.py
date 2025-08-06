
import json, os, uuid

class SewingPriceManager:
    """管理 data/sewing_prices.json 中的車工單價清單"""

    def __init__(self, data_path="data/sewing_prices.json"):
        self.data_path = data_path
        os.makedirs(os.path.dirname(self.data_path), exist_ok=True)
        if not os.path.exists(self.data_path):
            with open(self.data_path, 'w', encoding='utf-8') as f:
                json.dump([], f, ensure_ascii=False, indent=2)
        self.records = self._load()

    def _load(self):
        with open(self.data_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        return data if isinstance(data, list) else []

    def _save(self):
        with open(self.data_path, 'w', encoding='utf-8') as f:
            json.dump(self.records, f, ensure_ascii=False, indent=2)

    def get_all(self):
        return self.records

    def add(self, fabric, method, unit_price):
        rec = {
            "id": str(uuid.uuid4()),
            "fabric": fabric,
            "method": method,
            "unit_price": float(unit_price)
        }
        self.records.append(rec)
        self._save()
        return rec

    def update(self, rec_id, **kwargs):
        for r in self.records:
            if r["id"] == rec_id:
                r.update({
                    "fabric": kwargs.get("fabric", r["fabric"]),
                    "method": kwargs.get("method", r["method"]),
                    "unit_price": float(kwargs.get("unit_price", r["unit_price"]))
                })
                self._save()
                return r
        raise KeyError(f"找不到紀錄 ID: {rec_id}")

    def delete(self, rec_id):
        orig = len(self.records)
        self.records = [r for r in self.records if r["id"] != rec_id]
        if len(self.records) == orig:
            raise KeyError(f"找不到紀錄 ID: {rec_id}")
        self._save()

    def get_by_id(self, rec_id):
        """根據 ID 獲取紀錄"""
        for r in self.records:
            if r["id"] == rec_id:
                return r
        return None

    def get_fabrics(self):
        """獲取所有不重複的布料名稱"""
        return sorted(list(set(r['fabric'] for r in self.records)))

    def get_price(self, fabric, method):
        """根據布料和手法獲取單價"""
        for r in self.records:
            if r['fabric'] == fabric and r['method'] == method:
                return r['unit_price']
        return 0
