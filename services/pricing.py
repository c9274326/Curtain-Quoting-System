from dataclasses import dataclass, field
from typing import List, Any, Dict
from services.sewing_price_manager import SewingPriceManager

@dataclass
class SewingItem:
    """車工項目"""
    fabric: str
    method: str
    width: float
    height: float
    pieces: float
    unit_price: float
    subtotal: float

@dataclass
class SubItem:
    """附加項目或備註"""
    description: str
    quantity: float = 1
    unit_price: float = 0
    subtotal: float = 0

@dataclass
class ItemGroup:
    """代表一個貨號的群組"""
    item_number: str
    sewing_item: SewingItem
    sub_items: List[SubItem] = field(default_factory=list)

    @property
    def total(self) -> float:
        """計算整個貨號群組的總金額"""
        return self.sewing_item.subtotal + sum(item.subtotal for item in self.sub_items)

class PricingEngine:
    """計價引擎"""

    def __init__(self, config: Dict, sewing_price_manager: SewingPriceManager):
        self.config = config
        self.sewing_price_manager = sewing_price_manager

    def get_sewing_price(self, fabric: str, method: str) -> float:
        """獲取車工單價"""
        price = self.sewing_price_manager.get_price(fabric, method)
        if price is None:
            raise ValueError(f"找不到布料 '{fabric}' 搭配手法 '{method}' 的車工單價")
        return price

    def create_sewing_item(self, fabric: str, method: str, width: float, height: float, pieces: float) -> SewingItem:
        """建立車工項目並計算價格"""
        unit_price = self.get_sewing_price(fabric, method)
        subtotal = pieces * unit_price
        return SewingItem(
            fabric=fabric, method=method, width=width, height=height,
            pieces=pieces, unit_price=unit_price, subtotal=round(subtotal, 0)
        )
