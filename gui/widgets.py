import tkinter as tk
from tkinter import ttk
import re

class NumberEntry(ttk.Entry):
    """數字輸入框"""
    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)
        
        # 註冊驗證函數
        vcmd = (self.register(self.validate_number), '%P')
        self.config(validate='key', validatecommand=vcmd)
        
    def validate_number(self, value):
        """驗證數字輸入"""
        if value == "":
            return True
        try:
            float(value)
            return True
        except ValueError:
            return False


class CurrencyEntry(ttk.Entry):
    """貨幣輸入框"""
    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)
        
        # 註冊驗證函數
        vcmd = (self.register(self.validate_currency), '%P')
        self.config(validate='key', validatecommand=vcmd)
        
    def validate_currency(self, value):
        """驗證貨幣輸入"""
        if value == "":
            return True
        # 允許數字和小數點
        pattern = r'^[0-9]*\.?[0-9]*$'
        return bool(re.match(pattern, value))


class AutoCompleteCombobox(ttk.Combobox):
    """自動完成下拉選單"""
    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)
        
        self.bind('<KeyRelease>', self.on_keyrelease)
        
    def on_keyrelease(self, event):
        """按鍵釋放事件"""
        # 獲取輸入的文字
        typed = self.get()
        
        if typed == '':
            self['values'] = self.original_values
        else:
            # 過濾符合的選項
            filtered_values = []
            for item in self.original_values:
                if typed.lower() in item.lower():
                    filtered_values.append(item)
            self['values'] = filtered_values
            
    def set_values(self, values):
        """設定可選值"""
        self.original_values = values
        self['values'] = values
