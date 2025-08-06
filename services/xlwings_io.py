import xlwings as xw

class XlwingsManager:
    def __init__(self, config):
        self.config = config

    def open_quote_with_template(self, template_path: str, quote_data: dict):
        app = xw.App(visible=True)
        wb = app.books.open(template_path)
        ws = wb.sheets[0]
        # 本公司資訊
        ws.range('B2').value = quote_data['company']['company_name']
        ws.range('B3').value = quote_data['company']['company_phone']
        ws.range('B4').value = quote_data['company']['company_address']
        # 客戶資訊
        ws.range('D2').value = quote_data['customer']['name']
        ws.range('D3').value = quote_data['customer']['phone']
        ws.range('D4').value = quote_data['customer']['address']
        # 明細自動展開
        data = [[i['item'], i['quantity'], i['unit'], i['unit_price'], i['subtotal']]
                for i in quote_data['items']]
        ws.range('A7').options(expand='table').value = data
        ws.autofit()
        return wb
