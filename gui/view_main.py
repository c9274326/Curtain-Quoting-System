import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
from services.pricing import PricingEngine, ItemGroup, SubItem
from services.excel_io import ExcelManager
from services.xlwings_io import XlwingsManager
from services.customer_manager import CustomerManager
from services.sewing_price_manager import SewingPriceManager
from services.history_manager import HistoryManager
import uuid

class CurtainPricingApp:
    def __init__(self, root, config):
        self.root = root
        self.config = config
        self.customer_mgr = CustomerManager("data/customers.json")
        self.sewing_price_mgr = SewingPriceManager("data/sewing_prices.json")
        self.history_mgr = HistoryManager("data")
        self.pricing_engine = PricingEngine(config, self.sewing_price_mgr)
        self.excel_manager = ExcelManager(config)
        self.xlwings_manager = XlwingsManager(config)
        self.quote_items = []
        self.current_customer = None
        self.setup_gui()

    def setup_gui(self):
        self.root.title("窗簾計價系統 v2.5")
        self.root.geometry("1200x800")
        self.root.minsize(1000, 700)
        style = ttk.Style()
        style.configure('Title.TLabel', font=('Arial', 14, 'bold'))
        style.configure('Treeview.Heading', font=('Arial', 10, 'bold'))
        style.configure("Bold.Treeview", font=('Arial', 9, 'bold'))

        frame = ttk.Frame(self.root, padding=10)
        frame.grid(row=0, column=0, sticky="nsew")
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        frame.columnconfigure(1, weight=1)

        self.create_header(frame)
        self.create_customer_section(frame)
        self.create_sewing_price_manager_button(frame)
        self.create_input_section(frame)
        self.create_quote_section(frame)
        self.create_total_section(frame)
        self.create_action_buttons(frame)

        self.set_input_controls_state('disabled')
        self.set_controls_state([self.export_btn, self.open_btn, self.add_btn], 'disabled')

    def set_controls_state(self, widgets, state):
        for widget in widgets:
            widget.state([state])

    def set_input_controls_state(self, state):
        for widget in self.input_frame.winfo_children():
            try:
                if widget.winfo_children():
                    for sub_widget in widget.winfo_children():
                        sub_widget.state([state])
                widget.state([state])
            except tk.TclError:
                pass

    def create_header(self, parent):
        hf = ttk.LabelFrame(parent, text="本公司資訊", padding=10)
        hf.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 10))
        ttk.Label(hf, text=self.config.get('company_name', ''), style='Title.TLabel').grid(row=0, column=0, sticky="w")
        info = f"電話: {self.config.get('company_phone', '')} | 地址: {self.config.get('company_address', '')}"
        ttk.Label(hf, text=info).grid(row=1, column=0, sticky="w")

    def create_customer_section(self, parent):
        cf = ttk.LabelFrame(parent, text="客戶與案子管理", padding=10)
        cf.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0, 10))
        cf.columnconfigure(1, weight=1)

        cust_frame = ttk.Frame(cf)
        cust_frame.grid(row=0, column=0, sticky="w")
        ttk.Label(cust_frame, text="選擇客戶:").grid(row=0, column=0, sticky="w", padx=(0, 5))
        self.customer_var = tk.StringVar()
        self.customer_combo = ttk.Combobox(cust_frame, textvariable=self.customer_var, values=[c['name'] for c in self.customer_mgr.get_all()], state="readonly", width=25)
        self.customer_combo.grid(row=0, column=1, sticky="w")
        self.customer_combo.bind('<<ComboboxSelected>>', self.on_customer_selected)
        self.project_combo.bind('<<ComboboxSelected>>', self.on_project_selected)

        project_frame = ttk.Frame(cf)
        project_frame.grid(row=0, column=1, sticky="e")
        self.manage_project_btn = ttk.Button(project_frame, text="管理案子", command=self.open_project_manager, state="disabled")
        self.manage_project_btn.pack(side="left", padx=5)
        ttk.Label(project_frame, text="選擇案子:").pack(side="left", padx=(10, 5))
        self.project_var = tk.StringVar()
        self.project_combo = ttk.Combobox(project_frame, textvariable=self.project_var, state="readonly", width=20)
        self.project_combo.pack(side="left")

        self.customer_info = ttk.Label(cf, text="")
        self.customer_info.grid(row=1, column=0, columnspan=2, sticky="w", pady=(5, 0))

    def on_customer_selected(self, event=None):
        name = self.customer_var.get()
        customer = next((c for c in self.customer_mgr.get_all() if c['name'] == name), None)
        if not customer: return

        self.current_customer = customer
        self.customer_info.config(text=f"客戶: {customer['name']} | {customer['phone']} | {customer['address']}")
        self.manage_project_btn.state(['!disabled'])
        
        projects = self.customer_mgr.get_projects(customer["id"])
        project_names = [p["name"] for p in projects]
        self.project_combo.config(values=project_names)
        
        if project_names:
            self.project_combo.set(project_names[0])
        else:
            self.project_combo.set('')
        
        self.set_input_controls_state('!disabled')
        self.set_controls_state([self.export_btn, self.open_btn, self.add_btn], '!disabled')

        self.on_project_selected()

    def on_project_selected(self, event=None):
        if not self.current_customer: return
        project_id = self.project_var.get()
        if not project_id: 
            self.quote_items = []
        else:
            self.quote_items = self.history_mgr.load_project_items(self.current_customer['id'], project_id)
        
        self.refresh_quote_tree()
        self.update_totals()

    def open_project_manager(self):
        messagebox.showinfo("提示", "案子管理功能尚未實作。")

    def create_sewing_price_manager_button(self, parent):
        btn = ttk.Button(parent, text="車工單價管理", command=self.open_sewing_price_manager)
        btn.grid(row=2, column=0, columnspan=2, pady=(0, 10))

    def open_sewing_price_manager(self):
        win = tk.Toplevel(self.root)
        win.title("車工單價管理")
        win.geometry("500x400")

        cols = ("fabric", "method", "unit_price")
        tree = ttk.Treeview(win, columns=cols, show="headings")
        tree.heading("fabric", text="布料")
        tree.heading("method", text="手法")
        tree.heading("unit_price", text="單價")
        tree.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        
        for r in self.sewing_price_mgr.get_all():
            tree.insert("", "end", iid=r["id"], values=(r["fabric"], r["method"], r["unit_price"]))

        form_frame = ttk.Frame(win, padding=10)
        form_frame.grid(row=1, column=0, sticky='ew')
        
        vars = {"fabric": tk.StringVar(), "method": tk.StringVar(), "price": tk.StringVar()}
        ttk.Label(form_frame, text="布料:").grid(row=0, column=0, sticky='w')
        ttk.Entry(form_frame, textvariable=vars['fabric']).grid(row=0, column=1, sticky='ew')
        ttk.Label(form_frame, text="手法:").grid(row=1, column=0, sticky='w')
        ttk.Combobox(form_frame, textvariable=vars['method'], values=["一般簾", "蛇行簾"], state='readonly').grid(row=1, column=1, sticky='ew')
        ttk.Label(form_frame, text="單價:").grid(row=2, column=0, sticky='w')
        ttk.Entry(form_frame, textvariable=vars['price']).grid(row=2, column=1, sticky='ew')

        def on_select(e):
            sel_id = tree.selection()
            if sel_id:
                rec = self.sewing_price_mgr.get_by_id(sel_id[0])
                if rec:
                    vars['fabric'].set(rec['fabric'])
                    vars['method'].set(rec['method'])
                    vars['price'].set(rec['unit_price'])

        tree.bind("<<TreeviewSelect>>", on_select)

        btn_frame = ttk.Frame(win, padding=10)
        btn_frame.grid(row=2, column=0, sticky='ew')

        def action(func_name):
            try:
                if func_name == 'add':
                    rec = self.sewing_price_mgr.add(vars['fabric'].get(), vars['method'].get(), vars['price'].get())
                    tree.insert("", "end", iid=rec["id"], values=(rec["fabric"], rec["method"], rec["unit_price"]))
                elif func_name in ['update', 'delete']:
                    sel_id = tree.selection()
                    if not sel_id: return
                    sel_id = sel_id[0]
                    if func_name == 'update':
                        rec = self.sewing_price_mgr.update(sel_id, fabric=vars['fabric'].get(), method=vars['method'].get(), unit_price=vars['price'].get())
                        tree.item(sel_id, values=(rec["fabric"], rec["method"], rec["unit_price"]))
                    else:
                        self.sewing_price_mgr.delete(sel_id)
                        tree.delete(sel_id)
                self.refresh_sewing_prices()
            except Exception as e:
                messagebox.showerror("錯誤", str(e))

        ttk.Button(btn_frame, text="新增", command=lambda: action('add')).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="修改", command=lambda: action('update')).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="刪除", command=lambda: action('delete')).pack(side="left", padx=5)

        win.rowconfigure(0, weight=1); win.columnconfigure(0, weight=1)
        form_frame.columnconfigure(1, weight=1)

    def refresh_sewing_prices(self):
        fabrics = self.sewing_price_mgr.get_fabrics()
        self.fabric_combo.config(values=fabrics)

    def on_key_press(self, event):
        if event.char == '.':
            event.widget.insert(tk.INSERT, ".")
            return "break"

    def create_input_section(self, parent):
        self.input_frame = ttk.LabelFrame(parent, text="車工規格輸入", padding=10)
        self.input_frame.grid(row=3, column=0, sticky="new", padx=(0, 5), pady=(0, 10))
        self.input_frame.columnconfigure(1, weight=1)

        vcmd = (self.root.register(self.validate_float), '%P')
        self.input_vars = {
            "item_number": tk.StringVar(), "fabric": tk.StringVar(), "method": tk.StringVar(value="一般簾"),
            "width": tk.StringVar(), "height": tk.StringVar(), "pieces": tk.StringVar(value="1")
        }

        ttk.Label(self.input_frame, text="貨號:").grid(row=0, column=0, sticky=tk.W, pady=2)
        ttk.Entry(self.input_frame, textvariable=self.input_vars["item_number"]).grid(row=0, column=1, sticky="ew")
        
        ttk.Label(self.input_frame, text="布料:").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.fabric_combo = ttk.Combobox(self.input_frame, textvariable=self.input_vars["fabric"], values=self.sewing_price_mgr.get_fabrics(), state="readonly")
        self.fabric_combo.grid(row=1, column=1, sticky="ew")
        
        ttk.Label(self.input_frame, text="手法:").grid(row=2, column=0, sticky=tk.W, pady=2)
        method_frame = ttk.Frame(self.input_frame)
        method_frame.grid(row=2, column=1, sticky='ew')
        ttk.Radiobutton(method_frame, text="一般簾", variable=self.input_vars["method"], value="一般簾").pack(side='left', padx=5)
        ttk.Radiobutton(method_frame, text="蛇行簾", variable=self.input_vars["method"], value="蛇行簾").pack(side='left', padx=5)

        width_entry = ttk.Entry(self.input_frame, textvariable=self.input_vars["width"], validate='key', validatecommand=vcmd)
        width_entry.grid(row=3, column=1, sticky="ew")
        width_entry.bind("<Key>", self.on_key_press)
        ttk.Label(self.input_frame, text="寬度 (臺尺):").grid(row=3, column=0, sticky=tk.W, pady=2)

        height_entry = ttk.Entry(self.input_frame, textvariable=self.input_vars["height"], validate='key', validatecommand=vcmd)
        height_entry.grid(row=4, column=1, sticky="ew")
        height_entry.bind("<Key>", self.on_key_press)
        ttk.Label(self.input_frame, text="高度 (臺尺):").grid(row=4, column=0, sticky=tk.W, pady=2)

        pieces_entry = ttk.Entry(self.input_frame, textvariable=self.input_vars["pieces"], validate='key', validatecommand=vcmd)
        pieces_entry.grid(row=5, column=1, sticky="ew")
        pieces_entry.bind("<Key>", self.on_key_press)
        ttk.Label(self.input_frame, text="幅數:").grid(row=5, column=0, sticky=tk.W, pady=2)

        ttk.Label(self.input_frame, text="車工單價:").grid(row=6, column=0, sticky=tk.W, pady=2)
        self.sewing_price_var = tk.StringVar(value="NT$ 0")
        ttk.Label(self.input_frame, textvariable=self.sewing_price_var, font=('Arial', 10, 'bold')).grid(row=6, column=1, sticky="w")

        ttk.Label(self.input_frame, text="試算小計:").grid(row=7, column=0, sticky=tk.W, pady=2)
        self.trial_price_var = tk.StringVar(value="NT$ 0")
        ttk.Label(self.input_frame, textvariable=self.trial_price_var, font=('Arial', 10, 'bold')).grid(row=7, column=1, sticky="w")

        for var in self.input_vars.values():
            var.trace_add('write', self.trial_price)

        self.add_btn = ttk.Button(self.input_frame, text="新增貨號項目", command=self.add_item_group)
        self.add_btn.grid(row=8, column=0, columnspan=2, pady=10)

    def validate_float(self, value_if_allowed):
        if value_if_allowed == "":
            return True
        try:
            float(value_if_allowed)
            return True
        except ValueError:
            return False

    def trial_price(self, *args):
        try:
            fabric = self.input_vars["fabric"].get()
            method = self.input_vars["method"].get()
            pieces_str = self.input_vars["pieces"].get()
            if not all([fabric, method, pieces_str]):
                self.add_btn.state(['disabled'])
                return

            sewing_item = self.pricing_engine.create_sewing_item(
                fabric=fabric, method=method, width=0, height=0, pieces=float(pieces_str)
            )
            self.sewing_price_var.set(f"NT$ {sewing_item.unit_price:,.0f}")
            self.trial_price_var.set(f"NT$ {sewing_item.subtotal:,.0f}")
            self.add_btn.state(['!disabled'])
        except (ValueError, KeyError):
            self.sewing_price_var.set("N/A")
            self.trial_price_var.set("N/A")
            self.add_btn.state(['disabled'])

    def add_item_group(self):
        if not self.current_customer:
            messagebox.showerror("錯誤", "請先選擇一個客戶")
            return
        try:
            item_number = self.input_vars['item_number'].get()
            if not item_number:
                item_number = f"item-{uuid.uuid4().hex[:6]}"
            if any(item.item_number == item_number for item in self.quote_items):
                messagebox.showerror("錯誤", "貨號重複!")
                return

            sewing_item = self.pricing_engine.create_sewing_item(
                fabric=self.input_vars['fabric'].get(),
                method=self.input_vars['method'].get(),
                width=float(self.input_vars['width'].get() or 0),
                height=float(self.input_vars['height'].get() or 0),
                pieces=float(self.input_vars['pieces'].get() or 1)
            )
            
            new_item = ItemGroup(
                item_number=item_number,
                sewing_item=sewing_item
            )
            self.quote_items.append(new_item)
            self.history_mgr.save_project_items(self.current_customer['id'], self.project_var.get(), self.quote_items)
            self.refresh_quote_tree()
            self.update_totals()

        except Exception as e:
            messagebox.showerror("錯誤", f"新增項目失敗: {e}")

    def create_quote_section(self, parent):
        qf = ttk.LabelFrame(parent, text="報價明細", padding=10)
        qf.grid(row=3, column=1, rowspan=3, sticky="nsew", padx=(5, 0))
        qf.rowconfigure(0, weight=1)
        qf.columnconfigure(0, weight=1)

        cols = ('item_number', 'desc', 'spec', 'qty', 'price', 'subtotal')
        self.quote_tree = ttk.Treeview(qf, columns=cols, show="headings")
        self.quote_tree.heading('item_number', text='貨號')
        self.quote_tree.heading('desc', text='項目')
        self.quote_tree.heading('spec', text='規格')
        self.quote_tree.heading('qty', text='幅數')
        self.quote_tree.heading('price', text='單價')
        self.quote_tree.heading('subtotal', text='小計')
        self.quote_tree.column('item_number', width=80)
        self.quote_tree.column('desc', width=150)
        self.quote_tree.column('spec', width=120)
        self.quote_tree.column('qty', width=80)
        self.quote_tree.column('price', width=80)
        self.quote_tree.column('subtotal', width=80)
        self.quote_tree.grid(row=0, column=0, sticky="nsew")

        sb = ttk.Scrollbar(qf, orient="vertical", command=self.quote_tree.yview)
        self.quote_tree.configure(yscroll=sb.set)
        sb.grid(row=0, column=1, sticky="ns")

        btn_frame = ttk.Frame(qf)
        btn_frame.grid(row=1, column=0, columnspan=2, pady=5, sticky='ew')
        ttk.Button(btn_frame, text="新增子項目/備註", command=self.add_sub_item).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="刪除選中", command=self.delete_selected).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="清空明細", command=self.clear_quote).pack(side="left", padx=5)

    def refresh_quote_tree(self):
        for i in self.quote_tree.get_children():
            self.quote_tree.delete(i)

        for item in self.quote_items:
            sewing_item = item.sewing_item
            main_id = self.quote_tree.insert('', 'end', iid=item.item_number, values=(
                item.item_number,
                f"{sewing_item.fabric} ({sewing_item.method})",
                f"寬:{sewing_item.width} 高:{sewing_item.height}",
                sewing_item.pieces,
                f"NT$ {sewing_item.unit_price:,.0f}",
                f"NT$ {item.total:,.0f}"
            ))
            for sub_item in item.sub_items:
                sub_id = f"{item.item_number}-{uuid.uuid4().hex[:6]}"
                self.quote_tree.insert(main_id, 'end', iid=sub_id, values=(
                    "", f"  - {sub_item.description}", "", sub_item.quantity,
                    f"NT$ {sub_item.unit_price:,.0f}", f"NT$ {sub_item.subtotal:,.0f}"
                ))

    def add_sub_item(self):
        selected_ids = self.quote_tree.selection()
        if not selected_ids:
            messagebox.showwarning("注意", "請先選擇一個主項目")
            return
        
        parent_id = self.quote_tree.parent(selected_ids[0]) or selected_ids[0]
        main_item = next((item for item in self.quote_items if item.item_number == parent_id), None)
        if not main_item: return

        win = tk.Toplevel(self.root)
        win.title("新增子項目/備註")
        vars = {'desc': tk.StringVar(), 'qty': tk.StringVar(value='1'), 'price': tk.StringVar(value='0')}
        ttk.Label(win, text="說明:").pack(padx=10, pady=5)
        ttk.Entry(win, textvariable=vars['desc'], width=40).pack(padx=10)
        ttk.Label(win, text="數量:").pack(padx=10, pady=5)
        ttk.Entry(win, textvariable=vars['qty']).pack(padx=10)
        ttk.Label(win, text="單價:").pack(padx=10, pady=5)
        ttk.Entry(win, textvariable=vars['price']).pack(padx=10)
        
        def on_ok():
            try:
                qty = float(vars['qty'].get())
                price = float(vars['price'].get())
                sub_item = SubItem(
                    description=vars['desc'].get(),
                    quantity=qty,
                    unit_price=price,
                    subtotal=qty * price
                )
                main_item.sub_items.append(sub_item)
                self.history_mgr.save_project_items(self.current_customer['id'], self.project_var.get(), self.quote_items)
                self.refresh_quote_tree()
                self.update_totals()
                win.destroy()
            except ValueError:
                messagebox.showerror("錯誤", "數量和單價必須是數字")

        ttk.Button(win, text="確定", command=on_ok).pack(pady=10)
        win.grab_set()

    def delete_selected(self):
        selected_ids = self.quote_tree.selection()
        if not selected_ids: return
        if not messagebox.askyesno("確認", "確定要刪除選中的項目嗎？"): return

        for item_id in selected_ids:
            parent_id = self.quote_tree.parent(item_id)
            if parent_id: # It's a sub-item
                main_item = next((item for item in self.quote_items if item.item_number == parent_id), None)
                if main_item:
                    desc_to_delete = self.quote_tree.item(item_id)['values'][0].strip().lstrip('- ')
                    main_item.sub_items = [si for si in main_item.sub_items if si.description != desc_to_delete]
            else: # It's a main item
                self.quote_items = [item for item in self.quote_items if item.item_number != item_id]
        
        self.history_mgr.save(self.current_customer['id'], self.quote_items)
        self.refresh_quote_tree()
        self.update_totals()

    def clear_quote(self):
        if messagebox.askyesno("確認", "確定要清空所有報價明細嗎？"):
            self.quote_items.clear()
            self.history_mgr.save_project_items(self.current_customer['id'], self.project_var.get(), self.quote_items)
            self.refresh_quote_tree()
            self.update_totals()

    def create_total_section(self, parent):
        tf = ttk.LabelFrame(parent, text="價格總計", padding=10)
        tf.grid(row=6, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        tf.columnconfigure(1, weight=1)
        ttk.Label(tf, text="總計:", font=('Arial', 12, 'bold')).grid(row=0, column=0, sticky="w")
        self.total_var = tk.StringVar(value="NT$ 0")
        ttk.Label(tf, textvariable=self.total_var, font=('Arial', 12, 'bold'), foreground='red').grid(row=0, column=1, sticky="e")

    def update_totals(self):
        total = sum(item.total for item in self.quote_items)
        self.total_var.set(f"NT$ {total:,.0f}")

    def create_action_buttons(self, parent):
        bf = ttk.Frame(parent)
        bf.grid(row=7, column=0, columnspan=2, pady=10)
        self.export_btn = ttk.Button(bf, text="匯出報價單", command=self.export_quote)
        self.export_btn.pack(side="left", padx=5)
        self.open_btn = ttk.Button(bf, text="開啟Excel", command=self.open_excel)
        self.open_btn.pack(side="left", padx=5)

    def export_quote(self):
        messagebox.showinfo("提示", "匯出功能待更新以支援子項目。")

    def open_excel(self):
        messagebox.showinfo("提示", "開啟Excel功能待更新以支援子項目。")

    def open_project_manager(self):
        messagebox.showinfo("提示", "案子管理功能尚未實作。")
