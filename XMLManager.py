import tkinter as tk
from tkinter import ttk
import os
import xml.etree.ElementTree as ET
import codecs
from tkinter import messagebox
from xml.dom import minidom
import sys
import win32com.client as win32
import threading
import pythoncom
import html

class DatabaseBrowser(tk.Tk):
    def __init__(self):
        '''
        Data是汇总后表格；
        EmployeeInfo是人员表格；
        InstrumentInfo是仪表表格；
        LocationInfo是地点表格；
        TestItemInfo是测试项目表格；
        UncertaintyInfo是不确定度表格
        '''        
        super().__init__()
        self.title("XML Database Browser")
        self.geometry("1400x900")

        

        # 按钮/下拉菜单字体
        self.button_font_size = 14
        self.default_font = ("Helvetica", self.button_font_size)

        # 设置表头
        self.column_name_map = {
            "Summary Data": {
                "EmployeeID": "Employee Name",
                "ItemID": "Project Name",
                "SerialNo": "Instrument Info",
                "LocationID": "Location",
            },
            "Employee Data": {
                "EmployeeID": "Employee ID",
                "Name": "Employee Name",
            },
            "Instrument Data": {
                "Name": "Instrument Name",
                "Model": "Instrument Model",
                "SerialNo": "Serial No.",
                "Manufacturer": "Manufacturer",
                "CalDueDate": "Cal Due Date",
            },
            "Location Data": {
                "LocationID": "Location ID",
                "Address": "Address",
                "Type": "Address Type",
                "Function": "Address Function",
            },
            "Project Data": {
                "ItemID": "Project ID",
                "ItemName": "Project Name",
            },
            "Uncertainty Data":{
                "Name":"Bookmark",
                "Description":"Description",
                "Value":"Value",
                },
        }
        '''
        Summary Data
        Employee Data
        Instrument Data
        Location Data
        Project Data
        Uncertainty Data
        '''

        self.filename_to_table = {
            "Data.xml": "Summary Data",
            "EmployeeInfo.xml": "Employee Data",
            "InstrumentInfo.xml": "Instrument Data",
            "LocationInfo.xml": "Location Data",
            "TestItemInfo.xml": "Project Data",
            "UncertaintyInfo.xml": "Uncertainty Data"
        }

        self.load_config_file() # 加载文件

        # 设置不确定度表格中项目名称和书签名的映射
        self.project_filter_keywords = {
        }
        '''
        速度优化，在程序启动的时候自动遍历所有word文件，将文件路径和对应的书签列表存入字典
        这样在筛选后写入word时就不用遍历所有word文件，提升速度
        如果在程序开启后修改了任何word文件内容，需要在不确定度表格中按下重新扫描按钮来更新字典
        遍历word文件会在后台线程运行，不会影响写入以外的其他操作
        '''
        self.active_project_filters = set()  # 当前选中的项目（用于筛选）
        self.word_bookmark_map = {}  # 文件路径 -> 书签列表
        self.word_scan_complete = False
        self.word_scan_thread = threading.Thread(target=self.scan_word_bookmarks, daemon=True)
        self.word_scan_thread.start()


  
        # 顶部导航栏，选择表格、筛选汇总表格控件
        self.create_navbar()
        
        # 显示选中的表格
        self.current_table_label = tk.Label(self, text="", font=("Helvetica", 14))
        self.current_table_label.pack(pady=(0, 5)) 
        self.create_table()
        self.filtered_data = []
        
        # 添加、删除、编辑、保存按钮
        self.create_controls()
        self.all_selected = False
        self.modify_state = "Not Modified"
        # self._last_hovered = None
        self._original_tags = {} 
        self.switch_table("Summary Data")
        self.write_word_button.pack_forget()

        # 检查 data 文件夹是否存在，xml表格需要放在和程序同目录下的data文件夹
        """
        data_folder = os.path.join(os.path.dirname(sys.argv[0]), "data")
        if not os.path.exists(data_folder):
            messagebox.showerror("错误", "未找到 data 文件夹，请确保它与程序放在同一目录下。")
            sys.exit(1)
            """

    def create_button(self, parent, text, command, width=None):
        if width:
            return tk.Button(parent, text=text, command=command, font=("Helvetica", self.button_font_size), width=width)
        else:
            return tk.Button(parent, text=text, command=command, font=("Helvetica", self.button_font_size))

    # 汇总表格中文显示————————————————————————————————————————————————————————————————
    def get_display_column_name(self, table_name, column_key):
        """根据当前表格名和字段名获取显示名"""
        return self.column_name_map.get(table_name, {}).get(column_key, column_key)

    def get_employee_display(self, emp_id):
        """Employee Name"""
        info = self.engineer_info.get(emp_id, "")
        return f'{info} / {emp_id}' if info else emp_id

    def get_item_display(self, item_id):
        """Project Name"""
        info = self.item_info.get(item_id, "")
        return f'{info} / {item_id}' if info else item_id

    def get_serial_display(self, serial_no):
        """Instrument Info"""
        info = self.instrument_info.get(serial_no, "")
        return f'{info} / {serial_no}' if info else serial_no

    def get_location_display(self, loc_id):
        """Location"""
        info = self.location_info.get(loc_id, "")
        return f'{info} / {loc_id}' if info else loc_id
    # ——————————————————————————————————————————————————————————————————————————————————

    def create_navbar(self):
        """创建顶部按钮和下拉菜单"""
        navbar = tk.Frame(self, bg="#d9d9d9", height=50)
        navbar.pack(fill=tk.X, padx=5, pady=5)
        
        tables = ["Summary Data", "Employee Data", "Instrument Data", "Location Data", "Project Data", "Uncertainty Data"]
        for table in tables:
            self.create_button(navbar, table, lambda t=table: self.switch_table(t)).pack(side=tk.LEFT, padx=2)

        self.create_filter_controls(navbar)

        separator = ttk.Separator(self, orient='horizontal')
        separator.pack(fill='x', padx=5, pady=(0, 5))

        separator = ttk.Separator(self, orient='horizontal')
        separator.pack(fill='x', padx=5, pady=(0, 5))

    def create_table(self):
        """创建和初始化表格"""
        container = tk.Frame(self)
        container.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        style = ttk.Style(self)
        style.configure("Treeview", font=('Helvetica', 14))
        style.configure("Treeview.Heading", font=('Helvetica', 14, 'bold'))
        style.map("Treeview", background=[('selected', '#aec2dc')])  # 选中

        row_height = int(14 * 1.8)  # 行高
        style.configure("Treeview", rowheight=row_height)

        self.tree = ttk.Treeview(container, show="headings")
        self.tree.tag_configure('evenrow', background='#f5f5f5')
        self.tree.tag_configure('oddrow', background='#e0e0e0')
        # self.tree.tag_configure('hoverrow', background='#afc3dd') 

        vsb = ttk.Scrollbar(container, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(container, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        self._last_hovered = None 
        # self.tree.bind("<Leave>", self.on_leave_tree)
        self.tree.bind("<<TreeviewSelect>>", lambda e: self.update_status())
        self.tree.bind("<Button-1>", self.on_click_toggle_selection)


    def create_controls(self):
        """创建底部按钮和提示信息"""
        separator = ttk.Separator(self, orient='horizontal')
        separator.pack(fill='x', padx=5, pady=(0, 5))
        control_frame = tk.Frame(self, bg="#d9d9d9")
        control_frame.pack(fill=tk.X, padx=10, pady=10)

        # 底部按钮
        self.create_button(control_frame, "New Record", self.add_record).pack(side=tk.LEFT, padx=5)
        self.create_button(control_frame, "Delete", self.delete_record).pack(side=tk.LEFT, padx=5)
        self.create_button(control_frame, "Edit", self.edit_record).pack(side=tk.LEFT, padx=5)
        self.status_label = tk.Label(
            control_frame,
            text="Selected 0 Records   Not Modified",
            anchor="w",
            font=("Helvetica", 14),
            fg="gray",
            bg="#d9d9d9"
        )
        self.status_label.pack(side=tk.LEFT, padx=10)
        self.select_all_button.pack(side=tk.RIGHT, padx=5)
        # self.create_button(control_frame,"写入Word",self.write_to_word).pack(side=tk.RIGHT, padx=5)
        # word扫描进度条
        # word扫描进度条（挂在 control_frame，而不是不存在的 self.status_frame）
        self.scan_progress = ttk.Progressbar(control_frame, mode='indeterminate', length=180)
        self.scan_progress.pack(side=tk.RIGHT, padx=10)
        self.scan_progress.stop()
        self.scan_progress.pack_forget()

        self.write_word_button = self.create_button(control_frame, "Write to Word", self.write_to_word)
        self.write_word_button.pack(side=tk.RIGHT, padx=5) 
        self.create_button(control_frame, "Save", self.save_data).pack(side=tk.RIGHT, padx=5)
        self.font_size = 14 
        self.create_button(control_frame, "Zoom+", self.increase_font_size).pack(side=tk.RIGHT, padx=5)
        self.create_button(control_frame, "Zoom-", self.decrease_font_size).pack(side=tk.RIGHT, padx=5)
        # 重新扫描按钮
        self.rescan_button = self.create_button(control_frame, "Rescan Word", self.trigger_rescan)

    def switch_table(self, table_name):
        """
        根据选中的表格切换显示到对应的表格
        根据对应的切换后的xml刷新显示的内容
        在汇总表格显示筛选控件
        在不确定度表格显示“写入Word按钮”
        其余表格隐藏
        """
        print(f"正在切换到表格：{table_name}")
        self.current_table_label.config(text=f"当前表格：{table_name}")

        columns, data = self.load_table_data(table_name)
        self.columns = columns
        self.all_data = data.copy()

        self.tree["columns"] = columns

        current_table_map = self.column_name_map.get(table_name, {})
        for col in columns:
            display_name = current_table_map.get(col, col)
            self.tree.heading(col, text=display_name)
            self.tree.column(col, anchor="center")

        self.tree.delete(*self.tree.get_children())
        self._original_tags.clear()

        if table_name == "Uncertainty Data":
            self.write_word_button.pack(side=tk.RIGHT, padx=5)
            self.project_filter_button.pack(side=tk.RIGHT, padx=4)
            self.rescan_button.pack(side=tk.RIGHT, padx=10)
        else:
            self.write_word_button.pack_forget()
            self.project_filter_button.pack_forget()
            self.rescan_button.pack_forget()

        if table_name == "Summary Data":
            self.engineer_info = {e.attrib["EmployeeID"]: e.attrib.get("Name", "") for e in self.load_xml_root("EngineerInfo.xml").findall(".//Engineer")}
            self.item_info = {i.attrib["ItemID"]: i.attrib.get("ItemName", "") for i in self.load_xml_root("TestItemInfo.xml").findall(".//TestItem")}
            self.instrument_info = {ins.attrib["SerialNo"]: f'{ins.attrib.get("Name", "")} / {ins.attrib.get("Model", "")}' for ins in self.load_xml_root("InstrumentInfo.xml").findall(".//Instrument")}
            self.location_info = {loc.attrib["LocationID"]: loc.attrib.get("Address", "") for loc in self.load_xml_root("LocationInfo.xml").findall(".//Location")}
            self.update_summary_filters()
        self.show_filtered_data(self.all_data)

        if table_name == "Summary Data":
            employee_ids = self.extract_attributes_from_xml("Data.xml", "Summary", "EmployeeID")
            item_ids = self.extract_attributes_from_xml("Data.xml", "Summary", "ItemID")

            display_employee_list = [self.get_employee_display(eid) for eid in employee_ids]
            display_item_list = [self.get_item_display(iid) for iid in item_ids]

            self.filter_employee["values"] = sorted(display_employee_list)
            self.filter_item["values"] = sorted(display_item_list)

            self.filter_employee.set("")
            self.filter_item.set("")
            self.filter_employee.configure(state="readonly")
            self.filter_item.configure(state="readonly")
            self.filter_frame.pack(side=tk.RIGHT, padx=10)
        else:
            self.filter_employee.set("")
            self.filter_item.set("")
            self.filter_employee.configure(state="disabled")
            self.filter_item.configure(state="disabled")
            self.filter_frame.pack_forget()

        if table_name == "Uncertainty Data":
            self.refresh_uncertainty_map()


        self.adjust_column_widths()


    def load_table_data(self, table_name):
        """从xml读取数据，返回表格"""
        # 直接通过表格名称从 data_path_map 获取路径
        path = self.data_path_map.get(table_name)
    
        # 路径检查（合并重复检查）
        if not path:
            return ["Error"], [("The path is not declared in Config.xml", "")]
        if not os.path.exists(path):
            return ["Error"], [("XML file does not exist: " + path, "")]
    
        try:
            # 读取并修复 XML 内容
            with codecs.open(path, "r", encoding="gb2312", errors="replace") as f:
                xml_content = f.read()
            xml_content = xml_content.replace("&", "&amp;")  # 处理xml非法字符
            root = ET.fromstring(xml_content)
        
            # 解析 XML
            # root = ET.fromstring(xml_content)
        
            # 处理汇总表格的特殊结构
            if table_name == "Summary Data":
                summary_info = root.find("Summary_Info")
                if not summary_info:
                    return ["Error"], [("XML structure error：Summary_Info is needed", "")]
                
                entries = summary_info.findall("Summary")
                if not entries:
                    return ["Error"], [("The data does not exist in summary data", "")]
                
                # 提取列名和数据
                columns = list(entries[0].attrib.keys())
                data = [
                    [entry.attrib.get(col, "") for col in columns]
                    for entry in entries
                ]
            
            # 处理其他子表
            else:
                # 自动识别 XML 结构（属性或子节点）
                parent = list(root.iter())[1] if len(list(root.iter())) > 1 else root
                entries = list(parent)
            
                if not entries:
                    return ["Error"], [("The data does not exist", "")]
                
                first = entries[0]
                if first.attrib:  # 属性模式（如人员表格）
                    columns = list(first.attrib.keys())
                    data = [
                        [entry.attrib.get(col, "") for col in columns]
                        for entry in entries
                    ]
                else:  # 子节点模式（如不确定度表格）
                    columns = [elem.tag for elem in first]
                    data = [
                        [entry.find(tag).text if entry.find(tag) else "" for tag in columns]
                        for entry in entries
                    ]
        
            print(f"[DEBUG] {table_name} Column：", columns)
            print(f"[DEBUG] {table_name} First row data：", data[0] if data else "empty")
            return columns, data
        
        except ET.ParseError as e:
            print(f"XML analysis error（{path}）:", e)
            return ["Error"], [("XML analysis error，please check", "")]
        except Exception as e:
            print(f"Unknown error（{path}）:", e)
            return ["Error"], [(f"Can't read file：{str(e)}", "")]

        # ——————————————————————————————————————————————————————

    def add_record(self):
        """
        添加记录
        汇总表格使用下拉菜单添加，拥有批量添加功能
        其余表格使用输入框输入内容添加
        """
        columns = self.tree["columns"]

        if self.current_table_label.cget("text") == "Current Table: Summary Data":
            # 中文显示下拉菜单
            employee_ids = self.extract_attributes_from_xml("EngineerInfo.xml", "Engineer", "EmployeeID")
            item_ids = self.extract_attributes_from_xml("TestItemInfo.xml", "TestItem", "ItemID")
            serial_nos = self.extract_attributes_from_xml("InstrumentInfo.xml", "Instrument", "SerialNo")
            location_ids = self.extract_attributes_from_xml("LocationInfo.xml", "Location", "LocationID")

            display_to_employee_id = {self.get_employee_display(eid): eid for eid in employee_ids}
            display_to_item_id = {self.get_item_display(iid): iid for iid in item_ids}
            display_to_serial_no = {self.get_serial_display(sn): sn for sn in serial_nos}
            display_to_location_id = {self.get_location_display(lid): lid for lid in location_ids}

            options_map = {
                "EmployeeID": list(display_to_employee_id.keys()),
                "ItemID": list(display_to_item_id.keys()),
                "SerialNo": list(display_to_serial_no.keys()),
                "LocationID": list(display_to_location_id.keys())
            }

            top = tk.Toplevel(self)
            top.title("Add in bulk (Summary data)")
            self.update_idletasks()
            win_w = max(200 * len(columns) + 430, 900)
            win_h = 450
            x = self.winfo_x() + (self.winfo_width() - win_w) // 2
            y = self.winfo_y() + (self.winfo_height() - win_h) // 2
            top.geometry(f"{win_w}x{win_h}+{x}+{y}")

            current_table = self.current_table_label.cget("text").replace("Current Table:", "")
            for j, col in enumerate(columns):
                display_name = self.get_display_column_name(current_table, col)
                tk.Label(top, text=display_name, font=("Helvetica", 14, "bold")).grid(row=0, column=j + 1, padx=5, pady=5)

            entry_vars_list = []

            def render_rows():
                for widget in top.grid_slaves():
                    info = widget.grid_info()
                    row_num = int(info["row"])
                    if 1 <= row_num <= len(entry_vars_list):
                        widget.destroy()

                for row_idx, row_vars in enumerate(entry_vars_list):
                    for col_idx, col in enumerate(columns):
                        var = row_vars[col]
                        options = options_map.get(col, [])
                        box = ttk.Combobox(top, textvariable=var, values=options, state="readonly", width=25, font=self.default_font)
                        box.grid(row=row_idx + 1, column=col_idx + 1, padx=5, pady=5)

            def add_row():
                new_vars = {}
                if entry_vars_list:
                    prev = entry_vars_list[-1]
                    for col in columns:
                        new_vars[col] = tk.StringVar(value=prev[col].get())
                else:
                    for col in columns:
                        new_vars[col] = tk.StringVar()
                entry_vars_list.append(new_vars)
                render_rows()

            def remove_row():
                if len(entry_vars_list) > 1:
                    entry_vars_list.pop()
                    render_rows()

            def submit_batch():
                new_rows = []
                for row_vars in entry_vars_list:
                    values = []
                    for col in columns:
                        val = row_vars[col].get()
                        if col == "EmployeeID":
                            val = display_to_employee_id.get(val, val)
                        elif col == "ItemID":
                            val = display_to_item_id.get(val, val)
                        elif col == "SerialNo":
                            val = display_to_serial_no.get(val, val)
                        elif col == "LocationID":
                            val = display_to_location_id.get(val, val)
                        values.append(val)
                    if any(values):
                        new_rows.append(values)

                # 批量输入内部查重
                seen_in_batch = set()
                for row in new_rows:
                    if tuple(row) in seen_in_batch:
                        messagebox.showerror("Error", "Some data already exists, please check.", parent=top)
                        return
                    seen_in_batch.add(tuple(row))

                # 添加内容和整个表格查重
                for row in new_rows:
                    if self.is_duplicate_record(row):
                        messagebox.showerror("Error", "The data you are trying to add already exists, please check.", parent=top)
                        return

                for row in new_rows:
                    self.all_data.append(row)
                self.apply_filter()
                self.modify_state = "Modefied, not saved"
                self.update_status()
                top.destroy()


            add_row()

            control_frame = tk.Frame(top)
            control_frame.grid(row=999, column=0, columnspan=len(columns) + 1, pady=10)
            self.create_button(control_frame, "+", add_row, width=3).pack(side=tk.LEFT, padx=5)
            self.create_button(control_frame, "-", remove_row, width=3).pack(side=tk.LEFT, padx=5)
            self.create_button(control_frame, "Canccel", top.destroy).pack(side=tk.RIGHT, padx=10)
            self.create_button(control_frame, "Submit", submit_batch).pack(side=tk.RIGHT, padx=10)
            return

        # 非批量添加
        top = tk.Toplevel(self)
        top.title("Add record")
        self.update_idletasks()
        win_w, win_h = 500, min(30 + 50 * len(columns), 700)
        pos_x = self.winfo_x() + (self.winfo_width() - win_w) // 2
        pos_y = self.winfo_y() + (self.winfo_height() - win_h) // 2
        top.geometry(f"{win_w}x{win_h}+{pos_x}+{pos_y}")

        entry_vars = {}
        for i, col in enumerate(columns):
            display_name = self.get_display_column_name(self.current_table_label.cget("text").replace("Current Table:", ""), col)
            tk.Label(top, text=display_name).grid(row=i, column=0, padx=10, pady=5, sticky='e')
            var = tk.StringVar()
            tk.Entry(top, textvariable=var, width=40).grid(row=i, column=1, padx=10, pady=5)
            entry_vars[col] = var

        def submit_normal():
            new_values = [entry_vars[col].get() for col in columns]

            if self.is_duplicate_record(new_values):
                messagebox.showerror("Error", "The data you are trying to add already exists, please check.", parent=top)
                return

            self.all_data.append(new_values)
            self.show_filtered_data(self.all_data)
            self.modify_state = "Modified, not saved"
            self.update_status()
            top.destroy()

        button_row = len(columns) + 1
        button_frame = tk.Frame(top)
        button_frame.grid(row=button_row, column=0, columnspan=2, pady=10, sticky="e")
        self.create_button(button_frame, "Cancel", top.destroy).pack(side=tk.RIGHT, padx=5)
        self.create_button(button_frame, "Submit", submit_normal).pack(side=tk.RIGHT, padx=5)

        # ——————————————————————————————
    def delete_record(self):
        """
        删除选中记录
        """
        selected_items = self.tree.selection()
        if not selected_items:
            return
        confirm_win = tk.Toplevel(self)
        confirm_win.title("Confirm delete")
        confirm_win.resizable(False, False)
        self.update_idletasks()
        main_x = self.winfo_x()
        main_y = self.winfo_y()
        main_w = self.winfo_width()
        main_h = self.winfo_height()
        win_w, win_h = 300, 120
        pos_x = main_x + (main_w - win_w) // 2
        pos_y = main_y + (main_h - win_h) // 2
        confirm_win.geometry(f"{win_w}x{win_h}+{pos_x}+{pos_y}")

        tk.Label(confirm_win, text=f"Are you sure you want to delete {len(selected_items)} record(s) that you selected?", font=("Helvetica", 14)).pack(pady=20)

        btn_frame = tk.Frame(confirm_win)
        btn_frame.pack(pady=10)

        def confirm_delete():
            to_delete = [self.tree.item(item, "values") for item in selected_items]

            # 针对汇总表格用 extract_id_from_display 对比，子表格直接用字符串对比
            if self.current_table_label.cget("text") == "Current Table: Summary Data":
                new_all_data = []
                for row in self.all_data:
                    if not any(all(str(row[i]) == self.extract_id_from_display(sel[i]) for i in range(len(self.columns))) for sel in to_delete):
                        new_all_data.append(row)
                self.all_data = new_all_data
            else:
                # 直接用字符串对比
                new_all_data = []
                for row in self.all_data:
                    if not any(all(str(row[i]) == sel[i] for i in range(len(self.columns))) for sel in to_delete):
                        new_all_data.append(row)
                self.all_data = new_all_data

            self.show_filtered_data(self.all_data)
            self.modify_state = "Modified, not saved"
            self.update_status()
            confirm_win.destroy()

        self.create_button(btn_frame, "Delete", confirm_delete).pack(side=tk.LEFT, padx=10)
        self.create_button(btn_frame, "Cancel", confirm_win.destroy).pack(side=tk.RIGHT, padx=10)

    def edit_record(self):
        """编辑选中内容"""
        selected_items = self.tree.selection()
        if not selected_items:
            return

        old_values_list = [self.tree.item(item, "values") for item in selected_items]
        columns = self.tree["columns"]
        top = tk.Toplevel(self)
        top.title(f"Edit {len(selected_items)} record(s)")

        self.update_idletasks()
        win_w, win_h = 700, min(80 + 50 * len(columns), 700)
        pos_x = self.winfo_x() + (self.winfo_width() - win_w) // 2
        pos_y = self.winfo_y() + (self.winfo_height() - win_h) // 2
        top.geometry(f"{win_w}x{win_h}+{pos_x}+{pos_y}")

        entry_vars = {}

        is_summary = self.current_table_label.cget("text") == "Current Table: Summary Data"

        if is_summary:
            # 准备 ID 到显示值映射
            employee_ids = self.extract_attributes_from_xml("EngineerInfo.xml", "Engineer", "EmployeeID")
            item_ids = self.extract_attributes_from_xml("TestItemInfo.xml", "TestItem", "ItemID")
            serial_nos = self.extract_attributes_from_xml("InstrumentInfo.xml", "Instrument", "SerialNo")
            location_ids = self.extract_attributes_from_xml("LocationInfo.xml", "Location", "LocationID")

            display_to_employee_id = {self.get_employee_display(eid): eid for eid in employee_ids}
            display_to_item_id = {self.get_item_display(iid): iid for iid in item_ids}
            display_to_serial_no = {self.get_serial_display(sn): sn for sn in serial_nos}
            display_to_location_id = {self.get_location_display(lid): lid for lid in location_ids}

            options_map = {
                "EmployeeID": list(display_to_employee_id.keys()),
                "ItemID": list(display_to_item_id.keys()),
                "SerialNo": list(display_to_serial_no.keys()),
                "LocationID": list(display_to_location_id.keys())
            }

        is_single = len(selected_items) == 1
        current_table = self.current_table_label.cget("text").replace("Current Table:", "")

        for i, col in enumerate(columns):
            display_name = self.get_display_column_name(current_table, col)
            tk.Label(top, text=display_name).grid(row=i, column=0, padx=10, pady=5, sticky='e')

            current_values = self.tree.item(selected_items[0], "values")
            raw_value = current_values[i] if is_single else ""

            var = tk.StringVar()

            if is_summary and col in options_map:
                # 设置汇总表格每一列显示的内容
                if col == "EmployeeID":
                    var.set(self.get_employee_display(raw_value))
                elif col == "ItemID":
                    var.set(self.get_item_display(raw_value))
                elif col == "SerialNo":
                    var.set(self.get_serial_display(raw_value))
                elif col == "LocationID":
                    var.set(self.get_location_display(raw_value))
                else:
                    var.set(raw_value)

                entry = ttk.Combobox(top, textvariable=var, values=options_map[col], state="readonly", width=40, font=self.default_font)
            else:
                var.set(raw_value)
                entry = tk.Entry(top, textvariable=var, width=40)

            entry.grid(row=i, column=1, padx=10, pady=5)
            entry_vars[col] = var

        def submit():
            for item in selected_items:
                old_values = self.tree.item(item, "values")
                updated_row = []
                for i, col in enumerate(columns):
                    display_val = entry_vars[col].get().strip()
                    if display_val == "":
                        current_val = old_values[i]
                        if is_summary and col in ["EmployeeID", "ItemID", "SerialNo", "LocationID"]:
                            current_val = self.extract_id_from_display(current_val)
                        updated_row.append(current_val)
                    else:
                        if is_summary:
                            if col == "EmployeeID":
                                val = display_to_employee_id.get(display_val, display_val)
                            elif col == "ItemID":
                                val = display_to_item_id.get(display_val, display_val)
                            elif col == "SerialNo":
                                val = display_to_serial_no.get(display_val, display_val)
                            elif col == "LocationID":
                                val = display_to_location_id.get(display_val, display_val)
                            else:
                                val = display_val
                        else:
                            val = display_val
                        updated_row.append(val)

                for row in self.all_data:
                    if all(str(row[i]) == self.extract_id_from_display(old_values[i]) for i in range(len(self.columns))):
                        continue

                    compare_row = [self.extract_id_from_display(v) if is_summary and col in ["EmployeeID", "ItemID", "SerialNo", "LocationID"] else v for col, v in zip(columns, row)]
                    if compare_row == updated_row:
                        messagebox.showerror("Error", "This record already exists", parent=top)
                        return

                for idx, row in enumerate(self.all_data):
                    if all(str(row[i]) == self.extract_id_from_display(old_values[i]) for i in range(len(self.columns))):
                        self.all_data[idx] = updated_row
                        break

            if is_summary:
                self.apply_filter()
            else:
                self.show_filtered_data(self.all_data)
            self.modify_state = "Modified, not saved"
            self.update_status()
            top.destroy()




        button_frame = tk.Frame(top)
        button_frame.grid(row=len(columns), column=0, columnspan=2, pady=10, sticky="e")
        self.create_button(button_frame, "Cancel", top.destroy).pack(side=tk.RIGHT, padx=5)
        self.create_button(button_frame, "Submit", submit).pack(side=tk.RIGHT, padx=5)

        #——————————————————————————————

    def save_data(self):
        """
        保存当前表格数据
        汇总表格和不确定度表格只需要保存到自己的文件
        其他表格保存到自己的文件同时保存到汇总表格对应的子表格
        """
        '''
        Summary Data
        Employee Data
        Instrument Data
        Location Data
        Project Data
        Uncertainty Data
        '''

        table_name = self.current_table_label.cget("text").replace("Current Table:", "")
        path = self.data_path_map.get(table_name)

        if not path:
            messagebox.showerror("Error", f"Unaable to find the path of {table_name} in Config.xml")
            return

        try:
            if table_name == "Summary Data":
                with codecs.open(path, "r", encoding="gb2312") as f:
                    content = f.read()
                root = ET.fromstring(content)

                section = root.find("Summary_Info")
                if section is not None:
                    section.clear()
                else:
                    section = ET.SubElement(root, "Summary_Info")

                for row in self.all_data:
                    ET.SubElement(section, "Summary", attrib={col: val for col, val in zip(self.columns, row)})

                with codecs.open(path, "w", encoding="gb2312") as f:
                    f.write(prettify_xml(root))

            elif table_name == "Uncertainty Data":
                root_tag = "UncertaintyInfo"
                outer_tag = "UncertaintyInfo_Info"
                item_tag = "Uncertainty"

                outer_elem = ET.Element(outer_tag)
                for row in self.all_data:
                    # ET.SubElement(outer_elem, item_tag, attrib={col: val for col, val in zip(self.columns, row)})
                    ET.SubElement(outer_elem, item_tag, attrib={col: restore_escaped_amp(val) for col, val in zip(self.columns, row)})
                wrapper = ET.Element(root_tag)
                wrapper.append(outer_elem)

                with codecs.open(path, "w", encoding="gb2312") as f:
                    f.write(prettify_xml(wrapper))

                self.refresh_uncertainty_map()

            else:
                # 保存到自己的文件
                tag_map = {
                    "Employee Data": ("Engineer_Info", "Engineer"),
                    "Instrument Data": ("Instrument_Info", "Instrument"),
                    "Location Data": ("Location_Info", "Location"),
                    "Project Data": ("TestItem_Info", "TestItem")
                }

                outer_tag, item_tag = tag_map[table_name]
                root_tag = os.path.splitext(os.path.basename(path))[0]  # e.g. "EngineerInfo"

                outer_elem = ET.Element(outer_tag)
                for row in self.all_data:
                    # ET.SubElement(outer_elem, item_tag, attrib={col: val for col, val in zip(self.columns, row)})
                    ET.SubElement(outer_elem, item_tag, attrib={col: restore_escaped_amp(val) for col, val in zip(self.columns, row)})
                wrapper = ET.Element(root_tag)
                wrapper.append(outer_elem)

                with codecs.open(path, "w", encoding="gb2312") as f:
                    f.write(prettify_xml(wrapper))

                # 同步更新汇总表格
                summary_path = self.data_path_map.get("Summary Data")
                if not summary_path:
                    messagebox.showerror("Error", "Cannot find the path of Summary Data in Config.xml ")
                    return

                with codecs.open(summary_path, "r", encoding="gb2312") as f:
                    content = f.read()
                root = ET.fromstring(content)

                section = root.find(outer_tag)
                if section is not None:
                    section.clear()
                else:
                    section = ET.SubElement(root, outer_tag)

                for row in self.all_data:
                    # ET.SubElement(section, item_tag, attrib={col: val for col, val in zip(self.columns, row)})
                    ET.SubElement(section, item_tag, attrib={col: restore_escaped_amp(val) for col, val in zip(self.columns, row)})
                with codecs.open(summary_path, "w", encoding="gb2312") as f:
                    f.write(prettify_xml(root))

            messagebox.showinfo("Success", f"{table_name} saved")
            self.modify_state = "Saved"
            self.update_status()

        except Exception as e:
            messagebox.showerror("Error", f"Data not saved: {e}")


            # ————————————————————————————————————————————————————————
    def toggle_select_all(self):
        """全选/取消全选功能"""
        if not self.all_selected:
            # 全选
            for item in self.tree.get_children():
                self.tree.selection_add(item)
            self.all_selected = True
            self.select_all_button.config(text="Cancel Selection")
        else:
            # 取消全选
            self.tree.selection_remove(self.tree.selection())
            self.all_selected = False
            self.select_all_button.config(text="Select All")

    def extract_attributes_from_xml(self, filename, tag, attribute):
        table_name = self.filename_to_table.get(filename)
        path = self.data_path_map.get(table_name)
        if not path or not os.path.exists(path):
            print(f"[Error] {filename}: didn't find path")
            return []

        try:
            with codecs.open(path, "r", encoding="gb2312", errors="replace") as f:
                xml_content = f.read()
            xml_content = xml_content.replace("&", "&amp;")
            root = ET.fromstring(xml_content)
            return list({elem.attrib.get(attribute, "") for elem in root.findall(f".//{tag}") if attribute in elem.attrib})
        except Exception as e:
            print(f"[Error] {filename}: {e}")
            return []


    def update_status(self):
    # 更新底部状态显示
        selected_count = len(self.tree.selection())
        self.status_label.config(
            text=f"Selected {selected_count} record   {self.modify_state}"
        )

    def on_click_toggle_selection(self, event):
        """点击任意一行来选中，再次点击取消选中"""
        item_id = self.tree.identify_row(event.y)
        if not item_id:
            return

        if item_id in self.tree.selection():
            self.tree.selection_remove(item_id)
        else:
            self.tree.selection_add(item_id)

        self.update_status()

        return "break"

    def update_treeview_font(self):
        """根据字体大小设置行高"""
        style = ttk.Style(self)
        style.configure("Treeview", font=('Helvetica', self.font_size))
        style.configure("Treeview.Heading", font=('Helvetica', self.font_size, 'bold'))

        # 设置行高
        row_height = int(self.font_size * 1.8)
        style.configure("Treeview", rowheight=row_height)

    def increase_font_size(self):
        """字体放大"""
        if self.font_size < 25:
            self.font_size += 1
            self.update_treeview_font()


    def decrease_font_size(self):
        """字体缩小"""
        if self.font_size > 8:  
            self.font_size -= 1
            self.update_treeview_font()

    def create_filter_controls(self, parent_frame):
        """创建全选按钮、筛选控件"""
        top_right_frame = tk.Frame(parent_frame, bg="#d9d9d9")
        top_right_frame.pack(side=tk.RIGHT, padx=10)

        # 全选按钮
        self.select_all_button = self.create_button(top_right_frame, "Select All", self.toggle_select_all)
        self.select_all_button.pack(side=tk.RIGHT, padx=4)

        # 汇总表格筛选下拉菜单
        self.filter_frame = tk.Frame(top_right_frame, bg="#d9d9d9")
        self.filter_frame.pack(side=tk.LEFT)

        tk.Label(self.filter_frame, text="Employee Name: ",  font=("Helvetica", 14), bg="#d9d9d9").pack(side=tk.LEFT)
        self.filter_employee = ttk.Combobox(self.filter_frame, width=8, font=self.default_font,  state="readonly")
        self.filter_employee.pack(side=tk.LEFT, padx=2)

        tk.Label(self.filter_frame, text="Project Name: ",  font=("Helvetica", 14), bg="#d9d9d9").pack(side=tk.LEFT)
        self.filter_item = ttk.Combobox(self.filter_frame, width=8,  font=self.default_font, state="readonly")
        self.filter_item.pack(side=tk.LEFT, padx=2)


        self.create_button(self.filter_frame, "Filter", self.apply_filter).pack(side=tk.LEFT, padx=4)
        self.create_button(self.filter_frame, "Cancel Filter", self.reset_filter).pack(side=tk.LEFT)

        # 不确定度表格筛选按钮、重新扫描按钮
        self.project_filter_button = self.create_button(top_right_frame, "Project Filter", self.open_project_filter_window)
        # self.rescan_button = tk.Button(self.status_frame, text="重新扫描书签", font=self.default_font, command=self.trigger_rescan)
        '''
        if table_name == "不确定度表格":
            self.write_word_button.pack(side=tk.RIGHT, padx=5)
            self.project_filter_button.pack(side=tk.RIGHT, padx=4)  # 项目筛选按钮只在不确定度表格显示
        else:
            self.write_word_button.pack_forget()
            self.project_filter_button.pack_forget() 
            '''

            
        '''
        Summary Data
        Employee Data
        Instrument Data
        Location Data
        Project Data
        Uncertainty Data
        '''


    def apply_filter(self):
        """汇总表格筛选功能"""
        emp_display = self.filter_employee.get()
        item_display = self.filter_item.get()

        display_to_employee_id = {self.get_employee_display(eid): eid for eid in self.engineer_info.keys()}
        display_to_item_id = {self.get_item_display(iid): iid for iid in self.item_info.keys()}

        emp_id = display_to_employee_id.get(emp_display, "") if emp_display else ""
        item_id = display_to_item_id.get(item_display, "") if item_display else ""

        columns = self.columns
        filtered_data = []

        for row in self.all_data:
            raw_emp_id = row[columns.index("EmployeeID")]
            raw_item_id = row[columns.index("ItemID")]

            # 直接用 ID 比较
            if (not emp_id or raw_emp_id == emp_id) and (not item_id or raw_item_id == item_id):
                filtered_data.append(row)

        self.show_filtered_data(filtered_data)


    def reset_filter(self):
        """重置筛选"""
        self.filter_employee.set("")
        self.filter_item.set("")
        # columns, data = self.load_table_data("汇总表格")
        # self.show_filtered_data(columns, data)
        self.show_filtered_data(self.all_data)

    def show_filtered_data(self, data):
        """根据选中的筛选内容更新表格显示内容"""
        self.filtered_data = data  # 记录当前筛选状态下显示的数据
        columns = self.columns
        current_table = self.current_table_label.cget("text").replace("Current Table:", "")

        self.tree.delete(*self.tree.get_children())
        self._original_tags.clear()

        for idx, row in enumerate(data):
            tag = 'evenrow' if idx % 2 == 0 else 'oddrow'
            display_row = []

            for i, col in enumerate(columns):
                # val = row[i]
                # val = val.replace("&amp;", "&") 
                val = restore_xml_display(row[i])
                if current_table == "Summary Data":
                    if col == "EmployeeID":
                        val = self.get_employee_display(val)
                    elif col == "ItemID":
                        val = self.get_item_display(val)
                    elif col == "SerialNo":
                        val = self.get_serial_display(val)
                    elif col == "LocationID":
                        val = self.get_location_display(val)
                display_row.append(val)

            self.tree.insert("", "end", values=display_row, tags=(tag,))
        self.adjust_column_widths()

    def load_xml_root(self, filename):
        import re
        table_name = self.filename_to_table.get(filename)
        if not table_name:
            raise FileNotFoundError(f"Cannot find {filename} in filename_to_table")

        path = self.data_path_map.get(table_name)
        if not path or not os.path.exists(path):
            raise FileNotFoundError(f"Unable to find the path of {filename}")

        with codecs.open(path, "r", encoding="gb2312", errors="replace") as f:
            xml_content = f.read()

        xml_content = re.sub(r'&(?![a-zA-Z]+;)', '&amp;', xml_content)

        return ET.fromstring(xml_content)



    def adjust_column_widths(self):
        """根据内容长度自动调节表格宽度"""
        font = ('Helvetica', self.font_size)
        # 每个列计算最长的字符长度
        for col in self.tree["columns"]:
            max_text_length = len(self.tree.heading(col)["text"]) 
            for item in self.tree.get_children():
                cell_value = self.tree.item(item, "values")[self.tree["columns"].index(col)]
                if cell_value:
                    max_text_length = max(max_text_length, len(cell_value))
            est_width = max_text_length * 14 + 20  # 每个字14像素
            self.tree.column(col, width=est_width)

    def extract_id_from_display(self, display_val):
        """从汇总表格的中文显示值提取 ID，如果格式正确返回ID，否则返回原值"""
        if '/' in display_val:
            return display_val.split('/')[-1].strip()
        else:
            return display_val

    def is_duplicate_record(self, new_row):
        """
        根据当前表格自动判断新添加/修改的内容（new_row）是否已存在
        """
        if self.current_table_label.cget("text") == "Current Table: Summary Data":
            # 转换显示值为纯 ID
            converted_row = []
            for i, col in enumerate(self.columns):
                val = new_row[i]
                if col == "EmployeeID":
                    val = self.extract_id_from_display(val)
                elif col == "ItemID":
                    val = self.extract_id_from_display(val)
                elif col == "SerialNo":
                    val = self.extract_id_from_display(val)
                elif col == "LocationID":
                    val = self.extract_id_from_display(val)
                converted_row.append(val)

            return tuple(converted_row) in [tuple(row) for row in self.all_data]

        else:
            # 子表格直接对比
            return tuple(new_row) in [tuple(row) for row in self.all_data]

    def is_duplicate_record_edit(self, new_row, old_row):
        """
        编辑查重
        """
        if self.current_table_label.cget("text") == "Current Table: Summary Data":
            converted_new_row = []
            converted_old_row = []
            for i, col in enumerate(self.columns):
                new_val = new_row[i]
                old_val = old_row[i]
                if col in ["EmployeeID", "ItemID", "SerialNo", "LocationID"]:
                    new_val = self.extract_id_from_display(new_val)
                    old_val = self.extract_id_from_display(old_val)
                converted_new_row.append(new_val)
                converted_old_row.append(old_val)

            for row in self.all_data:
                if row == converted_old_row:
                    continue 
                if row == converted_new_row:
                    return True
            return False

        else:
            for row in self.all_data:
                if row == old_row:
                    continue
                if row == new_row:
                    return True
            return False

    def refresh_uncertainty_map(self):
        """
        刷新当前不确定度书签字典
        1，每次进入不确定度表格或修改不确定度表格
        2，刷新不确定度字典内容{Name：Value}
        3，写入word时通过该字典查找书签对应的要修改的值
        """
        if self.current_table_label.cget("text") != "Current Table: Uncertainty Data":
            return  # 只有打开不确定度表格时刷新字典

        try:
            self.uncertainty_map = {
                row[self.columns.index("Name")]: row[self.columns.index("Value")]
                for row in self.all_data
                if len(row) >= max(self.columns.index("Name"), self.columns.index("Value")) + 1
            }
            print("[Debug] Uncertainty hash table updated", self.uncertainty_map)
        except Exception as e:
            print(f"[Debug] Error while updating uncertainty hash table: {e}")

    def write_to_word(self):
        if self.current_table_label.cget("text") != "Current Table: Uncertainty Data":
            messagebox.showwarning("Warning", "Current table is not uncertainty table, cannot write to word")
            return

        if not self.filtered_data:
            messagebox.showwarning("Warning", "Cannot find anything to write")
            return

        word_folder = os.path.join(os.getcwd(), "word")
        if not os.path.exists(word_folder):
            messagebox.showerror("Error", f"Unable to find word file: {word_folder}")
            return

        def wait_for_scan_then_write():
            import pythoncom
            import time
            pythoncom.CoInitialize()

            # 1. 等待扫描完成（如果还在进行）
            while not self.word_scan_complete:
                print("Scanning word files...")
                time.sleep(0.1)

            # 2. 构建筛选后的 name->value 字典
            try:
                name_index = self.columns.index("Name")
                value_index = self.columns.index("Value")
                filtered_map = {
                    row[name_index]: row[value_index]
                    for row in self.filtered_data
                    if len(row) > max(name_index, value_index)
                }
            except Exception as e:
                self.after(0, lambda: messagebox.showerror("Error", f"{e}"))
                return

            if not filtered_map:
                self.after(0, lambda: messagebox.showwarning("Warning", "No bookmark found"))
                return

            word_app = None
            try:
                word_app = win32.gencache.EnsureDispatch('Word.Application')
                word_app.Visible = False

                write_count = 0
                for path, bookmarks in self.word_bookmark_map.items():
                    if not any(name in bookmarks for name in filtered_map):
                        continue  # 跳过与当前筛选无关的文件

                    try:
                        doc = word_app.Documents.Open(path)
                        modified = False
                        for name in filtered_map:
                            if name in bookmarks:
                                bookmark = doc.Bookmarks(name)
                                rng = bookmark.Range
                                rng.Text = filtered_map[name]
                                doc.Bookmarks.Add(name, rng)
                                modified = True
                                print(f"[Write] {os.path.basename(path)}：{name} → {filtered_map[name]}")
                        if modified:
                            doc.Save()
                            write_count += 1
                        doc.Close()
                    except Exception as e:
                        print(f"[Write failed] {path}: {e}")
                        continue

                if word_app:
                    word_app.Quit()

                self.after(0, lambda: messagebox.showinfo("Finished", f"Processed {write_count} Word file(s)"))
            except Exception as ex:
                self.after(0, lambda: messagebox.showerror("Error", f"Writing failed: {str(ex)}"))
                if word_app:
                    word_app.Quit()
            finally:
                pythoncom.CoUninitialize()

        # 启动后台写入线程
        threading.Thread(target=wait_for_scan_then_write, daemon=True).start()

        # threading.Thread(target=task, daemon=True).start()

    def load_config_file(self):
        """加载 Config.xml 中的数据文件路径与 Word 文件路径"""
        config_path = os.path.join(os.getcwd(), "Config.xml")
        if not os.path.exists(config_path):
            messagebox.showerror("Error", "Didn't find Config.xml")
            return

        try:
            with codecs.open(config_path, "r", encoding="gb2312") as f:
                xml_content = f.read().replace("&", "&amp;")  # 处理非法实体
            root = ET.fromstring(xml_content)

            self.data_path_map = {}
            self.word_file_paths = []

            data_info = root.find("Data_Info")
            if data_info is not None:
                for file_elem in data_info.findall("File"):
                    name = file_elem.attrib.get("Name")
                    path = file_elem.attrib.get("Path")
                    if name and path:
                        table_name = self.filename_to_table.get(name)
                        if table_name:
                            self.data_path_map[table_name] = path

            word_info = root.find("Word_Info")
            if word_info is not None:
                for file_elem in word_info.findall("File"):
                    path = file_elem.attrib.get("Path")
                    if path:
                        self.word_file_paths.append(path)

            print("[Debug] Table file mapping", self.data_path_map)
            print("[Debug] Found word file(s): ", self.word_file_paths)

        except Exception as e:
            messagebox.showerror("Error", f"Config.xml error: {e}")

    def open_project_filter_window(self):
        top = tk.Toplevel(self)
        top.title("Project Filter")
        self.update_idletasks()
        win_w, win_h = 300, 400
        pos_x = self.winfo_x() + (self.winfo_width() - win_w) // 2
        pos_y = self.winfo_y() + (self.winfo_height() - win_h) // 2
        top.geometry(f"{win_w}x{win_h}+{pos_x}+{pos_y}")

        selected_vars = {}
        for project, keyword in self.project_filter_keywords.items():
            var = tk.BooleanVar(value=project in self.active_project_filters)
            cb = tk.Checkbutton(top, text=project, variable=var, font=("Helvetica", 14))
            cb.pack(anchor='w', padx=10, pady=4)
            selected_vars[project] = var

        def apply_filters():
            self.active_project_filters = {
                proj for proj, var in selected_vars.items() if var.get()
            }
            self.apply_uncertainty_filter()
            top.destroy()

        tk.Button(top, text="Apply Filter", command=apply_filters, font=("Helvetica", 14)).pack(pady=10)

    def apply_uncertainty_filter(self):
        if self.current_table_label.cget("text") != "Current Table: Uncertainty Data":
            return

        if not self.active_project_filters:
            self.show_filtered_data(self.all_data)
            return

        keywords = [self.project_filter_keywords[p] for p in self.active_project_filters]
        filtered = []
        for row in self.all_data:
            name = row[self.columns.index("Name")]
            if any(k in name for k in keywords):
                filtered.append(row)

        self.show_filtered_data(filtered)

    def filtered_data_names(self):
        if self.current_table_label.cget("text") != "Current Table: Uncertainty Data":
            return []

        try:
            name_index = self.columns.index("Name")
            return [row[name_index] for row in self.filtered_data]
        except:
            return []

    def scan_word_bookmarks(self):
        """
        后台扫描所有 Word 文件中的书签，建立缓存。
        """
        import pythoncom
        pythoncom.CoInitialize()
        word_app = None

        try:
            self.word_bookmark_map.clear()  # 清空旧缓存，适配“重新扫描”情况

            word_app = win32.gencache.EnsureDispatch('Word.Application')
            word_app.Visible = False

            for path in self.word_file_paths:
                if not path.lower().endswith((".doc", ".docx")):
                    continue

                try:
                    doc = word_app.Documents.Open(path, False, True)  # ReadOnly=True
                    bookmarks = [b.Name for b in doc.Bookmarks]
                    self.word_bookmark_map[path] = bookmarks
                    doc.Close(False)
                    print(f"[Scan Success] {os.path.basename(path)} → {bookmarks}")
                except Exception as e:
                    print(f"[Scan Failed] {os.path.basename(path)}：{e}")
                    continue

            print("[Scan Success] Total found word files:", len(self.word_bookmark_map))
        except Exception as e:
            print(f"[Scan Error] {e}")
        finally:
            if word_app:
                word_app.Quit()
            pythoncom.CoUninitialize()

            self.word_scan_complete = True

            def finalize_ui():
                if hasattr(self, "scan_progress"):
                    self.scan_progress.stop()
                    self.scan_progress.pack_forget()
            self.after(0, finalize_ui)


    def trigger_rescan(self):
        """重新扫描word文件，更新字典"""
        if self.word_scan_thread and self.word_scan_thread.is_alive():
            messagebox.showinfo("Warning", "File scanning, please try again later.")
            return

        self.word_scan_complete = False
        self.scan_progress.pack(side=tk.RIGHT, padx=10)
        self.scan_progress.start()

        self.word_scan_thread = threading.Thread(target=self.scan_word_bookmarks, daemon=True)
        self.word_scan_thread.start()

    def update_summary_filters(self):
        """刷新汇总表格的筛选下拉菜单内容"""
        self.engineer_info_display = {k: f"{v} / {k}" for k, v in self.engineer_info.items()}
        self.item_info_display = {k: f"{v} / {k}" for k, v in self.item_info.items()}
        # self.instrument_info_display = {k: f"{v} / {k}" for k, v in self.instrument_info.items()}
        # self.location_info_display = {k: f"{v} / {k}" for k, v in self.location_info.items()}

        self.filter_employee["values"] = list(self.engineer_info_display.values())
        self.filter_item["values"] = list(self.item_info_display.values())
        # self.filter_serial["values"] = list(self.instrument_info_display.values())
        # self.filter_location["values"] = list(self.location_info_display.values())

def restore_escaped_amp(val):
    return val.replace("&amp;", "&")

def restore_xml_display(val: str) -> str:
    """将 XML 中的转义字符还原为原始字符"""
    if not isinstance(val, str):
        return val
    # 多次调用 unescape 防止 &amp;amp; 的多重嵌套
    for _ in range(2):
        val = html.unescape(val)
    return val


def prettify_xml(elem):
    """格式化 ElementTree XML 对象为带缩进的字符串（移除空行并添加注释）"""
    # rough_string = ET.tostring(elem, encoding="utf-8")
    # rough_string = ET.tostring(elem, encoding="gb2312")
    rough_string = ET.tostring(elem, encoding="unicode")
    reparsed = minidom.parseString(rough_string)
    # pretty = reparsed.toprettyxml(indent="  ", encoding="gb2312").decode("gb2312")
    pretty = reparsed.toprettyxml(indent="  ", encoding="gb2312").decode("gb2312", errors="replace")
    pretty = "\n".join([line for line in pretty.splitlines() if line.strip()])
    
    # 保存的时候再添加一遍注释
    comment_map = {
        'TestItem_Info': '<!--Keep This Format-->',
        'Engineer_Info': '<!--Keep This Format-->',
        'Location_Info': '<!--Keep This Format-->',
        'Instrument_Info': '<!--Keep This Format-->',
        'Summary_Info': '<!--Keep This Format-->'
    }
    
    import re
    for tag, comment in comment_map.items():
        pattern = rf'(<{tag}>\n)'
        replacement = rf'\1    {comment}\n'
        pretty = re.sub(pattern, replacement, pretty, count=1)
    
    return pretty

if __name__ == "__main__":
    app = DatabaseBrowser() 
    app.mainloop()