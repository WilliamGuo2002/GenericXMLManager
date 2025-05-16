import tkinter as tk
from tkinter import ttk, messagebox
import os
import xml.etree.ElementTree as ET

class XMLDatabaseBrowser(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("XML Database Browser")
        self.geometry("1200x800")

        self.data_folder = "data"
        self.current_file = None
        self.columns = []
        self.all_data = []

        self.create_ui()
        self.load_xml_files()

    def create_ui(self):
        # File selector
        self.file_combo = ttk.Combobox(self, state="readonly")
        self.file_combo.pack(pady=10)
        self.file_combo.bind("<<ComboboxSelected>>", self.on_file_selected)

        # Table
        self.tree = ttk.Treeview(self, show="headings")
        self.tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Control buttons
        btn_frame = tk.Frame(self)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="Add", command=self.add_record).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Delete", command=self.delete_record).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Save", command=self.save_data).pack(side=tk.LEFT, padx=5)

    def load_xml_files(self):
        if not os.path.exists(self.data_folder):
            os.makedirs(self.data_folder)
        files = [f for f in os.listdir(self.data_folder) if f.lower().endswith(".xml")]
        self.file_combo["values"] = files
        if files:
            self.file_combo.current(0)
            self.load_table_data(files[0])

    def on_file_selected(self, event):
        filename = self.file_combo.get()
        self.load_table_data(filename)

    def load_table_data(self, filename):
        path = os.path.join(self.data_folder, filename)
        try:
            tree = ET.parse(path)
            root = tree.getroot()
            entries = list(root)
            if not entries:
                self.columns = []
                self.all_data = []
            else:
                first = entries[0]
                self.columns = list(first.attrib.keys())
                self.all_data = [[entry.attrib.get(col, "") for col in self.columns] for entry in entries]
            self.refresh_table()
            self.current_file = filename
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load XML: {e}")

    def refresh_table(self):
        self.tree["columns"] = self.columns
        for col in self.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, anchor="center")
        self.tree.delete(*self.tree.get_children())
        for row in self.all_data:
            self.tree.insert("", "end", values=row)

    def add_record(self):
        messagebox.showinfo("Info", "Add function will be implemented in step 2.")

    def delete_record(self):
        selected = self.tree.selection()
        if not selected:
            return
        for item in selected:
            row = self.tree.item(item, "values")
            if row in self.all_data:
                self.all_data.remove(list(row))
        self.refresh_table()

    def save_data(self):
        if not self.current_file:
            return
        root = ET.Element("Data")
        for row in self.all_data:
            ET.SubElement(root, "Entry", attrib={col: val for col, val in zip(self.columns, row)})
        tree = ET.ElementTree(root)
        path = os.path.join(self.data_folder, self.current_file)
        tree.write(path, encoding="utf-8", xml_declaration=True)
        messagebox.showinfo("Saved", f"Data saved to {self.current_file}")
        
    # 在现有 XMLDatabaseBrowser_step1 基础上，增加通用添加功能

# 直接补全 add_record 函数

add_record_code = 
    def add_record(self):
        if not self.columns:
            messagebox.showwarning("Warning", "No file loaded or file is empty.")
            return

        top = tk.Toplevel(self)
        top.title("Add Record")
        entries = {}

        for i, col in enumerate(self.columns):
            tk.Label(top, text=col).grid(row=i, column=0, padx=10, pady=5, sticky='e')
            var = tk.StringVar()
            tk.Entry(top, textvariable=var, width=40).grid(row=i, column=1, padx=10, pady=5)
            entries[col] = var

        def submit():
            new_row = [entries[col].get() for col in self.columns]
            if any(not val.strip() for val in new_row):
                messagebox.showerror("Error", "All fields must be filled.", parent=top)
                return
            if new_row in self.all_data:
                messagebox.showerror("Error", "Record already exists.", parent=top)
                return
            self.all_data.append(new_row)
            self.refresh_table()
            top.destroy()

        tk.Button(top, text="Submit", command=submit).grid(row=len(self.columns), column=1, pady=10, sticky='e')


# 替换并保存到新的文件
with open("/mnt/data/XMLDatabaseBrowser_step2.py", "w", encoding="utf-8") as f:
    with open("/mnt/data/XMLDatabaseBrowser_step1.py", "r", encoding="utf-8") as base:
        content = base.read()
        # 用新的 add_record 替换原函数
        content = content.replace(
            'def add_record(self):\n        messagebox.showinfo("Info", "Add function will be implemented in step 2.")',
            add_record_code.strip()
        )
        f.write(content)

"/mnt/data/XMLDatabaseBrowser_step2.py"

        

if __name__ == "__main__":
    app = XMLDatabaseBrowser()
    app.mainloop()
