import pandas as pd
import tkinter as tk
from tkinter import ttk, simpledialog, filedialog
import tkinter.font as tkfont

# 讀取 Excel 文件
def load_data(file_path):
    return pd.read_excel(file_path)

# 更新表格顯示數據
def update_tree(tree, data):
    for row in tree.get_children():
        tree.delete(row)
    for idx, row in data.iterrows():
        tree.insert("", "end", values=list(row))

# 欄位選取值篩選
def filter_by_value(tree, data, column):
    value = simpledialog.askstring("Filter by Value", f"Enter value for {column}:")
    if value is not None:
        filtered_data = data[data[column].astype(str).str.contains(value, regex=False, na=False)]
        update_tree(tree, filtered_data)

# 全域字串查詢
def search_global(tree, data):
    query = simpledialog.askstring("Global Search", "Enter search query:")
    if query:
        filtered_data = data.apply(lambda col: col.astype(str).str.contains(query, regex=False, na=False))
        result = data[filtered_data.any(axis=1)]
        update_tree(tree, result)

# 選擇文件
def choose_file():
    file_path = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path:
        data = load_data(file_path)
        create_ui(data)

# 創建 UI 顯示數據
def create_ui(data):
    root = tk.Tk()
    root.title("Excel Data Viewer")
    root.geometry("800x600")  # 初始化大小，但可以調整

    frame = ttk.Frame(root, padding="3 3 12 12")
    frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
    frame.grid_columnconfigure(0, weight=1)
    frame.grid_rowconfigure(0, weight=1)

    root.grid_columnconfigure(0, weight=1)
    root.grid_rowconfigure(0, weight=1)

    tree = ttk.Treeview(frame, columns=list(data.columns), show="headings")
    tree.grid(row=0, column=0, sticky='nsew')

    # 滾動條
    scrollbar_vertical = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=tree.yview)
    scrollbar_vertical.grid(row=0, column=1, sticky='ns')
    tree.configure(yscrollcommand=scrollbar_vertical.set)

    # 添加水平滾動條
    scrollbar_horizontal = ttk.Scrollbar(frame, orient=tk.HORIZONTAL, command=tree.xview)
    scrollbar_horizontal.grid(row=1, column=0, sticky='ew')
    tree.configure(xscrollcommand=scrollbar_horizontal.set)

    # 添加列名
    for col in data.columns:
        tree.heading(col, text=col, command=lambda c=col: filter_by_value(tree, data, c))
        tree.column(col, anchor="w", stretch=tk.YES)

    # 添加數據行
    update_tree(tree, data)

    # 搜索按鈕
    search_btn = ttk.Button(frame, text="Global Search", command=lambda: search_global(tree, data))
    search_btn.grid(row=2, column=0, sticky='ew', pady=4)

    root.mainloop()

# 主功能
if __name__ == "__main__":
    choose_file()
