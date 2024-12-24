import pandas as pd
import tkinter as tk
from tkinter import messagebox
import tkinter.ttk as ttk
from tkinter import scrolledtext  # For better text display

# Load databases
inventory_db_path = 'E:\\Program\\WarehouseManageSystem\\database\\database1.xlsx'
borrow_return_db_path = 'E:\\Program\\WarehouseManageSystem\\database\\database2.xlsx'


try:
    inventory_df = pd.read_excel(inventory_db_path, engine='openpyxl')
    borrow_return_df = pd.read_excel(borrow_return_db_path, engine='openpyxl')
    # Ensure data types are consistent.  This prevents later errors.
    inventory_df['数量'] = inventory_df['数量'].astype(int)
    # Add similar type checking for other relevant columns as needed.
except FileNotFoundError:
    messagebox.showerror("错误", "没找到数据库，请检查文件路径是否正确。")
    exit()
except pd.errors.EmptyDataError:
    messagebox.showerror("错误", "数据库文件为空。")
    exit()
except Exception as e:
    messagebox.showerror("错误", f"加载数据库失败: {e}")
    exit()


def clean_text(text):
    """Removes brackets and quotes from a string."""
    text = str(text)  # Handle potential non-string types
    text = text.replace('[', '').replace(']', '').replace("'", "")
    return text.strip()


def populate_cabinet_menu(location):
    """Populates the cabinet_menu based on the selected location."""
    if location:
        try:
            location_df = inventory_df[inventory_df['存放位置'].str.startswith(location)]
            cabinets = location_df['存放位置'].str.split('-').str[1].unique()
            cleaned_cabinets = [clean_text(c) for c in cabinets]  # Clean cabinet names
            cabinet_menu['values'] = cleaned_cabinets
        except Exception as e:
            messagebox.showerror("Error", f"Error populating cabinet menu: {e}")
    else:
        cabinet_menu['values'] = []


def on_location_menu_select(event):
    selected_location = location_choice.get()
    populate_cabinet_menu(selected_location)
    cabinet_entry.delete(0, tk.END)  # Clear cabinet_entry when location changes

def on_cabinet_menu_select(event):
    selected_cabinet = cabinet_menu.get()
    cabinet_entry.delete(0, tk.END)
    cabinet_entry.insert(0, selected_cabinet)

def on_category_subcategory_select(event):
    selected_category = category_entry.get()
    selected_subcategory = subcategory_entry.get()

# Functions
def update_subcategory_options(*args):
    selected_category = category_entry.get().strip()
    if selected_category:
        subcategories = inventory_df[inventory_df['大类名称'] == selected_category]['小类名称'].unique()
        subcategory_menu['menu'].delete(0, 'end')
        for subcategory in subcategories:
            subcategory_menu['menu'].add_command(label=subcategory, command=tk._setit(subcategory_choice, subcategory, set_subcategory_from_dropdown))
            
def update_subcategory_options2(*args):
    selected_category2 = category_entry2.get().strip()
    if selected_category2:
        subcategories = inventory_df[inventory_df['大类名称'] == selected_category2]['小类名称'].unique()
        subcategory_menu2['menu'].delete(0, 'end')
        for subcategory in subcategories:
            subcategory_menu2['menu'].add_command(label=subcategory, command=tk._setit(subcategory_choice2, subcategory, set_subcategory_from_dropdown2))

def update_cabinet_options(*args):
    selected_location = location_choice.get().strip()
    if selected_location:
        cabinets = inventory_df[inventory_df['存放位置'].str.startswith(selected_location)]['存放位置'].apply(lambda x: x.split('-')[1]).unique()
        cabinet_menu['menu'].delete(0, 'end')
        for cabinet in cabinets:
            cabinet_menu['menu'].add_command(label=cabinet, command=tk._setit(cabinet_number, cabinet, set_cabinet_from_dropdown))

def set_category_from_dropdown(*args):
    category_entry.delete(0, tk.END)
    category_entry.insert(0, category_choice.get())
    update_subcategory_options()

def set_subcategory_from_dropdown(*args):
    subcategory_entry.delete(0, tk.END)
    subcategory_entry.insert(0, subcategory_choice.get())
    
def set_category_from_dropdown2(*args):
    category_entry2.delete(0, tk.END)
    category_entry2.insert(0, category_choice2.get())
    update_subcategory_options2()

def set_subcategory_from_dropdown2(*args):
    subcategory_entry2.delete(0, tk.END)
    subcategory_entry2.insert(0, subcategory_choice2.get())

def set_cabinet_from_dropdown(*args):
    cabinet_entry.delete(0, tk.END)
    cabinet_entry.insert(0, cabinet_number.get())
    
def set_status_from_dropdown(*args):
    status_entry.delete(0, tk.END)
    status_entry.insert(0, status_choice.get())

def set_remark_from_dropdown(*args):
    remark_entry.delete(0, tk.END)
    remark_entry.insert(0, remark_choice.get())

def calculate_and_display_totals():
    category = category_choice.get().strip()
    subcategory = subcategory_choice.get().strip()

    if not category or not subcategory:
        messagebox.showerror("Error", "请在第2行和第3行右边的下拉列表选择要找东西的大类别和小类别名称。")
        return

    # Filter the DataFrame for the selected category and subcategory
    filtered_df = inventory_df[(inventory_df['大类名称'] == category) & (inventory_df['小类名称'] == subcategory)]

    # Calculate total count
    total_count = filtered_df['数量'].sum()

    # Calculate good count (excluding '坏的' and '损坏')
    good_count = filtered_df[~filtered_df['备注'].isin(['坏的', '损坏', '旧版本'])]['数量'].sum()

    # Get unique storage locations
    storage_locations = ", ".join(filtered_df['存放位置'].unique())

    # Display the result
    result = f"大类别名字：{category} —— 小类别名字：{subcategory} ——仓库现有总数：{total_count}  —— 仓库现有好的个数：{good_count}  —— 存放位置：{storage_locations}"
    messagebox.showinfo("统计结果", result)

def view_inventory():
    inventory_window = tk.Toplevel(root)
    inventory_window.title("Inventory")
    text = tk.Text(inventory_window)
    text.pack()
    for index, row in inventory_df.iterrows():
        text.insert(tk.END, f"{row['大类名称']} - {row['小类名称']}: {row['数量']} 放在 {row['存放位置']}\n")

def view_borrow_return():
    borrow_return_window = tk.Toplevel(root)
    borrow_return_window.title("Borrow/Return Records")
    text = tk.Text(borrow_return_window)
    text.pack()
    for index, row in borrow_return_df.iterrows():
        text.insert(tk.END, f"人员: {row['保管人员']}, {row['借出物品大类名称']} - {row['借出物品小类名称']}: {row['借出物品数量']} ({row['物品状态']})\n")

def add_inventory_item():
    global inventory_df
    try:
        category = category_entry.get().strip().lower()
        subcategory = subcategory_entry.get().strip().lower()
        quantity = int(quantity_entry.get().strip())
        location = f"{location_choice.get().strip()}-{cabinet_number.get().strip()}-{description_entry.get().strip()}"
        remark = remark_entry.get().strip().lower()
        user_input_note = item_note_entry.get().strip() if item_note_entry.get().strip() else ""  # 如果用户没有输入，确保是空字符串

        if quantity <= 0:
            raise ValueError("请输入一个正数。")

        # Find matching items in the database, ignoring case and including item_note
        matching_items = inventory_df[
            (inventory_df['大类名称'].str.lower() == category) &
            (inventory_df['小类名称'].str.lower() == subcategory) &
            (inventory_df['备注'].str.lower() == remark) &
            (inventory_df['物品备注'].fillna('').str.lower() == user_input_note.lower())  # 处理数据库中的空值
        ]

        if not matching_items.empty:
            # Update quantity of existing item when all fields match
            inventory_df.loc[matching_items.index, '数量'] += quantity
            messagebox.showinfo("成功", "物品已存在，数量已更新。")
        else:
            # Add as new item if any field doesn't match
            new_item_data = {
                '大类名称': category,
                '小类名称': subcategory,
                '数量': quantity,
                '存放位置': location,
                '备注': remark,
                '物品备注': user_input_note
            }
            inventory_df = pd.concat([inventory_df, pd.DataFrame([new_item_data])], ignore_index=True)
            messagebox.showinfo("成功", "物品不存在，物品已添加!")

        # Save updated DataFrame to Excel
        inventory_df.to_excel(inventory_db_path, index=False, engine='openpyxl')

    except ValueError as e:
        messagebox.showerror("Error", f"Invalid input: {e}")
    except Exception as e:
        messagebox.showerror("Error", f"An unexpected error occurred: {e}")

def update_databases():
    try:
        borrower = borrower_entry.get()
        category = category_entry2.get()
        subcategory = subcategory_entry2.get()
        quantity = int(quantity_entry2.get())
        status = status_entry.get()
        remark = remark_entry2.get()

        if quantity <= 0:
            raise ValueError("不要乱写负数！fk你！.")

        # Update borrow/return database
        new_borrow_entry = {
            '借出物品大类名称': category,
            '借出物品小类名称': subcategory,
            '借出物品数量': quantity,
            '保管人员': borrower,
            '物品状态': status,
            '备注': remark
        }
        global borrow_return_df
        borrow_return_df = borrow_return_df.append(new_borrow_entry, ignore_index=True)
        borrow_return_df.to_excel(borrow_return_db_path, index=False, engine='openpyxl')

        # Update inventory database
        inventory_index = inventory_df[(inventory_df['大类名称'] == category) & (inventory_df['小类名称'] == subcategory) & (inventory_df['备注'] == remark)].index

        if not inventory_index.empty:
            if status == '借出':
                inventory_df.at[inventory_index[0], '数量'] -= quantity
            elif status in ['归还', '采购']:
                inventory_df.at[inventory_index[0], '数量'] += quantity
            elif status in ['交付', '损坏']:
                inventory_df.at[inventory_index[0], '数量'] -= quantity

        inventory_df.to_excel(inventory_db_path, index=False, engine='openpyxl')
        messagebox.showinfo("成功", "数据库更新成功！")
    except ValueError as e:
        messagebox.showerror("Error", f"Invalid input: {e}")

def search_borrower_items():
    borrower_name = search_entry.get()
    search_window = tk.Toplevel(root)
    search_window.title(f"Items borrowed by {borrower_name}")
    text = tk.Text(search_window)
    text.pack()

    # Filter records for the specified borrower
    borrower_records = borrow_return_df[borrow_return_df['保管人员'] == borrower_name]

    if borrower_records.empty:
        text.insert(tk.END, f"没有找到这个人： {borrower_name}.\n")
    else:
        current_count = {}
        delivered_count = {}
        damaged_count = {}

        for index, row in borrower_records.iterrows():
            category = row['借出物品大类名称']  # Get the category name
            subcategory = row['借出物品小类名称']
            quantity = row['借出物品数量']
            status = row['物品状态']

            item_key = (category, subcategory) # Use a tuple as key to store both category and subcategory

            if status == '借出':
                current_count[item_key] = current_count.get(item_key, 0) + quantity
            elif status == '归还':
                current_count[item_key] = current_count.get(item_key, 0) - quantity
            elif status == '交付':
                delivered_count[item_key] = delivered_count.get(item_key, 0) + quantity
            elif status == '损坏':
                damaged_count[item_key] = damaged_count.get(item_key, 0) + quantity

        # Display results.  Format output to include category.
        def format_item(count, category, subcategory):
            return f"{count} 个 {category}-{subcategory}"

        current_items = ", ".join([format_item(count, category, subcategory) for (category, subcategory), count in current_count.items()])
        delivered_items = ", ".join([format_item(count, category, subcategory) for (category, subcategory), count in delivered_count.items()])
        damaged_items = ", ".join([format_item(count, category, subcategory) for (category, subcategory), count in damaged_count.items()])


        text.insert(tk.END, f"当前名下还有：{current_items}。\n")
        text.insert(tk.END, f"交付：{delivered_items}。\n")
        text.insert(tk.END, f"损坏：{damaged_items}。\n")

# GUI setup
root = tk.Tk()
root.title("仓库管理系统-v0.2")

# 初始化变量
location_choice = tk.StringVar(value="627")
cabinet_number = tk.StringVar(value="")
category_choice = tk.StringVar(value="")
subcategory_choice = tk.StringVar(value="")
status_choice = tk.StringVar(value="借出")
category_choice2 = tk.StringVar()
subcategory_choice2 = tk.StringVar()
detailed_description = tk.StringVar()
remark_choice = tk.StringVar()
# 设置样式
style = ttk.Style()
style.configure('Title.TLabel', font=('Arial', 12, 'bold'))
style.configure('Header.TLabel', font=('Arial', 10))
style.configure('Alert.TLabel', foreground='red', font=('Arial', 9))
style.configure('Action.TButton', padding=5)

# 创建主框架
main_frame = ttk.Frame(root, padding="10")
main_frame.pack(fill=tk.BOTH, expand=True)

# 标题
title_frame = ttk.Frame(main_frame)
title_frame.pack(fill=tk.X, pady=(0, 10))
title_label = ttk.Label(title_frame, 
    text="仓库管理系统-给yhw点赞版", 
    style='Title.TLabel',
    anchor='center')
title_label.pack(fill=tk.X)

# 使用Notebook来组织不同功能区域
notebook = ttk.Notebook(main_frame)
notebook.pack(fill=tk.BOTH, expand=True)

# === 库存管理标签页 ===
inventory_frame = ttk.Frame(notebook, padding="10")
notebook.add(inventory_frame, text='库存管理')

# 库存输入区域
input_frame = ttk.LabelFrame(inventory_frame, text="添加物品", padding="10")
input_frame.pack(fill=tk.X, pady=(0, 10))

# 第一行：大类别名称
category_frame = ttk.Frame(input_frame)
category_frame.pack(fill=tk.X, pady=2)
ttk.Label(category_frame, text="大类别名称（可以输入或下拉）：", 
         style='Header.TLabel', width=30).pack(side=tk.LEFT)
category_entry = ttk.Entry(category_frame, width=30)
category_entry.pack(side=tk.LEFT, padx=5)
category_menu = ttk.OptionMenu(category_frame, category_choice, "",
                             *sorted(inventory_df['大类名称'].unique()))
category_menu.pack(side=tk.LEFT)

# 第二行：子类别名称
subcategory_frame = ttk.Frame(input_frame)
subcategory_frame.pack(fill=tk.X, pady=2)
ttk.Label(subcategory_frame, text="子类别名称（输入中文或者小写）：", 
         style='Header.TLabel', width=30).pack(side=tk.LEFT)
subcategory_entry = ttk.Entry(subcategory_frame, width=30)
subcategory_entry.pack(side=tk.LEFT, padx=5)
subcategory_menu = ttk.OptionMenu(subcategory_frame, subcategory_choice, "")
subcategory_menu.pack(side=tk.LEFT)

# 第三行：数量
quantity_frame = ttk.Frame(input_frame)
quantity_frame.pack(fill=tk.X, pady=2)
ttk.Label(quantity_frame, text="数量：", 
         style='Header.TLabel', width=30).pack(side=tk.LEFT)
quantity_entry = ttk.Entry(quantity_frame, width=30)
quantity_entry.pack(side=tk.LEFT, padx=5)

# 第四行：存放位置
location_frame = ttk.Frame(input_frame)
location_frame.pack(fill=tk.X, pady=2)
ttk.Label(location_frame, text="存放位置（门号-柜子号-细分描述）：", 
         style='Header.TLabel', width=30).pack(side=tk.LEFT)

# Location Combobox
location_combo = ttk.Combobox(location_frame, textvariable=location_choice,
                            values=["627", "629"], width=10, state='readonly')
location_combo.pack(side=tk.LEFT, padx=5)

# Cabinet Entry and Combobox
cabinet_entry = ttk.Entry(location_frame, textvariable=cabinet_number, width=15)
cabinet_entry.pack(side=tk.LEFT, padx=5)

cabinet_menu = ttk.Combobox(location_frame, width=15, state='readonly')
cabinet_menu.pack(side=tk.LEFT, padx=5)

# Description Entry
description_entry = ttk.Entry(location_frame, width=20)
description_entry.pack(side=tk.LEFT, padx=5)

# 第五行：描述符
remark_frame = ttk.Frame(input_frame)
remark_frame.pack(fill=tk.X, pady=2)
ttk.Label(remark_frame, text="描述符(好的，坏的，编号)：", 
         style='Header.TLabel', width=30).pack(side=tk.LEFT)
remark_entry = ttk.Entry(remark_frame, width=30)
remark_entry.pack(side=tk.LEFT, padx=5)
# 添加备注下拉菜单
# 处理备注选项：将所有值转换为字符串，并过滤掉空值
remark_options = inventory_df['备注'].fillna('').astype(str)  # 将NaN转换为空字符串
remark_options = sorted([x for x in remark_options.unique() if x != ''])  # 排序并移除空字符串

# 添加备注下拉菜单
remark_menu = ttk.OptionMenu(remark_frame, 
                            remark_choice, 
                            "", 
                            *remark_options,  # 使用处理后的选项
                            command=set_remark_from_dropdown)
remark_menu.pack(side=tk.LEFT)

# 新增第六行：物品备注
item_note_frame = ttk.Frame(input_frame)
item_note_frame.pack(fill=tk.X, pady=2)
ttk.Label(item_note_frame, text="物品备注（选填）：", 
         style='Header.TLabel', width=30).pack(side=tk.LEFT)
item_note_entry = ttk.Entry(item_note_frame, width=30)
item_note_entry.pack(side=tk.LEFT, padx=5)
item_note_entry.insert(0, "")  # 设置默认值为空字符串

# 按钮区域
button_frame = ttk.Frame(input_frame)
button_frame.pack(fill=tk.X, pady=10)
ttk.Button(button_frame, text="点击添加物品", 
          command=add_inventory_item, style='Action.TButton').pack(side=tk.LEFT, padx=5)
ttk.Button(button_frame, text="点击查找物品信息", 
          command=calculate_and_display_totals, style='Action.TButton').pack(side=tk.LEFT, padx=5)

# 警告信息
alert_frame = ttk.Frame(input_frame)
alert_frame.pack(fill=tk.X, pady=5)
ttk.Label(alert_frame, text="管理员操作注意：物品先出库，再入库，最后处理损坏或交付！", 
         style='Alert.TLabel').pack(side=tk.LEFT)
ttk.Label(alert_frame, text="不知道描述符请查看数据库或找管理员贴标签！", 
         style='Alert.TLabel').pack(side=tk.RIGHT)

# === 借还管理标签页 ===
borrow_return_frame = ttk.Frame(notebook, padding="10")
notebook.add(borrow_return_frame, text='借还管理')

# 借还管理区域
borrow_frame = ttk.LabelFrame(borrow_return_frame, text="借还管理", padding="10")
borrow_frame.pack(fill=tk.X, pady=(0, 10))

# 第一行：借还人员
borrower_frame = ttk.Frame(borrow_frame)
borrower_frame.pack(fill=tk.X, pady=2)
ttk.Label(borrower_frame, text="借还人员：", 
         style='Header.TLabel', width=30).pack(side=tk.LEFT)
borrower_entry = ttk.Entry(borrower_frame, width=30)
borrower_entry.pack(side=tk.LEFT, padx=5)

# 第二行：借还状态
status_frame = ttk.Frame(borrow_frame)
status_frame.pack(fill=tk.X, pady=2)
ttk.Label(status_frame, text="借还状态：", 
         style='Header.TLabel', width=30).pack(side=tk.LEFT)
status_entry = ttk.Entry(status_frame, width=30)
status_entry.pack(side=tk.LEFT, padx=5)

# Define a function to update the status_entry
def update_status_entry(*args):
    status_entry.delete(0, tk.END)
    status_entry.insert(0, status_choice.get())

# Bind the function to the status_choice variable
status_choice.trace("w", update_status_entry)

status_menu = ttk.OptionMenu(status_frame, status_choice, 
                           "默认","借出", "归还", "交付", "采购", "损坏")
status_menu.pack(side=tk.LEFT)

# 第三行：大类别名称
category2_frame = ttk.Frame(borrow_frame)
category2_frame.pack(fill=tk.X, pady=2)
ttk.Label(category2_frame, text="大类别名称：", 
         style='Header.TLabel', width=30).pack(side=tk.LEFT)
category_entry2 = ttk.Entry(category2_frame, width=30)
category_entry2.pack(side=tk.LEFT, padx=5)
category_menu2 = ttk.OptionMenu(category2_frame, category_choice2, "",
                              *sorted(inventory_df['大类名称'].unique()))
category_menu2.pack(side=tk.LEFT)

# 第四行：子类别名称
subcategory2_frame = ttk.Frame(borrow_frame)
subcategory2_frame.pack(fill=tk.X, pady=2)
ttk.Label(subcategory2_frame, text="子类别名称：", 
         style='Header.TLabel', width=30).pack(side=tk.LEFT)
subcategory_entry2 = ttk.Entry(subcategory2_frame, width=30)
subcategory_entry2.pack(side=tk.LEFT, padx=5)
subcategory_menu2 = ttk.OptionMenu(subcategory2_frame, subcategory_choice2, "")
subcategory_menu2.pack(side=tk.LEFT)

# 第五行：数量
quantity2_frame = ttk.Frame(borrow_frame)
quantity2_frame.pack(fill=tk.X, pady=2)
ttk.Label(quantity2_frame, text="数量：", 
         style='Header.TLabel', width=30).pack(side=tk.LEFT)
quantity_entry2 = ttk.Entry(quantity2_frame, width=30)
quantity_entry2.pack(side=tk.LEFT, padx=5)

# 第六行：描述符
remark2_frame = ttk.Frame(borrow_frame)
remark2_frame.pack(fill=tk.X, pady=2)
ttk.Label(remark2_frame, text="描述符（借出时的描述符）：", 
         style='Header.TLabel', width=30).pack(side=tk.LEFT)
remark_entry2 = ttk.Entry(remark2_frame, width=30)
remark_entry2.pack(side=tk.LEFT, padx=5)

# 更新按钮
update_button = ttk.Button(borrow_frame, text="点击更新数据库", 
                         command=update_databases, style='Action.TButton')
update_button.pack(pady=10)

# === 查询功能标签页 ===
search_frame = ttk.Frame(notebook, padding="10")
notebook.add(search_frame, text='查询功能')

# 查询区域
search_area = ttk.LabelFrame(search_frame, text="查询功能", padding="10")
search_area.pack(fill=tk.X, pady=(0, 10))

# 查询输入框和按钮
search_input_frame = ttk.Frame(search_area)
search_input_frame.pack(fill=tk.X, pady=5)
ttk.Label(search_input_frame, text="查找名下财产：", 
         style='Header.TLabel').pack(side=tk.LEFT)
search_entry = ttk.Entry(search_input_frame, width=30)
search_entry.pack(side=tk.LEFT, padx=5)
search_button = ttk.Button(search_input_frame, text="开始查找", 
                         command=search_borrower_items, style='Action.TButton')
search_button.pack(side=tk.LEFT, padx=5)

# 查看记录按钮
view_frame = ttk.Frame(search_area)
view_frame.pack(fill=tk.X, pady=10)
view_button = ttk.Button(view_frame, text="查看仓库物品详细信息", 
                       command=view_inventory, style='Action.TButton')
view_button.pack(side=tk.LEFT, padx=5)
view_borrow_return_button = ttk.Button(view_frame, text="查看借还记录", 
                                    command=view_borrow_return, style='Action.TButton')
view_borrow_return_button.pack(side=tk.LEFT, padx=5)

# 绑定事件
location_combo.bind('<<ComboboxSelected>>', on_location_menu_select)
cabinet_menu.bind('<<ComboboxSelected>>', on_cabinet_menu_select)
category_entry.bind("<FocusOut>", lambda event: on_category_subcategory_select(None))
subcategory_entry.bind("<FocusOut>", lambda event: on_category_subcategory_select(None))
# 绑定事件
category_choice.trace("w", set_category_from_dropdown)
subcategory_choice.trace("w", set_subcategory_from_dropdown)
# 绑定事件
category_choice.trace("w", set_category_from_dropdown)
subcategory_choice.trace("w", set_subcategory_from_dropdown)
category_choice2.trace("w", set_category_from_dropdown2)
subcategory_choice2.trace("w", set_subcategory_from_dropdown2)

# 启动主循环
root.mainloop()