import pandas as pd
import tkinter as tk
from tkinter import messagebox
import tkinter.ttk as ttk

# Load databases
inventory_db_path = 'E:\\Program\\WarehouseManageSystem\\database\\database1.xlsx'
borrow_return_db_path = 'E:\\Program\\WarehouseManageSystem\\database\\database2.xlsx'

try:
    inventory_df = pd.read_excel(inventory_db_path, engine='openpyxl')
    borrow_return_df = pd.read_excel(borrow_return_db_path, engine='openpyxl')
except FileNotFoundError:
    messagebox.showerror("错误", "没找到数据库，检查E盘数据是不是被删掉了.")
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
    cabinet_entry.insert(0, "")

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

def set_cabinet_from_dropdown(*args):
    cabinet_entry.delete(0, tk.END)
    cabinet_entry.insert(0, cabinet_number.get())

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
    result = f"大类别名字：{category} —— 小类别名字：{subcategory}   —— 仓库现有好的个数：{good_count}  —— 存放位置：{storage_locations}"
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

        if quantity <= 0:
            raise ValueError("请输入一个正数。")

        new_item_data = {
            '大类名称': category,
            '小类名称': subcategory,
            '数量': quantity,
            '存放位置': location,
            '备注': remark
        }

        # Find matching items in the database, ignoring case
        matching_items = inventory_df[
            (inventory_df['大类名称'].str.lower() == new_item_data['大类名称']) &
            (inventory_df['小类名称'].str.lower() == new_item_data['小类名称']) &
            (inventory_df['备注'].str.lower() == new_item_data['备注'])
        ]

        if not matching_items.empty:
            # Update quantity of existing item
            inventory_df.loc[matching_items.index, '数量'] += quantity
            messagebox.showinfo("成功", "物品已存在，数量已更新。")
        else:
            # Add new item to the DataFrame
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

        if quantity <= 0:
            raise ValueError("不要乱写负数！fk你！.")

        # Update borrow/return database
        new_borrow_entry = {
            '借出物品大类名称': category,
            '借出物品小类名称': subcategory,
            '借出物品数量': quantity,
            '保管人员': borrower,
            '物品状态': status,
            '备注': ''
        }
        global borrow_return_df
        borrow_return_df = borrow_return_df.append(new_borrow_entry, ignore_index=True)
        borrow_return_df.to_excel(borrow_return_db_path, index=False, engine='openpyxl')

        # Update inventory database
        inventory_index = inventory_df[(inventory_df['大类名称'] == category) & (inventory_df['小类名称'] == subcategory)].index

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
        text.insert(tk.END, f"No records found for {borrower_name}.\n")
    else:
        current_count = {}
        delivered_count = {}
        damaged_count = {}
        
        for index, row in borrower_records.iterrows():
            subcategory = row['借出物品小类名称']
            quantity = row['借出物品数量']
            status = row['物品状态']
            
            if status == '借出':
                current_count[subcategory] = current_count.get(subcategory, 0) + quantity
            elif status == '归还':
                current_count[subcategory] = current_count.get(subcategory, 0) - quantity
            elif status == '交付':
                delivered_count[subcategory] = delivered_count.get(subcategory, 0) + quantity
            elif status == '损坏':
                damaged_count[subcategory] = damaged_count.get(subcategory, 0) + quantity
        
        # Display results
        current_items = ", ".join([f"{count} 个 {name}" for name, count in current_count.items()])
        delivered_items = ", ".join([f"{name} 数量：{count}" for name, count in delivered_count.items()])
        damaged_items = ", ".join([f"{name}个数：{count}" for name, count in damaged_count.items()])
        
        text.insert(tk.END, f"当前名下还有：{current_items}。\n")
        text.insert(tk.END, f"交付{delivered_items}。\n")
        text.insert(tk.END, f"损坏{damaged_items}。\n")


# GUI setup
root = tk.Tk()
root.title("仓库管理系统-v0.1-yhw最帅版")
view_button = tk.Button(root, text="----------------------------------->点击按钮，守护最棒的kd大将军<----------------------------------")
view_button.pack()


# Frame for Inventory Management
inventory_frame = tk.Frame(root)
inventory_frame.pack(pady=10)

tk.Label(inventory_frame, text="大类别名称（NX，飞控...）").grid(row=0, column=0)
category_entry = tk.Entry(inventory_frame)
category_entry.insert(0, "") # Set default to empty string
category_entry.grid(row=0, column=1)

# Dropdown for category
category_choice = tk.StringVar()
category_menu = tk.OptionMenu(inventory_frame, category_choice, *inventory_df['大类名称'].unique(), command=set_category_from_dropdown)
category_menu.grid(row=0, column=2)

tk.Label(inventory_frame, text="子类别名称（8g核心板，底板，nxt v1, ...）").grid(row=1, column=0)
subcategory_entry = tk.Entry(inventory_frame)
subcategory_entry.insert(0, "")
subcategory_entry.grid(row=1, column=1)

# Dropdown for subcategory
subcategory_choice = tk.StringVar()
subcategory_menu = tk.OptionMenu(inventory_frame, subcategory_choice, "", command=set_subcategory_from_dropdown)
subcategory_menu.grid(row=1, column=2)

tk.Label(inventory_frame, text="数量").grid(row=2, column=0)
quantity_entry = tk.Entry(inventory_frame)
quantity_entry.insert(0, "")
quantity_entry.grid(row=2, column=1)

# Location selection using dropdowns
tk.Label(inventory_frame, text="存放位置（门号-柜子号-细分描述）").grid(row=3, column=0)
location_choice = tk.StringVar(value="627")
location_menu = tk.OptionMenu(inventory_frame, location_choice, "627", "629", command=update_cabinet_options)
location_menu.grid(row=3, column=1)

cabinet_number = tk.StringVar()
cabinet_entry = tk.Entry(inventory_frame, textvariable=cabinet_number)
cabinet_entry.grid(row=3, column=2)

# Dropdown for cabinet number
# Cabinet Menu
cabinet_number = tk.StringVar()
cabinet_menu = ttk.Combobox(inventory_frame, textvariable=cabinet_number, state='readonly')
cabinet_menu.grid(row=3, column=3)
cabinet_menu.bind("<<ComboboxSelected>>", on_cabinet_menu_select)

# Description Entry
description_entry = tk.Entry(inventory_frame)
description_entry.grid(row=3, column=5)  # Placed to the right of description label


# Bind events
location_choice.trace("w", lambda *args: on_location_menu_select(None)) # Trigger when location changes
category_entry.bind("<FocusOut>", lambda event: on_category_subcategory_select(None))
subcategory_entry.bind("<FocusOut>", lambda event: on_category_subcategory_select(None))

tk.Label(inventory_frame, text="备注(建议写上：好的，坏的，编号，便于数据库查询)").grid(row=4, column=0)
remark_entry = tk.Entry(inventory_frame)
remark_entry.insert(0, "")
remark_entry.grid(row=4, column=1)

add_button = tk.Button(inventory_frame, text="添加物品", command=add_inventory_item)
add_button.grid(row=5, columnspan=4, pady=5)

# Button to calculate and display totals
calculate_button = tk.Button(inventory_frame, text="查找物品和数量吧~", command=calculate_and_display_totals)
calculate_button.grid(row=6, columnspan=2, pady=5)

alert_button = tk.Button(inventory_frame, text="先出库，再入库，操作完再操作损坏或者交付!", command=calculate_and_display_totals)
alert_button.grid(row=6, columnspan=9, pady=5)

# Frame for Borrow/Return Management
borrow_return_frame = tk.Frame(root)
borrow_return_frame.pack(pady=10)

tk.Label(borrow_return_frame, text="借还人员").grid(row=0, column=0)
borrower_entry = tk.Entry(borrow_return_frame)
borrower_entry.grid(row=0, column=1)

tk.Label(borrow_return_frame, text="借还状态（借出，归还，交付，采购, 损坏）").grid(row=1, column=0)
status_entry = tk.Entry(borrow_return_frame)
status_entry.grid(row=1, column=1)

tk.Label(borrow_return_frame, text="大类别名称").grid(row=2, column=0)
category_entry2 = tk.Entry(borrow_return_frame)
category_entry2.grid(row=2, column=1)

tk.Label(borrow_return_frame, text="子类别名称").grid(row=3, column=0)
subcategory_entry2 = tk.Entry(borrow_return_frame)
subcategory_entry2.grid(row=3, column=1)

tk.Label(borrow_return_frame, text="数量").grid(row=4, column=0)
quantity_entry2 = tk.Entry(borrow_return_frame)
quantity_entry2.grid(row=4, column=1)

update_button = tk.Button(borrow_return_frame, text="点击更新数据库", command=update_databases)
update_button.grid(row=5, columnspan=2, pady=5)



# Frame for Search Functionality
search_frame = tk.Frame(root)
search_frame.pack(pady=10)

tk.Label(search_frame, text="查找你名下的财产").grid(row=0, column=0)
search_entry = tk.Entry(search_frame)
search_entry.grid(row=0, column=1)

search_button = tk.Button(search_frame, text="开始查找吧~", command=search_borrower_items)
search_button.grid(row=1, columnspan=2, pady=5)

# Frame for Viewing Records
view_frame = tk.Frame(root)
view_frame.pack(pady=10)

view_button = tk.Button(view_frame, text="查看仓库", command=view_inventory)
view_button.grid(row=0, column=0, padx=5)

view_borrow_return_button = tk.Button(view_frame, text="查看借还记录", command=view_borrow_return)
view_borrow_return_button.grid(row=0, column=1, padx=5)

root.mainloop()