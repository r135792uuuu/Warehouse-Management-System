import pandas as pd
import tkinter as tk
from tkinter import messagebox
import tkinter.ttk as ttk
from tkinter import scrolledtext  # For better text display
from datetime import datetime  # 添加这行来导入datetime
import sys
import uuid # 用于生成唯一的RequestID
import time # 用于时间戳
import os



# Load databases本地数据
# inventory_db_path = 'E:\\Program\\WarehouseManageSystem\\database\\database1.xlsx'
# borrow_return_db_path = 'E:\\Program\\WarehouseManageSystem\\database\\database2.xlsx'
# permissions_return_db_path = 'E:\\Program\\WarehouseManageSystem\\database\\permissions.xlsx'

# 飞书表格api相关。很麻烦，暂时放弃
# 知识库id ： https://pcnkfnfllq1a.feishu.cn/wiki/ZGMZwPVqriCrU6kJ8nQcfMExnqL
# excel表格链接： https://pcnkfnfllq1a.feishu.cn/wiki/KqjZw2BCdi4vx1k6re5cav1MnAe

# 后面开发将直接在NAS上维护。nas的管理中我设置了仅管理员权限可以访问。需要的话后面更改
# nas表格地址： //HILAB627_DS/database
inventory_db_path = '//HILAB627_DS/database/database1.xlsx'
borrow_return_db_path = '//HILAB627_DS/database/database2.xlsx'
permissions_return_db_path = '//HILAB627_DS/database/permissions.xlsx'
REQUESTS_DB_PATH = '//HILAB627_DS/database/requests.xlsx' # 新的请求文件路径

# 全局变量，用于存储登录用户和权限信息
LOGGED_IN_USER = None
USER_PERMISSION = None
USER_NAME = None
# 在主程序启动时获取传递过来的用户名和权限
if __name__ == "__main__":
    # 添加调试信息
    print(f"DEBUG [managermentSystem_user.py]: Received sys.argv: {sys.argv}")
    print(f"DEBUG [managermentSystem_user.py]: len(sys.argv): {len(sys.argv)}")
    if len(sys.argv) == 4:
        LOGGED_IN_USER = sys.argv[1]
        USER_PERMISSION = sys.argv[2]
        USER_NAME = sys.argv[3]
        print(f"用户程序已启动。登录用户: {LOGGED_IN_USER}, 权限: {USER_PERMISSION}, 姓名: {USER_NAME}")

        # 加载数据库
        try:
            inventory_df = pd.read_excel(inventory_db_path, engine='openpyxl')
            borrow_return_df = pd.read_excel(borrow_return_db_path, engine='openpyxl')
            # Ensure data types are consistent.  This prevents later errors.
            inventory_df['数量'] = inventory_df['数量'].astype(int)
            # Add similar type checking for other relevant columns as needed.
        except FileNotFoundError:
            messagebox.showerror("错误", "没找到数据库，请检查文件路径是否正确。")
            #exit()
        except pd.errors.EmptyDataError:
            messagebox.showerror("错误", "数据库文件为空。")
            #exit()
        except Exception as e:
            messagebox.showerror("错误", f"加载数据库失败: {e}")
            #exit()
    elif len(sys.argv) == 1: # 如果直接运行 managermentSystem.py 而没有参数
        print("主程序直接启动（未传递用户信息和权限）。")
        # 此处可以添加逻辑，例如：
        # 1. 强制退出并提示需要通过登录界面启动
        messagebox.showerror("启动错误", "请通过登录界面启动程序。")
        # 2. 或者以默认用户/受限权限运行（不推荐，除非有明确场景）
    elif len(sys.argv) == 2:
        print(f"错误：传递给主应用程序的参数数量不正确。Actual len(sys.argv) was {len(sys.argv)}") # 提供更详细的错误信息
        l1 = sys.argv[0]
        l2 = sys.argv[1]
        print(f"DEBUG [managermentSystem_user.py]: l1: {l1}")
        print(f"DEBUG [managermentSystem_user.py]: l2: {l2}")
        messagebox.showerror("启动错误", "启动参数错误。")
        sys.exit()
    else:
        print(f"错误：传递给主应用程序的参数数量不正确。Actual len(sys.argv) was {len(sys.argv)}") # 提供更详细的错误信息
        messagebox.showerror("启动错误", "启动参数错误。")
        sys.exit()




def load_requests_db():
    try:
        if os.path.exists(REQUESTS_DB_PATH):
            return pd.read_excel(REQUESTS_DB_PATH, engine='openpyxl')
        else:
            # 如果文件不存在，创建一个空的DataFrame并保存
            df = pd.DataFrame(columns=['RequestID', 'Timestamp', 'Username', 'UserFullName', 
                                       'ItemCategory', 'ItemSubcategory', 'ItemName', 'Quantity', 
                                       'RequestType', 'AdminActionStatus', 'AdminRemarks', 
                                       'UserNotified', 'OriginalBorrowRequestID'])
            df.to_excel(REQUESTS_DB_PATH, index=False, engine='openpyxl')
            return df
    except Exception as e:
        messagebox.showerror("错误", f"加载请求数据库失败: {e}")
        # sys.exit() # 根据用户要求，严重错误时退出
        return None # 或者返回None，让调用者处理

def save_requests_db(df):
    try:
        df.to_excel(REQUESTS_DB_PATH, index=False, engine='openpyxl')
    except Exception as e:
        messagebox.showerror("错误", f"保存请求数据库失败: {e}")

# 新增：用户提交借用请求的UI处理函数
def request_borrow_item_ui():
    global req_category_entry, req_subcategory_entry, req_item_name_entry, req_quantity_entry # 确保可以访问这些UI元素
    try:
        category = req_category_entry.get().strip()
        subcategory = req_subcategory_entry.get().strip()
        item_name = req_item_name_entry.get().strip() # 这应该是物品的唯一标识，如 '物品备注'
        quantity_str = req_quantity_entry.get().strip()

        if not all([category, subcategory, item_name, quantity_str]):
            messagebox.showerror("输入错误", "所有字段均为必填项。")
            return
        
        quantity = int(quantity_str)
        if quantity <= 0:
            messagebox.showerror("输入错误", "数量必须为正整数。")
            return

        if submit_operation_request(LOGGED_IN_USER, USER_NAME, category, subcategory, item_name, quantity, 'Borrow'):
            req_category_entry.delete(0, tk.END)
            req_subcategory_entry.delete(0, tk.END)
            req_item_name_entry.delete(0, tk.END)
            req_quantity_entry.delete(0, tk.END)
    except ValueError:
        messagebox.showerror("输入错误", "数量必须是有效的数字。")
    except Exception as e:
        messagebox.showerror("操作失败", f"提交借用请求时发生错误: {e}")

def submit_operation_request(username, user_full_name, item_category, item_subcategory, item_name, quantity, request_type, original_borrow_id=None):
    """
    用户提交归还、交付、损坏等操作请求。
    """
    requests_df = load_requests_db()
    if requests_df is None:
        return False

    new_request = {
        'RequestID': str(uuid.uuid4()),
        'Timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'Username': username,
        'UserFullName': user_full_name,
        'ItemCategory': item_category,
        'ItemSubcategory': item_subcategory,
        'ItemName': item_name, # 确保这是物品的唯一标识符之一
        'Quantity': quantity,
        'RequestType': request_type, # 'Return', 'Deliver', 'Damage'
        'AdminActionStatus': 'Pending',
        'AdminRemarks': '',
        'UserNotified': False,
        'OriginalBorrowRequestID': original_borrow_id
    }
    # requests_df = requests_df.append(new_request, ignore_index=True) # <--- 旧代码
    requests_df = pd.concat([requests_df, pd.DataFrame([new_request])], ignore_index=True) # <--- 修改后的代码
    save_requests_db(requests_df)
    messagebox.showinfo("请求已提交", f"{request_type} 请求已提交给管理员审批。")
    return True

def request_return_item_ui():
    global req_category_entry, req_subcategory_entry, req_item_name_entry, req_quantity_entry
    try:
        category = req_category_entry.get().strip()
        subcategory = req_subcategory_entry.get().strip()
        item_name = req_item_name_entry.get().strip() 
        quantity_str = req_quantity_entry.get().strip()

        if not all([category, subcategory, item_name, quantity_str]):
            messagebox.showerror("输入错误", "所有字段均为必填项。")
            return
        
        quantity = int(quantity_str)
        if quantity <= 0:
            messagebox.showerror("输入错误", "数量必须为正整数。")
            return
        
        if submit_operation_request(LOGGED_IN_USER, USER_NAME, category, subcategory, item_name, quantity, 'Return'):
            req_category_entry.delete(0, tk.END)
            req_subcategory_entry.delete(0, tk.END)
            req_item_name_entry.delete(0, tk.END)
            req_quantity_entry.delete(0, tk.END)
    except ValueError:
        messagebox.showerror("输入错误", "数量必须是有效的数字。")
    except Exception as e:
        messagebox.showerror("操作失败", f"提交归还请求时发生错误: {e}")

def request_deliver_item_ui():
    global req_category_entry, req_subcategory_entry, req_item_name_entry, req_quantity_entry
    try:
        category = req_category_entry.get().strip()
        subcategory = req_subcategory_entry.get().strip()
        item_name = req_item_name_entry.get().strip()
        quantity_str = req_quantity_entry.get().strip()

        if not all([category, subcategory, item_name, quantity_str]):
            messagebox.showerror("输入错误", "所有字段均为必填项。")
            return
        
        quantity = int(quantity_str)
        if quantity <= 0:
            messagebox.showerror("输入错误", "数量必须为正整数。")
            return

        if submit_operation_request(LOGGED_IN_USER, USER_NAME, category, subcategory, item_name, quantity, 'Deliver'):
            req_category_entry.delete(0, tk.END)
            req_subcategory_entry.delete(0, tk.END)
            req_item_name_entry.delete(0, tk.END)
            req_quantity_entry.delete(0, tk.END)
    except ValueError:
        messagebox.showerror("输入错误", "数量必须是有效的数字。")
    except Exception as e:
        messagebox.showerror("操作失败", f"提交交付请求时发生错误: {e}")


def request_damage_item_ui():
    global req_category_entry, req_subcategory_entry, req_item_name_entry, req_quantity_entry
    try:
        category = req_category_entry.get().strip()
        subcategory = req_subcategory_entry.get().strip()
        item_name = req_item_name_entry.get().strip()
        quantity_str = req_quantity_entry.get().strip()

        if not all([category, subcategory, item_name, quantity_str]):
            messagebox.showerror("输入错误", "所有字段均为必填项。")
            return
        
        quantity = int(quantity_str)
        if quantity <= 0:
            messagebox.showerror("输入错误", "数量必须为正整数。")
            return
        
        if submit_operation_request(LOGGED_IN_USER, USER_NAME, category, subcategory, item_name, quantity, 'Damage'):
            req_category_entry.delete(0, tk.END)
            req_subcategory_entry.delete(0, tk.END)
            req_item_name_entry.delete(0, tk.END)
            req_quantity_entry.delete(0, tk.END)
    except ValueError:
        messagebox.showerror("输入错误", "数量必须是有效的数字。")
    except Exception as e:
        messagebox.showerror("操作失败", f"提交损坏请求时发生错误: {e}")
        
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

def populate_description_menu(location_filter, cabinet_filter):
    """Populates the description_menu based on the selected category, subcategory, location and cabinet."""
    # 获取当前选择的大类和小类 (来自用于添加入库的输入框)
    current_category = category_entry.get().strip().lower()
    current_subcategory = subcategory_entry.get().strip().lower()

    # 复制一份 DataFrame 以免修改原始数据
    filtered_df = inventory_df.copy()

    # 1. 根据大类筛选
    if current_category:
        filtered_df = filtered_df[filtered_df['大类名称'].str.lower() == current_category]

    # 2. 根据小类筛选 (在已按大类筛选的基础上)
    if current_subcategory:
        filtered_df = filtered_df[filtered_df['小类名称'].str.lower() == current_subcategory]

    # 3. 根据位置和柜子筛选描述
    if location_filter and cabinet_filter:
        try:
            # 确保 '存放位置' 列是字符串类型，便于处理
            # 从已经按 category/subcategory 筛选过的 filtered_df 中提取存放位置信息
            parts = filtered_df['存放位置'].astype(str).str.split('-', n=2, expand=True)
            
            # expand=True 会创建新的列，如果分割数不足，则后续列为 None
            # 我们需要确保 parts 有足够的列，或者在访问前检查
            part0_series = parts[0] if 0 in parts.columns else pd.Series(dtype='str')
            part1_series = parts[1] if 1 in parts.columns else pd.Series(dtype='str')
            part2_series = parts[2] if 2 in parts.columns else pd.Series(dtype='str')

            # 创建筛选条件：第一部分匹配 location_filter 且 第二部分匹配 cabinet_filter
            mask = (part0_series.str.lower() == location_filter.lower()) & \
                   (part1_series.str.lower() == cabinet_filter.lower())
            
            descriptions_series = part2_series[mask].dropna()
            
            if not descriptions_series.empty:
                unique_descriptions = descriptions_series.unique()
                cleaned_descriptions = [clean_text(str(c)) for c in unique_descriptions if str(c).strip()]
            else:
                cleaned_descriptions = []
            
            description_menu['values'] = cleaned_descriptions
        except Exception as e:
            messagebox.showerror("错误", f"更新描述选项时出错: {e}")
            cleaned_descriptions = [] # 出错时清空
            description_menu['values'] = cleaned_descriptions
    else:
        # 如果位置或柜子未选择，则清空描述选项
        cleaned_descriptions = []
        description_menu['values'] = cleaned_descriptions
    
    description_menu.set('')  # 清空下拉菜单的当前选定值
    description_entry.delete(0, tk.END)  # 清空关联的输入框


def on_location_menu_select(event):
    selected_location = location_choice.get()
    populate_cabinet_menu(selected_location)
    cabinet_entry.delete(0, tk.END)  # Clear cabinet_entry when location changes
    # 当位置改变时，也应该清空描述选项，因为柜子选项会变，进而影响描述
    description_menu['values'] = []
    description_menu.set('')
    description_entry.delete(0, tk.END)

def on_cabinet_menu_select(event):
    selected_cabinet = cabinet_menu.get()
    selected_location = location_choice.get()
    # 使用选择的 location 和 cabinet 更新 description 菜单
    # populate_description_menu 会自动使用当前的 category_entry 和 subcategory_entry 值
    populate_description_menu(selected_location, selected_cabinet)
    cabinet_entry.delete(0, tk.END)
    cabinet_entry.insert(0, selected_cabinet)
    # 当柜子改变时，清空已选的描述，因为描述选项已更新
    description_menu.set('') 
    description_entry.delete(0, tk.END)

def on_description_menu_select(event):
    selected_description = description_menu.get()
    description_entry.delete(0, tk.END)
    description_entry.insert(0, selected_description)

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

def update_description_options(*args):
    selected_location = location_choice.get().strip()
    if selected_location:
        descriptions = inventory_df[inventory_df['存放位置'].str.startswith(selected_location)]['存放位置'].apply(lambda x: x.split('-')[2]).unique()
        description_menu['menu'].delete(0, 'end')
        for description in descriptions:
            description_menu['menu'].add_command(label=description, command=tk._setit(description_number, description, set_description_from_dropdown))

def set_category_from_dropdown(*args):
    category_entry.delete(0, tk.END)
    category_entry.insert(0, category_choice.get())
    update_subcategory_options()
    # 当大类改变后，也需要更新描述菜单
    populate_description_menu(location_choice.get(), cabinet_number.get())

def set_subcategory_from_dropdown(*args):
    subcategory_entry.delete(0, tk.END)
    subcategory_entry.insert(0, subcategory_choice.get())
    # 当小类改变后，也需要更新描述菜单
    populate_description_menu(location_choice.get(), cabinet_number.get())
    
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
    
def set_description_from_dropdown(*args):
    description_entry.delete(0, tk.END)
    description_entry.insert(0, description_number.get())

def set_status_from_dropdown(*args):
    status_entry.delete(0, tk.END)
    status_entry.insert(0, status_choice.get())

def set_remark_from_dropdown(*args):
    remark_entry.delete(0, tk.END)
    remark_entry.insert(0, remark_choice.get())

# 添加新的函数
def set_borrower_from_dropdown(*args):
    borrower_entry.delete(0, tk.END)
    borrower_entry.insert(0, borrower_choice.get())

# 添加新的函数
def set_search_item_from_dropdown(*args):
    search_item_entry.delete(0, tk.END)
    search_item_entry.insert(0, search_item_choice.get())
    update_search_subitem_options()

def set_search_subitem_from_dropdown(*args):
    search_subitem_entry.delete(0, tk.END)
    search_subitem_entry.insert(0, search_subitem_choice.get())

def update_search_subitem_options(*args):
    selected_item = search_item_entry.get().strip()
    if selected_item:
        try:
            subitems = inventory_df[inventory_df['大类名称'] == selected_item]['小类名称'].unique()
            # 注意：下面这行在您提供的代码中似乎被截断了 (search_subitem_m...)
            # 请确保它是完整的，例如：search_subitem_menu['menu'].delete(0, 'end')
            search_subitem_menu['menu'].delete(0, 'end') # 假设这是正确的代码，如果不是请根据您的原意修改
            if 'values' in search_subitem_menu.config(): # 检查是否为Combobox或类似控件
                 search_subitem_menu['values'] = []

            for subitem in subitems:
                if isinstance(search_subitem_menu, ttk.Combobox):
                    current_values = list(search_subitem_menu['values'])
                    current_values.append(subitem)
                    search_subitem_menu['values'] = current_values
                elif isinstance(search_subitem_menu, tk.OptionMenu): # 假设是OptionMenu
                     search_subitem_menu['menu'].add_command(label=subitem, command=tk._setit(search_subitem_choice, subitem, set_search_subitem_from_dropdown))
            
            if isinstance(search_subitem_menu, ttk.Combobox) and subitems.size > 0 :
                search_subitem_menu.current(0) # 默认选择第一个
            elif isinstance(search_subitem_menu, tk.OptionMenu) and subitems.size > 0:
                 search_subitem_choice.set(subitems[0]) # 默认选择第一个
            else: # 如果没有子项，清空
                if isinstance(search_subitem_menu, ttk.Combobox):
                    search_subitem_menu.set('')
                elif isinstance(search_subitem_menu, tk.OptionMenu):
                    search_subitem_choice.set('')

        except Exception as e:
            messagebox.showerror("错误", f"更新小类查询选项时出错: {e}")
            if 'values' in search_subitem_menu.config():
                search_subitem_menu['values'] = []
            search_subitem_menu.set('')
    else:
        if 'values' in search_subitem_menu.config():
            search_subitem_menu['values'] = []
        search_subitem_menu.set('')
        search_subitem_entry.delete(0, tk.END)


def search_item_records():
    item = search_item_entry.get().strip()
    subitem = search_subitem_entry.get().strip()
    
    if not item or not subitem:
        messagebox.showerror("错误", "请选择要查询的物品类别和子类别")
        return
        
    search_window = tk.Toplevel(root)
    search_window.title(f"{item}-{subitem}的借还记录与库存") # 更新窗口标题
    text = scrolledtext.ScrolledText(search_window)
    text.pack(fill=tk.BOTH, expand=True)

    # 查询当前库存总量
    try:
        # 筛选库存中匹配的物品，不区分备注和物品备注，只看大类和小类
        # 假设 '大类名称' 和 '小类名称' 在 inventory_df 中是准确的
        current_inventory = inventory_df[
            (inventory_df['大类名称'].str.lower() == item.lower()) &
            (inventory_df['小类名称'].str.lower() == subitem.lower())
        ]
        total_quantity_in_stock = current_inventory['数量'].sum()
        text.insert(tk.END, f"物品【{item} - {subitem}】当前仓库总剩余数量: {total_quantity_in_stock}\n")

        # 计算并显示备注为'好的'和'坏的'物品数量
        if not current_inventory.empty:
            good_items_quantity = current_inventory[current_inventory['备注'].str.lower() == '好的']['数量'].sum()
            bad_items_quantity = current_inventory[current_inventory['备注'].str.lower() == '坏的']['数量'].sum()
            text.insert(tk.END, f"  其中，备注为【好的】数量: {good_items_quantity}\n")
            text.insert(tk.END, f"  其中，备注为【坏的】数量: {bad_items_quantity}\n")

            # 显示物品的存放位置
            storage_locations = current_inventory['存放位置'].unique()
            if len(storage_locations) > 0:
                text.insert(tk.END, f"当前物品【{item} - {subitem}】的存放位置有:\n")
                for loc in storage_locations:
                    if pd.notna(loc) and str(loc).strip():
                        text.insert(tk.END, f"- {loc}\n")
                    else:
                        text.insert(tk.END, f"- (未指定位置)\n")
            else:
                text.insert(tk.END, "未找到该物品的存放位置信息。\n")
        else:
            text.insert(tk.END, "未在库存中找到该物品，无法显示备注和存放位置。\n")
        

        text.insert(tk.END, "---------------------------------------------------\n")
        text.insert(tk.END, "借还记录详情:\n")
    except Exception as e:
        text.insert(tk.END, f"查询库存数量、备注或存放位置时发生错误: {e}\n")
        text.insert(tk.END, "---------------------------------------------------\n")
        text.insert(tk.END, "借还记录详情:\n")
    
    # 筛选指定物品的借还记录
    item_records = borrow_return_df[
        (borrow_return_df['借出物品大类名称'] == item) & 
        (borrow_return_df['借出物品小类名称'] == subitem)
    ]
    
    if item_records.empty:
        text.insert(tk.END, f"未找到 {item}-{subitem} 的借还记录\n")
    else:
        for index, row in item_records.iterrows():
            date_val = row.get('日期', '无日期')
            borrower_val = row.get('保管人员', '无')
            category_val = row.get('借出物品大类名称', '无')
            subcategory_val = row.get('借出物品小类名称', '无')
            quantity_val = row.get('借出物品数量', 0)
            status_val = row.get('物品状态', '无')
            remark_val = row.get('备注', '无') # 获取备注信息
            
            text.insert(tk.END, 
                        f"日期: {date_val}, 人员: {borrower_val}, "
                        f"物品: {category_val} - {subcategory_val}, "
                        f"数量: {quantity_val}, 状态: {status_val}, 备注: {remark_val}\n")

### 主体功能函数
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
        text.insert(tk.END, f"{row['大类名称']} - {row['小类名称']}({row['备注']}): {row['数量']} 放在 {row['存放位置']}\n")

def view_borrow_return():
    borrow_return_window = tk.Toplevel(root)
    borrow_return_window.title("Borrow/Return Records")
    text = tk.Text(borrow_return_window)
    text.pack()
    for index, row in borrow_return_df.iterrows():
        text.insert(tk.END, f"人员: {row['保管人员']}, {row['借出物品大类名称']} - {row['借出物品小类名称']}: {row['借出物品数量']} ({row['物品状态']})\n")

def add_inventory_item():
    global inventory_df
    # 权限检查
    if USER_PERMISSION != '管理员':
        messagebox.showerror("权限错误", "您没有权限执行此操作。")
        return
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
            # messagebox.showinfo("成功", "物品已存在，数量已更新。")
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
            # messagebox.showinfo("成功", "物品不存在，物品已添加!")

        # Save updated DataFrame to Excel
        inventory_df.to_excel(inventory_db_path, index=False, engine='openpyxl')

        # Re-read the updated inventory database
        try:
            inventory_df = pd.read_excel(inventory_db_path, engine='openpyxl')
            # Ensure data types are consistent after reloading.
            if '数量' in inventory_df.columns:
                 inventory_df['数量'] = inventory_df['数量'].astype(int)
            # Potentially update UI elements that depend on inventory_df if necessary
            # For example, re-populating dropdowns:
            # populate_main_category_options() # Assuming such a function exists or is needed
            # populate_all_dropdowns_from_inventory() # A more generic function
            messagebox.showinfo("成功", "物品已添加/更新，并且数据库已刷新。") # Updated success message
        except Exception as e:
            messagebox.showerror("错误", f"成功保存物品，但刷新数据库时出错: {e}")
    except ValueError as e:
        messagebox.showerror("Error", f"Invalid input: {e}")
    except Exception as e:
        messagebox.showerror("Error", f"An unexpected error occurred: {e}")

def update_databases():
    try:
        borrower = USER_NAME
        if not borrower:
            messagebox.showerror("错误", "无法获取当前用户名，请重新登录。")
            return
        category = category_entry2.get()
        subcategory = subcategory_entry2.get()
        quantity = int(quantity_entry2.get())
        status = status_entry.get()
        remark = remark_entry2.get()
        # 获取当前日期并格式化为YYYYMMDD格式
        current_date = datetime.now().strftime('%Y%m%d')

        if quantity <= 0:
            raise ValueError("不要乱写负数！csn你！.")

        # Update borrow/return database
        new_borrow_entry = {
            '借出物品大类名称': category,
            '借出物品小类名称': subcategory,
            '借出物品数量': quantity,
            '保管人员': borrower,
            '物品状态': status,
            '备注': remark,
            '日期': current_date  # 添加日期字段
        }
        global borrow_return_df
        borrow_return_df = borrow_return_df.append(new_borrow_entry, ignore_index=True)
        borrow_return_df.to_excel(borrow_return_db_path, index=False, engine='openpyxl')

        # Update inventory database
        inventory_index = inventory_df[(inventory_df['大类名称'] == category) & (inventory_df['小类名称'] == subcategory) & (inventory_df['备注'] == remark)].index

        # if not inventory_index.empty:
        #     if status == '借出':
        #         inventory_df.at[inventory_index[0], '数量'] -= quantity
        #     elif status in ['归还', '采购']:
        #         inventory_df.at[inventory_index[0], '数量'] += quantity
        #     elif status in ['交付', '损坏']:
        #         inventory_df.at[inventory_index[0], '数量'] -= quantity

        # 0.3.1版本，注释掉用户的其他权限，只允许用户借出自己的物品。后续将结合仓库增加更多功能
        if not inventory_index.empty:
            if status == '借出':
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

        # Filter out items with current_count of 0
        current_items = ", ".join([
            format_item(count, category, subcategory)
            for (category, subcategory), count in current_count.items()
            if count >= 1  # Only include items with count >= 1
        ])
        delivered_items = ", ".join([
            format_item(count, category, subcategory)
            for (category, subcategory), count in delivered_count.items()
        ])
        damaged_items = ", ".join([
            format_item(count, category, subcategory)
            for (category, subcategory), count in damaged_count.items()
        ])

        text.insert(tk.END, f"当前名下还有：{current_items}。\n")
        text.insert(tk.END, f"交付：{delivered_items}。\n")
        text.insert(tk.END, f"损坏：{damaged_items}。\n")

def view_personal_records(borrower_name):
    if not borrower_name:
        messagebox.showerror("错误", "请输入要查询的人员姓名")
        return
        
    personal_window = tk.Toplevel(root)
    personal_window.title(f"{borrower_name}的借还记录")
    text = scrolledtext.ScrolledText(personal_window)
    text.pack(fill=tk.BOTH, expand=True)
    
    # 筛选指定人员的记录
    personal_records = borrow_return_df[borrow_return_df['保管人员'] == borrower_name]
    
    if personal_records.empty:
        text.insert(tk.END, f"未找到 {borrower_name} 的借还记录\n")
    else:
        for index, row in personal_records.iterrows():
            # 获取日期信息，如果没有日期字段则显示"未记录"
            date = row['日期'] if '日期' in row else "未记录"
            text.insert(tk.END, 
                      f"日期: {date}, 人员: {row['保管人员']}, "
                      f"{row['借出物品大类名称']} - {row['借出物品小类名称']}: "
                      f"{row['借出物品数量']} ({row['物品状态']})\n")

# 使用说明
def show_instructions():
    instructions_window = tk.Toplevel(root)
    instructions_window.title("仓库管理系统使用说明")
    instructions_window.geometry("700x600")
    
    # 主框架
    main_frame = ttk.Frame(instructions_window)
    main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
    
    # 创建带滚动条的Canvas
    canvas = tk.Canvas(main_frame)
    scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
    scrollable_frame = ttk.Frame(canvas)
    
    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(
            scrollregion=canvas.bbox("all")
        )
    )
    
    # 创建Canvas窗口
    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)
    
    # 绑定鼠标滚轮事件
    def _on_mousewheel(event):
        canvas.yview_scroll(int(-1*(event.delta/120)), "units")
    
    canvas.bind_all("<MouseWheel>", _on_mousewheel)
    
    # 标题样式
    title_style = ttk.Style()
    title_style.configure("Title.TLabel", font=('微软雅黑', 12, 'bold'), foreground='#333333')
    
    # 内容样式
    content_style = ttk.Style()
    content_style.configure("Content.TLabel", font=('微软雅黑', 10), foreground='#555555')
    
    # 警告样式
    warning_style = ttk.Style()
    warning_style.configure("Warning.TLabel", font=('微软雅黑', 10, 'bold'), foreground='red')
    
    # 添加标题
    ttk.Label(scrollable_frame, 
             text="仓库管理系统使用指南（请连接SYSU-Hilab的wifi）", 
             style="Title.TLabel").pack(pady=(0, 15))
    
    # 添加说明内容
    sections = [
        {
            "title": "1. 借还管理页面",
            "items": [
                ("1.1 基本操作", "其他同上。操作描述符的时候没有下拉菜单，因为需要处理损坏情况。需要留心，所以没有涉及下拉。", "black"),
                ("1.2 完整流程", "", "black"),
                ("1.2.1 借出物品", "操作借还人员A借出物品。", "black"),
                ("1.2.2 归还物品", "操作借还人员A归还物品。", "black"),
                ("1.2.3 损坏物品", "操作借还人员B借出物品，但是B损坏物品。", "black"),
                ("1.2.4 损坏处理", "借还借还人员B损坏物品，描述符填好的（因为是损坏了好的物品）。", "black"),
                ("1.2.5 交付物品", "操作借还人员C借出物品，然后C交付甲方。", "black"),
                ("1.2.6 交付处理", "借还借还人员C交付物品，描述符填具体交付的（一般不会交付坏的）。", "black"),
                ("1.3 采购功能", "后续将会和采购流程结合。", "black"),
                ("1.3.1 采购功能", "如果发现在大类别和小类别中没有找到对应的物品，那么就联系管理员入库。", "black"),
                ("1.3.2 采购功能", "如果能够直接通过下拉列表选择，那么直接选择采购就相当于入库了。", "black"),
            ]
        },
        {
            "title": "2. 查询功能页面",
            "items": [
                ("2.1 按人员查询", "可以查到某个人名下的所有物品，在手上的，损坏的，交付的等。", "black"),
                ("2.2 数据库信息", "可以直接读取数据库信息，也即是'查看仓库总表物品详细信息'和'查看仓库总表借还记录'，功能和库存管理中的'点击查找物品信息'是一样的，更完善。", "black"),
                ("2.3 按物品查询", "可以查找具体某个物品的详细位置和数量，包括借出归还损坏的人员信息。和库存管理中的'点击查找物品信息是一样的，更完善", "black")
            ]
        }
    ]
    
    for section in sections:
        # 添加章节标题
        ttk.Label(scrollable_frame, 
                 text=section["title"], 
                 style="Title.TLabel").pack(pady=(10, 5), anchor='w')
        
        # 添加章节内容
        for item in section["items"]:
            style = "Warning.TLabel" if item[2] == "red" else "Content.TLabel"
            ttk.Label(scrollable_frame, 
                     text=f"  {item[0]}: {item[1]}", 
                     style=style,
                     wraplength=650,
                     justify=tk.LEFT).pack(pady=2, anchor='w')
    
    # 添加底部说明
    ttk.Label(scrollable_frame, 
             text="\n如有任何问题，请联系系统管理员", 
             style="Title.TLabel").pack(pady=(20, 5))
    
    # 布局滚动区域
    canvas.pack(side="left", fill=tk.BOTH, expand=True)
    scrollbar.pack(side="right", fill="y")


# --- 处理管理员审批后的操作 ---
def process_approved_action(request_id, request_type, category, subcategory, item_name, quantity, username_of_requester):
    global inventory_df, borrow_return_df
    try:
        # 标记：这里 item_name 应该是能唯一识别 database1.xlsx 中物品的关键信息
        # 例如 '物品名称' 或 '描述符'
        
        # Flow 2: 用户归还物品 (管理员同意后)
        if request_type == 'Return':
            # 1. 增加 database1.xlsx 中对应物品的数量
            match_condition_inv = (inventory_df['大类名称'] == category) & \
                                  (inventory_df['小类名称'] == subcategory) & \
                                  (inventory_df['物品名称'] == item_name) # 假设 '物品名称'
            if not inventory_df[match_condition_inv].empty:
                item_idx_inv = inventory_df[match_condition_inv].index[0]
                inventory_df.loc[item_idx_inv, '数量'] += quantity
            else:
                messagebox.showerror("错误", f"归还失败：未在库存中找到物品 {item_name}")
                return

            # 2. 在 database2.xlsx 中添加归还记录或更新状态
            #    减少该用户名下的这个物品对应的数量 (通过添加'归还'记录实现)
            new_return_entry = {
                '借出物品大类名称': category,
                '借出物品小类名称': subcategory,
                '借出物品名称': item_name,
                '借出物品数量': quantity, # 归还的数量
                '保管人员': username_of_requester, # 这里的username_of_requester应为UserFullName
                '用户名': LOGGED_IN_USER, # 或者从request中获取请求者用户名
                '物品状态': '归还',
                '备注': f'管理员批准归还，RequestID: {request_id}',
                '日期': datetime.now().strftime('%Y%m%d'),
                'RequestID': request_id # 使用原始请求ID
            }
            borrow_return_df = borrow_return_df.append(new_return_entry, ignore_index=True)
            messagebox.showinfo("成功", "物品归还成功！")

        # Flow 3: 用户直接从仓库交付 (管理员同意后)
        elif request_type == 'Deliver':
            # 1. 减少 database1.xlsx 中对应物品的数量
            match_condition_inv = (inventory_df['大类名称'] == category) & \
                                  (inventory_df['小类名称'] == subcategory) & \
                                  (inventory_df['物品名称'] == item_name)
            if not inventory_df[match_condition_inv].empty:
                item_idx_inv = inventory_df[match_condition_inv].index[0]
                if inventory_df.loc[item_idx_inv, '数量'] >= quantity:
                    inventory_df.loc[item_idx_inv, '数量'] -= quantity
                else:
                    messagebox.showerror("错误", f"交付失败：库存不足 {item_name}")
                    return
            else:
                messagebox.showerror("错误", f"交付失败：未在库存中找到物品 {item_name}")
                return
            
            # 2. 在 database2.xlsx 中记录交付
            new_deliver_entry = {
                '借出物品大类名称': category,
                '借出物品小类名称': subcategory,
                '借出物品名称': item_name,
                '借出物品数量': quantity, # 交付的数量
                '保管人员': username_of_requester,
                '用户名': LOGGED_IN_USER,
                '物品状态': '交付',
                '备注': f'管理员批准直接交付，RequestID: {request_id}',
                '日期': datetime.now().strftime('%Y%m%d'),
                'RequestID': request_id
            }
            borrow_return_df = borrow_return_df.append(new_deliver_entry, ignore_index=True)
            messagebox.showinfo("成功", "物品交付成功！")
        
        # Flow 4: 用户借出后再交付 (管理员同意后)
        #   - 仓库数量在借出时已减少，此处无需再动 database1.xlsx
        #   - 主要是在 database2.xlsx 中将原 '借出' 记录的状态更新或新增 '交付' 记录
        #   - 用户的描述是 "直接将记录在该用户名下的这个物品属性改为交付，该用户名下的这个物品拥有数量减去对应数量，交付的物品数量加上对应数量"
        #   - 这意味着我们需要找到原始的 '借出' 记录，或者更简单地是添加一条新的 '交付' 记录，
        #     并在计算用户持有量时，'交付' 同样视为减少持有。
        #     为了审计和清晰，建议添加新的 '交付' 记录。
        #     如果需要严格对应原借出记录，则需要 `OriginalBorrowRequestID`。
        #     当前实现与Flow 3类似，只是database1.xlsx不操作。
        #     若要区分，可以在submit_operation_request时传递一个标志，或根据物品是否已在用户借出名下判断。
        #     为简化，此处假设 'Deliver' 请求总是指从仓库直接出，或用户选择已借出的物品进行交付时，
        #     UI会传递正确的上下文。如果物品已借出，则不操作database1。
        #     一个更健壮的做法是，交付请求应指明是“新交付”还是“从已借出转交付”。
        #     当前代码按“新交付”处理，如果物品已借出，则需要调整逻辑，例如不减库存。

        # Flow 5: 损坏操作 (管理员同意后)
        elif request_type == 'Damage':
            # 逻辑类似交付，减少库存 (如果是直接报损)，并在 database2.xlsx 中记录
            # 1. 减少 database1.xlsx 中对应物品的数量 (如果物品在库房)
            #    如果物品是用户已借出的，则不操作 database1.xlsx
            #    需要UI传递上下文或在请求中包含此信息
            match_condition_inv = (inventory_df['大类名称'] == category) & \
                                  (inventory_df['小类名称'] == subcategory) & \
                                  (inventory_df['物品名称'] == item_name)
            if not inventory_df[match_condition_inv].empty:
                item_idx_inv = inventory_df[match_condition_inv].index[0]
                if inventory_df.loc[item_idx_inv, '数量'] >= quantity: # 假设损坏的是库存品
                    inventory_df.loc[item_idx_inv, '数量'] -= quantity
                # else: # 如果是用户已借出的物品报损，则不应出现库存不足
                #    pass 
            # else:
                # pass # 如果是用户已借出的物品报损，库存中可能没有（或不应操作）
            
            # 2. 在 database2.xlsx 中记录损坏
            new_damage_entry = {
                '借出物品大类名称': category,
                '借出物品小类名称': subcategory,
                '借出物品名称': item_name,
                '借出物品数量': quantity, # 损坏的数量
                '保管人员': username_of_requester,
                '用户名': LOGGED_IN_USER,
                '物品状态': '损坏',
                '备注': f'管理员批准报损，RequestID: {request_id}',
                '日期': datetime.now().strftime('%Y%m%d'),
                'RequestID': request_id
            }
            borrow_return_df = borrow_return_df.append(new_damage_entry, ignore_index=True)
            messagebox.showinfo("成功", "物品损坏记录成功！")

        save_databases() # 保存所有更改

        # 更新请求状态为已通知用户
        requests_df = load_requests_db()
        if requests_df is not None:
            req_idx = requests_df[requests_df['RequestID'] == request_id].index
            if not req_idx.empty:
                requests_df.loc[req_idx, 'UserNotified'] = True
                save_requests_db(requests_df)

    except Exception as e:
        messagebox.showerror("错误", f"处理已批准的 {request_type} 操作时发生错误: {e}")

def check_pending_requests_status():
    """
    定期检查用户提交的请求是否有管理员的审批结果。
    """
    requests_df = load_requests_db()
    if requests_df is None:
        return

    user_pending_requests = requests_df[
        (requests_df['Username'] == LOGGED_IN_USER) & 
        (requests_df['AdminActionStatus'].isin(['Approved', 'Denied'])) & 
        (requests_df['UserNotified'] == False)
    ]

    for index, req in user_pending_requests.iterrows():
        if req['AdminActionStatus'] == 'Approved':
            messagebox.showinfo("请求批准", f"您的 {req['RequestType']} 请求 (ID: {req['RequestID']}) 已被管理员批准。")
            # 调用实际处理函数
            process_approved_action(req['RequestID'], req['RequestType'], req['ItemCategory'], 
                                    req['ItemSubcategory'], req['ItemName'], req['Quantity'], 
                                    req['UserFullName'])
        elif req['AdminActionStatus'] == 'Denied':
            messagebox.showwarning("请求被拒", f"您的 {req['RequestType']} 请求 (ID: {req['RequestID']}) 已被管理员拒绝。备注: {req['AdminRemarks']}")
            # 更新 UserNotified 状态
            idx = requests_df[requests_df['RequestID'] == req['RequestID']].index
            requests_df.loc[idx, 'UserNotified'] = True
    
    if not user_pending_requests.empty:
        save_requests_db(requests_df) # 保存 UserNotified 的更改


# 逐级搜索下拉菜单添加更新函数
def update_req_subcategory_options(*args):
    selected_category = req_category_entry.get().strip()
    if selected_category:
        subcategories = inventory_df[inventory_df['大类名称'] == selected_category]['小类名称'].unique()
        req_subcategory_menu['menu'].delete(0, 'end')
        for subcategory in subcategories:
            req_subcategory_menu['menu'].add_command(label=subcategory, 
                command=tk._setit(req_subcategory_choice, subcategory, set_req_subcategory_from_dropdown))

def update_req_description_options(*args):
    selected_category = req_category_entry.get().strip()
    selected_subcategory = req_subcategory_entry.get().strip()
    
    if selected_category and selected_subcategory:
        filtered_df = inventory_df[
            (inventory_df['大类名称'] == selected_category) & 
            (inventory_df['小类名称'] == selected_subcategory)
        ]
        descriptions = filtered_df['备注'].unique()
        req_description_menu['menu'].delete(0, 'end')
        for description in descriptions:
            req_description_menu['menu'].add_command(label=description, 
                command=tk._setit(req_description_choice, description, set_req_description_from_dropdown))

def set_req_category_from_dropdown(*args):
    req_category_entry.delete(0, tk.END)
    req_category_entry.insert(0, req_category_choice.get())
    update_req_subcategory_options()

def set_req_subcategory_from_dropdown(*args):
    req_subcategory_entry.delete(0, tk.END)
    req_subcategory_entry.insert(0, req_subcategory_choice.get())
    update_req_description_options()

def set_req_description_from_dropdown(*args):
    req_item_name_entry.delete(0, tk.END)
    req_item_name_entry.insert(0, req_description_choice.get())

# 确保有一个 save_databases() 函数
def save_databases():
    global inventory_df, borrow_return_df
    try:
        inventory_df.to_excel(inventory_db_path, index=False, engine='openpyxl')
        borrow_return_df.to_excel(borrow_return_db_path, index=False, engine='openpyxl')
        # messagebox.showinfo("成功", "数据库已保存。") # 可选：频繁保存时此提示可能过多
    except Exception as e:
        messagebox.showerror("错误", f"保存数据库失败: {e}")
        # sys.exit() # 根据用户要求，严重错误时退出
# 新增：周期性检查用户请求状态的函数
def check_pending_requests_status_periodic():
    """
    周期性检查当前登录用户的请求状态，并在有更新时通知用户。
    """
    global root # 需要访问全局的 root 窗口对象以进行 rescheduling
    
    if LOGGED_IN_USER is None:
        print("DEBUG: LOGGED_IN_USER is None, skipping periodic check.")
        # 即使没有登录用户，也应该重新安排下一次检查，以防后续登录
        if 'root' in globals() and root.winfo_exists(): # 检查 root 是否已定义且窗口存在
            root.after(300000, check_pending_requests_status_periodic) # 5分钟后再次检查
        return

    try:
        requests_df = load_requests_db()
        if requests_df is None or requests_df.empty:
            # print(f"DEBUG: No requests data found for user {LOGGED_IN_USER}.")
            if 'root' in globals() and root.winfo_exists():
                root.after(300000, check_pending_requests_status_periodic) # 重新安排
            return

        # 筛选当前用户未被通知的、且管理员已处理的请求
        # AdminActionStatus 可能的值: 'Pending', 'Approved', 'Rejected', 'Completed' (根据您的系统设计)
        # UserNotified 应该是布尔值 True/False
        user_requests_to_notify = requests_df[
            (requests_df['Username'] == LOGGED_IN_USER) &
            (requests_df['UserNotified'] == False) &
            (requests_df['AdminActionStatus'] != 'Pending') # 管理员已处理
        ]

        if not user_requests_to_notify.empty:
            notification_messages = []
            for index, row in user_requests_to_notify.iterrows():
                message = (
                    f"您的请求 '{row['RequestType']}' (物品: {row.get('ItemName', 'N/A')}, "
                    f"数量: {row.get('Quantity', 'N/A')}) "
                    f"已被管理员处理，状态: {row['AdminActionStatus']}."
                )
                if pd.notna(row['AdminRemarks']) and str(row['AdminRemarks']).strip():
                    message += f" 管理员备注: {row['AdminRemarks']}"
                
                notification_messages.append(message)
                
                # 更新为已通知
                requests_df.loc[index, 'UserNotified'] = True
            
            if notification_messages:
                messagebox.showinfo("请求状态更新", "\n\n".join(notification_messages))
                save_requests_db(requests_df) # 保存更新后的 UserNotified 状态

    except Exception as e:
        print(f"Error in check_pending_requests_status_periodic: {e}")
        # 不在此处显示 messagebox，避免过多弹窗，错误记录在控制台即可

    finally:
        # 无论如何，都重新安排下一次检查
        # 确保 root 仍然存在 (例如，用户可能已经关闭了窗口)
        if 'root' in globals() and root.winfo_exists():
            root.after(300000, check_pending_requests_status_periodic) # 例如，每5分钟检查一次 (300000毫秒)

global root, category_choice, subcategory_choice, location_choice, cabinet_number, description_number, status_choice, remark_choice, borrower_choice, search_item_choice, search_subitem_choice
global category_entry, subcategory_entry, location_entry, cabinet_entry, description_entry, quantity_entry, status_entry, remark_entry, borrower_entry, search_item_entry, search_subitem_entry
global subcategory_menu, cabinet_menu, description_menu, search_subitem_menu
# 新增请求操作相关的UI元素
global req_category_entry, req_subcategory_entry, req_item_name_entry, req_quantity_entry

# GUI setup
root = tk.Tk()
root.title("仓库管理系统-v0.2")

# 初始化变量
location_choice = tk.StringVar(value="627")
cabinet_number = tk.StringVar(value="")
description_number = tk.StringVar(value="")
category_choice = tk.StringVar(value="")
subcategory_choice = tk.StringVar(value="")
status_choice = tk.StringVar(value="借出")
category_choice2 = tk.StringVar()
subcategory_choice2 = tk.StringVar()
detailed_description = tk.StringVar()
remark_choice = tk.StringVar()
search_item_choice = tk.StringVar()
search_subitem_choice = tk.StringVar()
borrower_choice = tk.StringVar()
# 设置样式
style = ttk.Style()
style.configure('Title.TLabel', font=('Arial', 13, 'bold'))
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
    text="Hilab仓库管理系统", 
    style='Title.TLabel',
    anchor='center')
title_label.pack(fill=tk.X)

# 新增：显示用户和权限信息
user_info_frame = ttk.Frame(main_frame) # 创建一个新的框架来容纳用户信息
user_info_frame.pack(fill=tk.X, pady=(0, 5)) # 放置在标题下方，笔记本上方

# 检查 LOGGED_IN_USER 和 USER_PERMISSION 是否有值，避免显示 None
user_display_text = f"使用者：{LOGGED_IN_USER}" if LOGGED_IN_USER else "使用者：未登录"
permission_display_text = f"权限：{USER_PERMISSION}" if USER_PERMISSION else "权限：未知"

# 将用户信息和权限信息合并到同一个标签中，用 " - " 分隔
combined_info_text = f"{user_display_text}  -  {permission_display_text}"

info_label = ttk.Label(user_info_frame,
                    text=combined_info_text,
                    style='UserInfo.TLabel', # 可以定义一个新的样式，或者使用默认
                    anchor='center')
info_label.pack(fill=tk.X)

# 新增：使用说明按钮
# 确保 show_instructions 函数在您的代码中已经定义
instructions_button = ttk.Button(user_info_frame,
                                text="查看使用说明",
                                command=show_instructions) # 绑定到 show_instructions 函数
instructions_button.pack(pady=(5,0)) # 在用户信息下方添加一些垂直间距


# 使用Notebook来组织不同功能区域
notebook = ttk.Notebook(main_frame)
notebook.pack(fill=tk.BOTH, expand=True)



# === 查询功能标签页 ===
search_frame = ttk.Frame(notebook, padding="10")
notebook.add(search_frame, text='查询功能')

# 查询区域1
search_area = ttk.LabelFrame(search_frame, text="按人员查询", padding="10")
search_area.pack(fill=tk.X, pady=(0, 10))

# 查询输入框和按钮
search_input_frame = ttk.Frame(search_area)
search_input_frame.pack(fill=tk.X, pady=5)
ttk.Label(search_input_frame, text="输入人员姓名：", 
        style='Header.TLabel').pack(side=tk.LEFT)
search_entry = ttk.Entry(search_input_frame, width=30)
search_entry.pack(side=tk.LEFT, padx=5)
search_button = ttk.Button(search_input_frame, text="查找名下物品", 
                        command=search_borrower_items, style='Action.TButton')
search_button.pack(side=tk.LEFT, padx=5)

# 添加查看个人借还记录按钮
view_personal_button = ttk.Button(search_input_frame, text="查看借还记录", 
                                command=lambda: view_personal_records(search_entry.get()), 
                                style='Action.TButton')
view_personal_button.pack(side=tk.LEFT, padx=5)

# 查看记录按钮
view_frame = ttk.Frame(search_area)
view_frame.pack(fill=tk.X, pady=10)
view_button = ttk.Button(view_frame, text="查看仓库总表物品详细信息", 
                    command=view_inventory, style='Action.TButton')
view_button.pack(side=tk.LEFT, padx=5)
view_borrow_return_button = ttk.Button(view_frame, text="查看仓库总表借还记录", 
                                    command=view_borrow_return, style='Action.TButton')
view_borrow_return_button.pack(side=tk.LEFT, padx=5)

# 查询区域2：按物品查询
search_area2 = ttk.LabelFrame(search_frame, text="按物品查询", padding="10")
search_area2.pack(fill=tk.X, pady=(0, 10))

# 第一行：物品大类
search_item_frame = ttk.Frame(search_area2)
search_item_frame.pack(fill=tk.X, pady=2)
ttk.Label(search_item_frame, text="物品大类：", 
        style='Header.TLabel', width=15).pack(side=tk.LEFT)
search_item_entry = ttk.Entry(search_item_frame, width=30)
search_item_entry.pack(side=tk.LEFT, padx=5)
search_item_menu = ttk.OptionMenu(search_item_frame, 
                                search_item_choice, "",
                                *sorted(inventory_df['大类名称'].unique()),
                                command=set_search_item_from_dropdown)
search_item_menu.pack(side=tk.LEFT, padx=5)

# 第二行：物品子类
search_subitem_frame = ttk.Frame(search_area2)
search_subitem_frame.pack(fill=tk.X, pady=2)
ttk.Label(search_subitem_frame, text="物品子类：", 
        style='Header.TLabel', width=15).pack(side=tk.LEFT)
search_subitem_entry = ttk.Entry(search_subitem_frame, width=30)
search_subitem_entry.pack(side=tk.LEFT, padx=5)
search_subitem_menu = ttk.OptionMenu(search_subitem_frame, 
                                    search_subitem_choice,
                                    "",
                                    command=set_search_subitem_from_dropdown)
search_subitem_menu.pack(side=tk.LEFT, padx=5)

# 搜索按钮
search_item_button = ttk.Button(search_area2, 
                            text="搜索物品", 
                            command=search_item_records,
                            style='Action.TButton')
search_item_button.pack(pady=5)


# --- 新增：操作请求标签页 ---
requests_op_tab = ttk.Frame(notebook)
notebook.add(requests_op_tab, text='操作请求')

# Frame for request inputs
request_input_frame = ttk.LabelFrame(requests_op_tab, text="请求信息")
request_input_frame.pack(padx=10, pady=10, fill="x")

# 创建StringVar变量用于下拉菜单
req_category_choice = tk.StringVar()
req_subcategory_choice = tk.StringVar()
req_description_choice = tk.StringVar()

ttk.Label(request_input_frame, text="物品大类:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
req_category_entry = ttk.Entry(request_input_frame, width=30)
req_category_entry.grid(row=0, column=2, padx=5, pady=5, sticky="ew")
req_category_menu = ttk.OptionMenu(request_input_frame, req_category_choice, "", *sorted(inventory_df['大类名称'].unique()))
req_category_menu.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

ttk.Label(request_input_frame, text="物品小类:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
req_subcategory_entry = ttk.Entry(request_input_frame, width=30)
req_subcategory_entry.grid(row=1, column=2, padx=5, pady=5, sticky="ew")
req_subcategory_menu = ttk.OptionMenu(request_input_frame, req_subcategory_choice, "")
req_subcategory_menu.grid(row=1, column=1, padx=5, pady=5, sticky="ew")

ttk.Label(request_input_frame, text="物品名称/备注:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
req_item_name_entry = ttk.Entry(request_input_frame, width=30)
req_item_name_entry.grid(row=2, column=2, padx=5, pady=5, sticky="ew")
req_description_menu = ttk.OptionMenu(request_input_frame, req_description_choice, "")
req_description_menu.grid(row=2, column=1, padx=5, pady=5, sticky="ew")
ttk.Label(request_input_frame, text="(看你借的是好的还是坏的)").grid(row=2, column=3, padx=5, pady=5, sticky="w")

ttk.Label(request_input_frame, text="数量:").grid(row=3, column=0, padx=5, pady=5, sticky="w")
req_quantity_entry = ttk.Entry(request_input_frame, width=10)
req_quantity_entry.grid(row=3, column=1, padx=5, pady=5, sticky="w")

request_input_frame.columnconfigure(1, weight=1)
request_input_frame.columnconfigure(2, weight=1)

# Frame for request buttons
request_buttons_frame = ttk.Frame(requests_op_tab)
request_buttons_frame.pack(padx=10, pady=10, fill="x")

ttk.Button(request_buttons_frame, text="申请借用", command=request_borrow_item_ui).pack(side=tk.LEFT, padx=5, pady=5)
ttk.Button(request_buttons_frame, text="申请归还", command=request_return_item_ui).pack(side=tk.LEFT, padx=5, pady=5)
ttk.Button(request_buttons_frame, text="申请交付", command=request_deliver_item_ui).pack(side=tk.LEFT, padx=5, pady=5)
ttk.Button(request_buttons_frame, text="申请报损", command=request_damage_item_ui).pack(side=tk.LEFT, padx=5, pady=5)

# --- 结束：操作请求标签页 ---


# 绑定事件
req_category_choice.trace("w", lambda *args: set_req_category_from_dropdown())
req_subcategory_choice.trace("w", lambda *args: set_req_subcategory_from_dropdown())
req_description_choice.trace("w", lambda *args: set_req_description_from_dropdown())



# 绑定事件
category_choice.trace("w", set_category_from_dropdown)
subcategory_choice.trace("w", set_subcategory_from_dropdown)
category_choice.trace("w", set_category_from_dropdown)
subcategory_choice.trace("w", set_subcategory_from_dropdown)
category_choice2.trace("w", set_category_from_dropdown2)
subcategory_choice2.trace("w", set_subcategory_from_dropdown2)


# 新增的物品查询相关的事件绑定
search_item_choice.trace("w", lambda *args: set_search_item_from_dropdown())
search_subitem_choice.trace("w", lambda *args: set_search_subitem_from_dropdown())
borrower_choice.trace("w", lambda *args: set_borrower_from_dropdown())
search_item_entry.bind("<FocusOut>", lambda event: update_search_subitem_options())
# 为搜索按钮绑定回车键
search_item_entry.bind("<Return>", lambda event: search_item_records())
search_subitem_entry.bind("<Return>", lambda event: search_item_records())
# 为菜单选项绑定事件
search_item_menu.bind('<Button-1>', lambda event: update_search_subitem_options())

# 启动定时器检查请求状态
root.after(5000, check_pending_requests_status_periodic) # 启动周期性检查
# 启动主循环
root.mainloop()