import pandas as pd
import tkinter as tk
from tkinter import messagebox
import tkinter.ttk as ttk
from tkinter import scrolledtext  # For better text display
from datetime import datetime  # 添加这行来导入datetime
import sys
import uuid
import time
import os
from tkinter import simpledialog

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
current_requests_df = None # 初始化全局变量
global requests_tree # ttk.Treeview

# 在主程序启动时获取传递过来的用户名和权限
if __name__ == "__main__":
    if len(sys.argv) == 4:  # 脚本名 + 用户名 + 权限 + 姓名
        LOGGED_IN_USER = sys.argv[1]
        USER_PERMISSION = sys.argv[2]
        USER_NAME = sys.argv[3]
        print(f"主程序已启动。登录用户: {LOGGED_IN_USER}, 权限: {USER_PERMISSION}, 姓名: {USER_NAME}")
        # 你可以在这里根据 USER_PERMISSION 的值来控制程序的不同行为或界面显示
    elif len(sys.argv) == 1: # 如果直接运行 managermentSystem.py 而没有参数
        print("主程序直接启动（未传递用户信息和权限）。")
        # 此处可以添加逻辑，例如：
        # 1. 强制退出并提示需要通过登录界面启动
        messagebox.showerror("启动错误", "请通过登录界面启动程序。")
        # exit()
        # 2. 或者以默认用户/受限权限运行（不推荐，除非有明确场景）
    else:
        print("错误：传递给主应用程序的参数数量不正确。")
        messagebox.showerror("启动错误", "启动参数错误。")
        exit()

# 加载数据库
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

def refresh_database():
    """刷新所有数据库"""
    try:
        global inventory_df, borrow_return_df
        inventory_df = pd.read_excel(inventory_db_path, engine='openpyxl')
        borrow_return_df = pd.read_excel(borrow_return_db_path, engine='openpyxl')
        inventory_df['数量'] = inventory_df['数量'].astype(int)
        return True
    except Exception as e:
        messagebox.showerror("数据库刷新错误", f"刷新数据库时发生错误: {e}")
        return False
        
def load_requests_db():
    try:
        if os.path.exists(REQUESTS_DB_PATH):
            return pd.read_excel(REQUESTS_DB_PATH, engine='openpyxl')
        else:
            # 如果文件不存在，管理员端可以不创建，等待用户端创建
            # 或者也创建一个空的
            df = pd.DataFrame(columns=['RequestID', 'Timestamp', 'Username', 'UserFullName', 
                                       'ItemCategory', 'ItemSubcategory', 'ItemName', 'Quantity', 
                                       'RequestType', 'AdminActionStatus', 'AdminRemarks', 
                                       'UserNotified', 'OriginalBorrowRequestID'])
            # df.to_excel(REQUESTS_DB_PATH, index=False, engine='openpyxl') # 可选
            return df
    except Exception as e:
        messagebox.showerror("错误", f"加载请求数据库失败: {e}")
        return None

def save_requests_db(df):
    try:
        df.to_excel(REQUESTS_DB_PATH, index=False, engine='openpyxl')
    except Exception as e:
        messagebox.showerror("错误", f"保存请求数据库失败: {e}")

def populate_pending_requests_tree():
    global requests_tree, current_requests_df # 声明 current_requests_df 为全局变量
    # 清空 Treeview
    for i in requests_tree.get_children():
       requests_tree.delete(i)
    
    current_requests_df = load_requests_db() # 将加载的数据赋值给全局变量
    
    if current_requests_df is None:
        # 如果 load_requests_db 返回 None (例如加载时发生错误)
        # 将 current_requests_df 设置为一个空的 DataFrame 以防止后续函数出错
        current_requests_df = pd.DataFrame(columns=['RequestID', 'Timestamp', 'Username', 'UserFullName', 
                                       'ItemCategory', 'ItemSubcategory', 'ItemName', 'Quantity', 
                                       'RequestType', 'AdminActionStatus', 'AdminRemarks', 
                                       'UserNotified', 'OriginalBorrowRequestID'])
        # load_requests_db 函数内部应该已经显示了错误信息
        return # 如果数据加载失败，则不继续填充 Treeview

    # 确保 'AdminActionStatus' 列存在
    if 'AdminActionStatus' in current_requests_df.columns:
        pending_df = current_requests_df[current_requests_df['AdminActionStatus'] == 'Pending']
        
        for index, row in pending_df.iterrows():
           requests_tree.insert("", tk.END, values=(
               row.get('RequestID'), row.get('Timestamp'), row.get('UserFullName'), row.get('RequestType'),
               row.get('ItemCategory'), row.get('ItemSubcategory'), row.get('ItemName'), row.get('Quantity')
           ))
    else:
        # 如果DataFrame中缺少必要的列，则显示错误
        messagebox.showerror("数据错误", "请求数据文件缺少 'AdminActionStatus' 列，无法加载待处理请求。")
    # pass # UI填充逻辑 # 此行不再需要

def approve_request():
    global requests_tree, current_requests_df, inventory_df, borrow_return_df # 添加 inventory_df 和 borrow_return_df
    selected_item = requests_tree.focus()
    if not selected_item:
        messagebox.showwarning("选择错误", "请先选择一个请求。")
        return

    selected_values = requests_tree.item(selected_item)['values']
    request_id = selected_values[0]
    request_type = selected_values[3]
    item_category = selected_values[4]
    item_subcategory = selected_values[5]
    item_name = selected_values[6] # This is 'ItemName' from requests.xlsx, likely corresponds to '物品备注' or a unique identifier
    quantity = int(selected_values[7])
    user_full_name = selected_values[2] # UserFullName, corresponds to '保管人员'

    remarks = simpledialog.askstring("管理员备注", "请输入批准备注 (可选):")
    # remarks can be None if user cancels, or empty string if they don't type anything

    # Find the request in current_requests_df
    request_idx = current_requests_df[current_requests_df['RequestID'] == request_id].index
    if request_idx.empty:
        messagebox.showerror("错误", f"在请求列表中未找到请求ID: {request_id}")
        return

    # --- Database Update Logic based on RequestType ---
    try:
        if request_type == 'Borrow':
            # Decrease inventory
            # Find matching item in inventory_df
            # We need to match on '大类名称', '小类名称', and '物品备注' (which is item_name here)
            inventory_match_condition = (
                (inventory_df['大类名称'] == item_category) &
                (inventory_df['小类名称'] == item_subcategory) &
                (inventory_df['备注'] == item_name) # Assuming item_name from request is the '物品备注'
            )
            item_in_inventory_idx = inventory_df[inventory_match_condition].index

            if item_in_inventory_idx.empty:
                messagebox.showerror("库存错误", f"未在库存中找到物品: {item_category}-{item_subcategory}, 备注: {item_name}")
                return
            
            # Assuming only one such item entry, or update the first one found
            idx_to_update = item_in_inventory_idx[0] 
            current_quantity = inventory_df.loc[idx_to_update, '数量']

            if current_quantity < quantity:
                messagebox.showerror("库存不足", f"物品 {item_category}-{item_subcategory} ({item_name}) 库存 ({current_quantity}) 不足 {quantity}。")
                return
            inventory_df.loc[idx_to_update, '数量'] -= quantity

            # Add to borrow_return_df
            new_borrow_record = pd.DataFrame([{
                '借出物品大类名称': item_category,
                '借出物品小类名称': item_subcategory,
                '借出物品数量': quantity,
                '保管人员': user_full_name,
                '物品状态': '借出', # Or 'Borrowed'
                '备注': f"请求ID: {request_id}. {remarks if remarks else ''}", # Include admin remarks
                '日期': datetime.now().strftime('%Y%m%d %H:%M:%S') # More precise timestamp
            }])
            borrow_return_df = pd.concat([borrow_return_df, new_borrow_record], ignore_index=True)

        elif request_type == 'Return':
            # Increase inventory (assuming returned item is '好的')
            # Find matching item in inventory_df to increase its quantity
            # We need to match on '大类名称', '小类名称', and '物品备注' (which is item_name here)
            # And typically, '备注' should be '好的' for returned items, or we add to existing '好的' stock
            inventory_match_condition = (
                (inventory_df['大类名称'] == item_category) &
                (inventory_df['小类名称'] == item_subcategory) &
                (inventory_df['备注'] == item_name) # Assuming returns go to '好的' stock
            )
            item_in_inventory_idx = inventory_df[inventory_match_condition].index

            if item_in_inventory_idx.empty:
                # If no '好的' stock exists with this specific '物品备注', create a new entry or handle as error
                # For simplicity, let's assume we find one or it's an error for now.
                # A more robust solution might create a new inventory line if one doesn't exist.
                messagebox.showwarning("库存警告", f"未在库存中找到物品 {item_category}-{item_subcategory} ({item_name}) 标记为 '好的'. 将尝试添加到第一个匹配项或创建新条目。")
                # Fallback: try to find any item with same category/subcategory/item_name and add there
                fallback_condition = (
                    (inventory_df['大类名称'] == item_category) &
                    (inventory_df['小类名称'] == item_subcategory) &
                    (inventory_df['备注'] == item_name)
                )
                item_in_inventory_idx = inventory_df[fallback_condition].index
                if item_in_inventory_idx.empty:
                     messagebox.showerror("库存错误", f"无法归还：未在库存中找到物品: {item_category}-{item_subcategory}, 备注: {item_name}")
                     return


            idx_to_update = item_in_inventory_idx[0]
            inventory_df.loc[idx_to_update, '数量'] += quantity
            
            # Add to borrow_return_df
            new_return_record = pd.DataFrame([{
                '借出物品大类名称': item_category,
                '借出物品小类名称': item_subcategory,
                '借出物品数量': quantity, # Quantity returned
                '保管人员': user_full_name,
                '物品状态': '归还', # Or 'Returned'
                '备注': f"请求ID: {request_id}. {remarks if remarks else ''}",
                '日期': datetime.now().strftime('%Y%m%d %H:%M:%S')
            }])
            borrow_return_df = pd.concat([borrow_return_df, new_return_record], ignore_index=True)

        elif request_type == 'Deliver':
            # Decrease inventory (similar to Borrow)
            inventory_match_condition = (
                (inventory_df['大类名称'] == item_category) &
                (inventory_df['小类名称'] == item_subcategory) &
                (inventory_df['备注'] == item_name)
            )
            item_in_inventory_idx = inventory_df[inventory_match_condition].index
            if item_in_inventory_idx.empty:
                messagebox.showerror("库存错误", f"未在库存中找到物品: {item_category}-{item_subcategory}, 备注: {item_name}")
                return
            idx_to_update = item_in_inventory_idx[0]
            current_quantity = inventory_df.loc[idx_to_update, '数量']
            if current_quantity < quantity:
                messagebox.showerror("库存不足", f"物品 {item_category}-{item_subcategory} ({item_name}) 库存 ({current_quantity}) 不足 {quantity}。")
                return
            inventory_df.loc[idx_to_update, '数量'] -= quantity

            # Add to borrow_return_df
            new_deliver_record = pd.DataFrame([{
                '借出物品大类名称': item_category,
                '借出物品小类名称': item_subcategory,
                '借出物品数量': quantity,
                '保管人员': user_full_name, # Or 'N/A' if delivered out of system
                '物品状态': '交付', # Or 'Delivered'
                '备注': f"请求ID: {request_id}. {remarks if remarks else ''}",
                '日期': datetime.now().strftime('%Y%m%d %H:%M:%S')
            }])
            borrow_return_df = pd.concat([borrow_return_df, new_deliver_record], ignore_index=True)

        elif request_type == 'Damage':
            # Decrease inventory (similar to Borrow)
            # And potentially update the '备注' of the item in inventory_df if it's not fully depleted
            # Or move to a '损坏品' category if you have one
            inventory_match_condition = (
                (inventory_df['大类名称'] == item_category) &
                (inventory_df['小类名称'] == item_subcategory) &
                (inventory_df['备注'] == item_name) # Assuming damage happens to '好的' items
            )
            item_in_inventory_idx = inventory_df[inventory_match_condition].index

            if item_in_inventory_idx.empty:
                messagebox.showerror("库存错误", f"未在库存中找到可损坏的'好的'物品: {item_category}-{item_subcategory}, 备注: {item_name}")
                return

            idx_to_update = item_in_inventory_idx[0]
            current_quantity = inventory_df.loc[idx_to_update, '数量']

            if current_quantity < quantity:
                messagebox.showerror("库存不足", f"物品 {item_category}-{item_subcategory} ({item_name}) '好的'库存 ({current_quantity}) 不足 {quantity} 以标记为损坏。")
                return
            
            inventory_df.loc[idx_to_update, '数量'] -= quantity # Reduce '好的' quantity

            # Add/Update '坏的' stock for the same item
            damaged_stock_condition = (
                (inventory_df['大类名称'] == item_category) &
                (inventory_df['小类名称'] == item_subcategory) &
                (inventory_df['物品备注'] == item_name) & # Match the specific item
                (inventory_df['备注'] == '坏的')
            )
            damaged_item_idx = inventory_df[damaged_stock_condition].index
            if not damaged_item_idx.empty:
                inventory_df.loc[damaged_item_idx[0], '数量'] += quantity
            else: # Create new entry for '坏的' if it doesn't exist
                new_damaged_item_details = inventory_df.loc[idx_to_update].copy() # copy details from '好的' item
                new_damaged_item_details['数量'] = quantity
                new_damaged_item_details['备注'] = '坏的'
                # Ensure '存放位置' and '物品备注' are correctly copied
                # new_damaged_item_details['存放位置'] = inventory_df.loc[idx_to_update, '存放位置']
                # new_damaged_item_details['物品备注'] = inventory_df.loc[idx_to_update, '物品备注']
                inventory_df = pd.concat([inventory_df, pd.DataFrame([new_damaged_item_details])], ignore_index=True)


            # Add to borrow_return_df
            new_damage_record = pd.DataFrame([{
                '借出物品大类名称': item_category,
                '借出物品小类名称': item_subcategory,
                '借出物品数量': quantity,
                '保管人员': user_full_name, # Person reporting damage
                '物品状态': '损坏', # Or 'Damaged'
                '备注': f"请求ID: {request_id}. {remarks if remarks else ''}",
                '日期': datetime.now().strftime('%Y%m%d %H:%M:%S')
            }])
            borrow_return_df = pd.concat([borrow_return_df, new_damage_record], ignore_index=True)
        
        else:
            messagebox.showwarning("未知请求", f"未知的请求类型: {request_type}")
            return # Do not proceed if type is unknown

        # Save updated databases
        inventory_df.to_excel(inventory_db_path, index=False, engine='openpyxl')
        borrow_return_df.to_excel(borrow_return_db_path, index=False, engine='openpyxl')
        
        # Update request status in current_requests_df
        current_requests_df.loc[request_idx, 'AdminActionStatus'] = 'Approved'
        current_requests_df.loc[request_idx, 'AdminRemarks'] = remarks if remarks else ''
        # current_requests_df.loc[request_idx, 'UserNotified'] = False # Or True
        save_requests_db(current_requests_df)

        messagebox.showinfo("成功", f"请求 {request_id} ({request_type}) 已批准并处理。")
        populate_pending_requests_tree() # Refresh the list

    except Exception as e:
        messagebox.showerror("处理错误", f"批准请求 {request_id} 时发生错误: {e}")
        # Potentially revert changes if partial update occurred, though this is complex with Excel files
        # For now, just log/show error. Reload data to be safe.
        # inventory_df = pd.read_excel(inventory_db_path, engine='openpyxl') # Reload
        # borrow_return_df = pd.read_excel(borrow_return_db_path, engine='openpyxl') # Reload
        populate_pending_requests_tree() # Refresh to show current state

def deny_request():
    global requests_tree, current_requests_df, inventory_df, borrow_return_df # 确保 inventory_df 和 borrow_return_df 可用
    selected_item = requests_tree.focus()
    if not selected_item:
       messagebox.showwarning("选择错误", "请先选择一个请求。")
       return
    
    selected_values = requests_tree.item(selected_item)['values']
    request_id = selected_values[0]
    # request_type = selected_values[3] # RequestType
    # item_category = selected_values[4] # ItemCategory
    # item_subcategory = selected_values[5] # ItemSubcategory
    # item_name = selected_values[6] # ItemName (物品备注)
    # quantity = int(selected_values[7]) # Quantity
    # user_full_name = selected_values[2] # UserFullName

    remarks = simpledialog.askstring("管理员备注", "请输入拒绝备注 (必须):")
    if remarks is None: # 用户取消输入
        return
    if not remarks.strip():
        messagebox.showerror("错误", "拒绝备注不能为空。")
        return

    idx = current_requests_df[current_requests_df['RequestID'] == request_id].index
    if not idx.empty:
       current_requests_df.loc[idx, 'AdminActionStatus'] = 'Denied'
       current_requests_df.loc[idx, 'AdminRemarks'] = remarks
       # current_requests_df.loc[idx, 'UserNotified'] = False # 或者 True，取决于是否立即通知
       save_requests_db(current_requests_df)
       messagebox.showinfo("成功", f"请求 {request_id} 已拒绝。")
       populate_pending_requests_tree() #刷新列表
    else:
        messagebox.showerror("错误", f"未找到请求ID: {request_id}")

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
            search_subitem_menu['menu'].delete(0, 'end')
            for subitem in subitems:
                search_subitem_menu['menu'].add_command(
                    label=subitem, 
                    command=tk._setit(search_subitem_choice, subitem, set_search_subitem_from_dropdown))
        except Exception as e:
            messagebox.showerror("Error", f"更新子类别选项时出错: {e}")

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
        borrower = borrower_entry.get()
        category = category_entry2.get()
        subcategory = subcategory_entry2.get()
        quantity = int(quantity_entry2.get())
        status = status_entry.get()
        remark = remark_entry2.get()
        # 获取当前日期并格式化为YYYYMMDD格式
        current_date = datetime.now().strftime('%Y%m%d')

        if quantity <= 0:
            raise ValueError("不要乱写负数！fk你！.")

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
    if not refresh_database():
        return
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
        current_count = {}  # 当前借出的物品
        delivered_count = {}  # 已交付的物品
        damaged_count = {}  # 已损坏的物品
        remark_count = {}  # 存储不同备注的物品数量
        admin_remarks_count = {}  # 存储不同管理员审批备注的物品数量

        # 加载请求数据库以获取管理员备注
        requests_df = load_requests_db()

        for index, row in borrower_records.iterrows():
            category = row['借出物品大类名称']
            subcategory = row['借出物品小类名称']
            quantity = row['借出物品数量']
            status = row['物品状态']
            remark = row.get('备注', '无备注')  # 获取物品备注信息

            item_key = (category, subcategory)

            # 更新物品状态计数
            if status == '借出':
                current_count[item_key] = current_count.get(item_key, 0) + quantity
                # 更新备注计数
                if item_key not in remark_count:
                    remark_count[item_key] = {}
                remark_count[item_key][remark] = remark_count[item_key].get(remark, 0) + quantity

                # 查找并统计管理员审批备注
                if requests_df is not None and not requests_df.empty:
                    matching_requests = requests_df[
                        (requests_df['Username'] == borrower_name) &
                        (requests_df['ItemCategory'] == category) &
                        (requests_df['ItemSubcategory'] == subcategory) &
                        (requests_df['AdminActionStatus'] != 'Pending')
                    ]
                    if not matching_requests.empty:
                        if item_key not in admin_remarks_count:
                            admin_remarks_count[item_key] = {}
                        for _, req in matching_requests.iterrows():
                            if pd.notna(req['AdminRemarks']):
                                admin_remark = req['AdminRemarks']
                                admin_remarks_count[item_key][admin_remark] = admin_remarks_count[item_key].get(admin_remark, 0) + req['Quantity']

            elif status == '归还':
                current_count[item_key] = current_count.get(item_key, 0) - quantity
            elif status == '交付':
                delivered_count[item_key] = delivered_count.get(item_key, 0) + quantity
            elif status == '损坏':
                damaged_count[item_key] = damaged_count.get(item_key, 0) + quantity

        def format_item(count, category, subcategory):
            return f"{count} 个 {category}-{subcategory}"

        # 显示当前借出的物品及其详细信息
        for (category, subcategory), count in current_count.items():
            if count >= 1:
                text.insert(tk.END, f"\n当前名下物品还有：{category}-{subcategory}\n")
                text.insert(tk.END, f"  总数量：{count} 个\n")
                
                # 显示不同备注的物品数量
                if (category, subcategory) in remark_count:
                    text.insert(tk.END, "  详细备注信息：\n")
                    for remark, remark_quantity in remark_count[(category, subcategory)].items():
                        text.insert(tk.END, f"    - {remark}: {remark_quantity} 个\n")
                
                # 显示不同管理员审批备注的物品数量
                if (category, subcategory) in admin_remarks_count:
                    text.insert(tk.END, "  管理员审批备注统计：\n")
                    for admin_remark, admin_quantity in admin_remarks_count[(category, subcategory)].items():
                        text.insert(tk.END, f"    - {admin_remark}: {admin_quantity} 个\n")
                
                text.insert(tk.END, "-------------------\n")

        # 显示已交付的物品
        delivered_items = ", ".join([
            format_item(count, category, subcategory)
            for (category, subcategory), count in delivered_count.items()
        ])
        if delivered_items:
            text.insert(tk.END, f"\n已交付物品：{delivered_items}。\n")

        # 显示已损坏的物品
        damaged_items = ", ".join([
            format_item(count, category, subcategory)
            for (category, subcategory), count in damaged_count.items()
        ])
        if damaged_items:
            text.insert(tk.END, f"\n已损坏物品：{damaged_items}。\n")

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
            "title": "1. 库存管理页面",
            "items": [
                ("1.1 添加物品栏目", "带有下拉箭头的可以选择，选择后会自动填入。也可以手动输入，不会冲突。", "black"),
                ("1.2 存放位置下拉菜单", "带有自动索引功能。你选择了仓库中已有的物品和位置后，会自动筛选然后给你下一个下拉菜单的选项。也可以自己写，就是新增位置。", "black"),
                ("1.3 描述符", "是区分物品好坏的，因为涉及损坏，匹配等。物品备注随意输入，不会作为索引逻辑。", "black"),
                ("1.4 必填项", "前六行是必须有输入的，不输入无法入库操作。", "red"),
                ("1.5 添加物品", "输入完成后点击添加物品即可，会同步更新数据库。", "black"),
                ("1.6 查找物品信息", "只需要输入大类别名称和子类别名称就行，是一个简化的功能。", "black")
            ]
        },
        {
            "title": "2. 借还管理页面",
            "items": [
                ("2.1 基本操作", "其他同上。操作描述符的时候没有下拉菜单，因为需要处理损坏情况。需要留心，所以没有涉及下拉。", "black"),
                ("2.2 完整流程", "", "black"),
                ("2.2.1 借出物品", "操作借还人员A借出物品。", "black"),
                ("2.2.2 归还物品", "操作借还人员A归还物品。", "black"),
                ("2.2.3 损坏物品", "操作借还人员B借出物品，但是B损坏物品。", "black"),
                ("2.2.4 损坏处理", "借还借还人员B损坏物品，描述符填好的（因为是损坏了好的物品）。", "black"),
                ("2.2.5 交付物品", "操作借还人员C借出物品，然后C交付甲方。", "black"),
                ("2.2.6 交付处理", "借还借还人员C交付物品，描述符填具体交付的（一般不会交付坏的）。", "black"),
                ("2.3 采购功能", "目前正在开发中，和后续报销等流程结合在一起。", "black")
            ]
        },
        {
            "title": "3. 查询功能页面",
            "items": [
                ("3.1 按人员查询", "可以查到某个人名下的所有物品，在手上的，损坏的，交付的等。", "black"),
                ("3.2 数据库信息", "可以直接读取数据库信息，也即是'查看仓库总表物品详细信息'和'查看仓库总表借还记录'，功能和库存管理中的'点击查找物品信息'是一样的，更完善。", "black"),
                ("3.3 按物品查询", "可以查找具体某个物品的详细位置和数量，包括借出归还损坏的人员信息。和库存管理中的'点击查找物品信息是一样的，更完善", "black")
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

user_label = ttk.Label(user_info_frame,
                       text=f"{user_display_text}  -  {permission_display_text}",
                       style='UserInfo.TLabel', # 可以定义一个新的样式，或者使用默认
                       anchor='center')
user_label.pack(fill=tk.X)


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

# 创建一个新的容器来容纳两行输入控件
inputs_container = ttk.Frame(location_frame)
inputs_container.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

# 第一行输入控件
location_row1_frame = ttk.Frame(inputs_container)
location_row1_frame.pack(fill=tk.X)

# Location Combobox
location_combo = ttk.Combobox(location_row1_frame, textvariable=location_choice,
                            values=["627", "629"], width=10, state='readonly')
location_combo.pack(side=tk.LEFT, padx=(0, 5)) # Adjusted padding

# Cabinet Entry and Combobox
cabinet_menu = ttk.Combobox(location_row1_frame, textvariable=cabinet_number, width=10, state='readonly') # Added textvariable for consistency
cabinet_menu.pack(side=tk.LEFT, padx=5)
cabinet_entry = ttk.Entry(location_row1_frame, textvariable=cabinet_number, width=10)
cabinet_entry.pack(side=tk.LEFT, padx=5)

# 第二行输入控件
location_row2_frame = ttk.Frame(inputs_container)
location_row2_frame.pack(fill=tk.X, pady=(5,0)) # Add some padding on top of the second row

# Description Entry and Combobox
description_menu = ttk.Combobox(location_row2_frame, textvariable=description_number, width=13, state='readonly') # Added textvariable for consistency
description_menu.pack(side=tk.LEFT, padx=(0, 5)) # Adjusted padding, assuming it's the first in its row
description_entry = ttk.Entry(location_row2_frame, textvariable=description_number, width=13)
description_entry.pack(side=tk.LEFT, padx=5)

# 第五行：备注和大类名称
remark_frame = ttk.Frame(input_frame)
remark_frame.pack(fill=tk.X, pady=2)
ttk.Label(remark_frame, text="描述符(好的，坏的，无)：", 
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
ttk.Button(button_frame, text="管理者使用说明",
          command=show_instructions, style='Action.TButton').pack(side=tk.LEFT, padx=5) # 修改此按钮

# 警告信息
alert_frame = ttk.Frame(input_frame)
alert_frame.pack(fill=tk.X, pady=5)
ttk.Label(alert_frame, text="管理员操作注意：物品先出库，再入库，最后处理损坏或交付！", 
         style='Alert.TLabel').pack(side=tk.LEFT)
ttk.Label(alert_frame, text="前六行必须输入，连接SYSU-HILAB能够访问NAS方可使用！", 
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

# 添加人员下拉菜单
borrower_menu = ttk.OptionMenu(borrower_frame, 
                              borrower_choice,
                              "",
                              *sorted(borrow_return_df['保管人员'].unique()),
                              command=set_borrower_from_dropdown)
borrower_menu.pack(side=tk.LEFT, padx=5)

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
ttk.Label(remark2_frame, text="描述符（请查看上一页面下拉菜单）：", 
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

# --- Tab 5: 请求审批 (Request Approval) ---
requests_approval_tab = ttk.Frame(notebook)
notebook.add(requests_approval_tab, text='请求审批')

# Frame for Treeview and Scrollbar
requests_tree_frame = ttk.Frame(requests_approval_tab)
requests_tree_frame.pack(pady=10, padx=10, fill="both", expand=True)

# Treeview for requests
requests_cols = ('RequestID', 'Timestamp', 'UserFullName', 'RequestType', 
                 'ItemCategory', 'ItemSubcategory', 'ItemName', 'Quantity')
requests_tree = ttk.Treeview(requests_tree_frame, columns=requests_cols, show='headings')

for col in requests_cols:
    requests_tree.heading(col, text=col)
    if col == 'RequestID':
        requests_tree.column(col, width=220, anchor='w')
    elif col == 'Timestamp':
        requests_tree.column(col, width=130, anchor='center')
    elif col == 'UserFullName':
        requests_tree.column(col, width=80, anchor='center')
    elif col == 'RequestType':
        requests_tree.column(col, width=70, anchor='center')
    elif col == 'Quantity':
        requests_tree.column(col, width=50, anchor='e')
    else:
        requests_tree.column(col, width=100, anchor='w')

# Scrollbar for requests_tree
requests_scrollbar = ttk.Scrollbar(requests_tree_frame, orient="vertical", command=requests_tree.yview)
requests_tree.configure(yscrollcommand=requests_scrollbar.set)

requests_scrollbar.pack(side="right", fill="y")
requests_tree.pack(side="left", fill="both", expand=True)


# Frame for buttons
requests_buttons_frame = ttk.Frame(requests_approval_tab)
requests_buttons_frame.pack(pady=5, fill="x")

refresh_requests_button = ttk.Button(requests_buttons_frame, text="刷新列表", command=populate_pending_requests_tree)
refresh_requests_button.pack(side=tk.LEFT, padx=5)

approve_request_button = ttk.Button(requests_buttons_frame, text="批准选中项", command=approve_request)
approve_request_button.pack(side=tk.LEFT, padx=5)

deny_request_button = ttk.Button(requests_buttons_frame, text="拒绝选中项", command=deny_request)
deny_request_button.pack(side=tk.LEFT, padx=5)


# 绑定事件
location_combo.bind('<<ComboboxSelected>>', on_location_menu_select)
cabinet_menu.bind('<<ComboboxSelected>>', on_cabinet_menu_select)
description_menu.bind('<<ComboboxSelected>>', on_description_menu_select)
category_entry.bind("<FocusOut>", lambda event: on_category_subcategory_select(None))
subcategory_entry.bind("<FocusOut>", lambda event: on_category_subcategory_select(None))
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

# 启动主循环
root.mainloop()