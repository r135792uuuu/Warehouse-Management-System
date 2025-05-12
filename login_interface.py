import tkinter as tk
from tkinter import messagebox
import pandas as pd
import subprocess
import os

# 配置文件路径
# 注意：请确保此路径与 managermentSystem.py 中的路径一致，并且运行脚本的计算机有权访问此网络路径。
PERMISSIONS_DB_PATH = '//HILAB627_DS/database/permissions.xlsx'

# 获取当前脚本 (login_interface.py) 所在的目录
# 这有助于构建到 managermentSystem.py 的绝对路径，使其更具可移植性
try:
    # 当作为脚本运行时，__file__ 是定义的
    SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
except NameError:
    # 当在没有定义 __file__ 的环境中运行（例如 Jupyter cell 或某些IDE的控制台）
    # 我们假设当前工作目录是脚本所在的目录
    SCRIPT_DIR = os.getcwd()

MAIN_APP_SCRIPT_NAME = 'managermentSystem.py' # Default, will be overridden
# MAIN_APP_PATH will be set in launch_main_application

def verify_credentials(username, password):
    """
    验证用户凭据是否与 Excel 文件中的记录匹配。
    假设 Excel 文件 (permissions.xlsx) 的第一列是用户名，第二列是密码。
    并且这些列的表头分别是 '姓名' 和 '密码'。
    """
    try:
        # 读取 Excel 文件
        # 假设 Excel 文件有表头，列名为 '姓名' 和 '密码'
        # 如果 Excel 文件没有表头，请使用:
        # perms_df = pd.read_excel(PERMISSIONS_DB_PATH, engine='openpyxl', header=None)
        # 然后使用 perms_df.iloc[index, 0] 和 perms_df.iloc[index, 1] 进行比较
        perms_df = pd.read_excel(PERMISSIONS_DB_PATH, engine='openpyxl')

        # 确保比较时数据类型为字符串，并去除首尾空格
        input_username_str = str(username).strip()
        input_password_str = str(password).strip()

        for index, row in perms_df.iterrows():
            # 从 Excel 中获取用户名和密码，确保转换为字符串并去除空格
            # 使用 .get() 方法并在找不到列时提供默认空字符串，增加稳健性
            excel_username = str(row.get('姓名', '')).strip() # 修改这里： '用户名' -> '姓名'
            excel_password = str(row.get('密码', '')).strip()
            # print(f"Excel 用户名: {excel_username}, Excel 密码: {excel_password}")

            if excel_username == input_username_str and excel_password == input_password_str:
                # 可以在这里读取权限列 (假设为 '权限')，如果需要基于权限执行不同操作
                permission = str(row.get('权限', '')).strip()
                name = str(row.get('中文名', '')).strip() 
                # print(f"用户 {excel_username} 权限为: {permission}")
                return True, input_username_str, permission, name # 修改：返回用户名和权限
        return False, None, None, None # 修改：返回三个值
    except FileNotFoundError:
        messagebox.showerror("登录错误", f"权限文件未找到: {PERMISSIONS_DB_PATH}\n请检查路径是否正确以及文件是否存在。")
        return False, None, None # 修改：返回三个值
    except pd.errors.EmptyDataError:
        messagebox.showerror("登录错误", f"权限文件 {PERMISSIONS_DB_PATH} 为空。")
        return False, None, None # 修改：返回三个值
    except KeyError as e:
        messagebox.showerror("登录错误", f"权限文件中缺少必要的列名（应为 '用户名', '密码'）：{e}")
        return False, None, None # 修改：返回三个值
    except Exception as e:
        messagebox.showerror("登录错误", f"读取权限文件时发生未知错误: {e}")
        return False, None, None # 修改：返回三个值

def attempt_login():
    """
    获取用户输入的凭据，验证它们，如果成功则启动主应用程序。
    """
    username = username_entry.get()
    password = password_entry.get()
    print(f"输入的用户名: {username}, 输入的密码: {password}")

    if not username or not password:
        messagebox.showwarning("输入错误", "用户名和密码不能为空。")
        return

    login_successful, logged_in_username, user_permission, user_name = verify_credentials(username, password) # 修改：接收返回的用户名和权限

    if login_successful:
        messagebox.showinfo("登录成功", "凭据验证成功！正在启动主程序...")
        login_window.destroy()  # 关闭登录窗口
        launch_main_application(logged_in_username, user_permission, user_name) # 修改：传递用户名和权限
    else:
        messagebox.showerror("登录失败", "用户名或密码错误。")

def launch_main_application(logged_in_username, user_permission, user_name): # 修改：接受参数
    """
    启动主仓库管理系统应用程序。
    根据用户权限决定启动哪个脚本。
    """
    target_script_name = ""
    if user_permission == "admin":
        target_script_name = 'managermentSystem.py'
        print(f"管理员登录，启动管理员脚本: {target_script_name}")
    elif user_permission == "user":
        target_script_name = 'managermentSystem_user.py'
        print(f"普通用户登录，启动普通用户脚本: {target_script_name}")
    else:
        messagebox.showerror("权限错误", f"未知的用户权限: {user_permission}")
        return # 如果权限未知，则不启动任何程序

    if not target_script_name: # 再次检查，确保 target_script_name 已被设置
        messagebox.showerror("启动错误", "未能根据权限确定目标应用程序脚本。")
        return

    main_app_path = os.path.join(SCRIPT_DIR, target_script_name)

    try:
        # 使用 'python' 命令执行主脚本，并传递用户名和权限作为参数
        subprocess.Popen(['python', main_app_path, logged_in_username, user_permission, user_name])
    except FileNotFoundError:
        messagebox.showerror("启动错误", f"主应用程序脚本 '{target_script_name}' 未在目录 '{SCRIPT_DIR}' 中找到。")
    except Exception as e:
        messagebox.showerror("启动错误", f"启动主应用程序时出错: {e}")

# --- 创建登录界面的 GUI ---
login_window = tk.Tk()
login_window.title("仓库管理系统 - 登录")
login_window.geometry("350x200") # 设置窗口大小

# 设置窗口使其居中 (可选)
login_window.eval('tk::PlaceWindow . center')


main_frame = tk.Frame(login_window, padx=20, pady=20)
main_frame.pack(expand=True, fill=tk.BOTH)

# 用户名标签和输入框
tk.Label(main_frame, text="用户名:").grid(row=0, column=0, sticky=tk.W, pady=5)
username_entry = tk.Entry(main_frame, width=30)
username_entry.grid(row=0, column=1, pady=5)

# 密码标签和输入框
tk.Label(main_frame, text="密  码:").grid(row=1, column=0, sticky=tk.W, pady=5)
password_entry = tk.Entry(main_frame, show="*", width=30) # show="*" 使密码以星号显示
password_entry.grid(row=1, column=1, pady=5)

# 登录按钮
login_button = tk.Button(main_frame, text="登录", command=attempt_login, width=10)
login_button.grid(row=2, column=0, columnspan=2, pady=15)

# 允许通过回车键登录
username_entry.bind("<Return>", lambda event: password_entry.focus_set()) # 用户名回车后焦点移到密码框
password_entry.bind("<Return>", lambda event: attempt_login()) # 密码框回车后尝试登录

# 将焦点设置在用户名输入框
username_entry.focus_set()

login_window.mainloop()