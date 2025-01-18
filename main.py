import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog
import csv
import chardet
import json
import requests
import os
import configparser

def open_file():
    if not prompt_to_save_if_modified():  # 检查是否有未保存的修改，并提示保存
        return  # 如果用户取消操作，则返回，不继续打开文件
    file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv"), ("STI files", "*.sti")])
    if file_path:
        open_file_path(file_path)
        config, config_path = load_config()
        save_config(config, config_path, file_path)
    else:
        status_str_ver.set('未选择文件')
def save_file():
    file_path = filedialog.asksaveasfilename(defaultextension=".sti", filetypes=[("STI files", "*.sti")])
    if file_path:
        try:
            data = []
            for item in treeview.get_children():
                values = treeview.item(item, 'values')
                # 确保所有值都是字符串类型
                data.append([str(value) for value in values])
            with open(file_path, 'w', encoding='utf-8') as file:
                json.dump(data, file, ensure_ascii=False, indent=4)
            status_str_ver.set(f'文件成功保存 {file_path}')
            global modified
            modified = False  # 保存文件后设置为False
        except Exception as e:
            messagebox.showerror("错误", f"保存文件时出错: {e}")
            status_str_ver.set('保存文件失败')

def calculate_boxes_and_groups(quantity):
    items_per_group = 64
    groups_per_box = 27
    # 计算盒数
    total_boxes = quantity // (items_per_group * groups_per_box)
    # 计算组数
    total_groups = (quantity // items_per_group) - (total_boxes * groups_per_box)
    # 计算个数
    pieces = quantity - ((total_boxes * groups_per_box + total_groups) * items_per_group)
    return total_boxes, total_groups, pieces

def new_connection():
    messagebox.showinfo("新建连接", "新建连接功能尚未实现")

def disconnect():
    messagebox.showinfo("断开连接", "断开连接功能尚未实现")

def about():
    messagebox.showinfo("关于", "原理图材料列表查看器\n版本 1.0\n作者: SmailPang")

def on_right_click(event):
    # 获取点击位置的行
    iid = treeview.identify_row(event.y)
    if iid:
        # 选中该行
        treeview.selection_set(iid)
        # 获取当前行的“是否完成”状态
        completion_status = treeview.item(iid, 'values')[5]
        # 动态设置二级菜单的选项
        status_menu.entryconfig("未完成", state="normal" if completion_status != "未完成" else "disabled")
        status_menu.entryconfig("进行中", state="normal" if completion_status != "进行中" else "disabled")
        status_menu.entryconfig("已完成", state="normal" if completion_status != "已完成" else "disabled")
        # 弹出右键菜单
        right_click_menu.post(event.x_root, event.y_root)

def change_status(new_status):
    global modified
    selected_item = treeview.selection()
    if selected_item:
        for item in selected_item:
            treeview.item(item, values=(treeview.item(item, 'values')[0], treeview.item(item, 'values')[1], treeview.item(item, 'values')[2], treeview.item(item, 'values')[3], treeview.item(item, 'values')[4], new_status))
            if show_background_color:
                tag = 'in_progress' if new_status == "进行中" else 'completed' if new_status == "已完成" else 'not_completed'
                treeview.item(item, tags=(tag,))
            modified = True  # 修改状态后设置为True
            status_str_ver.set('当前修改未保存')
    else:
        messagebox.showinfo("提示", "请先选择一行")

def on_closing():
    if modified:  # 如果有修改未保存
        response = messagebox.askyesnocancel("保存", "是否保存当前修改？")
        if response is True:  # 点击“是”
            save_file()
            root.destroy()
        elif response is False:  # 点击“否”
            root.destroy()
        else:  # 点击“取消”
            pass
    else:
        root.destroy()

def prompt_to_save_if_modified():
    if modified:
        response = messagebox.askyesnocancel("保存", "当前文件有未保存的修改，是否保存？")
        if response is True:  # 点击“是”
            save_file()
        elif response is False:  # 点击“否”
            pass
        else:  # 点击“取消”
            return False  # 返回False表示操作被取消
    return True  # 返回True表示可以继续打开新文件

def clear_treeview():
    for i in treeview.get_children():
        treeview.delete(i)

def open_online_list():
    if not prompt_to_save_if_modified():  # 检查是否有未保存的修改，并提示保存
        return  # 如果用户取消操作，则返回，不继续打开在线列表
    url = simpledialog.askstring("输入", "请输入STI文件的直链:")
    if url:
        try:
            response = requests.get(url)
            response.raise_for_status()  # 确保请求成功
            data = response.json()
            clear_treeview()
            for item in data:
                item_name, quantity, boxes, groups, pieces, completion_status = item
                tag = 'completed' if completion_status == "已完成" else 'not_completed'
                treeview.insert('', 'end', values=[item_name, quantity, boxes, groups, pieces, completion_status], tags=(tag,))
            status_str_ver.set(f'在线列表成功打开 {url}')
        except requests.RequestException as e:
            messagebox.showerror("错误", f"下载文件时出错: {e}")
            status_str_ver.set('下载文件失败')
        except json.JSONDecodeError as e:
            messagebox.showerror("错误", f"解析文件时出错: {e}")
            status_str_ver.set('解析文件失败')

def load_config():
    config = configparser.ConfigParser()
    config_path = 'config.ini'
    if not os.path.exists(config_path):
        # 如果配置文件不存在，创建一个默认的配置文件
        config['DEFAULT'] = {'last_opened_file': ''}
        with open(config_path, 'w') as configfile:
            config.write(configfile)
    else:
        # 读取配置文件
        config.read(config_path)
    return config, config_path

def save_config(config, config_path, last_opened_file):
    config['DEFAULT']['last_opened_file'] = last_opened_file
    with open(config_path, 'w') as configfile:
        config.write(configfile)

def open_last_file(config):
    last_opened_file = config['DEFAULT']['last_opened_file']
    if last_opened_file and os.path.exists(last_opened_file):
        open_file_path(last_opened_file)
    else:
        messagebox.showinfo("提示", "没有找到上次打开的文件")

def open_file_path(file_path):
    if file_path.endswith('.sti'):
        try:
            with open(file_path, 'r', encoding='utf-8') as file:
                data = json.load(file)
                clear_treeview()
                for item in data:
                    item_name, quantity, boxes, groups, pieces, completion_status = item
                    tag = 'completed' if completion_status == "已完成" else 'not_completed'
                    treeview.insert('', 'end', values=[item_name, quantity, boxes, groups, pieces, completion_status], tags=(tag,))
            status_str_ver.set(f'文件成功打开 {file_path}')
            global modified
            modified = False  # 打开文件后设置为False
        except Exception as e:
            messagebox.showerror("错误", f"读取文件时出错: {e}")
            status_str_ver.set('读取文件失败')
    else:
        try:
            with open(file_path, 'rb') as f:
                result = chardet.detect(f.read())
                encoding = result['encoding']
            with open(file_path, mode='r', newline='', encoding=encoding) as file:
                reader = csv.reader(file)
                next(reader)  # 跳过原始表头
                clear_treeview()
                for row in reader:
                    item_name, quantity = row[0], row[1]
                    boxes, groups, pieces = calculate_boxes_and_groups(int(quantity))
                    treeview.insert('', 'end', values=[item_name, quantity, boxes, groups, pieces, "未完成"], tags=('not_completed',))
            status_str_ver.set(f'文件成功打开 {file_path}')
            modified = False  # 打开文件后设置为False
        except Exception as e:
            messagebox.showerror("错误", f"读取文件时出错: {e}")
            status_str_ver.set('读取文件失败')


def open_settings():
    global show_background_color
    settings_window = tk.Toplevel(root)
    settings_window.title('设置')
    settings_window.geometry('300x200')

    show_background_color_var = tk.BooleanVar(value=show_background_color)

    def save_settings():
        global show_background_color
        show_background_color = show_background_color_var.get()
        update_treeview_background()
        settings_window.destroy()

    tk.Label(settings_window, text='显示背景色').pack(pady=10)
    tk.Checkbutton(settings_window, variable=show_background_color_var).pack()
    tk.Button(settings_window, text='保存', command=save_settings).pack(pady=10)


def update_treeview_background():
    for item in treeview.get_children():
        completion_status = treeview.item(item, 'values')[5]
        if show_background_color:
            tag = 'completed' if completion_status == "已完成" else 'not_completed'
        else:
            tag = ''
        treeview.item(item, tags=(tag,))

root = tk.Tk()
root.title('原理图材料列表查看器')
root.geometry('800x547')

modified = False
show_background_color = True  # 默认显示背景色

menu = tk.Menu(root, tearoff=False)
file_menu = tk.Menu(menu, tearoff=False)
file_menu.add_command(label='打开', command=open_file)
file_menu.add_command(label='打开上一次的文件', command=lambda: open_last_file(load_config()[0]))
file_menu.add_command(label='打开在线列表', command=open_online_list)
file_menu.add_command(label='保存', command=save_file)
menu.add_cascade(label='文件', menu=file_menu)

editor_menu = tk.Menu(menu, tearoff=False)
status_menu_editor = tk.Menu(editor_menu, tearoff=False)
status_menu_editor.add_command(label="未完成", command=lambda: change_status("未完成"))
status_menu_editor.add_command(label="进行中", command=lambda: change_status("进行中"))
status_menu_editor.add_command(label="已完成", command=lambda: change_status("已完成"))
editor_menu.add_cascade(label="修改状态", menu=status_menu_editor)
menu.add_cascade(label='编辑', menu=editor_menu)

settings_menu = tk.Menu(menu, tearoff=False)
settings_menu.add_command(label='设置', command=open_settings)
menu.add_cascade(label='设置', menu=settings_menu)

about_menu = tk.Menu(menu, tearoff=False)
about_menu.add_command(label='关于', command=about)
menu.add_cascade(label='关于', menu=about_menu)

status_str_ver = tk.StringVar()
status_str_ver.set('请打开一个文件')
status_label = tk.Label(root, textvariable=status_str_ver, bd=1, relief=tk.SUNKEN, anchor=tk.W)
status_label.pack(side=tk.BOTTOM, fill=tk.X)

# 创建Treeview控件
treeview = ttk.Treeview(root, columns=("物品名", "数量", "盒数", "组数", "个数", "是否完成"), show='headings')
treeview.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

# 创建Scrollbar控件
scrollbar = ttk.Scrollbar(treeview, orient=tk.VERTICAL, command=treeview.yview)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

# 将Treeview与Scrollbar关联
treeview.configure(yscrollcommand=scrollbar.set)

# 初始化表头
custom_headers = ["物品名", "数量", "盒数", "组数", "个数", "是否完成"]
for header in custom_headers:
    treeview.heading(header, text=header)
    treeview.column(header, width=100, stretch=True)  # 列宽自适应

# 定义标签样式
treeview.tag_configure('completed', background='#a9d08e')  # 绿色
treeview.tag_configure('not_completed', background='#ff8181')  # 红色
treeview.tag_configure('in_progress', background='#ffd299')  # 进行中，

# 创建右键菜单
right_click_menu = tk.Menu(treeview, tearoff=False)
status_menu = tk.Menu(right_click_menu, tearoff=False)  # 创建二级菜单
right_click_menu.add_cascade(label="修改状态", menu=status_menu)
status_menu.add_command(label="未完成", command=lambda: change_status("未完成"))
status_menu.add_command(label="进行中", command=lambda: change_status("进行中"))
status_menu.add_command(label="已完成", command=lambda: change_status("已完成"))
# 更新treeview背景色
update_treeview_background()

# 绑定右键点击事件
treeview.bind("<Button-3>", on_right_click)

# 绑定窗口关闭事件
root.protocol("WM_DELETE_WINDOW", on_closing)

root.config(menu=menu)
root.mainloop()
