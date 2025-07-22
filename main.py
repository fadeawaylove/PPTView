import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import json
import subprocess
import sys


home_dir = os.path.expanduser('~')
app_data_dir = os.path.join(home_dir, '.PPTView')
os.makedirs(app_data_dir, exist_ok=True)

TAG_FILE = os.path.join(app_data_dir, 'tags.json')
CONFIG_FILE = os.path.join(app_data_dir, 'config.json')


def load_tags():
    if os.path.exists(TAG_FILE):
        with open(TAG_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}


def save_tags():
    with open(TAG_FILE, 'w', encoding='utf-8') as f:
        json.dump(tags, f, ensure_ascii=False, indent=2)


def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            return {}
    return {}


def save_config(cfg):
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


def scan_ppt_files(folder):
    ppt_files = []
    for root_dir, _, files in os.walk(folder):
        for f in files:
            if f.lower().endswith(('.ppt', '.pptx')):
                ppt_files.append(os.path.join(root_dir, f))
    return ppt_files


def choose_folder():
    folder = filedialog.askdirectory()
    if folder:
        config['last_folder'] = folder
        save_config(config)
        update_file_list(folder)


def update_file_list(folder, filename_keywords=None, tag_keywords=None):
    for row in file_tree.get_children():
        file_tree.delete(row)

    global current_folder, displayed_files
    current_folder = folder
    displayed_files = []

    ppt_files = scan_ppt_files(folder)

    for file in ppt_files:
        filename = os.path.basename(file)
        tag_list = tags.get(file, [])

        if filename_keywords:
            if not all(k.lower() in filename.lower() for k in filename_keywords):
                continue
        if tag_keywords:
            if not all(k in tag_list for k in tag_keywords):
                continue

        displayed_files.append(file)
        tag_str = " ".join(tag_list)
        file_tree.insert('', tk.END, values=(filename, tag_str))

    status_var.set(f"共找到 {len(displayed_files)} 个文件（已过滤）")


def get_selected_file():
    selection = file_tree.selection()
    if not selection:
        return None
    index = file_tree.index(selection[0])
    return displayed_files[index] if 0 <= index < len(displayed_files) else None


def add_tag():
    file_path = get_selected_file()
    if not file_path:
        return
    tag_input = simpledialog.askstring("添加标签", f"为文件添加标签（空格分隔）:")
    if tag_input:
        tag_list = tag_input.strip().split()
        tags[file_path] = list(set(tags.get(file_path, []) + tag_list))
        save_tags()
        on_search_event()


def delete_tag():
    file_path = get_selected_file()
    if not file_path:
        return
    if file_path not in tags:
        messagebox.showinfo("提示", "此文件没有任何标签。")
        return
    tag_input = simpledialog.askstring("删除标签",
                                       f"当前标签：{' '.join(tags[file_path])}\n请输入要删除的标签（空格分隔）:")
    if tag_input:
        to_delete = tag_input.strip().split()
        tags[file_path] = [t for t in tags[file_path] if t not in to_delete]
        if not tags[file_path]:
            tags.pop(file_path)
        save_tags()
        on_search_event()


def get_filename_keywords():
    query = filename_search_entry.get().strip()
    return query.split() if query else None


def get_tag_keywords():
    query = tag_search_entry.get().strip()
    return query.split() if query else None


def on_search_event(event=None):
    if current_folder:
        update_file_list(current_folder, get_filename_keywords(), get_tag_keywords())


def reset_search():
    filename_search_entry.delete(0, tk.END)
    tag_search_entry.delete(0, tk.END)
    if current_folder:
        update_file_list(current_folder)


def on_right_click(event):
    row = file_tree.identify_row(event.y)
    if row:
        file_tree.selection_set(row)
        context_menu.post(event.x_root, event.y_root)


def open_ppt(filepath):
    try:
        if sys.platform == "win32":
            os.startfile(filepath)
        elif sys.platform == "darwin":
            subprocess.call(["open", filepath])
        else:
            subprocess.call(["xdg-open", filepath])
    except Exception as e:
        messagebox.showerror("打开失败", f"无法打开文件:\n{e}")


def on_double_click(event):
    item_id = file_tree.focus()
    if not item_id:
        return
    index = file_tree.index(item_id)
    if 0 <= index < len(displayed_files):
        open_ppt(displayed_files[index])


# 初始化数据和界面
tags = load_tags()
config = load_config()
current_folder = ""
displayed_files = []

root = tk.Tk()
root.title("PPT标签管理器")
width, height = 950, 550
root.geometry(f"{width}x{height}")

# 居中窗口
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
x = (screen_width - width) // 2
y = (screen_height - height) // 2
root.geometry(f"{width}x{height}+{x}+{y}")

top_frame = tk.Frame(root)
top_frame.pack(fill=tk.X)

tk.Button(top_frame, text="选择文件夹", command=choose_folder).pack(side=tk.LEFT, padx=5, pady=5)
tk.Button(top_frame, text="添加标签", command=add_tag).pack(side=tk.LEFT, padx=5, pady=5)
tk.Button(top_frame, text="删除标签", command=delete_tag).pack(side=tk.LEFT, padx=5, pady=5)

# 新增提示Label + Entry组合

# top_frame已经创建了
tk.Button(top_frame, text="重置搜索", command=reset_search).pack(side=tk.RIGHT, padx=5, pady=5)

# 标签搜索 Label 和 Entry
tag_label = tk.Label(top_frame, text="标签搜索:")
tag_label.pack(side=tk.LEFT, padx=(5, 2), pady=5)
tag_search_entry = tk.Entry(top_frame, width=30)
tag_search_entry.pack(side=tk.LEFT, padx=(0, 15), pady=5)
tag_search_entry.bind("<Return>", on_search_event)

# 文件名搜索 Label 和 Entry
filename_label = tk.Label(top_frame, text="文件名搜索:")
filename_label.pack(side=tk.LEFT, padx=(5, 2), pady=5)
filename_search_entry = tk.Entry(top_frame, width=25)
filename_search_entry.pack(side=tk.LEFT, padx=(0, 10), pady=5)
filename_search_entry.bind("<Return>", on_search_event)

tree_frame = tk.Frame(root)
tree_frame.pack(fill=tk.BOTH, expand=True)

columns = ('文件名', '标签')
file_tree = ttk.Treeview(tree_frame, columns=columns, show='headings')
file_tree.heading('文件名', text='文件名', anchor='w')
file_tree.heading('标签', text='标签', anchor='w')
file_tree.column('文件名', width=350, anchor='w')
file_tree.column('标签', width=500, anchor='w')
file_tree.pack(fill=tk.BOTH, expand=True)

file_tree.bind("<Button-3>", on_right_click)
file_tree.bind("<Double-1>", on_double_click)

context_menu = tk.Menu(root, tearoff=0)
context_menu.add_command(label="添加标签", command=add_tag)
context_menu.add_command(label="删除标签", command=delete_tag)

status_var = tk.StringVar()
status_bar = tk.Label(root, textvariable=status_var, relief=tk.SUNKEN, anchor='w')
status_bar.pack(fill=tk.X, side=tk.BOTTOM)

# 启动时自动加载上次打开的文件夹（如果有效）
if 'last_folder' in config and os.path.isdir(config['last_folder']):
    current_folder = config['last_folder']
    update_file_list(current_folder)

root.mainloop()
