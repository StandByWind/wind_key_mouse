import time
import tkinter as tk
from tkinter import ttk, filedialog
from pynput.mouse import Listener as MouseListener
import os
from tkinter import messagebox
import ast
from pynput.mouse import Controller


data_folder = 'data_folder'
left_options = ["识别到__时左键点击","识别到__时右键点击","识别到__时滚动鼠标","识别到__时按下左键，移动后再松开左键","识别到__按下__键再松开","识别到__一直按下__键","识别到__松开__键","识别到__等待"]
mouse_listener = None
scroll_y = 0
change_x = 0
change_y = 0
png_list = []
key_list = []
action_list = []
action_pairs = []

def save_data():

    try:
        with open(os.path.join(data_folder, 'COM.data'), 'wb') as f:
           f.write(str(COM_entry.get()).encode())
    except Exception as e:
        messagebox.showinfo('提示', f"保存COM.data出错，错误信息: {e}")

    try:
        with open(os.path.join(data_folder, 'circle.data'), 'wb') as f:
           f.write(str(circle_entry.get()).encode())
    except Exception as e:
        messagebox.showinfo('提示', f"保存circle.data出错，错误信息: {e}")

    try:
        with open(os.path.join(data_folder, 'png_list.data'), 'wb') as f:
           f.write(str(png_list).encode())
    except Exception as e:
        messagebox.showinfo('提示', f"保存png_list.data出错，错误信息: {e}")

    try:
        with open(os.path.join(data_folder, 'key_list.data'), 'wb') as f:
           f.write(str(key_list).encode())
    except Exception as e:
        messagebox.showinfo('提示', f"保存key_list.data出错，错误信息: {e}")

    try:
        with open(os.path.join(data_folder, 'action_list.data'), 'wb') as f:
           f.write(str(action_list).encode())
    except Exception as e:
        messagebox.showinfo('提示', f"保存action_list.data出错，错误信息: {e}")

def clear_frame(frame):
    for widget in frame.winfo_children():
        widget.destroy()

def open_file_dialog(pos):
    global png_list
    file_path = filedialog.askopenfilename()
    parts = file_path.split('/')
    result = '/'.join(parts[-2:])
    png_list[pos] = result

def save_key(pos, action_entry):
    global key_list
    key = action_entry.get()
    key_list[pos] = key

def save_change_x_y(action_entry, _action_entry):
    global change_x, change_y
    change_x = action_entry.get()
    change_y = _action_entry.get()
    change_list = [change_x, change_y]
    try:
        with open(os.path.join(data_folder, 'change_x_y.data'), 'wb') as f:
            f.write(str(change_list).encode())
    except Exception as e:
        messagebox.showinfo('提示', f"出错，错误信息: {e}")

def on_scroll(x, y, dx, dy):
    global scroll_y
    scroll_y += dy

def start_listening():
    global mouse_listener, scroll_y
    scroll_y = 0
    mouse_listener = MouseListener(on_scroll=on_scroll)
    mouse_listener.start()

def stop_listening():
    global mouse_listener
    mouse_listener.stop()
    mouse_listener.join()
    try:
        with open(os.path.join(data_folder, 'mouse_scroll.data'), 'wb') as f:
            f.write(str(scroll_y).encode())
    except Exception as e:
        messagebox.showinfo('提示', f"出错，错误信息: {e}")

def test_scroll():
    global scroll_y
    try:
        with open(os.path.join(data_folder, 'mouse_scroll.data'), 'rb') as f:
            scroll_y = ast.literal_eval(f.read().decode())
        messagebox.showinfo('提示', f"请立刻切换到测试界面，1秒后开始测试")
        time.sleep(1)
        mouse = Controller()
        mouse.scroll(0, scroll_y)
    except Exception as e:
        messagebox.showinfo('提示', f"读取mouse_scroll.data文件出错，错误信息: {e}")

def create_file_button(frame):
    right_file_button = tk.Button(frame, text="选择文件")
    right_file_button.grid(row=frame.grid_info()["row"], column=1, padx=10, pady=5)
    right_file_button.config(command=lambda: open_file_dialog(frame.grid_info()["row"]))

def create_key_input(frame):
    action_label = tk.Label(frame, text="请输入按键名称")
    action_label.grid(row=frame.grid_info()["row"], column=2, padx=5, pady=5)
    action_entry_ = tk.Entry(frame, width=10)
    action_entry_.grid(row=frame.grid_info()["row"], column=3, padx=5, pady=5)
    right_file_button_ = tk.Button(frame, text="保存")
    right_file_button_.grid(row=frame.grid_info()["row"], column=4, padx=5, pady=5)
    right_file_button_.config(command=lambda: save_key(frame.grid_info()["row"], action_entry_))

def create_change_input(frame):
    action_label = tk.Label(frame, text="请输入移动的x值(没有填0)")
    action_label.grid(row=frame.grid_info()["row"], column=2, padx=5, pady=5)
    action_entry_ = tk.Entry(frame, width=10)
    action_entry_.grid(row=frame.grid_info()["row"], column=3, padx=5, pady=5)
    _action_label = tk.Label(frame, text="请输入移动的y值(没有填0)")
    _action_label.grid(row=frame.grid_info()["row"], column=4, padx=5, pady=5)
    _action_entry_ = tk.Entry(frame, width=10)
    _action_entry_.grid(row=frame.grid_info()["row"], column=5, padx=5, pady=5)
    right_file_button_ = tk.Button(frame, text="保存")
    right_file_button_.grid(row=frame.grid_info()["row"], column=6, padx=5, pady=5)
    right_file_button_.config(command=lambda: save_change_x_y(action_entry_, _action_entry_))

def left_combobox_select(left_combobox, right_frame):
    selected = left_combobox.get()
    pos = right_frame.grid_info()["row"]
    clear_frame(right_frame)

    action_mapping = {
        "识别到__时左键点击": ("click_left", lambda: None),
        "识别到__时右键点击": ("click_right", lambda: None),
        "识别到__时滚动鼠标": ("scroll", lambda: [
            create_file_button(right_frame),
            tk.Label(right_frame, text="请按键记录滚动数据").grid(row=pos, column=2, padx=5, pady=5),
            tk.Button(right_frame, text="开始记录", command=start_listening).grid(row=pos, column=3, padx=5, pady=5),
            tk.Button(right_frame, text="结束记录", command=stop_listening).grid(row=pos, column=4, padx=5, pady=5),
            tk.Button(right_frame, text="测试", command=test_scroll).grid(row=pos, column=5, padx=5, pady=5)
        ]),
        "识别到__时按下左键，移动后再松开左键": ("move", lambda: create_change_input(right_frame)),
        "识别到__按下__键再松开": ("key_short", lambda: create_key_input(right_frame)),
        "识别到__一直按下__键": ("key_long", lambda: create_key_input(right_frame)),
        "识别到__松开__键": ("key_release", lambda: create_key_input(right_frame)),
        "识别到__等待": ("wait", lambda: None)
    }

    action, setup_func = action_mapping.get(selected, (None, None))
    if action:
        action_list[pos] = action
        create_file_button(right_frame)
        setup_func()

def add_action():
    try:
        pos = int(pos_entry.get())
        if 0 <= pos <= len(action_pairs):
            left_combobox = ttk.Combobox(action_frame, values=left_options)
            right_frame = ttk.Frame(action_frame)
            left_combobox.bind("<<ComboboxSelected>>", lambda event: left_combobox_select(left_combobox, right_frame))

            # 插入新行，将后续行下移
            for i in range(len(action_pairs) - 1, pos - 1, -1):
                old_left, old_right = action_pairs[i]
                old_left.grid(row=i + 1, column=0, padx=10, pady=5)
                old_right.grid(row=i + 1, column=1, padx=5, pady=5)

            left_combobox.grid(row=pos, column=0, padx=10, pady=5)
            right_frame.grid(row=pos, column=1, padx=5, pady=5)
            action_pairs.insert(pos, (left_combobox, right_frame))
            action_list.insert(pos, "")
            png_list.insert(pos, "")
            key_list.insert(pos, "")
        else:
            messagebox.showinfo('提示', f"输入的位置无效，请输入有效的位置")
    except ValueError:
        pos = len(action_pairs)
        left_combobox = ttk.Combobox(action_frame, values=left_options)
        right_frame = ttk.Frame(action_frame)
        right_frame.grid(row=pos, column=1, padx=5, pady=5)
        left_combobox.bind("<<ComboboxSelected>>", lambda event: left_combobox_select(left_combobox,right_frame))
        left_combobox.grid(row=pos, column=0, padx=10, pady=5)

        action_pairs.insert(pos, (left_combobox, right_frame))
        action_list.insert(pos, "")
        png_list.insert(pos, "")
        key_list.insert(pos, "")

def delete_action():
    try:
        pos = int(pos_entry.get())
        if 0 <= pos <= len(action_pairs) - 1:
            left_combobox, right_frame = action_pairs.pop(pos)
            left_combobox.destroy()
            right_frame.destroy()

            action_list.pop(pos)
            png_list.pop(pos)
            key_list.pop(pos)

            for i in range(pos, len(action_pairs)):
                left, button = action_pairs[i]
                left.grid(row=i, column=0, padx=10, pady=5)
                button.grid(row=i, column=1, padx=10, pady=5)

        else:
            messagebox.showinfo('提示', f"输入的位置无效，请输入有效的位置")
    except ValueError:
        if len(action_pairs) > 0:
            pos = len(action_pairs) - 1
            left_combobox, right_frame = action_pairs.pop(pos)

            action_list.pop(pos)
            png_list.pop(pos)
            key_list.pop(pos)

            left_combobox.destroy()
            right_frame.destroy()
        else:
            messagebox.showinfo('提示', f"删完了")

def show_action_menu():
    add_action_window.deiconify()

def create_entry(label_text):
    frame = tk.Frame(add_action_window)
    frame.pack(pady=5, side=tk.TOP)

    label = tk.Label(frame, text=label_text)
    label.pack(pady=5, side=tk.LEFT)

    entry = tk.Entry(frame, width=10)
    entry.pack(pady=5, side=tk.RIGHT)

    return entry


add_action_window = tk.Tk(className='add_action_window')

pos_entry = create_entry("请输入位置 (从 0 开始,若空则在末尾添加或删除):")

circle_entry = create_entry("输入循环次数(不循环输入1):")

COM_entry = create_entry("若选择硬件执行请输入端口号，其他执行输入0")

button_frame = tk.Frame(add_action_window)
button_frame.pack(pady=5, side=tk.TOP)
add_button = tk.Button(button_frame, text="在指定位置添加", command=add_action)
add_button.pack(side = tk.LEFT, padx=10)
delete_button = tk.Button(button_frame, text="删除指定位置", command=delete_action)
delete_button.pack(side = tk.RIGHT, padx=10)

save_frame = tk.Frame(add_action_window)
save_frame.pack(pady=5, side=tk.TOP)
save_button = tk.Button(save_frame, text="保存", command=save_data)
save_button.pack(pady=5)

action_frame = tk.Frame(add_action_window)
action_frame.pack(pady=5, side=tk.TOP)

def window_close():
    add_action_window.withdraw()

add_action_window.protocol('WM_DELETE_WINDOW', window_close)
add_action_window.withdraw()
