from tkinter import Tk, Canvas
import mss
import numpy as np
import cv2
import os
from PIL import Image
from tkinter.messagebox import askyesno as ask
from tkinter import messagebox


pos = [0]
mode = 0
idx = 0
data_folder = 'data_folder'# 指定保存data文件的文件夹路径
png_folder = 'png_folder'# 指定保存png文件的文件夹路径


# 创建文件夹
try:
    os.mkdir(data_folder)# 创建文件夹
    messagebox.showinfo('提示', f"文件夹 {data_folder} 创建成功,在程序的同目录下")
except Exception as e:
    if not isinstance(e, FileExistsError):
        messagebox.showinfo('提示',f"创建文件夹{data_folder} 失败，错误信息: {e}")

try:
    os.mkdir(png_folder)# 创建文件夹
    messagebox.showinfo('提示', f"文件夹 {png_folder} 创建成功,在程序的同目录下")
except Exception as e:
    if not isinstance(e, FileExistsError):
        messagebox.showinfo('提示',f"创建文件夹{png_folder} 失败，错误信息: {e}")


def grab_big_zone():
    global mode
    mode = 0
    grab_window.deiconify()

def grab_small_zone():
    global mode
    mode = 5
    grab_window.deiconify()

def save_small_zone():
    try:
        with open(os.path.join(data_folder, 'small_zone_num.data'), 'wb') as f_:
            f_.write(str(idx).encode())
    except Exception as e_:
        messagebox.showinfo("提示", f"标志区域图片data文件保存失败，错误信息: {e_}")

def delete_png_files():
    global idx
    png_list = os.listdir(png_folder)
    for png_file in png_list:
        file_path = os.path.join(png_folder, png_file)
        try:
            os.remove(file_path)
        except Exception as e_:
            messagebox.showinfo("提示",f"删除文件 {file_path} 时出错: {e_}")
    idx = 0

def check_small_zone_num_data():
    global idx
    try:
        with open(os.path.join(data_folder, 'small_zone_num.data'), 'rb') as f_:
            flag_ = ask('提示', '找到标志区域data文件，是否使用原配置')
            if flag_:
                idx = int(f_.read().decode())
            else:
                messagebox.showinfo("提示", "将删除原标志区域图片！")
                delete_png_files()
                messagebox.showinfo('提示', '请将标志移动到不遮挡的位置, 点击确定后选择标志区域')
                grab_small_zone()
    except Exception as e_:
        if isinstance(e_, FileNotFoundError):
            messagebox.showinfo('提示', '请选择标志区域')
            grab_small_zone()
        else:
            messagebox.showinfo('提示', f"出错，错误信息: {e_}")

def mouse_press(event):
    global mode
    if mode == 0 or mode == 5:
        pos[1], pos[2] = event.x, event.y
        mode += 1

def mouse_move(event):
    if mode == 1 or mode == 6:
        if pos[0]:
            canvas.delete('all')#保证只出现一个矩形
        pos[3], pos[4] = event.x, event.y
        pos[0] = canvas.create_rectangle(pos[1], pos[2], pos[3], pos[4], fill='black')

def mouse_release(event):
    global idx
    pos[3], pos[4] = event.x, event.y
    if pos[1] > pos[3]:
        pos[1], pos[3] = pos[3], pos[1]
    if pos[2] > pos[4]:
        pos[2], pos[4] = pos[4], pos[2]
    grab_window.withdraw()
    zone = {"left": pos[1], "top": pos[2], "width": pos[3] - pos[1], "height": pos[4] - pos[2]}
    with mss.mss() as sct:
        if mode == 1:
            with open(os.path.join(data_folder, 'grab_window_position.data'), 'wb') as f_:
                f_.write(str(pos[1:])[1:-1].encode())
            png = sct.grab(zone)
            _png = np.asarray(png)
            cv2.imshow("Preview", _png)#显示预览图
            cv2.waitKey(0)
            cv2.destroyAllWindows()
            check_small_zone_num_data()
        elif mode == 6:
            png = sct.grab(zone)
            _png = np.asarray(png)
            cv2.imshow("Preview", _png)# 显示预览图
            cv2.waitKey(0)
            cv2.destroyAllWindows()
            _png_ = Image.fromarray(_png)# 将截图转换为 Pillow 图像对象
            image_path = os.path.join(png_folder, f"{idx}.png")# 生成保存图片的完整路径
            _png_.save(image_path)
            idx += 1
            save_small_zone()

grab_window = Tk(className='grab_window')
screen_width = grab_window.winfo_screenwidth()
screen_height = grab_window.winfo_screenheight()
canvas = Canvas(grab_window, width=screen_width, height=screen_height, bg='white')
canvas.pack(fill='both')  # 将 canvas 组件放置在其父容器中，并让它在水平和垂直方向上都填充所分配的空间
grab_window.attributes('-alpha', 0.3)  # 透明度
grab_window.attributes('-fullscreen', 1)  # 窗口最大化
grab_window.attributes('-topmost', 1)  # 窗口顶置 todo debug时请注释这一行
grab_window.config(cursor='crosshair')  # 将窗口鼠标指针样式设置为十字准线
grab_window.bind('<Button-1>', mouse_press)
grab_window.bind('<Motion>', mouse_move)
grab_window.bind('<ButtonRelease-1>', mouse_release)
grab_window.withdraw()

try:
    with open(os.path.join(data_folder, 'grab_window_position.data'), 'rb') as f:
        flag = ask('提示', '找到窗口位置文件，是否使用原配置')
        if flag:
            pos.extend(int(i) for i in f.read().decode().split(', '))
            check_small_zone_num_data()
        else:
            messagebox.showinfo('提示', '请将窗口移动到不遮挡的位置, 点击确定后选择区域')
            pos = [0, 0, 0, 0, 0]
            grab_big_zone()
except Exception as e:
    if isinstance(e, FileNotFoundError):
        messagebox.showinfo('提示', '请将窗口移动到不遮挡的位置, 点击确定后选择区域')
        pos = [0,0,0,0,0]
        grab_big_zone()
    else:
        messagebox.showinfo('提示', f"出错，错误信息: {e}")


