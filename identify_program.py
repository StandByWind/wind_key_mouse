import mss
import cv2
import numpy as np
import os
import ast
import time
import pyautogui
from tkinter import messagebox
from pynput.mouse import Controller
import threading
import subprocess

from PyComKB import SimuComKB


png_folder = 'png_folder'
data_folder = 'data_folder'
_png = None
mode = 0
location_list = []
threads = []
location_list_lock = threading.Lock()  # 创建线程锁
stop_event = threading.Event()


def search_small_zone(background, picture,fix_x,fix_y):
    global location_list, location_list_lock
    target = cv2.imread(picture)
    _target = cv2.cvtColor(target, cv2.COLOR_BGR2GRAY)
    result = cv2.matchTemplate(background, _target, cv2.TM_CCOEFF_NORMED)
    locations = np.where(result >= 0.85)
    locations = list(zip(*locations[::-1]))
    if len(locations) > 0:
        with location_list_lock:  # 使用线程锁保护共享资源
            for location in locations:
                location_list.append((location[0]+int(fix_x), location[1]+int(fix_y)))

def identify_target(big_zone, small_zone):
    global location_list, threads
    location_list = []
    background = cv2.cvtColor(big_zone, cv2.COLOR_BGR2GRAY)
    fix_list = small_zone.split("_")
    search_small_zone(background, small_zone,fix_list[2],fix_list[3])

def show_result():
    global _png
    pos = [0]
    with open(os.path.join(data_folder, 'grab_window_position.data'), 'rb') as f:
        pos.extend(int(i) for i in f.read().decode().split(', '))
    zone = {"left": pos[1], "top": pos[2], "width": pos[3], "height": pos[4]}
    with mss.mss() as sct:
        while True:
            png = sct.grab(zone)
            _png = np.array(png)
            png_path_list = os.listdir(png_folder)
            for i in range(len(png_path_list)):
                png_path_list[i] = os.path.join(png_folder, png_path_list[i])
            for picture in png_path_list:
                thread = threading.Thread(target=identify_target,args=(_png, picture))
                thread.start()
                threads.append(thread)
            for thread in threads:
                thread.join()
            for location in location_list:
                cv2.circle(_png, (location[0], location[1]), 5, (0, 0, 255), thickness=-1)
            cv2.imshow('press space to exit', _png)
            if cv2.waitKey(5) & 0xFF == ord(' '):
                cv2.destroyAllWindows()
                break

def start_identify():
    global _png
    pos = [0]
    with open(os.path.join(data_folder, 'grab_window_position.data'), 'rb') as f:
        pos.extend(int(i) for i in f.read().decode().split(', '))
    zone = {"left": pos[1], "top": pos[2], "width": pos[3],"height": pos[4]}
    with mss.mss() as sct:
        while not stop_event.is_set():
            png = sct.grab(zone)
            _png = np.array(png)

def read_circle():
    try:
        with open(os.path.join(data_folder, 'circle.data'), 'rb') as f:
            circle = ast.literal_eval(f.read().decode())
            return circle
    except Exception as e:
        messagebox.showinfo('提示', f"读取circle.data文件出错，错误信息: {e}")

def read_action_list():
    try:
        with open(os.path.join(data_folder, 'action_list.data'), 'rb') as f:
            action_list = ast.literal_eval(f.read().decode())
            return action_list
    except Exception as e:
        messagebox.showinfo('提示', f"读取action_list.data文件出错，错误信息: {e}")
        return None

def read_png_list():
    try:
        with open(os.path.join(data_folder, 'png_list.data'), 'rb') as f:
            png_list = ast.literal_eval(f.read().decode())
            return png_list
    except Exception as e:
        messagebox.showinfo('提示', f"读取png_list.data文件出错，错误信息: {e}")
        return None

def read_key_list():
    try:
        with open(os.path.join(data_folder, 'key_list.data'), 'rb') as f:
            key_list = ast.literal_eval(f.read().decode())
            return key_list
    except Exception as e:
        messagebox.showinfo('提示', f"读取key_list.data文件出错，错误信息: {e}")
        return None

def read_change_x_y():
    try:
        with open(os.path.join(data_folder, 'key_list.data'), 'rb') as f:
            change_x_y = ast.literal_eval(f.read().decode())
            change_x = change_x_y[0]
            change_y = change_x_y[1]
            return change_x, change_y
    except Exception as e:
        messagebox.showinfo('提示', f"读取key_list.data文件出错，错误信息: {e}")
        return None

def read_mouse_scroll():
    try:
        with open(os.path.join(data_folder, 'mouse_scroll.data'), 'rb') as f:
            scroll_y = ast.literal_eval(f.read().decode())
            return scroll_y
    except Exception as e:
        messagebox.showinfo('提示', f"读取mouse_scroll.data文件出错，错误信息: {e}")
        return None

def read_COM():
    try:
        with open(os.path.join(data_folder, 'COM.data'), 'rb') as f_:
            COM = ast.literal_eval(f_.read().decode())
        InputKey = SimuComKB(f"COM{COM}")
        return InputKey
    except Exception as e_:
        messagebox.showinfo('提示', f"读取COM.data文件出错，错误信息: {e_}")
        return None

def wait_for_location(picture):
    while not location_list:
        time.sleep(0.5)
        identify_target(_png, picture)

def start_action():
    action_list = read_action_list()
    png_list = read_png_list()
    if mode == 3:
        InputKey = read_COM()
    pos = [0]
    with open(os.path.join(data_folder, 'grab_window_position.data'), 'rb') as f:
        pos.extend(int(i) for i in f.read().decode().split(', '))
    for i in range (len(action_list)):
        identify_target(_png, png_list[i])
        if action_list[i] == "click_left":
            wait_for_location(png_list[i])
            if mode == 1:
                pyautogui.moveTo(pos[1] + location_list[0][0], pos[2] + location_list[0][1])
                pyautogui.click()
            elif mode == 2:
                try:
                    subprocess.run(f"adb shell input tap {location_list[0][0]} {location_list[0][1]}", shell=True, check=True)
                except subprocess.CalledProcessError as e:
                    messagebox.showinfo("提示",f"执行点击命令时出错: {e}")
            elif mode == 3:
                InputKey.M_MoveTo(pos[1] + location_list[0][0], pos[2] + location_list[0][1])
                InputKey.M_BClick("l")
            time.sleep(0.5)

        elif action_list[i] == "click_right":
            wait_for_location(png_list[i])
            if mode == 1:
                pyautogui.moveTo(pos[1] + location_list[0][0], pos[2] + location_list[0][1])
                pyautogui.rightClick()
            elif mode == 2:
                pass
            elif mode == 3:
                InputKey.M_MoveTo(pos[1] + location_list[0][0], pos[2] + location_list[0][1])
                InputKey.M_BClick("r")
            time.sleep(0.5)

        elif action_list[i] == "scroll":
            scroll_y = read_mouse_scroll()
            wait_for_location(png_list[i])
            if mode == 1:
                mouse = Controller()
                mouse.scroll(0, scroll_y)
            elif mode == 2:
                pass
            elif mode == 3:
                time.sleep(1)
                InputKey.scroll(scroll_y)
            time.sleep(0.5)

        elif action_list[i] == "move":
            change_x, change_y = read_change_x_y()
            wait_for_location(png_list[i])
            if mode == 1:
                pyautogui.moveTo(pos[1] + location_list[0][0], pos[2] + location_list[0][1])
                pyautogui.mouseDown()
                pyautogui.moveTo(pos[1] + location_list[0][0] + change_x, pos[2] + location_list[0][1] + change_y)
                pyautogui.mouseUp()
            elif mode == 2:
                try:
                    subprocess.run(f"adb shell input swipe {location_list[0][0]} {location_list[0][1]} {location_list[0][0] + change_x} {location_list[0][1] + change_y} {1000}", shell=True, check=True)
                except subprocess.CalledProcessError as e:
                    messagebox.showinfo("提示",f"执行滑动命令时出错: {e}")
            elif mode == 3:
                InputKey.M_MoveTo(pos[1] + location_list[0][0], pos[2] + location_list[0][1])
                InputKey.M_BDown("l")
                InputKey.M_MoveR(change_x, change_y)
                InputKey.M_BUp()
            time.sleep(0.5)

        elif action_list[i] == "key_short":
            key_list = read_key_list()
            wait_for_location(png_list[i])
            if mode == 1:
                pyautogui.press(key_list[i])
            elif mode == 2:
                pass
            elif mode == 3:
                InputKey.KB_Click(key_list[i])
            time.sleep(0.5)

        elif action_list[i] == "key_long":
            key_list = read_key_list()
            wait_for_location(png_list[i])
            if mode == 1:
                pyautogui.keyDown(key_list[i])
            elif mode == 2:
                pass
            elif mode == 3:
                InputKey.KB_Down(key_list[i])
            time.sleep(0.5)

        elif action_list[i] == "key_release":
            key_list = read_key_list()
            wait_for_location(png_list[i])
            if mode == 1:
                pyautogui.keyUp(key_list[i])
            elif mode == 2:
                pass
            elif mode == 3:
                InputKey.KB_Up()
            time.sleep(0.5)

        elif action_list[i] == "wait":
            while location_list:
                time.sleep(0.5)

def start_all(x):
    global mode
    mode = x

    stop_event.clear()

    identify_thread = threading.Thread(target=start_identify)
    identify_thread.start()

    time.sleep(0.5)

    for i in range(read_circle()):
        action_thread = threading.Thread(target=start_action)
        action_thread.start()
        action_thread.join()

    stop_event.set()

    identify_thread.join()

