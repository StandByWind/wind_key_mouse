import tkinter as tk
from tkinter import Menu
import sys

import grab_window_program
import identify_program
import add_action_window_program


menu_window = tk.Tk(className='menu')
screen_width = grab_window_program.screen_width
menu_width = int(screen_width/3)
menu_height = 10
menu_window.geometry(f'{menu_width}x{menu_height}+{menu_width}+0')#设置菜单大小和位置
menu_window.resizable(True, True)
menu_window.attributes('-toolwindow', 1)#将窗口设置为工具窗口
menu_window.attributes('-topmost', 1)
menu_window.protocol('WM_DELETE_WINDOW', sys.exit)#设置关闭菜单时程序退出
menu = Menu(menu_window)#创建菜单
menu.add_command(label='选择区域', command=grab_window_program.grab_big_zone)
menu.add_command(label='选择标志区域', command=grab_window_program.grab_small_zone)
menu.add_command(label='创建操作组', command=add_action_window_program.show_action_menu)
menu.add_command(label='展示识别结果', command=identify_program.show_result)
menu.add_command(label='window执行', command=lambda : identify_program.start_all(1))
menu.add_command(label='安卓执行', command=lambda : identify_program.start_all(2))
menu.add_command(label='硬件执行', command=lambda : identify_program.start_all(3))
menu_window.config(menu=menu)#将一个已经创建好的菜单对象关联到指定的窗口上
menu_window.mainloop()


#pyinstaller -F -w

