# 找到进程目录，打开它。然后置于活动前台。

import psutil
import os
import time
import pygetwindow as gw
import subprocess
import ctypes

import psutil

import win32gui
import win32con

def get_window_handle(process_name):
    # 获取当前置顶窗口的标题
    foreground_handle = win32gui.GetForegroundWindow()  # 获取当前置顶窗口的句柄
    foreground_title = win32gui.GetWindowText(foreground_handle)  # 获取当前置顶窗口的标题

    if foreground_title != process_name:  # 如果当前置顶窗口的标题不等于指定的进程名
        # 获取所有窗口的句柄和标题
        window_titles = {}

        def callback(hwnd, _):
            if win32gui.IsWindowVisible(hwnd):  # 判断窗口是否可见
                title = win32gui.GetWindowText(hwnd)  # 获取窗口的标题
                if title:
                    window_titles[hwnd] = title  # 将窗口的句柄和标题保存到字典中

        win32gui.EnumWindows(callback, None)  # 枚举所有窗口，调用回调函数获取窗口句柄和标题

        for hwnd, title in window_titles.items():  # 遍历保存的窗口句柄和标题
            if title == process_name:  # 如果窗口的标题等于指定的进程名
                if win32gui.IsIconic(hwnd):  # 如果窗口被最小化
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)  # 恢复窗口显示
                win32gui.SetForegroundWindow(hwnd)  # 将窗口设置为前台窗口
                break  # 找到指定窗口后终止循环
        else:
            return False  # 如果未找到指定进程名的窗口，返回 False

    return True  # 找到指定进程名的窗口，返回 True

if __name__ == "__main__":
    # 调用函数
    get_window_handle('微信')
