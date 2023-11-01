# 找到进程目录，打开它。然后置于活动前台。

import win32gui
import psutil
import os
import time
import pygetwindow as gw
import subprocess
import ctypes
import psutil
import win32gui
import win32con
import win32process
import psutil

def get_window_handle(process_title):
    # 获取当前置顶窗口的标题
    foreground_handle = win32gui.GetForegroundWindow()  # 获取当前置顶窗口的句柄
    foreground_title = win32gui.GetWindowText(foreground_handle)  # 获取当前置顶窗口的标题

    if foreground_title != process_title:  # 如果当前置顶窗口的标题不等于指定的进程名
        # 获取所有窗口的句柄和标题
        window_titles = {}

        def callback(hwnd, _):
            if win32gui.IsWindowVisible(hwnd):  # 判断窗口是否可见
                title = win32gui.GetWindowText(hwnd)  # 获取窗口的标题
                if title:
                    window_titles[hwnd] = title  # 将窗口的句柄和标题保存到字典中

        win32gui.EnumWindows(callback, None)  # 枚举所有窗口，调用回调函数获取窗口句柄和标题

        for hwnd, title in window_titles.items():  # 遍历保存的窗口句柄和标题
            if title == process_title:  # 如果窗口的标题等于指定的进程名
                if win32gui.IsIconic(hwnd):  # 如果窗口被最小化
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)  # 恢复窗口显示
                win32gui.SetForegroundWindow(hwnd)  # 将窗口设置为前台窗口
                break  # 找到指定窗口后终止循环
        else:
            return False  # 如果未找到指定进程名的窗口，返回 False

    return True  # 找到指定进程名的窗口，返回 True

# 获取置顶窗口的信息、句柄、进程id、进程名、进程标题
def get_foreground_window_info():
    foreground_handle = win32gui.GetForegroundWindow()  # 获取当前置顶窗口的句柄
    _, process_id = win32process.GetWindowThreadProcessId(foreground_handle)  # 获取当前置顶窗口的进程ID
    if process_id <= 0:
        return None, None, None, None  # 如果进程ID无效，返回 None

    try:
        process = psutil.Process(process_id)
        process_name = process.name()  # 获取当前置顶窗口的进程名
        window_title = win32gui.GetWindowText(foreground_handle)  # 获取当前置顶窗口的标题
    except (psutil.NoSuchProcess, psutil.AccessDenied):
        process_name = None
        window_title = None

    return foreground_handle, process_id, process_name, window_title


def get_window_handle_name(process_name):
    # 获取当前置顶窗口的标题
    foreground_handle = win32gui.GetForegroundWindow()  # 获取当前置顶窗口的句柄
    process_id = win32process.GetWindowThreadProcessId(foreground_handle)[1]
    process = psutil.Process(process_id)
    foreground_name = process.name()
    # print(foreground_name)

    if foreground_name != process_name:  # 如果当前置顶窗口的标题不等于指定的进程名
        # 获取所有窗口的句柄和标题
        window_names = {}
        def get_process_name(hwnd):
            _, process_id = win32process.GetWindowThreadProcessId(hwnd)
            try:
                process = psutil.Process(process_id)
                return process.name()
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                return None
        def callback(hwnd, _):
            if win32gui.IsWindowVisible(hwnd):  # 判断窗口是否可见
                name = get_process_name(hwnd)  # 获取窗口的标题
                if name:
                    window_names[hwnd] = name  # 将窗口的句柄和标题保存到字典中

        win32gui.EnumWindows(callback, None)  # 枚举所有窗口，调用回调函数获取窗口句柄和标题

        # print(window_names)
        for hwnd, name in window_names.items():  # 遍历保存的窗口句柄和标题
            if name == process_name:  # 如果窗口的标题等于指定的进程名
                if win32gui.IsIconic(hwnd):  # 如果窗口被最小化
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)  # 恢复窗口显示
                win32gui.SetForegroundWindow(hwnd)  # 将窗口设置为前台窗口
                break  # 找到指定窗口后终止循环
        else:
            return False  # 如果未找到指定进程名的窗口，返回 False

    return True  # 找到指定进程名的窗口，返回 True




if __name__ == "__main__":
    # 调用函数
    get_window_handle('原神')       # 这是标题
    get_window_handle_name('YuanShen.exe')  # 这是进程名
    time.sleep(1)
    handle, pid, name, title = get_foreground_window_info()
    if handle:
        print("窗口句柄:", handle)
        print("进程ID:", pid)
        print("进程名:", name)
        print("窗口标题:", title)
    else:
        print("未找到当前置顶窗口")