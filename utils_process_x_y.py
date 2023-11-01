from ctypes import windll
import pygetwindow as gw
from PIL import ImageGrab
import psutil

# 定义函数 get_width，接受窗口标题作为参数
def get_width(title):
    window = gw.getWindowsWithTitle(title)[0]  # 获取指定标题的窗口对象
    print(window)
    hwnd = window._hWnd  # 获取窗口的句柄

    # 获取活动窗口的大小
    window_rect = window.width, window.height

    user32 = windll.user32  # 加载 user32.dll 库
    desktop_width = user32.GetSystemMetrics(0)  # 获取主显示器的宽度
    desktop_height = user32.GetSystemMetrics(1)  # 获取主显示器的高度

    # 单显示器屏幕宽度和高度:
    img = ImageGrab.grab()  # 进行屏幕截图
    width, height = img.size  # 获取截图的宽度和高度

    scaling = round(width / desktop_width * 100) / 100  # 计算缩放比例

    # 计算出真实分辨率
    real_width = int(window_rect[0])    # 不包含DPI的宽
    real_height = int(window_rect[1])   # 不包含DPI的高
    borderless = True if real_width * scaling == 1920 else False  # 判断窗口是否无边框
    left_border = (real_width * scaling - 1920) / 2  # 计算左边框的偏移量
    up_border = (real_height * scaling - 1080) - left_border  # 计算上边框的偏移量
    real_width1 = 1920  # 真实宽度
    real_height1 = 1080  # 真实高度

    print(f"真实尺寸: 宽{real_width} x 高{real_height} x DPI{scaling} x 无边框{borderless} x 左边框{left_border} x 上边框-标题{up_border}")
    print('缩放尺寸：',int(real_width * scaling ) ,'x',int(real_height * scaling))
    

    # 排除缩放干扰
    windll.user32.SetProcessDPIAware()  # 告诉操作系统不要对进程进行 DPI 缩放
    
    return int(real_width * scaling ) , int(real_height * scaling) ,scaling ,borderless

# 该代码能够知道进程的运行时间，从而知道开机时间？
import win32gui
import win32process
import psutil

def get_process_dimensions(process_name):
    process_width = 0
    process_height = 0

    def callback(hwnd, lParam):
        nonlocal process_width, process_height
        thread_id, process_id = win32process.GetWindowThreadProcessId(hwnd)
        if process_id == lParam:
            rect = win32gui.GetClientRect(hwnd)
            # 获取进程的宽度和高度
            process_width = rect[2]
            process_height = rect[3]
            return False
        return True

    for process in psutil.process_iter():
        if process.name() == process_name:
            try:
                win32gui.EnumWindows(callback, process.pid)
                break
            except Exception as e:
                print(f"Error: {e}")

    user32 = windll.user32  # 加载 user32.dll 库
    desktop_width = user32.GetSystemMetrics(0)  # 获取主显示器的宽度
    desktop_height = user32.GetSystemMetrics(1)  # 获取主显示器的高度

    # 单显示器屏幕宽度和高度:
    img = ImageGrab.grab()  # 进行屏幕截图
    width, height = img.size  # 获取截图的宽度和高度

    scaling = round(width / desktop_width * 100) / 100  # 计算缩放比例

    # 计算出真实分辨率
    real_width = int(process_width)    # 不包含DPI的宽
    real_height = int(process_height)   # 不包含DPI的高
    borderless = True if real_width * scaling == 1920 else False  # 判断窗口是否无边框
    left_border = (real_width * scaling - 1920) / 2  # 计算左边框的偏移量
    up_border = (real_height * scaling - 1080) - left_border  # 计算上边框的偏移量
    real_width1 = 1920  # 真实宽度
    real_height1 = 1080  # 真实高度

    # 输出计算结果
    print(f"真实尺寸: 宽{real_width} x 高{real_height} x DPI{scaling} x 无边框{borderless} x 左边框{left_border} x 上边框-标题{up_border}")
    print('缩放尺寸：',int(real_width * scaling ) ,'x',int(real_height * scaling))
    windll.user32.SetProcessDPIAware()

    return int(real_width * scaling ) , int(real_height * scaling) ,scaling ,borderless



if __name__ == "__main__":
    #get_width('Genshin Impact')   
    get_process_dimensions("YuanShen.exe")
    # process_width, process_height = get_process_dimensions("YuanShen.exe")
    # print(f"进程 YuanShen.exe 的宽度和高度分别为：{process_width} x {process_height}")