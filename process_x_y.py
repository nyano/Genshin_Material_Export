from ctypes import windll
import pygetwindow as gw
from PIL import ImageGrab

# 定义函数 get_width，接受窗口标题作为参数
def get_width(title):
    window = gw.getWindowsWithTitle(title)[0]  # 获取指定标题的窗口对象
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

    # 输出计算结果
    print(f"Real resolution: {real_width} x {real_height} x {scaling} x {borderless}")
    print(real_width)    # 不包含DPI的宽
    print(real_height)   # 不包含DPI的高
    print(scaling)          # DPI
    print(borderless)       # 是否为无边框，True则 =1920
    print(left_border)      # 左边框？为什么要有左边框
    print(up_border)        #上边框，高度，标题栏
    print('*' * 20)
    print(real_width * scaling, real_height * scaling)
    

    # 排除缩放干扰
    windll.user32.SetProcessDPIAware()  # 告诉操作系统不要对进程进行 DPI 缩放
    
    return real_width , real_height ,scaling ,borderless


if __name__ == "__main__":
    get_width('微信')   
