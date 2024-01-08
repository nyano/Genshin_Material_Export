#
'''获得屏幕尺寸、真实屏幕尺寸、DPI、窗口大小
创建时间：2023-12-18
修改时间：2023-12-19
'''
'''调用该库的模块

'''


import ctypes
from PIL import ImageGrab
import io
import win32api
import pygetwindow as gw
import win32gui
# import pyautogui      # 注意，导入后会造成DPI问题 get_screen_resolution()

class Display():
    '''该类用于获得屏幕分辨率
    001 capture_screen()该方法通过截图来获得主屏幕的分辨率，输入空，输出：主屏幕的宽，高    (int,int)
    002 get_screen_resolution()该方法通过win32api来获得屏幕分辨率，没有算入DPI 输入：空 输出：去掉DPI的宽，高(int,int)
    003 get_monitor_dpi()该方法通过截图尺寸/API获得的尺寸获得DPI的值 输入：空 输出：DPI(float)   输出为显示器列表 ，第一个显示器为monitor_list[0][1]
    类方法 windows_size_by_title(title) 该方法通过窗口标题来获得窗口大小(附带恢复与置前) 输入：窗口标题 输出：窗口宽，高，离屏幕左，离屏幕上，离屏幕下  int,int,int,int,int
    类方法 get_screen_size() 直接输出屏幕尺寸相关信息 输入空，输出：real_width屏幕宽(int),real_height屏幕高(int),width缩放宽(int),height缩放高(int),dpi缩放比(float)
    '''
    # 001 截图获得主屏幕分辨率
    def capture_screen(self) -> (int,int):
        '''该方法通过截图来获得主屏幕的分辨率
        输入：空
        输出：主屏幕的宽，高
        '''
        screenshot = ImageGrab.grab()           # 截取整个屏幕
        width, height = screenshot.size         # 获取图像尺寸
        #print(f"图像尺寸：{width} x {height}")  # 显示图像尺寸
        return width,height
        # screenshot.save("screenshot.png")     # 保存截图（可选）

    # 002 通过WIN32API获得主屏幕实际分辨率
    
    def get_screen_resolution(self) -> (int,int):    # 也可以用ctypes
        '''该方法经过DPI计算后的实际分辨率，主屏幕（非显示器分辨率）
        输入：空
        输出：经过DPI的宽，高
        '''
        pic_width,pic_height = self.capture_screen()
        monitor_list = self.get_monitor_dpi()
        dpi = monitor_list[0][1]
        width = pic_width / dpi
        height = pic_height / dpi

        return int(width), int(height)
        '''另一种方式获得实际分辨率（经过DPI后）
        # 定义 GetSystemMetricsForDpi 函数
        def get_system_metrics_for_dpi(metric, dpi):
            user32 = ctypes.windll.user32
            return user32.GetSystemMetricsForDpi(metric, dpi)

        # 获取主显示器的水平像素数
        horizontal_pixels = get_system_metrics_for_dpi(0, 96)
        print(f"Horizontal Pixels (DPI 96): {horizontal_pixels}")

        # 获取主显示器的垂直像素数
        vertical_pixels = get_system_metrics_for_dpi(1, 96)
        print(f"Vertical Pixels (DPI 96): {vertical_pixels}")

        '''
    # 003 获得主屏幕DPI
    def get_monitor_dpi(self) -> list:
        '''获得是多显示器列表序号和DPI值，第一个是主显示器'''
        # 定义常量 PROCESS_PER_MONITOR_DPI_AWARE 和 MDT_EFFECTIVE_DPI
        PROCESS_PER_MONITOR_DPI_AWARE = 2
        MDT_EFFECTIVE_DPI = 0
        # 获取 ctypes.windll.shcore 模块
        shcore = ctypes.windll.shcore
        
        # 使用 win32api.EnumDisplayMonitors() 获取所有显示器的信息
        monitors = win32api.EnumDisplayMonitors()
        
        # 设置进程为支持每个显示器独立的 DPI 设置
        hresult = shcore.SetProcessDpiAwareness(PROCESS_PER_MONITOR_DPI_AWARE)
        #assert hresult == 0  # 确保设置 DPI 意识成功
        
        # 用于存储 DPI 信息的变量
        dpiX = ctypes.c_uint()
        dpiY = ctypes.c_uint()
        monitor_list = []
        # 遍历每个显示器并获取 DPI 信息
        for i, monitor in enumerate(monitors):
            shcore.GetDpiForMonitor(
                monitor[0].handle,  # 获取监视器句柄
                MDT_EFFECTIVE_DPI,   # 获取有效 DPI
                ctypes.byref(dpiX),  # 存储水平 DPI 的变量
                ctypes.byref(dpiY),   # 存储垂直 DPI 的变量
            )
            monitor_list.append((i, dpiX.value / 96))

            # 打印 DPI 信息
            #print(f"Monitor {i} (hmonitor: {monitor[0]}) = dpiX: {dpiX.value}, dpiY: {dpiY.value}")
        #print(monitor_list)  #列表[(tuple元组)]
        #print(type(monitor_list[0]))
        #print('第一个显示器的DPI',monitor_list[0][1])
        return monitor_list
    
    # 004 窗体恢复、激活（置前）、窗口尺寸和定位
    '''
    '''
    @classmethod
    def windows_size_by_title(self,window_title) -> (int,int,int,int,int):
        '''
        前置：窗体需要置前
        输入：窗口标题
        输出：窗口宽，高，离屏幕左，离屏幕上，离屏幕下  int,int,int,int,int
        窗体的截图大小 = 设置大小 = 无DPI大小=主屏幕分辨率
        win32api获得的大小=计入DPI的大小
        注意，先精准判断，是否存在相同标题的窗体，如果没有，则返回 包含关系的第一个窗体
        '''
        windows = gw.getWindowsWithTitle(window_title)      # [Win32Window(hWnd=138484)]

        if windows:
            for w in windows:
                if w.title == window_title:
                    window = w
                    break
            else:
                # 如果没有完全相等的窗体，输出第一个窗体
                window = windows[0]
        else:
            return 0 , 0 , 0 , 0 , 0        # 未能获得
        window.restore()        # 如果最小化，则自动恢复窗口
        window.activate()       # 激活（置前）
        width,height,left,top,bottom = window.width,window.height,window.left,window.top,window.bottom
        #print('窗口宽：',width,'窗口高：',height,'窗口离屏幕左：',left,'窗口离屏幕上：',top,'窗口离屏幕下：',bottom)    # 与DPI无关，等同截屏尺寸。当主屏幕有DPI后，窗口在非DPI屏上，需要除以DPI
        # 截取窗口区域的屏幕图像
        #screenshot = pyautogui.screenshot(region=(window.left, window.top, window.width, window.height))
        # 保存截图到文件
        #screenshot.save("window_screenshot.png")
        return width , height , left , top , bottom     # 窗口宽，高，离屏幕左，离屏幕上，离屏幕下（下没啥用）int
    
    @classmethod
    def get_screen_size(cls) -> (int,int,int,int,float):
        '''
        输出：空
        输入：real_width屏幕宽(int),real_height屏幕高(int),width缩放宽(int),height缩放高(int),dpi缩放比(float)
        '''
        screen = Display()
        real_width,real_height = screen.capture_screen()
        width,height = screen.get_screen_resolution()
        dpi = screen.get_screen_dpi()
        return real_width,real_height,width,height,dpi
    
if __name__ == "__main__":
    display = Display()
    #width,height = display.capture_screen()    # 该方法获得是实际显示器分辨率
    a,b = display.get_screen_resolution()        # 获得的是缩放DPI后的尺寸
    print('结果',a,b)
    dpi = display.get_monitor_dpi()
    #dpi = display.get_screen_dpi()
    print(dpi)
    #Display.get_screen_size()
    #display.windows_size_by_title('微信')
    #Display.windows_size_by_title('微信')

