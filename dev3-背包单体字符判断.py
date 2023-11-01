# 该工具用于采集图片和文字识别获取 - 物产志-提瓦特物产/战利品
# 设为通用
# 01. 【标准部件】判断是否前置窗口为原神get_window_handle()  get_width()
# 02. 【标准部件】利用pyautogui对当前窗口（全屏）进行截取，获取窗口化的窗口的实际坐标进行计算 软件部分截图：py_screenshot(screenshot_png,pic_save=True) 全屏截图：py_all_screenshot(screenshot_png,pic_save=True)
# 03. 【标准部件】对截图的指定部分进行截取crop_and_save_image(screenshot, x, y, width, height, save_path)
# 04. 【标准部件】对屏幕指定位置进行单击和中键移动py_click(x,y,width,height) py_scroll(rows)
# 05. 【标准部件】图片与图片进行比对，确定被比对的图在截图中的位置 find_image_coordinates(image_path, template_path)
# 06. 【标准部件】文字识别 text_ocr(pic_ocr)
# 07. 测试：对单个文字进行获取
# 07. 实操：物产志中循环文字提取并输出图片

import os
import time
import shutil
import keyboard
from PIL import Image

def on_key_release(event):
    if event.name == 'f3':
        # 按下 F3 键后终止程序
        print("按下 F3 键，程序退出。")
        os._exit(0)

# 注册按键释放事件的回调函数
keyboard.on_release(on_key_release)

'''
图片大小为114*114（与背包相比少10px，124x124）
第一个位置
左上 140,125

第二个
左上 269,125    差异129 横向


下方
左上 140,286  差异161

文字位置
左上    1322,666    高228
左下    1322,894                【下方文字位置为900高】
'''
#'''
print('执行01步骤***************************************')
print('########### 01. 【标准部件】判断是否前置窗口为原神')
from utils_find_process import get_window_handle
from utils_find_process import get_window_handle_name
from utils_find_process import get_foreground_window_info
from utils_process_x_y import get_width
from utils_process_x_y import get_process_dimensions

lag = 'chinese'
"""
# 其他语言的标题不是原神！【【【【【【【【【【【【备忘问题】】】】】】】】】】】】
if get_window_handle('原神'):   # 将标题名为原神的进程前置
    # print('你都没开原神！还是改语言了？')
    lag = 'chinese'
else:
    get_window_handle_name('YuanShen.exe')
    time.sleep(0.5)
    handle, pid, name, title = get_foreground_window_info()
    if title == 'Genshin Impact':
        # print('not open Genshin Impact?')
        lag = 'english'
    else:
        lag = 'other'
    # os._exit(1)     # 没做其他语言判断
print(lag)
if lag == 'chinese':
    real_width , real_height ,scaling ,borderless = get_width('原神')
elif lag == 'english':
    real_width , real_height ,scaling ,borderless = get_process_dimensions("YuanShen.exe")
else:
    real_width , real_height ,scaling ,borderless = get_process_dimensions("YuanShen.exe")
time.sleep(0.5)
if not real_width == 1920 or not real_height == 1080:
    print('窗口不是1920X1080')
    os._exit(1)
"""

print('执行02步骤***************************************')
print('########### 02. 【标准部件】利用pyautogui对当前窗口进行截取，获取窗口化的窗口的实际坐标进行计算')
# 问题：窗口必须在主界面，否则无法截取，显示为全黑
import pygetwindow as gw
import pyautogui
# 获取当前活动窗口
active_window = gw.getActiveWindow()            # <Win32Window left="0", top="0", width="1920", height="1080", title="原神">    # 偏移和尺寸
# 获取窗口位置和大小
window_left = active_window.left        # 相对于屏幕左上角的水平坐标
window_top = active_window.top          # 相对于屏幕左上角的垂直坐标
window_width = active_window.width      # 当前窗口的宽度
window_height = active_window.height    # 当前窗口的高度
# 进行截图
screenshot = pyautogui.screenshot(region=(window_left, window_top, window_width, window_height))        # 元组，window_left和window_top设为0则为全屏截图
# 当前窗口信息 <Win32Window left="1912", top="-8", width="1936", height="1056", title="dev2.py - 100 - Visual Studio Code [管理员]">
# 多屏下，为1屏时，x正确。在其他屏时，相对1屏的位置（从1屏开始算轴，不论1屏的位置）
print('当前窗口信息',active_window)
print('相对于屏幕左上角的水平坐标',window_left)
print('相对于屏幕左上角的垂直坐标',window_top)
print('当前窗口的宽度',window_width)
print('当前窗口的高度',window_height)


def py_screenshot(screenshot_png,pic_save=True):  # 截图：'test-screenshot.png'，仅限软件截图而非全屏截图,做坐标判断时，需要加入全屏偏移
    screenshot_pic = pyautogui.screenshot(region=(window_left, window_top, window_width, window_height))
    if pic_save == True:
        screenshot_pic.save(screenshot_png)     # 保存到图片，find_image_coordinates方法需要使用图片路径，否则需要加上判断
    return screenshot_pic       # 注意！输出的是图片的值，而不是图片的路径
# py_screenshot('test-screenshot.png')

def py_all_screenshot(screenshot_png,pic_save=True):  # 截图：'test-screenshot.png' ,全屏而不是单个软件部分，用于单击坐标获取
    screenshot_pic = pyautogui.screenshot()
    if pic_save == True:
        screenshot_pic.save(screenshot_png)     # 保存到图片
    return screenshot_pic       # 注意！输出的是图片的值，而不是图片的路径
# py_screenshot('test-screenshot.png')

print('执行03步骤***************************************')
print('########### 03. 【标准部件】对截图的指定部分进行截取')
def crop_and_save_image(screenshot, x, y, width, height, save_path):        # 从前端获得A. screenshot截图，B C.定义x、y坐标，D E.宽高矩形，F. 保存位置
    # 判断输入类型是图像路径还是图像对象
    if isinstance(screenshot, str):  # 如果是字符串，则加载图像
        screenshot = Image.open(screenshot)
    # 根据给定的坐标和尺寸进行截取
    cropped_image = screenshot.crop((x, y, x + width, y + height))
    # 保存截取的图片
    cropped_image.save(save_path)
# crop_and_save_image(screenshot, 140, 125, 114, 114, 'test-cropped_image.png')

print('执行04步骤***************************************')
print('########### 04. 【标准部件】对屏幕指定位置进行单击和中键移动')
# 单击操作
def py_click(x,y,width,height):     # x坐标，y坐标=左上角顶点坐标，width、height=宽高中心点增值，也可以设为0
    click_loc = (x + width + window_left,y + height + window_top)    # 左上角顶点+中心点增值 + 窗口与屏幕的差距
    # print('单击位置',click_loc)
    pyautogui.click(click_loc) 
    return click_loc
# py_click(140, 125, 114/2, 114/2)

# 滚动操作
def py_scroll(rows):
    for i in range(0,rows):                    # 右侧为绕过的行数（含本行）,物产志提瓦特物产首页5行-堇瓜
        for i in range(0,9):                # 素材的一行，物产志和素材一致,9次滚动
            pyautogui.scroll(-1,x=0,y=0)    # 只定义正值（向上）和负值（向下），x水平正值向右，负值向左。y正值向上，负值向下
    pyautogui.scroll(1,x=0,y=0)             # 返回向上一次
    time.sleep(0.5)
# py_scroll(5)    

print('执行05步骤***************************************')
print('########### 05. 【标准部件】图片与图片进行比对，确定被比对的图在截图中的位置')
import cv2
import numpy as np
def find_image_coordinates(image_path, template_path):      # image_path = 屏幕截图 template_path = 对比图,添加了，输入内容判断是图片还是图片的路径
    # 判断image_path和template_path是否为字符串路径
    if isinstance(image_path, str):
        # 读取图像文件
        image = cv2.imread(image_path)
    else:
        # image_path为图像变量
        image = np.array(image_path)
        image = cv2.cvtColor(image, cv2.COLOR_RGB2BGR)

    if isinstance(template_path, str):
        # 读取模板文件
        template = cv2.imread(template_path)
    else:
        # template_path为图像变量
        template = np.array(template_path)
        template = cv2.cvtColor(template, cv2.COLOR_RGB2BGR)
    # 使用模板匹配算法查找模板在图像中的位置
    result = cv2.matchTemplate(image, template, cv2.TM_CCOEFF_NORMED)
    _, max_val, _, max_loc = cv2.minMaxLoc(result)      # max_val是相似度，max_loc为坐标（元组）
    # 获取模板的宽度和高度
    template_width = template.shape[1]
    template_height = template.shape[0]

    # 计算模板的左上角和右下角坐标
    top_left = max_loc      # 左上角坐标
    bottom_right = (top_left[0] + template_width, top_left[1] + template_height)    # 右下角坐标，依赖比对图的大小，如果有差异、拉伸则错误，114或125
    print('比对相似度为：',max_val, '坐标为：',top_left, bottom_right)
    return max_val,top_left, bottom_right

print('执行06步骤***************************************')
print('########### 06. 【标准部件】文字识别')
import cnocr
def text_ocr(pic_ocr,single=False):
    current_dir = os.getcwd()   # 当前目录，要把模型放在当前目录，model目录
    # cnocr
    #print('载入OCR模型')
    if lag == 'english':
        det_model_name = "en_PP-OCRv3_det"
    elif lag == 'chinese':
        det_model_name = "ch_PP-OCRv3_det"
    else:
        det_model_name = "en_PP-OCRv3_det"              # 【【【【【注意！其他语言没有指定OCR模型！可能全错】】】】】
    det_root = os.path.join(current_dir, "model/cnocr")
    # cnstd
    #print('载入文字识别模型')
    rec_model_name = "densenet_lite_114-fc"
    rec_root = os.path.join(current_dir, "model/cnstd")
    #print('文字识别启用')
    # 使用 OCR 模型进行文本识别，仅数字后面加上, cand_alphabet='0123456789'
    ocr = cnocr.CnOcr(det_model_name=det_model_name, rec_model_name=rec_model_name, det_root=det_root, rec_root=rec_root)
    if single == True:
        # 获取识别结果中的文本值 # 单行内容  
        text = ocr.ocr_for_single_line(pic_ocr)
    else:
        text = ocr.ocr(pic_ocr)
    return text


print('执行07步骤***************************************')
print('########### 07. 实操：物产志中循环文字提取并输出图片')
# screenshot_pic = py_screenshot('./dev3/screenshot.png')                              # B.步骤2，截图窗口并保存图片
screenshot_pic = './dev3/screenshot-1.png'    # 截图？或选用图片
crop_pic = './dev3/crop_cnocr.png'
crop_and_save_image(screenshot_pic, 1310, 120, 490, 60, crop_pic)    # C.步骤3，对截图中进行文字截图，固定位置
single = False
textocr = text_ocr(crop_pic,single)                                      # D. 步骤6，对截图中部分进行文字识别并输出，textocr[0]['text']
print(textocr)
if single == False:
    ocr_word = textocr[0]['text']
    ocr_score = round(textocr[0]['score'],2)
else:
    ocr_word = textocr['text']
    ocr_score = round(textocr['score'],2)
print(ocr_word,ocr_score)

'''
待办：
1. 如果某个素材未获得，则可能中断【识别OCR空值】
'''