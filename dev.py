# 首先找到坐标系
# 宽度8次相加
'''
图片大小为124x124
左上	118,121
右上	241,121		差异123 横向
左下	118,244		差异123 上下
右下	241,244

第二个
左上	265,121     差异x 147 ,与右上间隔24
右上	387,121

下方
左上	118,296		差异175 上下(与左上)，差异52（与左下）

该程序是获取单系列的所有图片
1.首先获取4行数据*8张图片，取正方形，左上坐标(x,y)，右上x+123，左下y+123，右下x+123,y+123。
第一行8图的坐标为
命名规则为1-1-001.png，行-列-编号
4行后，获取4-1的图片名称，进行中键滚动，然后进行图片识别定位，获取4-1的坐标，获取下一行的坐标，进行下一行匹配
如果匹配后行坐标没有变化，或者向下截取时超过了屏幕尺寸y>1080，则终止循环



'''



# 01.首先判断是否存在dev子文件夹，如存在则删除它，然后再创建。如不存在，则创建它。
import shutil
import os
folder_name = "dev"
if os.path.exists(folder_name):
    shutil.rmtree(folder_name)
os.mkdir(folder_name)


# 02.判断是否前置窗口为原神
from find_process import get_window_handle
from process_x_y import get_width

if not get_window_handle('原神'):   # 将标题名为原神的进程前置
    print('你都没开原神！')
    os._exit(1)

real_width , real_height ,scaling ,borderless = get_width('原神')
if not real_width == 1920 or not real_height == 1080:
    print('窗口不是1920X1080')
    os._exit(1)
    
# 03.识别dev.png在屏幕截图中的坐标，并输出
# 并保存到dev文件夹中，文件名为0-screenshot.png，
print('开始前，请确认是否已经把所有（新）标记取消掉，鼠标定位到首个')

# 01会删掉目录注意！
import pyautogui
# 进行屏幕截图
screenshot = pyautogui.screenshot()
# 保存截图为文件
screenshot.save('./dev/0-screenshot.png')

import cv2

# 比对函数
def find_image_coordinates(image_path, template_path):
    # 读取原始图像和模板图像
    image = cv2.imread(image_path)
    template = cv2.imread(template_path)

    # 使用模板匹配算法查找模板在图像中的位置
    result = cv2.matchTemplate(image, template, cv2.TM_CCOEFF_NORMED)
    _, max_val, _, max_loc = cv2.minMaxLoc(result)      # max_val是相似度，max_loc为坐标（元组）

    # 获取模板的宽度和高度
    template_width = template.shape[1]
    template_height = template.shape[0]

    # 计算模板的左上角和右下角坐标
    top_left = max_loc
    bottom_right = (top_left[0] + template_width, top_left[1] + template_height)
    print('比对相似度为：',max_val)
    if max_val < 0.75:
        print('未能有效匹配，请确认比对的图片（首次为dev.png）或游戏屏幕是否正确')
        print('退出')
        os._exit(0)

    return top_left, bottom_right

# 图像路径
image_path = './dev/0-screenshot.png'
template_path = './dev.png'

# 查找模板在图像中的坐标
top_left, bottom_right = find_image_coordinates(image_path, template_path)

# 输出坐标
print(f'左上坐标: {top_left}, 右下坐标: {bottom_right}')


# 04.
'''

定义3个变量，num，rows，columns，初始值均为数字1，后续截取的图片文件名格式以num-rows-columns.png为准，如001-1-1.png。
其中num的格式化为3位数字，每截取一张图自增1。rows表示该图片行数，每坐标向下移动一次，自增1.columns表示该图片的列数，每坐标向右移动一次，自增1。
定义一个方法。在此方法中，定义一个变量top_res，是一个坐标118,121。首先屏幕截图，在屏幕截图中，以此坐标作为左上角顶点。
截取一个123px的正方形，保存到dev目录中，文件名如001-1-1.png，横坐标向右移动147px，再次截取123px正方形，文件名如002-1-2.png。以此类推，共向右移动并截取7次。
然后更新top_res的值，向下移动175px，重复截取正方形并向右移动截取。以此类推，向下移动共3次。
'''
'''
完成后，再将坐标向下移动一次175px，截取一次123px正方形，保存为0-temp.png，然后滚动鼠标中键向下移动5px，然后重新进行屏幕截图，利用opencv确认0-temp.png在屏幕截图中的位置，将左上角顶点坐标作为top_res的值
重新开始之前的循环
'''

import cv2
from PIL import ImageGrab
import numpy as np
import pyautogui
import time

def capture_and_crop_images():
    num = 1   # 自增编号，格式化输出
    rows = 1    # 行
    columns = 1 # 列
    top_res = top_left  # (118, 121)
    square_size = 123   # 正方形b
    move_distance = 147 # 横向移动
    down_move_distance = 175  # 坐标向下移动距离
    a = 0

    # 屏幕截图
    screenshot_img = ImageGrab.grab()
    for _ in range(15):  # 重复2次
        for _ in range(3):  # 坐标向下移动3次,行数
            for _ in range(8):  # 向右移动8次，列数
                # print("当前坐标:" ,top_res)
                # 截取正方形，+123px
                cropped_square = screenshot_img.crop((top_res[0], top_res[1], top_res[0] + square_size, top_res[1] + square_size))
                # 生成文件名
                file_name = f"{num:03d}-{rows}-{columns}.png"
                # 保存图像到 dev 目录
                cv2.imwrite(f"./dev/{file_name}", cv2.cvtColor(np.array(cropped_square), cv2.COLOR_RGB2BGR))

                num += 1
                columns += 1
                top_res = (top_res[0] + move_distance, top_res[1])

            rows += 1
            columns = 1
            top_res = (118, top_res[1] + down_move_distance)   # Y轴移动，首行

        if top_res[1] >= 930:
            print('无法移动了，停止')
            break
        # 保存0-temp.png
        # top_res = (118, top_res[1] + down_move_distance)    # 第四行的首行
        print('*' * 20)
        print(f'当前为第{a}次移动，当前首行坐标',top_res,'首行为',rows,'最后一个文件为：',file_name)
        temp_img = screenshot_img.crop((top_res[0], top_res[1], top_res[0] + square_size, top_res[1] + square_size))
        temp_img.save('./dev/0-temp.png')           # 当前需要比对的首行，确认坐标
        temp_img.save(f'./dev/0-temp-{a}.png')
        a += 1

        # 移动屏幕操作
        click_loc = (top_res[0]+0,top_res[1]-50)    # 在首行，向上移动
        pyautogui.click(click_loc)  # 单击
        time.sleep(0.5)
        for i in range(1,4):    # 3行
            for i in range(1,11):
                pyautogui.scroll(-10,x=0,y=0) # 向下滚动72个像素
        pyautogui.scroll(10,x=0,y=0)    # 再向上移动一次
        # 移动鼠标中键滚动5px向下
        # pyautogui.scroll(-5)
        click_loc2 = (50,50)       # 返回点击位置，向左移动，减少鼠标对截图的影响.左上角
        pyautogui.click(click_loc2)  # 单击
        print(f'已进行第{a}次移动后，单击坐标',click_loc)
        if click_loc[1] >= 710:       # 在最后一行时退出循环
            print('无法移动了，停止')
            break

        time.sleep(2)   # 阻止移动缓慢造成截图失败

        # 重新截图，界面已经移动了
        screenshot_img = ImageGrab.grab()
        screenshot_img_path = f'./dev/0-screenshot-{a}.png'
        screenshot_img.save(screenshot_img_path)
        image_path = screenshot_img_path
        template_path = './dev/0-temp.png'
        find_res, bottom_right = find_image_coordinates(image_path, template_path)
        top_res = list(top_res)
        top_res[1] = find_res[1]
        top_res[0] = 118    # 手动设置首行为x=118
        top_res = tuple(top_res)
        print('重新截图并比对后，当前比对的坐标为',top_res, '比对的图片为',image_path)   # 应该为118，？ 首行坐标系，包括本行，第二次应该为第四行

        # 删除临时图像
        cv2.imwrite('./dev/0-temp.png', np.zeros((square_size, square_size, 3), dtype=np.uint8))  # 清空临时图像内容，阻止对后续的影响



    # 打印最终的 num、rows 和 columns 值
    print(f"Final num: {num}")
    print(f"Final rows: {rows}")
    print(f"Final columns: {columns}")

# 调用方法
capture_and_crop_images()
