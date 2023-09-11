# 首先找到坐标系
# 宽度8次相加
# 单个 128X132
# 横向，146.5 +相加
# 高度 176 4排

import os
import cv2
from PIL import Image
import os
import pyautogui
import numpy as np
import time

dev_image_path = './dev'

if not os.path.exists(dev_image_path):
   os.makedirs(dev_image_path)

for filename in os.listdir(dev_image_path):
   file_path = os.path.join(dev_image_path, filename)
   if os.path.isfile(file_path):
       os.remove(file_path)
       
# 构建完整的图片路径
image_path = 'dev.png'
# 加载当前图片
current_image = cv2.imread(image_path) 
# 加载当前屏幕截图
screenshot = pyautogui.screenshot()
screenshot = np.array(screenshot)
screenshot = cv2.cvtColor(screenshot, cv2.COLOR_RGB2BGR)
result = cv2.matchTemplate(current_image, screenshot, cv2.TM_CCOEFF_NORMED)
min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)
print("匹配图为", image_path, "匹配值为", max_val, "匹配坐标为", max_loc)

from PIL import ImageGrab

def capture_and_save_image(location, width, height, path, image_name):
    # 根据给定的位置和尺寸截取屏幕图像
    left = location[0]
    top = location[1]
    right = left + width
    bottom = top + height
    screen_image = ImageGrab.grab(bbox=(left, top, right, bottom))

    # 保存图像到指定路径
    image_path = f"{path}/{image_name}.png"
    screen_image.save(image_path)
    print(f"Saved image: {image_path}")

# 设置初始参数
# max_loc = (147, 560)

rows = 0
now_loc = max_loc

# 根据要求进行截图和保存，重复 4 次
for rows in range(0, 4):
    # 更新 max_loc 的高度
    now_loc = (max_loc[0], max_loc[1] + 176 * rows)         # 修改高度+176

    for i in range(1, 9):
        # 截取宽128，高132的图像
        capture_and_save_image(now_loc, 128, 132, dev_image_path, f"{rows+1}-{i}")

        # 更新 max_loc 的宽度
        now_loc = (now_loc[0] + 146.5, now_loc[1])        # 横向+146.5px，重复8次
