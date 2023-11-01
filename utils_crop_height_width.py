# 对材料图像部分进行高宽截取，保证坐标系正确,目前看仅用于滚动时候y轴确保
"""
以1px为一个点，输出图片的5x5每个点的颜色。
除第一行外，其余各点RGB如果同时与上一行对应列的点RGB相差不大于5，则输出是同一个颜色。
输出时，按一行的点输出为一列

绿80-100 ，110-130,100-120

蓝90-100,100-120,140-150

紫100-120,90-110,140-170

白100-110,100-110,100-120

符合范围：80-120,90-130,100-170
其他 110-180,100-200,130-210
110-120,100-130,130-170
(178, 175, 202)
(154, 161, 185)
(141, 162, 158)
(179, 194, 190)
(124, 126, 131)
(179, 194, 190)
(116, 109, 164)


打开一张图片，以左上角为顶点，横向读取每个像素点的RGB颜色，输出该像素点的RGB颜色。如果R值不在80-120范围内、G不在90-130范围内或B不在100-170范围内，三个要求任意一个未满足，则该点视为不符合要求，输出该点是否符合要求
如果该点同时满足三个要求，则以该点为顶点，向下读取5个像素点的RGB颜色，如果这5个像素点均符合要求，则输出该顶点的坐标以该点所在行为顶点，裁剪为新图片。
然后，以左上角为顶点，竖向读取每个像素点的RGB颜色，输出该像素点的RGB颜色。如果R值不在80-120范围内、G不在90-130范围内或B不在100-170范围内，三个要求任意一个未满足，则该点视为不符合要求，输出该点是否符合要求
如果该点同时满足三个要求，则以该点为顶点，向右读取5个像素点的RGB颜色，如果这5个像素点均符合要求，则输出该顶点的坐标以该点所在列为顶点，裁剪为新图片。
"""


from PIL import Image

def check_pixel_compliance_height(image_path, column, row):
    # 打开图片
    if isinstance(image_path, str):
        # image_path 是图片路径，打开图片
        image = Image.open(image_path)
    else:
        # image_path 是图像对象，直接使用
        image = image_path
    # 获取图片的尺寸
    width, height = image.size
    # 遍历每个像素点
    for y in range(row, height):
        for x in range(column, width):
            pixel = image.getpixel((x, y))
            r, g, b = pixel
            # 检查像素的颜色是否在所需范围内
            r_compliant = 80 <= r <= 120
            g_compliant = 90 <= g <= 130
            b_compliant = 100 <= b <= 170
            # 输出像素点的RGB值以及是否符合要求
            # print(f"坐标({x}, {y}) 的 RGB 值: R={r}, G={g}, B={b}，符合要求: {r_compliant and g_compliant and b_compliant}")
            # 如果像素点同时满足三个要求
            if r_compliant and g_compliant and b_compliant:
                # 检查下方5个像素点是否都符合要求
                region_compliant = True
                for i in range(y + 1, y + 6):
                    pixel = image.getpixel((x, i))
                    r, g, b = pixel
                    if not (80 <= r <= 120 and 90 <= g <= 130 and 100 <= b <= 170):
                        region_compliant = False
                        break
                # 如果下方5个像素点都符合要求，则输出顶点坐标并裁剪图片
                if region_compliant:
                    print("满足要求的顶点坐标：", (x, y))
                    cropped_height = height - y - 1
                    cropped_image = image.crop((0, y + 1, width, height))
                    new_image = Image.new("RGB", (width, cropped_height), (255, 255, 255))
                    new_image.paste(cropped_image, (0, 0))
                    new_image.save(image_path)
                    print("裁剪后的图片已保存。")
                    return
    print(f"未找到满足要求的顶点。{image_path}")

def check_pixel_compliance_width(image_path, column, row):
    # 打开图片
    if isinstance(image_path, str):
        # image_path 是图片路径，打开图片
        image = Image.open(image_path)
    else:
        # image_path 是图像对象，直接使用
        image = image_path
    # 获取图片的尺寸
    width, height = image.size
    # 遍历每个像素点
    for x in range(column, width):
        for y in range(row, height):
            pixel = image.getpixel((x, y))
            r, g, b = pixel

            # 检查像素的颜色是否在所需范围内
            r_compliant = 80 <= r <= 120
            g_compliant = 90 <= g <= 130
            b_compliant = 100 <= b <= 170

            # 输出像素点的RGB值以及是否符合要求
            # print(f"坐标({x}, {y}) 的 RGB 值: R={r}, G={g}, B={b}，符合要求: {r_compliant and g_compliant and b_compliant}")

            # 如果像素点同时满足三个要求
            if r_compliant and g_compliant and b_compliant:
                # 检查右侧5个像素点是否都符合要求
                region_compliant = True
                for i in range(y + 1, y + 6):
                    pixel = image.getpixel((x, i))
                    r, g, b = pixel
                    if not (80 <= r <= 120 and 90 <= g <= 130 and 100 <= b <= 170):
                        region_compliant = False
                        break

                # 如果右侧5个像素点都符合要求，则输出顶点坐标并裁剪图片
                if region_compliant:
                    print("满足要求的顶点坐标：", (x, y))
                    cropped_width = width - x - 1
                    cropped_image = image.crop((x + 1, 0, width, height))
                    new_image = Image.new("RGB", (cropped_width, height), (255, 255, 255))
                    new_image.paste(cropped_image, (0, 0))
                    new_image.save(image_path)
                    print("裁剪后的图片已保存。")
                    return
    print(f"未找到满足要求的顶点。{image_path}")

if __name__ == "__main__":
    fold = './01-test'
    # f'{fold}/a.png'
    # 图片路径
    image_path = f'{fold}/bbb.png'
    # 以左侧列为起点的顶点坐标
    column = 0      # 开始列
    row = 0         # 开始行，0,0表示左上角起步
    # 检查像素点并进行裁剪
    check_pixel_compliance_height(image_path, column, row)      # 裁剪Y轴
    check_pixel_compliance_width(image_path, column, row)       # 裁剪X轴