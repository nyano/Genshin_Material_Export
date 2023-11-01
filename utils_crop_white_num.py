# 数字部分进行裁剪，能够扩大化数字部分的截图，进行定位
"""
字的白色部分色系(233,229,220)
字的白色系部分高度27，满色24px
x轴 122px

打开图片，以左下角为顶点，横向读取20个像素点，输出每个像素点的像素点的RGB颜色。如果该行中有任意2个以上像素点RGB颜色在(233,229,220)各值的上下20范围内，
则记录该顶点坐标，向上截取高为27px，宽度与图片一致，保存为新图片
"""

from PIL import Image

def process_image_white_num(image_path,to_new_image):
    # 打开图片
    if isinstance(image_path, str):
        # image_path 是图片路径，打开图片
        image = Image.open(image_path)
    else:
        # image_path 是图像对象，直接使用
        image = image_path
    width, height = image.size

    # 初始化记录顶点坐标的列表
    top_vertex = None

    # 从左下角向上读取每个像素点的RGB颜色
    for y in range(height-1, -1, -1):
        # 初始化计数器
        count_within_range = 0

        # 从左到右读取每个像素点的RGB颜色,横向移动
        for x in range(width):
            pixel = image.getpixel((x, y))
            r, g, b = pixel

            # 检查RGB颜色是否在指定范围内
            if 213 <= r <= 253 and 209 <= g <= 249 and 200 <= b <= 240:
                count_within_range += 1

                # 输出每个像素点的RGB颜色
                # print(f"坐标：({x}, {y}), RGB颜色：{pixel}")

        # 如果该行中有2个以上像素点在范围内，记录顶点坐标
        if count_within_range >= 2:
            top_vertex = (0, y)
            break

    if top_vertex:
        # 截取顶点坐标上方27px高度的图像并保存
        x, y = top_vertex
        new_image = image.crop((0, y-27, width, y))
        # new_image.show()
        new_image.save(to_new_image)
    return to_new_image

if __name__ == "__main__":
    # 调用函数并传入图像路径
    fold = './01-test'
    # f'{fold}/a.png'
    # 图片路径
    image_path = f'{fold}/bbb.png'
    to_new_image = f'{fold}/ddd.png'
    process_image_white_num(image_path,to_new_image)
