from PIL import Image
import xlsxwriter
import os
from collections import Counter
import numpy as np
from colormath.color_objects import sRGBColor, LabColor
from colormath.color_diff import delta_e_cie2000
from colormath.color_conversions import convert_color
import math
# from colour import RGB_to_Lab
import numpy

def patch_asscalar(a):
    return a.item()

setattr(numpy, "asscalar", patch_asscalar)
color_dict = [
    [(99, 170, 171), (130, 192, 206), (126, 187, 217), (163, 191, 218), (185, 193, 220), (152, 176, 214), (225, 141, 174), (226, 153, 183), (238, 187, 189), (183, 154, 186), (185, 174, 212), (163, 153, 202)],
    [(59, 181, 170), (76, 173, 218), (60, 167, 206), (74, 171, 208), (110, 144, 211), (151, 167, 213), (230, 59, 106), (212, 66, 112), (222, 103, 135), (214, 97, 171), (169, 112, 170), (124, 102, 171)],
    [(39, 161, 178), (33, 149, 174), (36, 136, 192), (62, 131, 188), (48, 116, 187), (98, 128, 182), (208, 60, 112), (190, 31, 86), (145, 56, 123), (152, 45, 128), (131, 63, 158), (125, 89, 152)],
    [(63, 160, 145), (49, 141, 146), (56, 111, 183), (65, 112, 199), (49, 88, 157), (76, 114, 151), (171, 62, 95), (163, 59, 105), (135, 68, 134), (125, 74, 143), (100, 74, 134), (70, 66, 123)],
    [(194, 215, 217), (188, 220, 195), (185, 212, 186), (212, 203, 69), (169, 151, 121), (168, 149, 131), (235, 222, 111), (232, 207, 128), (241, 186, 115), (245, 147, 94), (237, 137, 132), (241, 100, 108)],
    [(89, 187, 126), (169, 213, 129), (186, 213, 107), (188, 188, 22), (156, 94, 67), (150, 94, 75), (235, 197, 76), (249, 182, 69), (247, 150, 52), (248, 115, 91), (245, 69, 67), (233, 72, 71)],
    [(54, 155, 74), (42, 186, 50), (88, 180, 29), (155, 184, 30), (115, 74, 52), (104, 74, 72), (238, 185, 74), (250, 128, 56), (244, 98, 40), (204, 48, 41), (206, 71, 89), (146, 58, 55)],
    [(48, 75, 67), (31, 93, 79), (37, 116, 76), (124, 137, 54), (83, 83, 62), (105, 84, 83), (219, 161, 75), (245, 97, 61), (225, 51, 56), (200, 45, 56), (171, 40, 56), (126, 55, 67)],
    [(220, 184, 177), (233, 182, 166), (239, 193, 165), (239, 207, 187), (236, 234, 227), (234, 232, 228), (218, 195, 190), (216, 197, 164), (222, 207, 158), (193, 170, 145), (203, 178, 151), (195, 185, 175)],
    [(231, 159, 124), (218, 150, 103), (236, 170, 116), (244, 194, 138), (240, 218, 173), (226, 215, 193), (185, 173, 176), (151, 147, 129), (187, 152, 99), (173, 119, 85), (156, 133, 112), (160, 159, 154)],
    [(209, 122, 67), (208, 116, 55), (213, 154, 93), (152, 117, 88), (136, 129, 123), (130, 124, 117), (138, 114, 118), (94, 100, 81), (177, 124, 73), (156, 86, 62), (140, 111, 96), (145, 136, 129)],
    [(127, 85, 67), (90, 63, 58), (91, 71, 65), (77, 72, 73), (58, 56, 58), (73, 70, 71), (127, 97, 115), (66, 74, 79), (161, 95, 54), (161, 77, 65), (121, 91, 85), (104, 100, 100)]
]


def rgb_to_hex(rgb):
    r, g, b = rgb
    return "#{:02x}{:02x}{:02x}".format(r, g, b)

def most_frequent_color(image_path):
    global color_dict
    image = Image.open(image_path)
    top = 0
    bottom = 0
    item_width = image.width / 12
    item_height = image.height / 12
    res = []
    for x in range(12):
        left = 0
        right = 0
        bottom = bottom + item_height
        row = []
        for y in range(12):
            right = right + item_width
            cropped_image = image.crop((left, top, right, bottom))
            left = left + item_width
            pixels = cropped_image.getdata()
            r_sum, g_sum, b_sum = 0, 0, 0
            num_pixels = len(pixels)
            for pixel in pixels:
                r_sum += pixel[0]
                g_sum += pixel[1]
                b_sum += pixel[2]
            r_avg = r_sum // num_pixels
            g_avg = g_sum // num_pixels
            b_avg = b_sum // num_pixels
            # res.append(f'rgb({r_avg}, {g_avg}, {b_avg})')
            row.append((r_avg, g_avg, b_avg))
        res.append(row)
        top = top + item_height
    color_dict = res
    # 创建 Excel 文件
    workbook = xlsxwriter.Workbook('colors.xlsx')
    worksheet = workbook.add_worksheet()
    for row_index, row in enumerate(res):
        for i, col in enumerate(row):
            color = rgb_to_hex(col)
            # 将颜色值写入 Excel 单元格
            cell_format = workbook.add_format(
                {'bg_color': color})
            worksheet.write(row_index, i, color, cell_format)
    workbook.close()
    return res


def closest_color(target_color, color_array):
    min_distance = float('inf')
    closest_color_found = None
    closest_color_position = None
    target_color = np.array(target_color)
    for i, color in enumerate(color_array):
        color = np.array(color)
        distance = np.sqrt(np.sum((target_color - color) ** 2))
        if distance < min_distance:
            min_distance = distance
            closest_color_found = color
            closest_color_position = i

    return closest_color_found, [closest_color_position // 12, closest_color_position % 12]


def find_closest_color_ciede2000(target_color, color_array):
    """
    使用CIEDE2000色差公式在颜色数组中找到与目标颜色最接近的颜色。

    参数:
    color_array (list of tuples): 包含多个颜色元组的列表，每个元组表示一个颜色，格式为 (R, G, B)。
    target_color (tuple): 目标颜色，格式为 (R, G, B)。

    返回:
    tuple: 最接近目标颜色的颜色元组。
    """
    min_distance = float('inf')
    closest_color = None
    # 假设这是目标颜色的RGB值
    # 先将RGB值转换为sRGBColor对象（如果颜色值范围是0-255，设置is_upscaled=True）
    target_srgb = sRGBColor(target_color[0], target_color[1], target_color[2], is_upscaled=True)
    # 再将sRGBColor对象转换为LabColor对象
    target_lab = convert_color(target_srgb, LabColor)
    for color in color_array:
        current_srgb = sRGBColor(color[0], color[1], color[2], is_upscaled=True)
        current_lab = convert_color(current_srgb, LabColor)
        # 使用CIEDE2000色差公式计算距离
        distance = delta_e_cie2000(target_lab, current_lab)
        if distance < min_distance:
            min_distance = distance
            closest_color = color
    closest_color_position = color_array.index(closest_color)
    return closest_color, [closest_color_position // 12, closest_color_position % 12]

def traverse_directory_os(directory):
    for root, dirs, files in os.walk(directory):
        print(f"Current directory: {root}")

        for file_name in files:
            print(f"File: {file_name}")
            file_path = os.path.join(root, file_name)
            # 打开图片
            img = Image.open(file_path)
            pixels = img.load()
            width, height = img.size

            # 创建 Excel 文件
            workbook = xlsxwriter.Workbook(file_name + '.xlsx')
            worksheet = workbook.add_worksheet()
            flat_color = [c for sub_list in color_dict for c in sub_list]

            # 遍历图片的每个像素点
            # 定义块大小，例如 10x10 的像素块
            BLOCK_SIZE = 24
            # 遍历图片的行，按块处理
            for block_row in range(0, height, BLOCK_SIZE):
                for block_col in range(0, width, BLOCK_SIZE):
                    # 用于统计块内颜色信息，这里以计算平均颜色为例
                    r_sum, g_sum, b_sum = 0, 0, 0
                    num_pixels = 0
                    # 遍历块内的每个像素
                    for row in range(block_row, min(block_row + BLOCK_SIZE, height)):
                        for col in range(block_col, min(block_col + BLOCK_SIZE, width)):
                            color = pixels[col, row]
                            r_sum += color[0]
                            g_sum += color[1]
                            b_sum += color[2]
                            num_pixels += 1
                    # 计算块内平均颜色
                    avg_color = (r_sum // num_pixels, g_sum // num_pixels, b_sum // num_pixels)
                    # result, position = closest_color(avg_color, flat_color)
                    result, position = find_closest_color_ciede2000(avg_color, flat_color)
                    cell_format = workbook.add_format(
                        {'bg_color': rgb_to_hex((result[0], result[1], result[2]))})
                    # 将块对应的单元格区域一次性写入 Excel，这里假设使用的是 xlsxwriter 库，
                    # 根据块的起始行列坐标和块大小来确定写入区域
                    worksheet.write_row(block_row // BLOCK_SIZE, block_col // BLOCK_SIZE,
                                        [str(position)] * BLOCK_SIZE * BLOCK_SIZE, cell_format)
            workbook.close()

def main():
    # most_frequent_color('./colors2png.png')
    traverse_directory_os('./imgs')


if __name__ == '__main__':
    main()

