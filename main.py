import math

from PIL import Image, ImageFile
import xlsxwriter
import os
import numpy as np
from colormath.color_objects import sRGBColor, LabColor
from colormath.color_diff import delta_e_cie2000
from colormath.color_conversions import convert_color
# from colour import delta_E
import numpy
# from colour.difference import delta_E_CIE1976, delta_E_CIE2000, delta_E_CIE1994, delta_E_CMC, delta_E_DIN99


def patch_asscalar(a):
    return a.item()


setattr(numpy, "asscalar", patch_asscalar)
color_dict = [
    [(99, 170, 171), (130, 192, 206), (126, 187, 217), (163, 191, 218), (185, 193, 220), (152, 176, 214),
     (225, 141, 174), (226, 153, 183), (238, 187, 189), (183, 154, 186), (185, 174, 212), (163, 153, 202)],
    [(59, 181, 170), (76, 173, 218), (60, 167, 206), (74, 171, 208), (110, 144, 211), (151, 167, 213), (230, 59, 106),
     (212, 66, 112), (222, 103, 135), (214, 97, 171), (169, 112, 170), (124, 102, 171)],
    [(39, 161, 178), (33, 149, 174), (36, 136, 192), (62, 131, 188), (48, 116, 187), (98, 128, 182), (208, 60, 112),
     (190, 31, 86), (145, 56, 123), (152, 45, 128), (131, 63, 158), (125, 89, 152)],
    [(63, 160, 145), (49, 141, 146), (56, 111, 183), (65, 112, 199), (49, 88, 157), (76, 114, 151), (171, 62, 95),
     (163, 59, 105), (135, 68, 134), (125, 74, 143), (100, 74, 134), (70, 66, 123)],
    [(194, 215, 217), (188, 220, 195), (185, 212, 186), (212, 203, 69), (169, 151, 121), (168, 149, 131),
     (235, 222, 111), (232, 207, 128), (241, 186, 115), (245, 147, 94), (237, 137, 132), (241, 100, 108)],
    [(89, 187, 126), (169, 213, 129), (186, 213, 107), (188, 188, 22), (156, 94, 67), (150, 94, 75), (235, 197, 76),
     (249, 182, 69), (247, 150, 52), (248, 115, 91), (245, 69, 67), (233, 72, 71)],
    [(54, 155, 74), (42, 186, 50), (88, 180, 29), (155, 184, 30), (115, 74, 52), (104, 74, 72), (238, 185, 74),
     (250, 128, 56), (244, 98, 40), (204, 48, 41), (206, 71, 89), (146, 58, 55)],
    [(48, 75, 67), (31, 93, 79), (37, 116, 76), (124, 137, 54), (83, 83, 62), (105, 84, 83), (219, 161, 75),
     (245, 97, 61), (225, 51, 56), (200, 45, 56), (171, 40, 56), (126, 55, 67)],
    [(220, 184, 177), (233, 182, 166), (239, 193, 165), (239, 207, 187), (236, 234, 227), (234, 232, 228),
     (218, 195, 190), (216, 197, 164), (222, 207, 158), (193, 170, 145), (203, 178, 151), (195, 185, 175)],
    [(231, 159, 124), (218, 150, 103), (236, 170, 116), (244, 194, 138), (240, 218, 173), (226, 215, 193),
     (185, 173, 176), (151, 147, 129), (187, 152, 99), (173, 119, 85), (156, 133, 112), (160, 159, 154)],
    [(209, 122, 67), (208, 116, 55), (213, 154, 93), (152, 117, 88), (136, 129, 123), (130, 124, 117), (138, 114, 118),
     (94, 100, 81), (177, 124, 73), (156, 86, 62), (140, 111, 96), (145, 136, 129)],
    [(127, 85, 67), (90, 63, 58), (91, 71, 65), (77, 72, 73), (58, 56, 58), (73, 70, 71), (127, 97, 115), (66, 74, 79),
     (161, 95, 54), (161, 77, 65), (121, 91, 85), (104, 100, 100)]
]
flat_color = [c for sub_list in color_dict for c in sub_list]
color_name_dict = [
    ["B10", "C2", "C3", "C13", "D16", "D17", "A15", "A3", "A11", "A9", "F14", "F12"],
    ["B6", "C4", "C10", "C17", "D1", "D11", "A4", "A13", "A6", "F1", "F2", "F3"],
    ["C15", "C11", "C5", "C6", "C7", "D2", "A5", "A10", "A7", "F13", "F9", "F6"],
    ["B19", "B7", "C8", "C9", "D3", "C16", "A8", "A14", "F4", "F5", "F8", "F7"],
    ["C14", "B20", "C1", "B18", "M5", "M6", "A15", "A3", "A11", "A9", "F14", "F12"],
    ["B3", "B16", "B13", "B1", "G13", "F10", "A4", "A13", "A6", "F1", "F2", "F3"],
    ["B5", "B4", "B2", "B14", "G7", "F11", "A5", "A10", "A7", "F13", "F9", "F6"],
    ["B15", "B12", "B8", "B17", "B11", "G8", "A8", "A14", "F4", "F5", "F8", "F7"],
    ["E15", "E1", "E14", "E11", "H2", "H1", "H8", "G15", "A2", "H13", "G16", "H9"],
    ["A12", "G3", "G2", "G1", "A1", "H12", "H10", "M1", "G11", "G4", "M4", "H14"],
    ["G6", "G5", "G9", "M9", "H3", "H4", "M10", "M2", "G12", "M13", "M7", "H11"],
    ["G14", "M12", "G17", "H5", "H6", "H7", "M11", "M3", "G10", "M14", "M8", "M15"]
]
flat_color_name = [c for sub_list in color_name_dict for c in sub_list]


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


def closest_color_delta_E(target_color, color_array):
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
    for color in color_array:
        # 使用CIEDE2000色差公式计算距离
        distance = delta_E(target_color, color)
        if distance < min_distance:
            min_distance = distance
            closest_color = color
    closest_color_position = color_array.index(closest_color)
    return closest_color, [closest_color_position // 12, closest_color_position % 12]


def closest_color_cie_1976(target_color, color_array):
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
    target_srgb = sRGBColor(target_color[0], target_color[1], target_color[2], is_upscaled=True)
    # 再将sRGBColor对象转换为LabColor对象
    target_lab = convert_color(target_srgb, LabColor)
    for color in color_array:
        current_srgb = sRGBColor(color[0], color[1], color[2], is_upscaled=True)
        current_lab = convert_color(current_srgb, LabColor)
        # 使用CIEDE2000色差公式计算距离
        distance = delta_E_CIE1976(target_color, color)
        if distance < min_distance:
            min_distance = distance
            closest_color = color
    closest_color_position = color_array.index(closest_color)
    return closest_color, [closest_color_position // 12, closest_color_position % 12]


def closest_color_cie_1994(target_color, color_array):
    min_distance = float('inf')
    closest_color = None
    target_srgb = sRGBColor(target_color[0], target_color[1], target_color[2], is_upscaled=True)
    # 再将sRGBColor对象转换为LabColor对象
    target_lab = convert_color(target_srgb, LabColor)
    for color in color_array:
        current_srgb = sRGBColor(color[0], color[1], color[2], is_upscaled=True)
        current_lab = convert_color(current_srgb, LabColor)
        # 使用CIEDE2000色差公式计算距离
        distance = delta_E_CIE1994(target_color, color)
        if distance < min_distance:
            min_distance = distance
            closest_color = color
    closest_color_position = color_array.index(closest_color)
    return closest_color, [closest_color_position // 12, closest_color_position % 12]


def closest_color_cie_2000(target_color, color_array):
    min_distance = float('inf')
    closest_color = None
    target_srgb = sRGBColor(target_color[0], target_color[1], target_color[2], is_upscaled=True)
    # 再将sRGBColor对象转换为LabColor对象
    target_lab = convert_color(target_srgb, LabColor)
    for color in color_array:
        current_srgb = sRGBColor(color[0], color[1], color[2], is_upscaled=True)
        current_lab = convert_color(current_srgb, LabColor)
        # 使用CIEDE2000色差公式计算距离
        distance = delta_E_CIE2000(target_color, color)
        if distance < min_distance:
            min_distance = distance
            closest_color = color
    closest_color_position = color_array.index(closest_color)
    return closest_color, [closest_color_position // 12, closest_color_position % 12]


def closest_color_cie_cmc(target_color, color_array):
    min_distance = float('inf')
    closest_color = None
    target_srgb = sRGBColor(target_color[0], target_color[1], target_color[2], is_upscaled=True)
    # 再将sRGBColor对象转换为LabColor对象
    target_lab = convert_color(target_srgb, LabColor)
    for color in color_array:
        current_srgb = sRGBColor(color[0], color[1], color[2], is_upscaled=True)
        current_lab = convert_color(current_srgb, LabColor)
        distance = delta_E_CMC(target_color, color)
        if distance < min_distance:
            min_distance = distance
            closest_color = color
    closest_color_position = color_array.index(closest_color)
    return closest_color, [closest_color_position // 12, closest_color_position % 12]


def closest_color_cie_din99(target_color, color_array):
    min_distance = float('inf')
    closest_color = None
    target_srgb = sRGBColor(target_color[0], target_color[1], target_color[2], is_upscaled=True)
    # 再将sRGBColor对象转换为LabColor对象
    target_lab = convert_color(target_srgb, LabColor)
    for color in color_array:
        current_srgb = sRGBColor(color[0], color[1], color[2], is_upscaled=True)
        current_lab = convert_color(current_srgb, LabColor)
        distance = delta_E_DIN99(target_color, color)
        if distance < min_distance:
            min_distance = distance
            closest_color = color
    closest_color_position = color_array.index(closest_color)
    return closest_color, [closest_color_position // 12, closest_color_position % 12]


def get_colour_distance(rgb_1, rgb_2):
    # 将rgb_1中的元素转换为np.int32类型（如果是元组、列表等可迭代结构）
    R_1, G_1, B_1 = [np.int32(x) if isinstance(x, (np.uint8, np.ndarray)) else x for x in rgb_1]
    # 将rgb_2中的元素转换为np.int32类型（同样根据其可能的结构进行转换）
    R_2, G_2, B_2 = [np.int32(x) if isinstance(x, (np.uint8, np.ndarray)) else x for x in rgb_2]
    rmean = (R_1 + R_2) / 2
    R = R_1 - R_2
    G = G_1 - G_2
    B = B_1 - B_2
    return math.sqrt((2 + rmean / 256) * (R ** 2) + 4 * (G ** 2) + (2 + (255 - rmean) / 256) * (B ** 2))


def getColourDistance(target_color, color_array):
    min_distance = float('inf')
    closest_color = None
    for color in color_array:
        distance = get_colour_distance(target_color, color)
        if distance < min_distance:
            min_distance = distance
            closest_color = color
    closest_color_position = color_array.index(closest_color)
    return closest_color, [closest_color_position // 12, closest_color_position % 12]


def generate_color_excel(img: ImageFile, file_name: str):
    pixels = img.load()
    width, height = img.size

    # 创建 Excel 文件
    workbook = xlsxwriter.Workbook(file_name + '.xlsx')
    worksheet = workbook.add_worksheet()

    # 遍历图片的每个像素点
    # 定义块大小，例如 10x10 的像素块
    unique_pixels = list(get_unique_pixels(img, file_name))
    pixels_color_map = {}
    for color in unique_pixels:
        result9, position9 = getColourDistance((color[0], color[1], color[2]), flat_color)
        pixels_color_map[color] = [result9,position9]

    BLOCK_SIZE = 10
    # 遍历图片的行，按块处理
    for block_row in range(0, height, BLOCK_SIZE):
        for block_col in range(0, width, BLOCK_SIZE):
            color = pixels[block_col,block_row]
            result, position = pixels_color_map.get((color[0], color[1], color[2]))
            cell_format = workbook.add_format(
                {'bg_color': rgb_to_hex((result[0], result[1], result[2]))})
            # 将块对应的单元格区域一次性写入 Excel，这里假设使用的是 xlsxwriter 库，
            # 根据块的起始行列坐标和块大小来确定写入区域
            worksheet.write_row(block_row // BLOCK_SIZE, block_col // BLOCK_SIZE,
                                [color_name_dict[position[0]][position[1]]] * BLOCK_SIZE * BLOCK_SIZE, cell_format)
    workbook.close()


def traverse_directory_os(directory):
    for root, dirs, files in os.walk(directory):
        print(f"Current directory: {root}")

        for file_name in files:
            print(f"File: {file_name}")
            file_path = os.path.join(root, file_name)
            # 打开图片
            img = Image.open(file_path)
            generate_color_excel(img, file_name)


def traverse_directory_os2(directory):
    for root, dirs, files in os.walk(directory):
        print(f"Current directory: {root}")

        for file_name in files:
            print(f"File: {file_name}")
            file_path = os.path.join(root, file_name)
            # 打开图片
            img = Image.open(file_path)
            pixels = list(get_unique_pixels(img, file_name))[:100]
            # 创建 Excel 文件
            workbook = xlsxwriter.Workbook(file_name + '.xlsx')
            worksheet = workbook.add_worksheet()
            worksheet.write(0, 0, '图片色彩')
            worksheet.write(0, 1, 'ciede算法')
            worksheet.write(0, 2, '欧几里得算法')
            worksheet.write(0, 3, 'delta_E算法')
            worksheet.write(0, 4, 'CIE1976算法')
            worksheet.write(0, 5, 'CIE1994算法')
            worksheet.write(0, 6, 'CIE2000算法')
            worksheet.write(0, 7, 'CMC算法')
            worksheet.write(0, 8, 'DIN99算法')
            worksheet.write(0, 9, '百度的算法')
            for row_index, p in enumerate(pixels):
                result1, position1 = find_closest_color_ciede2000((p[0], p[1], p[2]), flat_color)
                result2, position2 = closest_color((p[0], p[1], p[2]), flat_color)
                result3, position3 = closest_color_delta_E((p[0], p[1], p[2]), flat_color)
                result4, position4 = closest_color_cie_1976((p[0], p[1], p[2]), flat_color)
                result5, position5 = closest_color_cie_1994((p[0], p[1], p[2]), flat_color)
                result6, position6 = closest_color_cie_2000((p[0], p[1], p[2]), flat_color)
                result7, position7 = closest_color_cie_cmc((p[0], p[1], p[2]), flat_color)
                result8, position8 = closest_color_cie_din99((p[0], p[1], p[2]), flat_color)
                result9, position9 = getColourDistance((p[0], p[1], p[2]), flat_color)

                cell_format = workbook.add_format(
                    {'bg_color': rgb_to_hex((p[0], p[1], p[2]))})
                worksheet.write(row_index + 1, 0, f'rgb({p[0]}, {p[1]}, {p[2]})', cell_format)

                cell_format = workbook.add_format(
                    {'bg_color': rgb_to_hex((result1[0], result1[1], result1[2]))})
                worksheet.write(row_index + 1, 1, f'rgb({result1[0]}, {result1[1]}, {result1[2]})', cell_format)

                cell_format = workbook.add_format(
                    {'bg_color': rgb_to_hex((result2[0], result2[1], result2[2]))})
                worksheet.write(row_index + 1, 2, f'rgb({result2[0]}, {result2[1]}, {result2[2]})', cell_format)

                cell_format = workbook.add_format(
                    {'bg_color': rgb_to_hex((result3[0], result3[1], result3[2]))})
                worksheet.write(row_index + 1, 3, f'rgb({result3[0]}, {result3[1]}, {result3[2]})', cell_format)

                cell_format = workbook.add_format(
                    {'bg_color': rgb_to_hex((result4[0], result4[1], result4[2]))})
                worksheet.write(row_index + 1, 4, f'rgb({result4[0]}, {result4[1]}, {result4[2]})', cell_format)

                cell_format = workbook.add_format(
                    {'bg_color': rgb_to_hex((result5[0], result5[1], result5[2]))})
                worksheet.write(row_index + 1, 5, f'rgb({result5[0]}, {result5[1]}, {result5[2]})', cell_format)

                cell_format = workbook.add_format(
                    {'bg_color': rgb_to_hex((result6[0], result6[1], result6[2]))})
                worksheet.write(row_index + 1, 6, f'rgb({result6[0]}, {result6[1]}, {result6[2]})', cell_format)

                cell_format = workbook.add_format(
                    {'bg_color': rgb_to_hex((result7[0], result7[1], result7[2]))})
                worksheet.write(row_index + 1, 7, f'rgb({result7[0]}, {result7[1]}, {result7[2]})', cell_format)

                cell_format = workbook.add_format(
                    {'bg_color': rgb_to_hex((result8[0], result8[1], result8[2]))})
                worksheet.write(row_index + 1, 8, f'rgb({result8[0]}, {result8[1]}, {result8[2]})', cell_format)

                cell_format = workbook.add_format(
                    {'bg_color': rgb_to_hex((result9[0], result9[1], result9[2]))})
                worksheet.write(row_index + 1, 9, f'rgb({result9[0]}, {result9[1]}, {result9[2]})', cell_format)

            workbook.close()


def get_unique_pixels(image: ImageFile, file_name: str):
    """
    获取图片中所有不重复的像素颜色值（以RGB元组形式）。

    参数:
    image: 图片文件。

    返回:
    set: 包含所有不重复像素颜色值（RGB元组）的集合。
    """
    # 将图片转换为numpy数组，方便处理像素数据
    pixels = np.array(image)
    unique_pixels = set()
    # 获取图片的高度和宽度
    height, width = pixels.shape[:2]
    for row in range(height):
        for col in range(width):
            # 获取每个像素点的RGB颜色值，以元组形式表示
            pixel = tuple(pixels[row, col][:3])
            unique_pixels.add(pixel)
    return unique_pixels


def main():
    # most_frequent_color('./colors2png.png')
    traverse_directory_os('./imgs')
    # traverse_directory_os2('./imgs')


if __name__ == '__main__':
    main()
