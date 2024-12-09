from PIL import Image
import xlsxwriter
import os
from collections import Counter
import numpy as np

color_dict = []

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
            BLOCK_SIZE = 10
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
                    result, position = closest_color(avg_color, flat_color)
                    cell_format = workbook.add_format(
                        {'bg_color': rgb_to_hex((result[0], result[1], result[2]))})
                    # 将块对应的单元格区域一次性写入 Excel，这里假设使用的是 xlsxwriter 库，
                    # 根据块的起始行列坐标和块大小来确定写入区域
                    worksheet.write_row(block_row // BLOCK_SIZE, block_col // BLOCK_SIZE,
                                        [str(position)] * BLOCK_SIZE * BLOCK_SIZE, cell_format)
            workbook.close()

def main():
    most_frequent_color('./colors2png.png')
    traverse_directory_os('./imgs')


if __name__ == '__main__':
    main()

