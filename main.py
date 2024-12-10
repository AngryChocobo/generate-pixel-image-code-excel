from algorithms import *


# 兼容 numpy
def patch_asscalar(a):
    return a.item()


setattr(numpy, "asscalar", patch_asscalar)

# 缓存用户拥有的拼豆的颜色（根据产品图片识别得到，无需每次都处理）
user_color_matrix = [
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
flat_user_color_list = [c for sub_list in user_color_matrix for c in sub_list]
# 拼豆颜色与编码的关系，通过截图OCR得到
user_color_name_matrix = [
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
flat_color_name = [c for sub_list in user_color_name_matrix for c in sub_list]


def rgb_to_hex(rgb):
    r, g, b = rgb
    return "#{:02x}{:02x}{:02x}".format(r, g, b)


def generate_user_color_martix(image_path):
    """生成用户拼豆颜色的矩阵"""
    global user_color_matrix
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
            row.append((r_avg, g_avg, b_avg))
        res.append(row)
        top = top + item_height
    user_color_matrix = res
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


def generate_color_excel(img: ImageFile, file_name: str):
    pixels = img.load()
    width, height = img.size

    # 创建 Excel 文件
    workbook = xlsxwriter.Workbook(file_name + '.xlsx')

    # 遍历图片的每个像素点
    unique_pixels = list(get_unique_pixels(img, file_name))

    # 定义块大小，例如 10x10 的像素块，每个像素都处理的话，需要的拼豆太多
    BLOCK_SIZE = 10
    algorithms = [baidu_algorithms, find_closest_color_ciede2000, euclidean_algorithm, closest_color_delta_E,
                  closest_color_cie_1976, closest_color_cie_2000, closest_color_cie_1994, closest_color_cie_cmc,
                  closest_color_cie_din99]
    for algorithm in algorithms:
        worksheet = workbook.add_worksheet(name=algorithm.__name__)
        # 每种算法缓存自己去重后像素的近似值
        pixels_color_map = {}
        for color in unique_pixels:
            result, position = algorithm((color[0], color[1], color[2]), flat_user_color_list)
            pixels_color_map[color] = [result, position]
        # 遍历图片的行，按块处理
        for block_row in range(0, height, BLOCK_SIZE):
            for block_col in range(0, width, BLOCK_SIZE):
                color = pixels[block_col, block_row]
                # 直接读缓存，无需每个像素都重新计算近似值
                result, position = pixels_color_map.get((color[0], color[1], color[2]))
                cell_format = workbook.add_format(
                    {'bg_color': rgb_to_hex((result[0], result[1], result[2]))})
                # 将块对应的单元格区域一次性写入 Excel，这里假设使用的是 xlsxwriter 库，
                # 根据块的起始行列坐标和块大小来确定写入区域
                # 同时写入对应拼豆的编码、背景色
                worksheet.write_row(block_row // BLOCK_SIZE, block_col // BLOCK_SIZE,
                                    [user_color_name_matrix[position[0]][position[1]]] * BLOCK_SIZE * BLOCK_SIZE,
                                    cell_format)
    workbook.close()
    print(file_name + ': 生成拼豆excel完毕')


def traverse_directory(directory, handler):
    for root, dirs, files in os.walk(directory):
        print(f"当前目录: {root}")

        for file_name in files:
            print(f"读取文件: {file_name}")
            file_path = os.path.join(root, file_name)
            # 打开图片
            img = Image.open(file_path)
            handler(img, file_name)


def generate_algorithmes_compare_excel(img: ImageFile, file_name: str):
    """横向对比各个算法对于颜色寻找近似值的效果"""
    # 找到图片去重之后的颜色即可，取前100条足以看出算法的效果
    pixels = list(get_unique_pixels(img, file_name))[:100]
    # 创建 Excel 文件
    workbook = xlsxwriter.Workbook('算法效果' + file_name + '.xlsx')
    worksheet = workbook.add_worksheet()
    algorithme_list = [('ciede算法', find_closest_color_ciede2000), ('欧几里得算法', euclidean_algorithm),
                       ('delta_E算法', closest_color_delta_E), ('CIE1976算法', closest_color_cie_1976),
                       ('CIE1994算法', closest_color_cie_1994), ('CIE2000算法', closest_color_cie_2000),
                       ('CMC算法', closest_color_cie_cmc), ('DIN99算法', closest_color_cie_din99),
                       ('百度的算法', baidu_algorithms)]
    worksheet.write(0, 0, '原图色彩')
    for i, (name, method) in enumerate(algorithme_list):
        # 表头
        worksheet.write(0, i + 1, name)
        for row_index, p in enumerate(pixels):
            # 原图色彩
            cell_format = workbook.add_format(
                {'bg_color': rgb_to_hex((p[0], p[1], p[2]))})
            worksheet.write(row_index + 1, 0, f'rgb({p[0]}, {p[1]}, {p[2]})', cell_format)

            result1, position1 = method((p[0], p[1], p[2]), flat_user_color_list)
            # 算法色彩
            cell_format = workbook.add_format(
                {'bg_color': rgb_to_hex((result1[0], result1[1], result1[2]))})
            worksheet.write(row_index + 1, 1, f'rgb({result1[0]}, {result1[1]}, {result1[2]})', cell_format)
    workbook.close()
    print(file_name + ': 生成算法效果对比完毕')


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
    # generate_user_color_martix('./colors2png.png')
    traverse_directory('./imgs', generate_color_excel)
    print('---')
    traverse_directory('./imgs', generate_algorithmes_compare_excel)
    input("按回车键退出...")


if __name__ == '__main__':
    main()

# pyinstaller -D main.py 生成exe