import math

from PIL import Image, ImageFile
import xlsxwriter
import os
import numpy as np
from colormath.color_objects import sRGBColor, LabColor
from colormath.color_diff import delta_e_cie2000
from colormath.color_conversions import convert_color
from colour import delta_E
import numpy
from colour.difference import delta_E_CIE1976, delta_E_CIE2000, delta_E_CIE1994, delta_E_CMC, delta_E_DIN99

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

def baidu_algorithms(target_color, color_array):
    min_distance = float('inf')
    closest_color = None
    for color in color_array:
        distance = get_colour_distance(target_color, color)
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


def euclidean_algorithm(target_color, color_array):
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
