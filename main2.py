import numpy as np


def closest_color(target_color, color_array):
    min_distance = float('inf')
    closest_color_found = None
    target_color = np.array(target_color)
    for color in color_array:
        color = np.array(color)
        distance = np.sqrt(np.sum((target_color - color) ** 2))
        if distance < min_distance:
            min_distance = distance
            closest_color_found = color
    return closest_color_found

def main():
    color_array = [(255, 0, 0), (0, 255, 0), (0, 0, 255)]
    target_color = (200, 0, 0)
    result = closest_color(target_color, color_array)
    print("最接近的颜色：", result)


if __name__ == '__main__':
    main()
