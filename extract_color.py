import cv2 as cv
import numpy as np

di = {'lower': {'blue': np.array([110, 43, 46]), 'yellow': np.array([26, 43, 46]), 'green': np.array([35, 43, 46])},
      'upper': {'blue': np.array([124, 255, 255]), 'yellow': np.array([34, 255, 255]),
                'green': np.array([77, 255, 255])}}


def extract_color(pic, path, color):
    # 读取图片
    img = cv.imread(pic)
    # 转换到hsv空间
    hsv = cv.cvtColor(img, cv.COLOR_BGR2HSV)
    # 蓝色范围
    lower = di['lower'][color]
    upper = di['upper'][color]
    mask = cv.inRange(hsv, lower, upper)
    res = cv.bitwise_and(img, img, mask=mask)
    cv.imwrite(path, res)


def extract_red(pic, path):
    # m使用inRange方法，拼接mask0,mask1
    img = cv.imdecode(np.fromfile(pic, dtype=np.uint8), -1)
    img_hsv = cv.cvtColor(img, cv.COLOR_BGR2HSV)
    # 区间1
    lower_red = np.array([0, 43, 46])
    upper_red = np.array([10, 255, 255])
    mask0 = cv.inRange(img_hsv, lower_red, upper_red)
    # 区间2
    lower_red = np.array([156, 43, 46])
    upper_red = np.array([180, 255, 255])
    mask1 = cv.inRange(img_hsv, lower_red, upper_red)
    # 拼接两个区间
    mask = mask0 + mask1
    # 提取红色区域
    res = cv.bitwise_and(img, img, mask=mask)
    cv.imwrite(path, res)
