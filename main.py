# @Description: Heros of the Storm - Tracer Technology - Level 7 Talent 3 Auto Reload
# @Author:      PetrelPine [https://github.com/PetrelPine]

import win32gui, win32ui, win32con
from ctypes import windll
from PIL import Image
import time
import key  # my own module
# import winsound
# import numpy
# import cv2

# 获取后台窗口的句柄，注意后台窗口不能最小化
hWnd = win32gui.FindWindow("Heroes of the Storm", None)  # 窗口的类名可以用Visual Studio的SPY++工具获取
# 获取句柄窗口的大小信息
left, top, right, bot = win32gui.GetWindowRect(hWnd)
width = int((right - left) * 1.25)  # 需要x1.25，因为windows启用了125%的缩放
height = int((bot - top) * 1.25)  # 需要x1.25，因为windows启用了125%的缩放

# 如果要截图到打印设备：
# 最后一个int参数：0 - 保存整个窗口，1 - 只保存客户区。如果PrintWindow成功函数返回值为 1
# result = windll.user32.PrintWindow(hWnd, saveDC.GetSafeHdc(), 0)
# print(result) # PrintWindow成功则输出 1
#
# 保存图像
# 方法一：windows api 保存
# 保存bitmap到文件
# saveBitMap.SaveBitmapFile(saveDC, "scr\\img_Winapi.bmp")
#
# 方法二(第一部分)：PIL保存
# 获取位图信息
# bmp_info = saveBitMap.GetInfo()
# bmp_str = saveBitMap.GetBitmapBits(True)
# 生成图像
# img_PIL = Image.frombuffer('RGB', (bmp_info['bmWidth'], bmp_info['bmHeight']), bmp_str, 'raw', 'BGRX', 0, 1)
# 方法二（第二部分）：PIL保存
# PrintWindow成功，保存到文件，显示到屏幕
# img_PIL.save("scr\\img_PIL.png")  # 保存
# img_PIL.show()  # 显示
#
# 方法三（第一部分）：opencv + numpy保存
# 获取位图信息
# signedIntsArray = saveBitMap.GetBitmapBits(True)
# 方法三（第二部分）：opencv+numpy保存
# PrintWindow成功，保存到文件，显示到屏幕
# im_opencv = numpy.frombuffer(signedIntsArray, dtype='uint8')
# im_opencv.shape = (height, width, 4)
# cv2.cvtColor(im_opencv, cv2.COLOR_BGRA2RGB)
# cv2.imwrite("im_opencv.jpg", im_opencv, [int(cv2.IMWRITE_JPEG_QUALITY), 100])  # 保存
# cv2.namedWindow('im_opencv')  # 命名窗口
# cv2.imshow("im_opencv", im_opencv)  # 显示
# cv2.waitKey(0)
# cv2.destroyAllWindows()
#
# 内存释放
# win32gui.DeleteObject(saveBitMap.GetHandle())
# saveDC.DeleteDC()
# mfcDC.DeleteDC()
# win32gui.ReleaseDC(hWnd, hWndDC)


# 检查当前状态是否需要按下D键
def check_status():
    rely = 0  # 可信度

    # 右下D键位置
    # x, y = 1132, 1051
    # r, g, b = 8, 251, 62
    pos_1 = img_PIL.getpixel((1132, 1051))  # 右下角D键状态判断
    if 0 < pos_1[0] < 18 and 241 < pos_1[1] < 255 and 52 < pos_1[2] < 72:
        rely += 1

    # '正' 字位置
    # x, y = 930, 882
    # x, y = 938, 882
    # r, g, b = 255, 255, 255 (white)
    zheng_1 = img_PIL.getpixel((930, 882))
    zheng_2 = img_PIL.getpixel((938, 882))
    if 240 < zheng_1[0] and 240 < zheng_1[1] and 240 < zheng_1[2]:
        rely += 1
    if 240 < zheng_2[0] and 240 < zheng_2[1] and 240 < zheng_2[2]:
        rely += 1

    # '填' 字位置
    # x, y = 983, 880
    # x, y = 988, 880
    # x, y = 983, 875
    # r, g, b = 255, 255, 255 (white)
    tian_1 = img_PIL.getpixel((983, 880))
    tian_2 = img_PIL.getpixel((988, 880))
    tian_3 = img_PIL.getpixel((983, 875))
    if 240 < tian_1[0] and 240 < tian_1[1] and 240 < tian_1[2]:
        rely += 1
    if 240 < tian_2[0] and 240 < tian_2[1] and 240 < tian_2[2]:
        rely += 1
    if 240 < tian_3[0] and 240 < tian_3[1] and 240 < tian_3[2]:
        rely += 1

    # rely最大值为6，大于等于5就算是换弹前半部分
    if rely >= 5:
        # load time = 0.750 s
        time.sleep(0.350)
        return 1  # 需要按下D
    else:
        return 0  # 不需要按下D


# main loop
print('Start in 5 seconds.')
time.sleep(5)
for _ in range(20000):
    # 返回句柄窗口的设备环境，覆盖整个窗口，包括非客户区，标题栏，菜单，边框
    hWndDC = win32gui.GetWindowDC(hWnd)
    # 创建设备描述表
    mfcDC = win32ui.CreateDCFromHandle(hWndDC)
    # 创建内存设备描述表
    saveDC = mfcDC.CreateCompatibleDC()
    # 创建位图对象准备保存图片
    saveBitMap = win32ui.CreateBitmap()
    # 为bitmap开辟存储空间
    saveBitMap.CreateCompatibleBitmap(mfcDC, width, height)
    # 将截图保存到saveBitMap中
    saveDC.SelectObject(saveBitMap)
    # # 保存bitmap到内存设备描述表
    saveDC.BitBlt((0, 0), (width, height), mfcDC, (0, 0), win32con.SRCCOPY)

    # 获取位图信息
    bmp_info = saveBitMap.GetInfo()
    bmp_str = saveBitMap.GetBitmapBits(True)
    # 生成图像
    img_PIL = Image.frombuffer('RGB', (bmp_info['bmWidth'], bmp_info['bmHeight']), bmp_str, 'raw', 'BGRX', 0, 1)

    # 内存释放
    win32gui.DeleteObject(saveBitMap.GetHandle())
    saveDC.DeleteDC()
    mfcDC.DeleteDC()
    win32gui.ReleaseDC(hWnd, hWndDC)

    if check_status() == 1:
        key.PressKey(0x20)
        time.sleep(0.025)
        key.ReleaseKey(0x20)
        print('D Confirm!')
        # winsound.Beep(500, 300)
        time.sleep(1.250)  # atk time = 1.250 s
    else:
        print('Not Confirm!')
        time.sleep(0.175)
