import ctypes
import sys
import time
import cv2
import numpy as np
import win32con
import win32gui
import win32ui
from PIL import Image


class GameControl():
    def __init__(self, window_name):
        self.hwnd = win32gui.FindWindow(0, window_name)

    def window_full_shot(self, file_name=None):
        """
        窗口截图
            :param file_name=None: 截图文件的保存名称
            :return: file_name为空则返回至self.screen
        """
        try:
            l, t, r, b = win32gui.GetWindowRect(self.hwnd)
            # 39和16为Window与Client高和宽的差值
            h = b - t - 39
            w = r - l - 16
            hwindc = win32gui.GetWindowDC(self.hwnd)
            srcdc = win32ui.CreateDCFromHandle(hwindc)
            memdc = srcdc.CreateCompatibleDC()
            bmp = win32ui.CreateBitmap()
            bmp.CreateCompatibleBitmap(srcdc, w, h)
            memdc.SelectObject(bmp)
            memdc.BitBlt((0, 0), (w, h), srcdc, (8, 31), win32con.SRCCOPY)
            if file_name != None:
                bmp.SaveBitmapFile(memdc, file_name)
                srcdc.DeleteDC()
                memdc.DeleteDC()
                win32gui.ReleaseDC(self.hwnd, hwindc)
                win32gui.DeleteObject(bmp.GetHandle())
                return
            else:
                signedIntsArray = bmp.GetBitmapBits(True)
                img = np.fromstring(signedIntsArray, dtype='uint8')
                img.shape = (h, w, 4)
                srcdc.DeleteDC()
                memdc.DeleteDC()
                win32gui.ReleaseDC(self.hwnd, hwindc)
                win32gui.DeleteObject(bmp.GetHandle())
                #cv2.imshow("image", cv2.cvtColor(img, cv2.COLOR_BGRA2GRAY))
                # cv2.waitKey(0)
                self.screen = cv2.cvtColor(img, cv2.COLOR_BGRA2GRAY)
                return self.screen
        except:
            pass

    def find_img(self, img_template_path):
        """
        查找图片
            :param img_template_path: 欲查找的图片路径
            :return: (maxVal,maxLoc) maxVal为相关性，越接近1越好，maxLoc为得到的坐标
        """
        img_src = self.screen
        img_template = cv2.imread(img_template_path, 0)
        res = cv2.matchTemplate(img_src, img_template, cv2.TM_CCOEFF_NORMED)
        minVal, maxVal, minLoc, maxLoc = cv2.minMaxLoc(res)
        # print(maxLoc)
        return maxVal, maxLoc

    def activate_window(self):
        '''
        激活窗口
        我知道该用user32.SwitchToThisWindow(self.hwnd, True)，但是不好使，原因不明
        '''
        user32 = ctypes.WinDLL('user32.dll')
        user32.SwitchToThisWindow(self.hwnd, False)
        user32.SetForegroundWindow(self.hwnd)

    def quit_game(self):
        """
        退出游戏
        """
        self.takescreenshot()  # 保存一下现场
        win32gui.SendMessage(self.hwnd, win32con.WM_DESTROY, 0, 0)  # 退出游戏
        sys.exit(0)

    def takescreenshot(self):
        '''
        截图
        '''
        img_src_path = 'img\\full.png'
        self.window_full_shot(img_src_path)

# 测试用


def main():
    yys = GameControl(u'阴阳师-网易游戏')


if __name__ == '__main__':
    main()
