from game_ctl import *
import win32com.client
import os
import threading

import logsystem

# 初始化对象
log = logsystem.WriteLog()
yys = GameControl(u'阴阳师-网易游戏')
log.writeinfo('Registration successful')


def is_admin():
    '''
    UAC申请，获得管理员权限
    '''
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False


def get_screen():
    '''
    每隔2s获取一次截图，截图缓存至yys.screen中
    '''
    while True:
        yys.window_full_shot()
        time.sleep(2)


def wait_game_img(img_path):
    """
    等待游戏图像，每隔2s查询1次，查询成功后休息10再继续
        :param img_path: 图片路径
    """
    try:
        while True:
            maxVal, maxLoc = yys.find_img(img_path)
            if maxVal > 0.97:
                yys.activate_window()
                log.writeinfo('Figure found: '+img_path)
                time.sleep(8)
            time.sleep(2)
    except:
        # 如果错误则输出警告
        log.writeinfo("Threading error! " + img_path)


def start():
    """
    启动
    """
    # 设置图片路径
    info = "img"

    # 获取所有待查询的图片
    listfile = os.listdir(info)

    # 激活窗口
    yys.activate_window()
    log.writeinfo('Activation successful')
    time.sleep(1)

    # 启动截图
    t = threading.Thread(target=get_screen)
    t.start()
    time.sleep(1)

    # 对每一个待查询的图片启动一个单一的线程
    for figure in listfile:
        if figure[-4:] == '.png':
            path = 'img\\'+figure
            log.writeinfo('Files: ' + path)
            t = threading.Thread(
                target=wait_game_img, args=(path,))
            t.start()


if __name__ == "__main__":
    log.writeinfo('python version: %s', sys.version)

    try:
        # 检测管理员权限
        if is_admin():
            log.writeinfo('UAC pass')

            # 开始
            start()

        else:
            ctypes.windll.shell32.ShellExecuteW(
                None, "runas", sys.executable, __file__, None, 1)
    except KeyboardInterrupt:
        log.writeinfo('terminated')
        os._exit(0)
    else:
        pass
