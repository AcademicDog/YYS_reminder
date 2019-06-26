from big_boss import *
from PyQt5.QtGui import QTextCursor
from PyQt5.QtWidgets import QApplication, QMainWindow
from Ui_big_boss import *
import logging

class GuiLogger(logging.Handler):
    def emit(self, record):
        self.edit.append(self.format(record))  # implementation of append_line omitted
        self.edit.moveCursor(QTextCursor.End)

class MyMainWindow(QMainWindow):
    def __init__(self, parent = None):
        super(MyMainWindow, self).__init__(parent)
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)

        self.ui.textEdit.ensureCursorVisible()
        
        # 将日志转存至UI
        h = GuiLogger()
        h.edit = self.ui.textEdit
        logger = logging.getLogger()
        logger.setLevel(logging.INFO)
        logger.addHandler(h)

    def start_onmyoji(self):
        start()

    def stop_onmyoji(self):
        logging.info('Quitting!')
        os._exit(0)

if __name__=="__main__":  
    
    try:
        # 检测管理员权限
        if is_admin():
            log.writeinfo('UAC pass')

            app = QApplication(sys.argv)
            myWin = MyMainWindow()
            myWin.show()
            sys.exit(app.exec_())

        else:
            ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 1)
    except:
        pass