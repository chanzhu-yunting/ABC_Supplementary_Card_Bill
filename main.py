# -*- coding:utf-8 -*-
"""
作者: 禅主
时间: 2025/01/14 10:54
"""
from mainWindow import *
from PyQt5.QtWidgets import QApplication, QStyleFactory

if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle(QStyleFactory.create('fusion'))
    window = myMainWindow()
    window.setWindowTitle("ABC附属卡账单计算工具")
    window.setWindowIcon(QIcon("icon/main.png"))  # 设置窗口图标
    window.show()
    sys.exit(app.exec_())
