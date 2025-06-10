# -*- coding:utf-8 -*-
"""
作者: 禅主
时间: 2025/01/14 10:54
"""
import os
import gc
import sys
import datetime
import winreg as reg
from cryptography.fernet import Fernet
from analyse_email import analyse_deal_mess
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon, QFont
from PyQt5.QtWidgets import QLabel, QMessageBox, QApplication, QFileDialog
from ui_mainWindow import *


class myMainWindow(QtWidgets.QMainWindow):

    def __init__(self):
        super(myMainWindow, self).__init__()
        # 初始化一些变量
        self.file_email_path = None
        self.generate_state = None
        self.mainUi = Ui_MainWindow()
        self.mainUi.setupUi(self)

        # 初始化状态条
        self.initStatusBar()
        # 设置字体
        font = QFont()
        font.setFamily("Courier New")  # 设置字体为等宽字体
        font.setPointSize(12)  # 设置字体大小
        self.mainUi.plainTextEdit.setFont(font)
        # 设置只读属性
        self.mainUi.plainTextEdit.setReadOnly(True)

        self.mainUi.resetBtn.setEnabled(False)
        self.mainUi.chooseInputFileBtn.clicked.connect(self.inputINPFiles)
        self.mainUi.generateBtn.clicked.connect(self.start_progress)
        self.mainUi.resetBtn.clicked.connect(self.reset)

    def initStatusBar(self):
        # 添加一个显示永久信息的标签控件
        self.mainUi.info = QLabel(self)
        self.mainUi.info.setText(u"© 2025 禅主 All rights reserved.")
        self.mainUi.info.setAlignment(Qt.AlignRight)
        self.mainUi.statusbar.addPermanentWidget(self.mainUi.info)
        self.setStatusBar(self.mainUi.statusbar)

    def inputINPFiles(self):
        """
            读取eml文件,
            :return:
        """
        print("欢迎使用软件！")
        # 打开文件选择对话框
        file_dialog = QFileDialog()
        file_dialog.setFileMode(QFileDialog.ExistingFiles)
        file_dialog.setNameFilter("Email File (*.eml)")
        if file_dialog.exec_():
            # 获取所选文件的路径列表
            self.file_email_path = file_dialog.selectedFiles()[0]
            # 获取文件夹路径
            if self.file_email_path:
                self.mainUi.inputFileEdit.setText(os.path.basename(self.file_email_path))


    def start_progress(self):

        card_number = self.mainUi.templateEdit.text().strip().replace(' ', '')
        if card_number and len(card_number) == 4:
            reply = QMessageBox.information(
                self,
                "提示!",
                f"尾号{card_number}是否为附属卡？",
                QMessageBox.Yes,
                QMessageBox.No,
            )
            if reply == QMessageBox.Yes:
                card_number += '附'
        elif len(card_number) != 4:
            QMessageBox.critical(
                self,
                "错误!",
                "卡号后四位异常，请检查！",
                QMessageBox.Yes | QMessageBox.No
            )
            return False
        else:
            QMessageBox.critical(
                self,
                "错误!",
                "卡号为空，请检查！",
                QMessageBox.Yes | QMessageBox.No
            )
            return False
        if self.file_email_path:
            pass
        else:
            QMessageBox.critical(
                self,
                "错误!",
                "请选择对应的电子对账单文件！",
                QMessageBox.Yes | QMessageBox.No
            )
            return False

        self.generate_state = False
        self.mainUi.generateBtn.setText("分析中...")
        self.mainUi.generateBtn.setEnabled(False)
        self.mainUi.resetBtn.setEnabled(False)
        acc_mess = analyse_deal_mess(self.file_email_path, card_number)
        if acc_mess is None:
            self.mainUi.plainTextEdit.appendPlainText(
                f"未检索到尾号 {card_number} 的农行账户数据！")
        elif isinstance(acc_mess, list):
            if len(acc_mess) == 2:
                total_balance = acc_mess[0]
                amounts = acc_mess[1]
                self.mainUi.plainTextEdit.appendPlainText(
                    f"您尾号 {card_number} 的农行账户累计入账金额为：{total_balance:.2f}")
                self.mainUi.plainTextEdit.appendPlainText(f"您尾号 {card_number} 的农行账户上月总计消费 {len(amounts)} 笔")
                print(f"每笔交易金额：{amounts}")
            else:
                self.mainUi.plainTextEdit.appendPlainText(f"计算异常！")
        else:
            self.mainUi.plainTextEdit.appendPlainText(f"计算失败！")
        self.generate_state = True
        self.mainUi.generateBtn.setText("已完成")

        self.mainUi.resetBtn.setEnabled(True)
        QApplication.processEvents()  # Process events to update UI
        return True

    def reset(self):
        self.mainUi.generateBtn.setEnabled(True)
        self.mainUi.resetBtn.setEnabled(False)
        self.mainUi.inputFileEdit.clear()
        self.mainUi.templateEdit.clear()
        self.mainUi.generateBtn.setText("计算")
        self.generate_state = False
        self.file_email_path = None  # contain the file name
        self.mainUi.plainTextEdit.clear()

    def closeEvent(self, event):  # 函数名固定不可变
        reply = QtWidgets.QMessageBox.question(
            self,
            "Notice!",
            "Are you sure to exit?",
            QtWidgets.QMessageBox.Yes,
            QtWidgets.QMessageBox.No,
        )
        if reply == QtWidgets.QMessageBox.Yes:
            self.generate_state = False
            self.file_email_path = None  # contain the file name
            gc.collect()
            event.accept()  # 关闭窗口
        else:
            event.ignore()  # 忽视点击X事件


########################################################################
# execute app
########################################################################
if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    window = myMainWindow()
    window.setWindowTitle("ABC附属卡账单计算工具")
    window.setWindowIcon(QIcon("icon/main.png"))  # 设置窗口图标
    window.show()
    sys.exit(app.exec_())
########################################################################
# end
########################################################################
