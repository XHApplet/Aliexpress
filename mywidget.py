# -*- coding:utf-8 -*-
"""
@Author: lamborghini
@Date: 2018-04-12 19:27:27
@Desc: 
"""

import sys
import mainwidget
import mycrawler
from PyQt5 import QtWidgets


class CMyWidget(QtWidgets.QMainWindow, mainwidget.Ui_MainWindow):
    def __init__(self):
        super(CMyWidget, self).__init__()
        self.m_ALiExpress = None
        self.setupUi(self)
        self.InitUI()
        self.InitConnect()


    def InitUI(self):
        self.comboBox.addItems(["厨房用品", "收纳用品"])
        self.pushButton.setText("开始")

    def InitConnect(self):
        self.pushButton.clicked.connect(self.Start)

    def Start(self):
        iIndex = self.comboBox.currentIndex()
        if iIndex == 0:
            url = "https://www.aliexpress.com/category/100005652/bakeware/"
        else:
            url = "https://www.aliexpress.com/category/1541/home-storage-organization/"
        print(url)


def Show():
    app = QtWidgets.QApplication(sys.argv)
    g_Obj = CMyWidget()
    g_Obj.show()
    sys.exit(app.exec_())
