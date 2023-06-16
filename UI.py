# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'UI.ui'
#
# Created by: PyQt5 UI code generator 5.15.4
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(612, 308)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.tabWidget = QtWidgets.QTabWidget(self.centralwidget)
        self.tabWidget.setGeometry(QtCore.QRect(0, 0, 611, 301))
        self.tabWidget.setObjectName("tabWidget")
        self.widget = QtWidgets.QWidget()
        self.widget.setObjectName("widget")
        self.getCsvBtn = QtWidgets.QPushButton(self.widget)
        self.getCsvBtn.setGeometry(QtCore.QRect(490, 90, 101, 171))
        self.getCsvBtn.setObjectName("getCsvBtn")
        self.wrongCsv = QtWidgets.QTextBrowser(self.widget)
        self.wrongCsv.setGeometry(QtCore.QRect(310, 90, 171, 171))
        self.wrongCsv.setObjectName("wrongCsv")
        self.LogOutCsv = QtWidgets.QTextBrowser(self.widget)
        self.LogOutCsv.setGeometry(QtCore.QRect(20, 90, 281, 171))
        self.LogOutCsv.setObjectName("LogOutCsv")
        self.label_5 = QtWidgets.QLabel(self.widget)
        self.label_5.setGeometry(QtCore.QRect(310, 70, 151, 16))
        self.label_5.setObjectName("label_5")
        self.label_6 = QtWidgets.QLabel(self.widget)
        self.label_6.setGeometry(QtCore.QRect(20, 70, 54, 12))
        self.label_6.setObjectName("label_6")
        self.dataXlsxFileShow = QtWidgets.QTextBrowser(self.widget)
        self.dataXlsxFileShow.setGeometry(QtCore.QRect(20, 10, 461, 51))
        self.dataXlsxFileShow.setObjectName("dataXlsxFileShow")
        self.dataXlsxFileBtn = QtWidgets.QPushButton(self.widget)
        self.dataXlsxFileBtn.setGeometry(QtCore.QRect(490, 10, 101, 51))
        self.dataXlsxFileBtn.setObjectName("dataXlsxFileBtn")
        self.tabWidget.addTab(self.widget, "")
        self.tab = QtWidgets.QWidget()
        self.tab.setObjectName("tab")
        self.dataFileBtn = QtWidgets.QPushButton(self.tab)
        self.dataFileBtn.setGeometry(QtCore.QRect(490, 10, 101, 41))
        self.dataFileBtn.setObjectName("dataFileBtn")
        self.saveFileBtn = QtWidgets.QPushButton(self.tab)
        self.saveFileBtn.setGeometry(QtCore.QRect(490, 60, 101, 41))
        self.saveFileBtn.setObjectName("saveFileBtn")
        self.label_2 = QtWidgets.QLabel(self.tab)
        self.label_2.setGeometry(QtCore.QRect(310, 110, 151, 16))
        self.label_2.setObjectName("label_2")
        self.getOutputBtn = QtWidgets.QPushButton(self.tab)
        self.getOutputBtn.setGeometry(QtCore.QRect(490, 130, 101, 131))
        self.getOutputBtn.setObjectName("getOutputBtn")
        self.LogOut = QtWidgets.QTextBrowser(self.tab)
        self.LogOut.setGeometry(QtCore.QRect(20, 130, 281, 131))
        self.LogOut.setObjectName("LogOut")
        self.wrongLogOut = QtWidgets.QTextBrowser(self.tab)
        self.wrongLogOut.setGeometry(QtCore.QRect(310, 130, 171, 131))
        self.wrongLogOut.setObjectName("wrongLogOut")
        self.label = QtWidgets.QLabel(self.tab)
        self.label.setGeometry(QtCore.QRect(20, 110, 54, 12))
        self.label.setObjectName("label")
        self.dataFileShow = QtWidgets.QTextBrowser(self.tab)
        self.dataFileShow.setGeometry(QtCore.QRect(20, 10, 461, 41))
        self.dataFileShow.setObjectName("dataFileShow")
        self.saveFileShow = QtWidgets.QTextBrowser(self.tab)
        self.saveFileShow.setGeometry(QtCore.QRect(20, 60, 461, 41))
        self.saveFileShow.setObjectName("saveFileShow")
        self.tabWidget.addTab(self.tab, "")
        MainWindow.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        self.tabWidget.setCurrentIndex(0)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "排车便利工具"))
        self.getCsvBtn.setText(_translate("MainWindow", "导出数据模板CSV"))
        self.label_5.setText(_translate("MainWindow", "待补充信息："))
        self.label_6.setText(_translate("MainWindow", "处理信息："))
        self.dataXlsxFileBtn.setText(_translate("MainWindow", "选择execel文件"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.widget), _translate("MainWindow", "输出PODFather信息导入文件"))
        self.dataFileBtn.setText(_translate("MainWindow", "选择数据文件夹"))
        self.saveFileBtn.setText(_translate("MainWindow", "选择输出文件夹"))
        self.label_2.setText(_translate("MainWindow", "待补充信息："))
        self.getOutputBtn.setText(_translate("MainWindow", "输出备货单与PDF"))
        self.label.setText(_translate("MainWindow", "处理信息："))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab), _translate("MainWindow", "输出备货表与贴纸PDF"))
