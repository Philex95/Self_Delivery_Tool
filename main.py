import datetime
import openpyxl
import pandas as pd
import sys
import function
from PyQt5 import QtCore, QtWidgets
from PyQt5.QtWidgets import QApplication,QMainWindow
import traceback
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

        self.getOutputBtn.clicked.connect(self.OutPutDataClick)
        self.saveFileBtn.clicked.connect(self.SelectSaveBtnClick)
        self.dataFileBtn.clicked.connect(self.SelectFileBtnClick)

        self.getCsvBtn.clicked.connect(self.OutPutCsvClick)
        self.dataXlsxFileBtn.clicked.connect(self.SelectXlsxFileBtnClick)


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

    def SelectSaveBtnClick(self):
        floder = QtWidgets.QFileDialog.getExistingDirectory(None,"选取文件夹",'.\\')  # 起始路径
        self.saveFileShow.setText(floder)
        self.savePath = floder + '/'

    def SelectFileBtnClick(self):
        floder = QtWidgets.QFileDialog.getExistingDirectory(None,"选取文件夹",'.\\')  # 起始路径
        self.dataFileShow.setText(floder)
        self.filePath = floder

    def SelectXlsxFileBtnClick(self):
        filename , _= QtWidgets.QFileDialog.getOpenFileName(None,"选取文件夹",'.\\',"Xlsx Files(*.xlsx)")  # 起始路径
        self.dataXlsxFileShow.setText(filename)

    def OutPutDataClick(self):
        if '/' not in self.saveFileShow.toPlainText():
            self.saveFileShow.setText("没有选择保存路径！")
            return
        elif '/' not in self.dataFileShow.toPlainText():
            self.dataFileShow.setText("没有选择数据文件路径！")
            return

        filePath = self.dataFileShow.toPlainText()
        savePath = self.saveFileShow.toPlainText() + '/'

        goodDataBook = openpyxl.load_workbook('./信息/货物属性表.xlsx')
        goodDataSheet = goodDataBook.worksheets[0]
        goodDic = function.GetGoodsInfo(goodDataSheet)
        self.LogOut.clear()
        self.wrongLogOut.clear()
        file = function.ReadDataFromFile(filePath, '.csv')
        for item in file:
            try:
                self.CallLogOutDetail("正在读取文件：" + str(item) + "........")
                runName = file[item]
                tarCSV = pd.read_csv(item, index_col=0)
                tarCSV.to_excel(savePath + runName + '.xlsx')
                dfCol = pd.read_excel('./信息/目标列表.xlsx', index_col=0)
                arrCol = dfCol['目标列名'].tolist()

                data = function.LoadDataGroup(tarCSV, arrCol, goodDic, self.CallWorngLogOutDetail)
                dataGrp = function.dataGroupBy(data)

                self.CallLogOutDetail("正在输出备货文件：" + runName + ".xlsx........")
                function.CreateXlsx(dataGrp, runName, savePath)
                self.CallLogOutDetail("正在输出备货文件：" + runName + ".pdf........")
                function.CreateWordAndPDF(dataGrp, arrCol, runName, savePath, self.CallWorngLogOutDetail)
            except Exception as e:
                erroMessage = '''exception:%s
                \ntraceBackFormat:%s
                \ndate:%s
                ''' % (str(e), str(traceback.format_exc()), datetime.datetime.today().strftime("%Y-%m-%d, %H:%M:%S"))
                with open(".\\error.txt", 'w+') as f:
                    f.write(erroMessage)
        self.CallLogOutDetail("所有文件已经输出完成！")

    def OutPutCsvClick(self):
        if '/' not in self.dataXlsxFileShow.toPlainText():
            self.dataXlsxFileShow.setText("没有选择数据文件！")
            return

        filePath = self.dataXlsxFileShow.toPlainText()
        prodPath = './信息/货物属性表.xlsx'
        prodDataWorkBook = openpyxl.load_workbook(prodPath)
        prodDataSheet = prodDataWorkBook.worksheets[0]
        prodDic = function.GetProductWeight(prodDataSheet)
        self.wrongCsv.clear()
        self.LogOutCsv.clear()
        function.CreateCSV(filePath, prodDic, self.CallLogOutCSVDetail, self.CallWorngLogOutCSVDetail)

    def CallLogOutDetail(self, detail):
        self.LogOut.append(detail)
        QApplication.processEvents()

    def CallWorngLogOutDetail(self,detail):
        self.wrongLogOut.append(detail)
        QApplication.processEvents()

    def CallLogOutCSVDetail(self, detail):
        self.LogOutCsv.append(detail)
        QApplication.processEvents()

    def CallWorngLogOutCSVDetail(self,detail):
        self.wrongCsv.append(detail)
        QApplication.processEvents()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWin = QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(myWin)
    myWin.show()
    sys.exit(app.exec_())



