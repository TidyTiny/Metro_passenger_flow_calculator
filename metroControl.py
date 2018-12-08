import metroDesign
from PyQt5 import QtWidgets, QtGui
from PyQt5.QtGui import *
from PyQt5.QtCore import *
from PyQt5.QtWidgets import *
import numpy as np
import matplotlib
matplotlib.use("Qt5Agg")
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
import matplotlib.pyplot as plt
from pandas import read_excel
import sys
import xlwt

# 设计递归深度为1,000,000
sys.setrecursionlimit(1000000)

# 定义关于界面
class HelpWindow(QWidget):
    def __init__(self, parent = None):
        super(HelpWindow, self).__init__(parent)
        self.resize(400,200)
        self.setWindowTitle('欢迎来到帮助')
        self.setWindowIcon(QIcon('icon/person_add_128px_1182113_easyicon.net.ico'))
        self.add_position_layout()

    def add_position_layout(self):
        label = QLabel("当前版本：1.0\n完成时间：2018.10.21\n谢谢使用",self)
        label.move(40,50)

    def handle_click_help(self):
        if not self.isVisible():
            self.show()

# 定义上行断面客流绘图类
class MyFigureUp(FigureCanvas):
    def __init__(self,width=5, height=4, dpi=100):
        self.fig = Figure(figsize=(width, height), dpi=dpi)
        super(MyFigureUp,self).__init__(self.fig)
        self.axes = self.fig.add_subplot(111)
        self.Draw()
        
    def Draw(self):
        url = "savedData.xls"
        updata = read_excel(url)
        upmat = updata.values[4]
        x = np.arange(0,17,1)
        self.axes.bar(x, upmat)

# 定义下行断面客流绘图类      
class MyFigureDown(FigureCanvas):
    def __init__(self,width=5, height=4, dpi=100):
        self.fig = Figure(figsize=(width, height), dpi=dpi)
        super(MyFigureDown,self).__init__(self.fig)
        self.axes = self.fig.add_subplot(111)
        self.Draw()
        
    def Draw(self):
        url = "savedData.xls"
        updata = read_excel(url)
        upmat = updata.values[3]
        x = np.arange(0,17,1)
        self.axes.bar(x, upmat)
        
			
class MainWindow(object):
    def __init__(self):
        app = QtWidgets.QApplication(sys.argv)
        MainWindow = QtWidgets.QMainWindow()
        self.ui = metroDesign.Ui_MainWindow()
        self.ui.setupUi(MainWindow)
		
		# 这是“帮助”按钮的弹窗：
        h = HelpWindow()
        self.ui.pushButton_16.clicked.connect(h.show)
		
		# 定义计算，保存，上行绘图，下行绘图按钮
        self.ui.pushButton_6.clicked.connect(self.OnButtonCompute)
        self.ui.pushButton_7.clicked.connect(self.SaveFile)
        self.ui.pushButton_20.clicked.connect(self.DrawUp)
        self.ui.pushButton_21.clicked.connect(self.DrawDown)
		
        MainWindow.show()
        sys.exit(app.exec_())
	 
    # 定义计算函数
    def OnButtonCompute(self, event):
        url = self.ui.lineEdit.text()
        data = read_excel(url, header = None)
        odmat = data.values
        m = odmat.shape[0]
        arrup = odmat.copy()    
        arrdown = odmat.copy()
        for i in range(m):
            arrdown[i, i:m] = 0
            arrup[i, 0:i] = 0
        downdebus = np.sum(arrdown, axis= 0)
        downaboard = np.sum(arrdown, axis= 1)
        updebus = np.sum(arrup, axis= 0)
        upaboard = np.sum(arrup, axis= 1)

        downdebusStr = str(downdebus).replace("\n", "")
        downaboardStr = str(downaboard).replace("\n", "")
        updebusStr = str(updebus).replace("\n", "")
        upaboardStr = str(upaboard).replace("\n", "")
        
        rowname = ["下行上车数", "下行下车数", "上行上车数", "上行下车数", "下行断面客流量", "上行断面客流量"]
        columnname = ["北客站","北苑","运动公园","行政中心","凤城五路","市图书馆","大明宫西","龙首原","安远门","北大街","钟楼","永宁门","南稍门","体育场","小寨","纬一街","会展中心"]
        self.ui.tableWidget.setVerticalHeaderLabels(rowname)
        self.ui.tableWidget.setHorizontalHeaderLabels(columnname)
       
        down = np.zeros(m)
        up = np.zeros(m)
        down[0] = downdebus[0]
        up[0] = upaboard[0]
        for i in range(1, m-1):
            down[i] = down[i-1]  + downdebus[i] - downaboard[i]
            up[i] = up[i-1]  - updebus[i] + upaboard[i]       

        str1 = [ss for ss in downdebusStr.strip("[ ]").split(" ") if ss != ""]
        str2 = [ss for ss in downaboardStr.strip("[ ]").split(" ") if ss != ""]
        str3 = [ss for ss in updebusStr.strip("[ ]").split(" ") if ss != ""]
        str4 = [ss for ss in upaboardStr.strip("[ ]").split(" ") if ss != ""]
        str5 = [ss for ss in str(down).strip("[]").split(" ") if ss != ""]
        str6 = [ss for ss in str(up).strip("[]").split(" ") if ss != ""]
       
        for i in range(m):
            self.ui.tableWidget.setItem(0, i, QTableWidgetItem(str1[i]))
            self.ui.tableWidget.setItem(1, i, QTableWidgetItem(str2[i]))
            self.ui.tableWidget.setItem(2, i, QTableWidgetItem(str3[i]))
            self.ui.tableWidget.setItem(3, i, QTableWidgetItem(str4[i]))
            self.ui.tableWidget.setItem(4, i, QTableWidgetItem(str5[i]))
            self.ui.tableWidget.setItem(5, i, QTableWidgetItem(str6[i]))
            
        return None
    
    # 定义保存数据函数
    def SaveFile(self):
        filename, type = QFileDialog.getSaveFileName(None, 'Save File', '', ".xls(*.xls)")
        wbk = xlwt.Workbook()
        self.sheet = wbk.add_sheet("sheet", cell_overwrite_ok=True)
        self.add2()
        wbk.save(filename)
    
    #定义保存函数
    def add2(self):
        row = 0
        col = 0         
        for i in range(self.ui.tableWidget.columnCount()):
            for x in range(self.ui.tableWidget.rowCount()):
                try:             
                    teext = str(self.ui.tableWidget.item(row, col).text())
                    self.sheet.write(row, col, teext)
                    row += 1
                except AttributeError:
                    row += 1
            row = 0
            col += 1
    
    # 实例化上行绘图类       
    def DrawUp(self):
        self.F = MyFigureUp(width=3, height=2, dpi=100)
        self.ui.verticalLayout.addWidget(self.F)
    
    # 实例化下行绘图类
    def DrawDown(self):
        self.F = MyFigureDown(width=3, height=2, dpi=100)
        self.ui.verticalLayout_2.addWidget(self.F)
		
if __name__ == "__main__":
    MainWindow()
