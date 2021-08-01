import sys, os

from PyQt5.QtCore import QDir

from ui import *
from PyQt5.QtWidgets import QApplication, QMainWindow, QMessageBox, QTableView, QFileDialog
from PyQt5 import QtSql
from PyQt5 import QtCore
import xlwt

class QSqlCenteredTableModel(QtSql.QSqlTableModel):
    def __init__(self):
        QtSql.QSqlTableModel.__init__(self)

    def data(self, index, role=None):
        if role == QtCore.Qt.TextAlignmentRole:
            return QtCore.Qt.AlignCenter
        return QtSql.QSqlTableModel.data(self, index, role)

class form(QMainWindow):
    def __init__(self, scaleRate):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self, scaleRate)
        self.scaleRate = scaleRate
        self.db = QtSql.QSqlDatabase.addDatabase('QSQLITE')
        self.db.setDatabaseName('fieldlist.db')
        self.model = QSqlCenteredTableModel()
        self.model.setTable('field')
        self.query = QtSql.QSqlQuery(self.db)
        self.query.exec('select * from location')
        self.location = []
        while self.query.next():
            self.location.append(self.query.value("Name"))
        print(self.location)
        self.ui.lineEdit.addItems(self.location)
        self.ui.lineEdit_2.addItems(self.location)
        self.model.setEditStrategy(QtSql.QSqlTableModel.OnFieldChange)
        self.model.setFilter(('Date like "' + self.ui.monthControl.text() + '-%%"'))

        self.model.setHeaderData(0, QtCore.Qt.Horizontal, "序号")
        self.model.setHeaderData(1, QtCore.Qt.Horizontal,"出发地")
        self.model.setHeaderData(2, QtCore.Qt.Horizontal, "目的地")
        self.model.setHeaderData(3, QtCore.Qt.Horizontal, "日期")
        self.model.setHeaderData(4, QtCore.Qt.Horizontal,"品名")
        self.model.setHeaderData(5, QtCore.Qt.Horizontal, "数量(件)")
        self.model.setHeaderData(6, QtCore.Qt.Horizontal, "重量(KG)")
        self.model.setHeaderData(7, QtCore.Qt.Horizontal, "单价")
        self.model.setHeaderData(8, QtCore.Qt.Horizontal, "总重(吨)")
        self.model.setHeaderData(9, QtCore.Qt.Horizontal, "金额")
        self.model.setSort(3, QtCore.Qt.AscendingOrder)
        self.model.setSort(2, QtCore.Qt.AscendingOrder)
        self.model.setSort(1, QtCore.Qt.AscendingOrder)

        self.ui.tableWidget.setModel(self.model)
        self.ui.tableWidget.setColumnHidden(0, True)
        self.ui.tableWidget.setColumnWidth(1, int(60 * scaleRate))
        self.ui.tableWidget.setColumnWidth(2, int(60 * scaleRate))
        self.ui.tableWidget.setColumnWidth(3, int(100 * scaleRate))
        self.ui.tableWidget.setColumnWidth(4, int(80 * scaleRate))
        self.ui.tableWidget.setColumnWidth(5, int(60 * scaleRate))
        self.ui.tableWidget.setColumnWidth(6, int(65 * scaleRate))
        self.ui.tableWidget.setColumnWidth(7, int(60 * scaleRate))
        self.ui.tableWidget.setColumnWidth(8, int(65 * scaleRate))
        self.ui.tableWidget.setColumnWidth(9, int(70 * scaleRate))
        self.model.select()

        self.ui.pushButton.clicked.connect(self.addToDb)
        self.show()
        self.ui.pushButton_2.clicked.connect(self.editlocation)
        self.ui.pushButton_3.clicked.connect(self.delrow)
        self.ui.pushButton_4.clicked.connect(self.export)
        self.ui.lineEdit_count.textChanged.connect(self.updateprice)
        self.ui.lineEdit_weight.textChanged.connect(self.updateprice)
        self.ui.lineEdit_price.textChanged.connect(self.updateprice)
        self.ui.monthControl.dateChanged.connect(self.updatemonth)
        self.i = self.model.rowCount()
        self.ui.lcdNumber.display(self.i)
        print(self.ui.tableWidget.currentIndex().row())

    def updateprice(self):
        try:
            self.ui.lineEdit_total.setText(str(float(self.ui.lineEdit_count.text()) * float(self.ui.lineEdit_weight.text()) / 1000))
        except Exception:
            self.ui.lineEdit_total.setText("0")
        try:
            self.ui.lineEdit_earn.setText(str(float(self.ui.lineEdit_total.text()) * float(self.ui.lineEdit_price.text())))
        except Exception:
            self.ui.lineEdit_earn.setText("0")

    def updatemonth(self):
        self.model.setFilter(('Date like "' + self.ui.monthControl.text() + '-%%"'))
        self.model.setHeaderData(1, QtCore.Qt.Horizontal, "出发地")
        self.model.setHeaderData(2, QtCore.Qt.Horizontal, "目的地")
        self.model.setHeaderData(3, QtCore.Qt.Horizontal, "日期")
        self.model.setHeaderData(4, QtCore.Qt.Horizontal, "品名")
        self.model.setHeaderData(5, QtCore.Qt.Horizontal, "数量(件)")
        self.model.setHeaderData(6, QtCore.Qt.Horizontal, "重量(KG)")
        self.model.setHeaderData(7, QtCore.Qt.Horizontal, "单价")
        self.model.setHeaderData(8, QtCore.Qt.Horizontal, "总重(吨)")
        self.model.setHeaderData(9, QtCore.Qt.Horizontal, "金额")
        self.ui.tableWidget.setModel(self.model)
        self.model.select()
        self.show()
        self.i = self.model.rowCount()
        self.ui.lcdNumber.display(self.i)

    def editlocation(self):
        text, okPressed = QtWidgets.QInputDialog.getText(self, "编辑运货地点", "地点:", QtWidgets.QLineEdit.Normal, ",".join(self.location))
        if okPressed and text != '':
            self.location = text.split(',')
            self.ui.lineEdit.clear()
            self.ui.lineEdit_2.clear()
            self.ui.lineEdit.addItems(self.location)
            self.ui.lineEdit_2.addItems(self.location)
            self.query = QtSql.QSqlQuery(self.db)
            self.query.exec('delete from location')
            temp = []
            for item in self.location:
                temp.append("('" + item + "')")
            self.query.exec('insert into location values ' + ','.join(temp))

    def addToDb(self):
        print(self.i)
        self.model.insertRows(self.i,1)
        self.model.setData(self.model.index(self.i,1),self.ui.lineEdit.currentText())
        self.model.setData(self.model.index(self.i,2), self.ui.lineEdit_2.currentText())
        self.model.setData(self.model.index(self.i,4), self.ui.lineEdit_3.text())
        self.model.setData(self.model.index(self.i,3), self.ui.monthControl.text() + "-" + self.ui.dateEdit.text())
        self.model.setData(self.model.index(self.i,5), self.ui.lineEdit_count.text())
        self.model.setData(self.model.index(self.i,6), self.ui.lineEdit_weight.text())
        self.model.setData(self.model.index(self.i,7), self.ui.lineEdit_total.text())
        self.model.setData(self.model.index(self.i,8), self.ui.lineEdit_price.text())
        self.model.setData(self.model.index(self.i,9), self.ui.lineEdit_earn.text())
        self.model.submitAll()
        self.ui.tableWidget.setModel(self.model)
        self.model.select()
        self.show()
        self.i = self.model.rowCount()
        self.ui.lcdNumber.display(self.i)
        self.updatemonth()

    def export(self):
        curPath = QDir.currentPath()  # 获取系统当前目录
        dlgTitle = "保存 " + self.ui.monthControl.text() + " 运费文件"
        filt = "Excel (*.xls);"
        filename, filtUsed = QFileDialog.getSaveFileName(self, dlgTitle, curPath, filt)
        if os.path.isdir(filename) or os.path.split(filename)[1] == "":
            msg_box = QMessageBox(QMessageBox.Warning, '警告', '未指定文件名，导出取消！')
            msg_box.exec_()
            return
        self.model.setSort(3, QtCore.Qt.AscendingOrder)
        self.model.setSort(2, QtCore.Qt.AscendingOrder)
        self.model.setSort(1, QtCore.Qt.AscendingOrder)
        self.model.select()
        book = xlwt.Workbook()
        sheet = book.add_sheet(self.ui.monthControl.text() + ' 报表')
        title = ["日期", "出发地", "目的地", "品名", "数量(件)", "重量(KG)", "总重(吨)", "单价", "金额"]
        title_index = [3, 1, 2, 4, 5, 6, 8, 7, 9]
        for i in range(0, self.model.rowCount() + 1):
            if i == 0:
                for index, item in enumerate(title_index):
                    try:
                        sheet.write(i, index, title[index])
                    except:
                        continue
            else:
                for index, item in enumerate(title_index):
                    try:
                        print(self.model.record(i-1).value(item))
                        sheet.write(i, index, self.model.record(i-1).value(item))
                    except:
                        continue
        # Add Tax and Total
        i = self.model.rowCount() + 2
        borders = xlwt.Borders()  # 创建边框对象Create Borders
        borders.bottom = xlwt.Borders.MEDIUM
        borders.bottom_colour = 0x0
        style = xlwt.XFStyle()  # Create Style   #创建样式对象
        style.borders = borders
        for index, item in enumerate(title):
            if index == 0:
                sheet.write(i, index, '总计', style)
            elif index == len(title) - 1:
                sheet.write(i, index, xlwt.Formula('SUM(I2:I' + str(self.model.rowCount() + 1) + ')'), style)
            else:
                sheet.write(i, index, '', style)
        i = i + 1
        sheet.write(i, len(title) - 2, '税率')
        sheet.write(i, len(title) - 1, '0.09')
        i = i + 1
        sheet.write(i, int(len(title)/2), '发票加税9%')
        sheet.write(i, int(len(title) / 2)+1, xlwt.Formula('I' + str(i-1) + '*(I' + str(i) + '+1)'))
        book.save(filename)
        QMessageBox.information(self, '提示', '导出成功！',
                                    QMessageBox.Yes)
        self.show()


    def delrow(self):
        if self.ui.tableWidget.currentIndex().row() > -1:
            self.model.removeRow(self.ui.tableWidget.currentIndex().row())
            self.i -= 1
            self.model.select()
            self.ui.lcdNumber.display(self.i)
        else:
            QMessageBox.question(self, '信息', "请选择需要删除的一行", QMessageBox.Ok)
            self.show()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    scaleRate = app.screens()[0].logicalDotsPerInch() / 96
    frm = form(scaleRate)
    sys.exit(app.exec_())
