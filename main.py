# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file '01.ui'
#
# Created by: PyQt5 UI code generator 5.14.0
#
# WARNING! All changes made in this file will be lost!
from check_db import CheckModel
from openpyxl import load_workbook
import excel_reader
import check_db
from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        # main window
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(1000, 600)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        # select file button
        self.select_file_btn = QtWidgets.QPushButton(self.centralwidget)
        self.select_file_btn.setGeometry(QtCore.QRect(450, 50, 91, 31))
        self.select_file_btn.setStyleSheet("font-size: 15px;")
        self.select_file_btn.setObjectName("pushButton")
        MainWindow.setCentralWidget(self.centralwidget)
        # list of new check
        # self.new_check_list_show = QtWidgets.QListWidget(self.centralwidget)
        # self.new_check_list_show.setHidden(True)
        # self.new_check_list_show.setGeometry(QtCore.QRect(0, 90, 1000, 571))
        # self.new_check_list_show.setSelectionMode(QtWidgets.QAbstractItemView.MultiSelection)
        # self.new_check_list_show.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        # self.new_check_list_show.setLayoutDirection(QtCore.Qt.RightToLeft)
        # self.new_check_list_show.setObjectName("list_show")
        # self.new_check_list_show.setStyleSheet("background: red;\n"
        #                              "color: white;\n"
        #                              "font-weight: bold;\n"
        #                              "font-size: 20px;")
        ##


        self.new_check_table_show = QtWidgets.QTableWidget(self.centralwidget)
        self.new_check_table_show.setLayoutDirection(QtCore.Qt.RightToLeft)
        self.new_check_table_show.setGeometry(QtCore.QRect(0, 90, 1000, 571))
        self.new_check_table_show.setEditTriggers(QtWidgets.QAbstractItemView.DoubleClicked)
        self.new_check_table_show.setSelectionMode(QtWidgets.QAbstractItemView.MultiSelection)
        self.new_check_table_show.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.new_check_table_show.horizontalHeader().setSortIndicatorShown(True)
        self.new_check_table_show.setObjectName("tableWidget")
        # self.new_check_table_show.setColumnCount(1)
        # self.new_check_table_show.setRowCount(3)
        # item = QtWidgets.QTableWidgetItem()
        # self.new_check_table_show.setVerticalHeaderItem(0, item)
        # item = QtWidgets.QTableWidgetItem()
        # self.new_check_table_show.setVerticalHeaderItem(1, item)
        # item = QtWidgets.QTableWidgetItem()
        # self.new_check_table_show.setVerticalHeaderItem(2, item)

        self.selected_rows = []

        font = QtGui.QFont()
        # font.setFamily("Times New Roman")
        font.setBold(True)
        font.setWeight(75)

        self.new_check_table_show.setColumnCount(8)
        # self.new_check_table_show.setRowCount(3)
        self.new_check_table_show.setHidden(True)
        for item_index in range(8):
            item = QtWidgets.QTableWidgetItem()
            item.setTextAlignment(QtCore.Qt.AlignCenter)
            item.setFont(font)
            self.new_check_table_show.setHorizontalHeaderItem(item_index, item)



        # submit record button
        self.submit_record_btn = QtWidgets.QPushButton(self.centralwidget)
        self.submit_record_btn.setGeometry(QtCore.QRect(300, 50, 91, 31))
        self.submit_record_btn.setStyleSheet("font-size: 15px;")
        self.submit_record_btn.setObjectName("submit_record_btn")
        self.submit_record_btn.setHidden(True)
        MainWindow.setCentralWidget(self.centralwidget)



        self.retranslateUi(MainWindow)
        # signals and functions
        self.select_file_btn.clicked.connect(self.file_select)
        # self.new_check_list_show.itemClicked.connect(self.change_item_background_style)
        # self.new_check_list_show.itemClicked.connect(self.unchange_item_background_style)
        self.submit_record_btn.clicked.connect(self.submit_selected_record_in_db)


        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.select_file_btn.setText(_translate("MainWindow", "انتخاب فایل"))
        self.submit_record_btn.setText(_translate("MainWindow", "غیربنفش"))
        item = self.new_check_table_show.horizontalHeaderItem(0)
        item.setText(_translate("MainWindow", "شماره"))
        item = self.new_check_table_show.horizontalHeaderItem(1)
        item.setText(_translate("MainWindow", "مبلغ"))
        item = self.new_check_table_show.horizontalHeaderItem(2)
        item.setText(_translate("MainWindow", "اسناد دریافتی"))
        item = self.new_check_table_show.horizontalHeaderItem(3)
        item.setText(_translate("MainWindow", "وضعیت"))
        item = self.new_check_table_show.horizontalHeaderItem(4)
        item.setText(_translate("MainWindow", "تاریخ چک"))
        item = self.new_check_table_show.horizontalHeaderItem(5)
        item.setText(_translate("MainWindow", "تاریخ دریافت چک"))
        item = self.new_check_table_show.horizontalHeaderItem(6)
        item.setText(_translate("MainWindow", "نام بانک"))
        item = self.new_check_table_show.horizontalHeaderItem(7)
        item.setText(_translate("MainWindow", "تاریخ ثبت(اختیاری)"))

    def file_select(self):
        self.file_select = QtWidgets.QFileDialog.getOpenFileName(MainWindow, 'Please give the excel file Mr. Kooti')
        sheet = excel_reader.open_excel(self.file_select[0])
        # remove_duplicate_records_from_excel()
        self.fill_table_items(sheet)

    def fill_list_items(self, list_item):
        self.new_check_list_show.clear()
        self.new_check_list_show.setHidden(False)
        _translate = QtCore.QCoreApplication.translate
        # self.new_check_list_show = list_item
        for case in list_item:
            ## add item
            item = QtWidgets.QListWidgetItem()
            self.new_check_list_show.addItem(item)
            item.setToolTip(str(case['number']))

            ## set text for each item
            # item = self.new_check_list_show.item(self.new_check_list_show.index(case['number']))
            item.setText(_translate("MainWindow", str(
                f"{case['number']} - {case['amount']} - {case['recieved_docs']} - {case['condition']} - {case['date_check']} - {case['date_recieved_ckeck']} - {case['bank_name']}"
            )))
        self.submit_record_btn.setHidden(False)

    def fill_table_items(self, list_item):
        # self.new_check_table_show.clear()
        self.new_check_table_show.setHidden(False)
        _translate = QtCore.QCoreApplication.translate
        # self.new_check_list_show = list_item
        self.new_check_table_show.setRowCount(len(list_item) + 1)
        for item_index in range(len(list_item)):  # 0 1 2 3 4 5 6 7
            ## add record to table
            font = QtGui.QFont()
            font.setPointSize(9)
            item = QtWidgets.QTableWidgetItem()
            item.setTextAlignment(QtCore.Qt.AlignCenter)
            item.setText(str(item_index+1))
            self.new_check_table_show.setVerticalHeaderItem(item_index, item)
            for record_index in range(len(list_item[item_index])):
                item = QtWidgets.QTableWidgetItem()
                item.setText(str(list_item[item_index][record_index]))
                self.new_check_table_show.setItem(item_index, record_index, item)  # row, column, item


        self.submit_record_btn.setHidden(False)

    def change_item_background_style(self):
        item = QtWidgets.QListWidgetItem()
        brush = QtGui.QBrush(QtGui.QColor(170, 180, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        for item in self.new_check_list_show.selectedItems():
            item.setBackground(brush)

    def submit_selected_record_in_db(self):
        selected_rows = self.new_check_table_show.selectionModel().selectedRows()
        print(selected_rows[0].__dir__())
        for selected_record in selected_rows:

        #     print('--->', selected_record)
            single_record = []
            # row_number = selected_record.row()
            for cell in range(8):
                cell_text = self.new_check_table_show.item(selected_record.row(), cell)
                if cell_text:
                    single_record.append(cell_text.text())
                else:
                    single_record.append(None)

            excel_reader.submit_record_in_db(single_record)

        # self.new_check_table_show.setHidden(True)
        self.new_check_table_show.clearContents()
        self.submit_record_btn.setHidden(True)


if __name__ == "__main__":
    import sys
    check_db.initial_db()
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())

