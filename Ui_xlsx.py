# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'p_test.ui'
#
# Created by: PyQt5 UI code generator 5.12.2
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui, QtWidgets
import sys

class Ui_Main_ui(object):
    def __init__(self):
        self.cb_list = ["BEMF","High Speed","Torque vs Current",]

    def setupUi(self, Main_ui):
        self.Main_ui = Main_ui
        Main_ui.setObjectName("Main_ui")
        Main_ui.resize(400, 204)
        self.formLayout = QtWidgets.QFormLayout(Main_ui)
        self.formLayout.setObjectName("formLayout")
        self.gridLayout = QtWidgets.QGridLayout()
        self.gridLayout.setObjectName("gridLayout")
        self.Tdms_path = QtWidgets.QLineEdit(Main_ui)
        self.Tdms_path.setObjectName("Tdms_path")
        self.gridLayout.addWidget(self.Tdms_path, 2, 0, 1, 1)
        self.toolButton = QtWidgets.QToolButton(Main_ui)
        self.toolButton.setObjectName("toolButton")
        self.gridLayout.addWidget(self.toolButton, 2, 1, 1, 1)
        self.Tdms_label = QtWidgets.QLabel(Main_ui)
        self.Tdms_label.setObjectName("Tdms_label")
        self.gridLayout.addWidget(self.Tdms_label, 1, 0, 1, 1)
        self.formLayout.setLayout(0, QtWidgets.QFormLayout.SpanningRole, self.gridLayout)
        self.gridLayout_2 = QtWidgets.QGridLayout()
        self.gridLayout_2.setObjectName("gridLayout_2")
        self.toolButton_2 = QtWidgets.QToolButton(Main_ui)
        self.toolButton_2.setObjectName("toolButton_2")
        self.gridLayout_2.addWidget(self.toolButton_2, 2, 1, 1, 1)
        self.Excel_path = QtWidgets.QLineEdit(Main_ui)
        self.Excel_path.setObjectName("Excel_path")
        self.gridLayout_2.addWidget(self.Excel_path, 2, 0, 1, 1)
        self.Excel_label = QtWidgets.QLabel(Main_ui)
        self.Excel_label.setObjectName("Excel_label")
        self.gridLayout_2.addWidget(self.Excel_label, 1, 0, 1, 1)
        self.formLayout.setLayout(1, QtWidgets.QFormLayout.SpanningRole, self.gridLayout_2)
        self.cb = QtWidgets.QComboBox(Main_ui)
        self.cb.setObjectName("cb")
        self.cb.addItems(self.cb_list)
        self.formLayout.setWidget(2, QtWidgets.QFormLayout.SpanningRole, self.cb)
        self.Process_btn = QtWidgets.QPushButton(Main_ui)
        self.Process_btn.setObjectName("Process_btn")
        self.formLayout.setWidget(3, QtWidgets.QFormLayout.FieldRole, self.Process_btn)

        self.retranslateUi(Main_ui)
        QtCore.QMetaObject.connectSlotsByName(Main_ui)


        # 设置触发
        self.toolButton.clicked.connect(self.openTDMSfile)
        self.toolButton_2.clicked.connect(self.openExcelfile)
        self.Process_btn.clicked.connect(self.Process)

    def retranslateUi(self, Main_ui):
        _translate = QtCore.QCoreApplication.translate
        Main_ui.setWindowTitle(_translate("Main_ui", "Form"))
        self.toolButton.setText(_translate("Main_ui", "..."))
        self.Tdms_label.setText(_translate("Main_ui", "TDMS source"))
        self.toolButton_2.setText(_translate("Main_ui", "..."))
        self.Excel_label.setText(_translate("Main_ui", "Excel source"))
        self.Process_btn.setText(_translate("Main_ui", "Process"))

    def openTDMSfile(self):
        fname= QtWidgets.QFileDialog.getOpenFileName()
        self.Tdms_path.setText(fname[0])

    def openExcelfile(self):
        fname= QtWidgets.QFileDialog.getOpenFileName(self.Main_ui,'Open File','','Excel(*.xlsx)')
        self.Excel_path.setText(fname[0])

    def Process(self):
        if self.Tdms_path.text() and self.Excel_path:
            pass
        else:
            QtWidgets.QMessageBox.about(self.Main_ui,"Tips","请先选择路径！")

if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    main_ui = QtWidgets.QWidget()
    ui = Ui_Main_ui()
    ui.setupUi(main_ui)
    main_ui.show()
    sys.exit(app.exec_())