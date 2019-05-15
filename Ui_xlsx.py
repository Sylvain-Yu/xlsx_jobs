# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'p_test.ui'
#
# Created by: PyQt5 UI code generator 5.12.2
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtGui import QIcon
from excel import Excel
from nptdms import TdmsFile
from tdms import TDMS
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from draw import Draw
import sys
import os

class Ui_Main_ui(object):
    def __init__(self):
        self.cb_list = ["BEMF","High Speed","Short Circuit","Continue Torque","Winding Heating"]

    def setupUi(self, Main_ui):
        self.Main_ui = Main_ui
        Main_ui.setObjectName("Main_ui")
        Main_ui.resize(400, 204)
        Main_ui.setWindowIcon(QIcon("icon_co.png"))
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
        self.cb_2 = QtWidgets.QComboBox(Main_ui)
        self.cb_2.setObjectName("cb_2")
        self.cb_2.addItems(["Instantly Data","Meta Data"])
        self.formLayout.setWidget(3, QtWidgets.QFormLayout.SpanningRole, self.cb_2)
        self.Process_btn = QtWidgets.QPushButton(Main_ui)
        self.Process_btn.setObjectName("Process_btn")
        self.formLayout.setWidget(4, QtWidgets.QFormLayout.FieldRole, self.Process_btn)

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
        fname= QtWidgets.QFileDialog.getOpenFileName(QtWidgets.QWidget(),'Open File','','TDMS(*.tdms)')
        self.Tdms_path.setText(fname[0])

    def openExcelfile(self):
        fname= QtWidgets.QFileDialog.getOpenFileName(QtWidgets.QWidget(),'Open File','','Excel(*.xlsx *XLSX)')
        self.Excel_path.setText(fname[0])

    def Process(self):
        if os.path.isfile(self.Tdms_path.text()) and os.path.isfile(self.Excel_path.text()):
            filename = self.Tdms_path.text()
            Excel_filename = self.Excel_path.text()
            group = self.cb_2.currentText()
            if self.cb.currentText() == "BEMF":
                sheetname = "Motor BEMF"
                listforname = ["MB_Command.Speed","MA_Command.Torque","U-RMS.Voltage",\
                "V-RMS.Voltage","W-RMS.Voltage","U-PP.RMS.Voltage",\
                "V-PP.RMS.Voltage","W-PP.RMS.Voltage"]
                listforposition = ["E","F","G","H"]
                Dict_temp = TDMS(filename,group,listforname).Read_Tdms()
                tips = Excel(Dict_temp,Excel_filename,sheetname,listforposition).WriteBEMF()

            elif self.cb.currentText() =="High Speed":
                sheetname = "High Speed"
                listforname = ["MB_Command.Speed"]
                listforposition = ["J","K"]
                Dict_temp = TDMS(filename,group,listforname).Read_Tdms()
                picname = filename[:-5]+".png"
                RTD,i = Draw(filename,picname).drawXmin_returnRTD(5)
                tips = Excel(Dict_temp,Excel_filename,sheetname,listforposition).WriteHighSpeed(RTD,picname,i)

            elif self.cb.currentText() =="Torque vs Current":
                pass

            elif self.cb.currentText() == "Short Circuit":
                sheetname = "Short circuit"
                listforname = ["MB_Command.Speed","Sensor-Torque",\
                "U-RMS.Current","U-F.Current","V-RMS.Current","V-F.Current",\
                "W-RMS.Current","W-F.Current","MA-Motor TEMP"]
                listforposition = ["D","E","F","G"]
                Dict_temp = TDMS(filename , group , listforname).Read_Tdms()
                Excel(Dict_temp,Excel_filename,sheetname,listforposition).WriteSc()

            elif self.cb.currentText() == "Continue Torque":
                sheetname = "Cont. Torque Curve"
                listforname = ["MB_Command.Speed","MA_Command.Torque","Sensor-Torque",\
                "SUM/AVG-RMS.Voltage","SUM/AVG-F.Voltage","SUM/AVG-RMS.Current","SUM/AVG-F.Current",\
                "SUM/AVG-Kwatts","SUM/AVG-F.Kwatts","SUM/AVG-PF","SUM/AVG-F.PF","DC Current",\
                "MA-RTD 1","MA-RTD 2"] 
                listforposition = ["G","I","K"]
                Dict_temp = TDMS(filename,group,listforname).Read_Tdms()
                Excel(Dict_temp,Excel_filename,sheetname,listforposition).WriteConti()
            elif self.cb.currentText() =="Winding Heating":
                sheetname = "Winding Heating"
                listforname = ["MA_Command.Torque","Sensor-Torque",\
                "U-PP.RMS.Voltage","V-PP.RMS.Voltage","W-PP.RMS.Voltage",\
                "SUM/AVG-RMS.Current"]#no need for RTD
                listforposition = ["L","M"]
                Dict_temp = TDMS(filename, group, listforname).Read_Tdms()
                picname = filename[:-5]+".png"
                RTD,i = Draw(filename,picname).drawXmin_returnRTD(8) #8 imply 8 minutes
                Excel(Dict_temp, Excel_filename, sheetname, listforposition).WriteWinding(RTD,picname,i)
            else:
                tips = "No setting for current Combox Text"
            QtWidgets.QMessageBox.about(self.Main_ui,"Tips",tips)
        else:
            QtWidgets.QMessageBox.about(self.Main_ui,"Tips","Please choose the path first!")

if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    main_ui = QtWidgets.QWidget()
    ui = Ui_Main_ui()
    ui.setupUi(main_ui)
    main_ui.show()
    sys.exit(app.exec_())