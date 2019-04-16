#! /usr/bin/python3
# -*- coding:UTF-8 -*-
import os
from openpyxl import load_workbook
from nptdms import TdmsFile
from openpyxl.drawing.image import Image
# defined by myself
from draw import Draw
from Jobs import Jobs
from datapath import DataPath
from excel import Excel
from tdms import TDMS


# Start Process

Excel_filename = DataPath().file_path_set()
Job_list = ["退出","BEMF","Continue Torque","High Speed","Winding Heating"]
while True:
    try:
        Job_num = Jobs(Job_list).Select()
        print("正在尝试操作: %s ..."%Job_list[Job_num])
        if Job_list[Job_num] == "BEMF":
            filename, group = DataPath().data_get()
            sheetname = "2.Motor BEMF"
            listforname = ["MB_Command.Speed","MA_Command.Torque","U-RMS.Voltage",\
            "V-RMS.Voltage","W-RMS.Voltage","U-PP.RMS.Voltage",\
            "V-PP.RMS.Voltage","W-PP.RMS.Voltage"]
            listforposition = ["E","F","G","H"]
            Dict_temp = TDMS(filename,group,listforname).Read_Tdms()
            Excel(Dict_temp,Excel_filename,sheetname,listforposition).WriteBEMF()
        elif Job_list[Job_num] == "Continue Torque":
            filename, group = DataPath().data_get()
            sheetname = "4. Cont. Torque Curve"
            listforname = ["MB_Command.Speed","MA_Command.Torque","Sensor-Torque",\
            "SUM/AVG-RMS.Voltage","SUM/AVG-F.Voltage","SUM/AVG-RMS.Current","SUM/AVG-F.Current",\
            "SUM/AVG-Kwatts","SUM/AVG-F.Kwatts","SUM/AVG-PF","SUM/AVG-F.PF","DC Current"] 
            listforposition = ["G","I","K"]
            Dict_temp = TDMS(filename,group,listforname).Read_Tdms()
            Excel(Dict_temp,Excel_filename,sheetname,listforposition).WriteConti()
        elif Job_list[Job_num] == "High Speed":
            filename, group = DataPath().data_get()
            sheetname = "5. High Speed"
            listforname = ["MA-RTD 1","MA-RTD 2"]
            listforposition = ["J","K"]
            Dict_temp = TDMS(filename,group,listforname).Read_Tdms()
            Excel(Dict_temp,Excel_filename,sheetname,listforposition).WriteHighSpeed()
        elif Job_list[Job_num] == "Winding Heating":
            filename, group = DataPath().data_get()
            sheetname = "8. Winding Heating"
            listforname = ["MA_Command.Torque","Sensor-Torque",\
            "U-PP.RMS.Voltage","V-PP.RMS.Voltage","W-PP.RMS.Voltage",\
            "SUM/AVG-RMS.Current"]#no need for RTD
            listforposition = ["L","M"]
            Dict_temp = TDMS(filename, group, listforname).Read_Tdms()
            picname = filename[:-5]+".png"
            RTD,i = Draw(filename,picname).draw8min_returnRTD()
            Excel(Dict_temp, Excel_filename, sheetname, listforposition).WriteWinding(RTD,picname,i)
        else:
            print("关闭中 ...")
            break
    except PermissionError:
        print("数据处理失败，请先关闭需要操作的文件\n..............................................\n")