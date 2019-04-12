#! /usr/bin/python3
# -*- coding:UTF-8 -*-
import os
from openpyxl import load_workbook
from nptdms import TdmsFile

class TDMS():# read TDMS files and dict it.
    def __init__(self,filename,group,listforname):
        self.filename = filename
        self.group = group
        self.listforname = listforname

    def Read_Tdms(self):
        Data_List = []
        tdmsfile = TdmsFile(self.filename)
        for name in self.listforname:
            Data_List.append(tdmsfile.object(self.group,name).data)
        Dict_temp = dict(list(zip(self.listforname,Data_List)))
        return Dict_temp

class Excel():  # read and write the xlsx and process.
    def __init__(self,Dict_temp,filename,sheetname,listforposition):
        self.Dict_temp = Dict_temp
        self.filename = filename
        self.sheetname = sheetname
        self.listforposition = listforposition
        self.wb = load_workbook(self.filename)
        self.current_sheet = self.wb[self.sheetname]

    def WriteBEMF(self):
        for v in Dict_temp["MB_Command.Speed"]:
            speed_range = [1000,2000,3000,4000]
            try:
                i = speed_range.index(v)
            except ValueError:
                status = "BEMF 数据处理失败 !\n"
            else:
                position = self.listforposition[i]+"14"
                self.current_sheet[position] = self.Dict_temp["U-RMS.Voltage"][i]
                position = self.listforposition[i]+"15"
                self.current_sheet[position] = self.Dict_temp["V-RMS.Voltage"][i]
                position = self.listforposition[i]+"16"
                self.current_sheet[position] = self.Dict_temp["W-RMS.Voltage"][i]
                position = self.listforposition[i]+"20"
                self.current_sheet[position] = self.Dict_temp["U-PP.RMS.Voltage"][i]
                position = self.listforposition[i]+"21"
                self.current_sheet[position] = self.Dict_temp["V-PP.RMS.Voltage"][i]
                position = self.listforposition[i]+"22"
                self.current_sheet[position] = self.Dict_temp["W-PP.RMS.Voltage"][i]
                status = "BEMF 数据处理完成 !\n"
        self.wb.save(self.filename)
        print(status)

    def Read_One_Value(self,single_pos):
        value = self.current_sheet[single_pos]
        return value

#仅写入耐久的数据
    def WriteConti_value(self):
        i = 0
        print(self.Dict_temp["MB_Command.Speed"])
        for speed in self.Dict_temp["MB_Command.Speed"]:
            if speed == self.Conti_Speed:
                position = self.listforposition[self.j]+"36"
                self.current_sheet[position] = self.Dict_temp["SUM/AVG-RMS.Voltage"][i]
                position = self.listforposition[self.j]+"37"
                self.current_sheet[position] = self.Dict_temp["SUM/AVG-F.Voltage"][i]
                position = self.listforposition[self.j]+"38"
                self.current_sheet[position] = self.Dict_temp["SUM/AVG-RMS.Current"][i]
                position = self.listforposition[self.j]+"39"
                self.current_sheet[position] = self.Dict_temp["SUM/AVG-F.Current"][i]
                position = self.listforposition[self.j]+"40"
                self.current_sheet[position] = self.Dict_temp["SUM/AVG-Kwatts"][i]
                position = self.listforposition[self.j]+"41"
                self.current_sheet[position] = self.Dict_temp["SUM/AVG-F.Kwatts"][i]
                position = self.listforposition[self.j]+"42"
                self.current_sheet[position] = self.Dict_temp["SUM/AVG-PF"][i]
                position = self.listforposition[self.j]+"43"
                self.current_sheet[position] = self.Dict_temp["SUM/AVG-F.PF"][i]
                position = self.listforposition[self.j]+"26"
                self.current_sheet[position] = self.Dict_temp["MA_Command.Torque"][i]
                position = self.listforposition[self.j]+"27"
                self.current_sheet[position] = self.Dict_temp["Sensor-Torque"][i]
                position = self.listforposition[self.j]+"33"
                self.current_sheet[position] = self.Dict_temp["DC Current"][i]
                print("已找到转速为%s,第%s个数列"%(speed,i+1))
                i += 1
        if i == 0:
            print("没有找到转速为%s的数据\n"%speed)
        else:
            self.wb.save(self.filename)
            print("完成数据处理\n")

# 连续扭矩试验处理数据，属于源为Instantly
    def WriteConti(self):
        listforNum = range(len(self.listforposition))
        SpeedPositionList = [ x + "6" for x in self.listforposition]
        for self.j,Speed_pos in zip(listforNum,SpeedPositionList):
            self.Conti_Speed = self.Read_One_Value(Speed_pos)
            self.WriteConti_value()

# 为高转速试验处理数据
    def WriteHighSpeed(self):
        point = 6
        i = 0
        zipped_list = list(zip(self.Dict_temp["MA-RTD 1"],self.Dict_temp["MA-RTD 2"]))
        for RTD1 ,RTD2 in zipped_list:
            if RTD1 > 55 or RTD2 > 55:
                break
            else:
                i += 1
        k = i + 300
        self.current_sheet["F20"] = zipped_list[k][0]
        self.current_sheet["G20"] = zipped_list[k][1]
        self.current_sheet["F14"] = zipped_list[i][0]
        self.current_sheet["G14"] = zipped_list[i][0]
        while i < len(self.Dict_temp["MA-RTD 1"]):
            pos1 = self.listforposition[0] + str(point)
            pos2 = self.listforposition[1] + str(point)
            self.current_sheet[pos1] = zipped_list[i][0]
            self.current_sheet[pos2] = zipped_list[i][1]
            point += 1
            i += 1
        print("已处理完高转速试验数据\n")
        self.wb.save(self.filename)

    def WriteWinding(self): #unfinished
        point = 10
        i = 0
        zipped_list = list(zip(self.Dict_temp["MA-RTD 1"],self.Dict_temp["MA-RTD 2"]))
        while i < 8*60:
            pos1 = self.listforposition[0] + str(point)
            pos2 = self.listforposition[1] + str(point)
            self.current_sheet[pos1] = zipped_list[i][0]
            self.current_sheet[pos2] = zipped_list[i][1]
            point += 1 
            i += 1
        self.current_sheet["E20"] = zipped_list[0][0] #winding temp at start point:RTD 1
        self.current_sheet["G20"] = zipped_list[0][1] #winding temp at start point:RTD 2
        self.current_sheet["E21"] = zipped_list[480][0] #winding temp at 8 min:RTD 1
        self.current_sheet["G21"] = zipped_list[480][1] #winding temp at 8 min:RTD 2
 #       self.current_sheet["E16"] =  # Torque Command at start point
 #       self.current_sheet["G16"] =  # Torque Command at end ponts
        print("已处理完成温升试验数据\n")
        self.wb.save(self.filename)

class DataPath():# input value set
    def file_path_set(self):
        file_path = os.listdir(".")
        Excel_path_list = []
        for Excel in file_path:
            if ".xlsx" in Excel.lower():
                Excel_path_list.append(Excel)
        num = range(0,len(Excel_path_list))
        file_dict = dict(zip(num,Excel_path_list))
        for Key, Value in file_dict.items():
            print("{0}:{1}".format(Key,Value))
        while True:
            filename_num = input("请选择需要操作的Excel\n")
            filename_num = int(filename.strip())
            if filename_num >= 0 and filename_num < len(Excel_path_list):
                filename = file_dict[filename_num]
                break
            else:
                print("请输入 0~%d"%(len(Excel_path_list)-1))
        return filename

    def data_get(self):
        dir_list = os.listdir(".")
        listforTDMS = []
        for f in dir_list:
            if ".tdms" in f.lower():
                listforTDMS.append(f)
        len_listforTDMS = len(listforTDMS)
        num = range(0,len_listforTDMS)
        file_name_dict =dict(zip(num,listforTDMS))
        print("以下为本文件夹内可操作文件：")
        for key, value in file_name_dict.items():
            print("{0}:{1}".format(key,value))
        while True:
            filename_num = input("请选择tdms格式的源文件件\n")
            filename_num = int(filename_num)
            if filename_num >= 0 and filename_num < len_listforTDMS:
                filename = file_name_dict[filename_num]
                break
            else:
                print("请输入 0~ %d"%(len_listforTDMS-1))
        choice = input("请选择需要进行数据处理的组[0/1]:\n0.Instantly Data(默认) 1.Meta Data\n")
        if int(choice.strip()) == 0:
            group = "Instantly Data"
        else:
            group ="Meta Data"
        return (filename,group)

class Jobs():
    def __init__(self,Job_list):
        self.Job_list = Job_list
        self.Jobs_numlist = range(0,len(self.Job_list))
        self.Jobs_dict = dict(zip(self.Jobs_numlist,self.Job_list))

    def Select(self):
        Job_num = 0
        print("请输入要操作的内容:[默认 0]")
        for key,value in self.Jobs_dict.items():
           print("{0}:{1}".format(key,value))
        Job_num = int(input())
        return Job_num


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
        elif Jobs[Job_num] == "Winding Heating":
            filename, group = DataPath().data_get()
            sheetname = "8. Winding Heating"
            listforname = ["MA_Command.Torque","Sensor-Torque","MA-RTD 1","MA-RTD 2",\
            "U-PP.RMS.Voltage","V-PP.RMS.Voltage","W-PP.RMS.Voltage"]
            listforposition = ["L","M"]
            Dict_temp = TDMS(filename, group, listforname).Read_Tdms()
            Excel(Dict_temp, Excel_filename, sheetname, listforposition).WriteWinding()
        else:
            print("关闭中 ...")
            break
    except PermissionError:
        print("数据处理失败，请先关闭需要操作的文件\n\
..............................................\n")