from openpyxl import load_workbook
#import numpy as np
from nptdms import TdmsFile

class TDMS():# read TDMS files and dict it.
    def __init__(self,filename,group,listforname):
        self.filename = filename
        self.group = group
        self.listforname = listforname

    def Read_Tdms(self):
        Data_List = []
        tdmsfile = TdmsFile(self.filename+".tdms")

        for name in self.listforname:
            Data_List.append(tdmsfile.object(self.group,name).data)
       # print(list(zip(self.listforname,Data_List)))
        Dict_temp = dict(list(zip(self.listforname,Data_List)))
     #   print(Dict_temp)
        return Dict_temp

class Excel():  # read and write the xlsx and process.
    def __init__(self,Dict_temp,filename,sheetname,listforposition):
        self.Dict_temp = Dict_temp
        self.filename = filename
        self.sheetname = sheetname
        self.listforposition = listforposition
        self.wb = load_workbook(self.filename)
    def WriteBEMF(self):
            current_sheet = self.wb[self.sheetname]

        for v in Dict_temp["MB_Command.Speed"]:
            speed_range = [1000,2000,3000,4000]
            i = speed_range.index(v)
            if i != -1:
                position = self.listforposition[i]+"14"
                current_sheet[position] = self.Dict_temp["U-RMS.Voltage"][i]
                position = self.listforposition[i]+"15"
                current_sheet[position] = self.Dict_temp["V-RMS.Voltage"][i]
                position = self.listforposition[i]+"16"
                current_sheet[position] = self.Dict_temp["W-RMS.Voltage"][i]
            else:
                print("%srpm haven't any value"%speed_range[i])
        wb.save(Excel_filename)

filename = "M21P3-s-19nbli- A01-006-190322$BEMF"

class Info_input():# input value set
    choice = input("please choose the group[1/2]:/n1.Instantly Data(default) 2.Meta Data/n")
    if choice.strip() !="2":
        group = "Instantly Data"
    else:
        group ="Meta Data"


# Start Process

Info_input()
listforname = ["MB_Command.Speed","MA_Command.Torque","U-RMS.Voltage","V-RMS.Voltage","W-RMS.Voltage"]
listforposition = ["E","F","G","H"]
Excel_filename = "测试表格2小时.XLSX"
sheetname = "2.Motor BEMF"
Dict_temp = TDMS(filename,group,listforname).Read_Tdms()
Excel(Dict_temp,Excel_filename,sheetname,listforposition).WriteBEMF()