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
            try:
                i = speed_range.index(v)
            except ValueError:
                status = "BEMF file Write Failed !"
            else:
                position = self.listforposition[i]+"14"
                current_sheet[position] = self.Dict_temp["U-RMS.Voltage"][i]
                position = self.listforposition[i]+"15"
                current_sheet[position] = self.Dict_temp["V-RMS.Voltage"][i]
                position = self.listforposition[i]+"16"
                current_sheet[position] = self.Dict_temp["W-RMS.Voltage"][i]
                position = self.listforposition[i]+"20"
                current_sheet[position] = self.Dict_temp["U-F.Voltage"][i]
                position = self.listforposition[i]+"21"
                current_sheet[position] = self.Dict_temp["V-F.Voltage"][i]
                position = self.listforposition[i]+"22"
                current_sheet[position] = self.Dict_temp["W-F.Voltage"][i]
                status = "BEMF file Write successful !"
        self.wb.save(Excel_filename)
        print(status)



class T_input():# input value set
    def BEMF_input(self):
        dir_list = os.listdir(".")
        listforTDMS = []
        for f in dir_list:
            if ".tdms" in f.lower():
                listforTDMS.append(f)
        len_listforTDMS = len(listforTDMS)
        num = range(0,len_listforTDMS)
        file_name_dict =dict(zip(num,listforTDMS))

        for key, value in file_name_dict.items():
            print("{0}:{1}".format(key,value))
        
        while True:
            filename_num = input("please choose the Num for Source BEMF file(.tdms)!\n")
            filename_num = int(filename_num)

            if filename_num >= 0 and filename_num < len_listforTDMS:
                filename = file_name_dict[filename_num]
                break
            else:
                print("you should choose 0~ %d"%(len_listforTDMS-1))

        choice = input("please choose the group[1/2]:\n1.Instantly Data(default) 2.Meta Data\n")
        if int(choice.strip()) != 2:
            group = "Instantly Data"
        else:
            group ="Meta Data"

        return (filename,group)

class Jobs():
    def __init__(self,Job_list):
        self.Job_list = Job_list
        self.Jobs_numlist = range(0,len(self.Job_list))
        self.Jobs_dict = dict(zip(self.Jobs_numlist,self.Job_list))

    def tip(self):
        Job_num = 0
        print("请输入要操作的内容 [默认 0]")
        for key,value in self.Jobs_dict.items():
           print("{0}:{1}".format(key,value))
        Job_num = int(input())
        return Job_num
# Start Process
Job_list = ["BEMF","KK","AA"]
Job_num = Jobs(Job_list).tip()

print("Trying to operate: %s"%Job_list[Job_num])
filename , group = T_input().BEMF_input()
listforname = ["MB_Command.Speed","MA_Command.Torque","U-RMS.Voltage","V-RMS.Voltage","W-RMS.Voltage","U-F.Voltage","V-F.Voltage","W-F.Voltage"]
listforposition = ["E","F","G","H"]
Excel_filename = "测试表格2小时.XLSX"
sheetname = "2.Motor BEMF"
Dict_temp = TDMS(filename,group,listforname).Read_Tdms()
Excel(Dict_temp,Excel_filename,sheetname,listforposition).WriteBEMF()
