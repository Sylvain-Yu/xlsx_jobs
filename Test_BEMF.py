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

    def Read_One_Value(self,single_pos):
        current_sheet = self.wb[self.sheetname]
        value = current_sheet[single_pos]
        return value

    def WriteConti_value(self):
        i = 0
        for speed in self.Dict_temp["MB_Command.Speed"]:
            if speed == self.Conti_Speed:
                position = self.listforposition[self.j]+"36"
                current_sheet[position] = self.Dict_temp["SUM/AVG-RMS.Voltage"][i]
                position = self.listforposition[self.j]+"37"
                current_sheet[position] = self.Dict_temp["SUM/AVG-F.Voltage"][i]
                position = self.listforposition[self.j]+"38"
                current_sheet[position] = self.Dict_temp["SUM/AVG-RMS.Current"][i]
                position = self.listforposition[self.j]+"39"
                current_sheet[position] = self.Dict_temp["SUM/AVG-F.Current"][i]
                position = self.listforposition[self.j]+"40"
                current_sheet[position] = self.Dict_temp["SUM/AVG-Kwatts"][i]
                position = self.listforposition[self.j]+"41"
                current_sheet[position] = self.Dict_temp["SUM/AVG-F.Kwatts"][i]
                position = self.listforposition[self.j]+"42"
                current_sheet[position] = self.Dict_temp["SUM/AVG-PF"][i]
                position = self.listforposition[self.j]+"43"
                current_sheet[position] = self.Dict_temp["SUM/AVG-F.PF"][i]
                print("已找到转速为%s,第%s个数列"%(s,i+1))
                i = i + 1
        if i == 0:
            print("没有找到转速为%s的数据"%s)

    def WriteConti(self):
        listforNum = range(len(self.listforposition))
        SpeedPositionList = [ x + "6" for x in self.listforposition]
        for self.j,Speed_pos in zip(listforNum,SpeedPositionList):
            self.Conti_Speed = self.Read_One_Value(Speed_pos)
            self.WriteConti_value()



class T_input():# input value set
    def Data_input(self):
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

        choice = input("请选择需要进行数据处理的组[1/2]:\n0.Instantly Data(默认) 1.Meta Data\n")
        if int(choice.strip()) != 0:
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
Job_list = ["BEMF","Continue Torque","AA"]
Job_num = Jobs(Job_list).Select()
print("正在尝试操作: %s ..."%Job_list[Job_num])

if Job_list[Job_num] == "BEMF":
    filename, group = T_input().Data_input()
    Excel_filename = "测试表格2小时.XLSX"
    sheetname = "2.Motor BEMF"
    listforname = ["MB_Command.Speed","MA_Command.Torque","U-RMS.Voltage",\
    "V-RMS.Voltage","W-RMS.Voltage","U-F.Voltage","V-F.Voltage","W-F.Voltage"]
    listforposition = ["E","F","G","H"]
    Dict_temp = TDMS(filename,group,listforname).Read_Tdms()
    Excel(Dict_temp,Excel_filename,sheetname,listforposition).WriteBEMF()
elif Job_list[Job_num] == "Continue Torque":
    filename, group = T_input().Data_input()
    #直流电流未添加，待处理
    Excel_filename = "测试表格2小时.XLSX"
    sheetname = "4.Cont.Torque Curve"
    listforname =["MB_Command.Speed","MA_Command.Torque","Sensor-Torque",\
    "SUM/AVG-RMS.Voltage","SUM/AVG-F.Voltage","SUM/AVG-RMS.Current","SUM/AVG-F.Current",\
    "SUM/AVG-Kwatts","SUM/AVG-F.Kwatts","SUM/AVG-PF","SUM/AVG-F.PF"] 
    listforposition["G","I"]
    Dict_temp = TDMS(filename,group,listforname).Read_Tdms()
    Excel(Dict_temp,Excel_filename,sheetname,listforposition).WriteConti()

else:
    pass

