from openpyxl import load_workbook
from openpyxl.drawing.image import Image

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

    def WriteWinding(self,RTD,picname,i): #unfinished
        self.current_sheet["E20"] = RTD[0] #winding temp at start point:RTD 1
        self.current_sheet["G20"] = RTD[1] #winding temp at start point:RTD 2
        self.current_sheet["E21"] = RTD[2]
        self.current_sheet["G21"] = RTD[3]
        p = 0
        for t in self.Dict_temp["MA_Command.Torque"]:
            if t > 0:
                break
            else:
                p += 1
        self.current_sheet["E16"] = self.Dict_temp["MA_Command.Torque"][p]
        self.current_sheet["G16"] = self.Dict_temp["MA_Command.Torque"][i]
        self.current_sheet["E17"] = self.Dict_temp["Sensor-Torque"][p+10]
        self.current_sheet["G17"] = self.Dict_temp["Sensor-Torque"][i]
        self.current_sheet["E15"] = self.Dict_temp["SUM/AVG-RMS.Current"][p+10]/1.414
        self.current_sheet["G15"] = self.Dict_temp["SUM/AVG-RMS.Current"][i]/1.414
        img = Image(picname)
        self.current_sheet.add_image(img,"N7")
        print("已处理完成温升试验数据\n")
        self.wb.save(self.filename)