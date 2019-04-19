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
        print("正在处理反电动势数据中，请等待...")
        for v in self.Dict_temp["MB_Command.Speed"]:
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
        value = self.current_sheet[single_pos].value
        return value

#仅写入耐久的数据
    def WriteConti_value(self):
        i = 0
        j = 0 # 用来判断是否找到对应参数 >0 为找到
        # print(self.Dict_temp["MB_Command.Speed"])
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
                position = self.listforposition[self.j]+"22"
                self.current_sheet[position] = self.Dict_temp["MA-RTD 1"][i]
                position = self.listforposition[self.j]+"23"
                self.current_sheet[position] = self.Dict_temp["MA-RTD 2"][i]


                j += 1
            i += 1
        if j == 0:
            print("没有找到转速为%s的数据\n"%self.Conti_Speed)
        else:
            print("已找到转速为%s,找到%s个，已选取最后一组数据"%(self.Conti_Speed,j))
# 连续扭矩试验处理数据，属于源为Instantly
    def WriteConti(self):
        print("正在处理连续扭矩数据中，请等待...")
        listforNum = range(len(self.listforposition))
        SpeedPositionList = [ x + "6" for x in self.listforposition]
        for self.j,Speed_pos in zip(listforNum,SpeedPositionList):
            self.Conti_Speed = self.Read_One_Value(Speed_pos)
            self.WriteConti_value()
        self.save()

# 为高转速试验处理数据
    def WriteHighSpeed(self,RTD,picname,i):
        print("正在处理高速数据中，请等待...")
        self.current_sheet["F14"] = RTD[0] # 开始温度RTD1
        self.current_sheet["G14"] = RTD[1] # 开始温度RTD2
        self.current_sheet["F20"] = RTD[2] # 5分钟后RTD1温度
        self.current_sheet["G20"] = RTD[3] # 5分钟后RTD2温度
        p = 0
        for s in self.Dict_temp["MB_Command.Speed"]:
            if s > 0:
                break
            else:
                p += 1
        else:
            p = 0
        img = Image(picname)
        self.current_sheet.add_image(img,"L6")
        print("已处理完高转速试验数据\n")
        self.wb.save(self.filename)

    def WriteSc_value(self):
        i = 0
        j = 0 # 用来判断是否找到对应参数 >0 为找到
        for speed in self.Dict_temp["MB_Command.Speed"]:
            if speed == self.target_speed:
                position = self.listforposition[self.j] + "10"
                self.current_sheet[position] = self.Dict_temp["U-RMS.Current"][i]
                position = self.listforposition[self.j] + "11"
                self.current_sheet[position] = self.Dict_temp["U-F.Current"][i]
                position = self.listforposition[self.j] + "12"
                self.current_sheet[position] = self.Dict_temp["V-RMS.Current"][i]
                position = self.listforposition[self.j] + "13"
                self.current_sheet[position] = self.Dict_temp["V-F.Current"][i]
                position = self.listforposition[self.j] + "14"
                self.current_sheet[position] = self.Dict_temp["W-RMS.Current"][i]
                position = self.listforposition[self.j] + "15"
                self.current_sheet[position] = self.Dict_temp["W-F.Current"][i]
                position = self.listforposition[self.j] + "20"
                self.current_sheet[position] = self.Dict_temp["Sensor-Torque"][i]
                position = self.listforposition[self.j] + "21"
                self.current_sheet[position] = self.Dict_temp["MA-Motor TEMP"][i]
                j += 1
            i += 1
        if j == 0:
            print("没有找到转速为%s的数据\n"%self.target_speed)
        else:
            print("已找到转速为%s,找到%s个，已选取最后一组数据"%(self.target_speed,j))

    def save(self):
        self.wb.save(self.filename)
        print("完成数据处理\n")

    def WriteSc(self):
        print("正在处理短路电流数据中，请等待...")
        listforNum = range(len(self.listforposition))
        SpeedPositionList = [ x + "9" for x in self.listforposition]
        for self.j ,Speed_pos in zip(listforNum,SpeedPositionList):
            self.target_speed = self.Read_One_Value(Speed_pos)
            self.WriteSc_value()
        self.save()

    def WriteWinding(self,RTD,picname,i): #unfinished
        print("正在处理连续扭矩数据中，请等待...")
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
        self.current_sheet["E15"] = self.Dict_temp["SUM/AVG-RMS.Current"][p+10]/1.732
        self.current_sheet["G15"] = self.Dict_temp["SUM/AVG-RMS.Current"][i]/1.732
        img = Image(picname)
        self.current_sheet.add_image(img,"N7")
        print("已处理完成温升试验数据\n")
        self.wb.save(self.filename)