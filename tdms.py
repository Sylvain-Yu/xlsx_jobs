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
