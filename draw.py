#! /usr/bin/python3
# -*- coding:UTF-8 -*-
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import numpy as np
from nptdms import TdmsFile
from datetime import datetime

class Draw():
    def __init__(self,filepath,picname):
        self.filepath = filepath
        self.picname = picname

    def readTDMS(self):
        tdms = TdmsFile(self.filepath)
        self.RTD1_list = tdms.object("Meta Data","MA-RTD 1").data
        self.RTD2_list = tdms.object("Meta Data","MA-RTD 2").data
        Time = tdms.object("Meta Data","Time").data
        Torque_CMD = tdms.object("Meta Data","MA_Command.Torque").data
        t0 = Torque_CMD[80]
        j = 0
        for t in Torque_CMD:
            if t >= t0:
                break
            else:
                j += 1
        else:
            j = 0
        self.j = j
        self.RTD1_list = self.RTD1_list[j:]
        self.RTD2_list = self.RTD2_list[j:]
        Time = Time[j:]
        # 时间需要年月日，默认设置为2019年1月1日
        zero = datetime(2019,1,1)
        zero = mdates.date2num(zero)
        T0 = mdates.date2num(Time[0])
        Time = mdates.date2num(Time)
        self.Relative_time = [t -T0 + zero for t in Time]

    def drawmax(self):
        self.readTDMS()
        RTD_gap = (abs(self.RTD1_list - self.RTD2_list)).max()
        RTD1_max = self.RTD1_list.max()
        RTD2_max = self.RTD2_list.max()
        if RTD1_max >= RTD2_max:
            RTD_max_index = np.argwhere(self.RTD1_list ==RTD1_max)[-1]
            RTD_max = RTD1_max
        else:
            RTD_max_index = np.argwhere(self.RTD2_list ==RTD2_max)[-1]
            RTD_max = RTD2_max
        T_max = self.Relative_time[int(RTD_max_index)]
        tip = "max T="+str(RTD_max)+ "℃"+\
            "\nmax Gap =%.2f"%RTD_gap
        plt_xy = (self.Relative_time[int(RTD_max_index)],RTD_max)
        plt_xytext = (self.Relative_time[int(RTD_max_index*2//3)], \
        	self.RTD1_list[0]+(RTD_max-self.RTD1_list[0])*3//4)

        Arg = mdates.DateFormatter("%H:%M:%S")
        f,ax = plt.subplots(1,1)
        ax.annotate(s = tip, xy = plt_xy, xytext = plt_xytext,\
            arrowprops = dict(facecolor="red",shrink=0.05),)

        ax.set(xlabel="Time",ylabel="Temperature [℃]",title=self.picname[:-4])
        ax.plot(self.Relative_time,self.RTD1_list,label="RTD 1")
        ax.plot(self.Relative_time,self.RTD2_list,label="RTD 2")
        ax.plot(T_max,RTD_max,"ro")
        # ax.xaxis_date()
        ax.xaxis.set_major_formatter(Arg)
        ax.yaxis.grid(True,which="major")
        plt.gcf().autofmt_xdate()
        plt.legend()
        plt.savefig(self.picname)
        return 0

# @ 8 min info and draw the pic and return rtd value
    def drawXmin_returnRTD(self,minute):
        self.readTDMS()
        m = datetime(2019,1,1,0,minute)
        self.minute = mdates.date2num(m)
        i = 0
        for t in self.Relative_time:
            if self.minute < t:
                break
            else:
                i += 1
        self.i = i
        self.RTD1_max_xmin = self.RTD1_list[i] # max@ x minute
        self.RTD2_max_xmin = self.RTD2_list[i]
        self.RTD_gap = abs(self.RTD1_max_xmin - self.RTD2_max_xmin)
        self.RTD_max = (self.RTD1_max_xmin if self.RTD1_max_xmin >= self.RTD2_max_xmin\
         else self.RTD2_max_xmin)
        self.RTD1_0 = self.RTD1_list[0]
        self.RTD2_0 = self.RTD2_list[0]
        #draw pict
        plt_xy = (self.Relative_time[self.i],self.RTD_max)
        tip = "max T @"+str(minute)+"min ="+str(self.RTD_max)+ "℃"+\
            "\nGap =%.2f"%self.RTD_gap
        plt_xytext = ((self.Relative_time[self.i*2//3]), \
            self.RTD1_list[0]+(self.RTD_max-self.RTD1_list[0])*3//4)
        Arg = mdates.DateFormatter("%H:%M:%S")
        f,ax = plt.subplots(1,1)
        ax.annotate(s = tip, xy = plt_xy, xytext = plt_xytext,\
            arrowprops = dict(facecolor="red",shrink=0.05),)
        ax.set(xlabel="Time",ylabel="Temperature [℃]",title=self.picname[:-4])
        ax.plot(self.Relative_time,self.RTD1_list,label="RTD 1")
        ax.plot(self.Relative_time,self.RTD2_list,label="RTD 2")
        ax.plot(self.minute,self.RTD_max,"ro")
        # ax.xaxis_date()
        ax.xaxis.set_major_formatter(Arg)
        ax.yaxis.grid(True,which="major")
        plt.gcf().autofmt_xdate()
        plt.legend()
        plt.savefig(self.picname)
        #plt.show()
        return (self.RTD1_0,self.RTD2_0,self.RTD1_max_xmin,self.RTD2_max_xmin),self.i+self.j