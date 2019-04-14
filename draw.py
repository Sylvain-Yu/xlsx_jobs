#! /usr/bin/python3
# -*- coding:UTF-8 -*-
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from nptdms import TdmsFile
import numpy as np
from datetime import datetime

filepath = "./M21P3-s-19nbli- A01-006-190322$BEMF.tdms"
tdms = TdmsFile(filepath)
RTD1_list = tdms.object("Meta Data","MA-RTD 1").data
RTD2_list = tdms.object("Meta Data","MA-RTD 2").data
Time = tdms.object("Meta Data","Time").data
# 时间需要年月日，默认设置为2019年1月1日
zero = datetime(2019,1,1)
zero = mdates.date2num(zero)
T0 = mdates.date2num(Time[0])
Time = mdates.date2num(Time)
Relative_time = [t -T0 + zero for t in Time]

RTD_gap = (abs(RTD1_list - RTD2_list)).max()
RTD1_max = RTD1_list.max()
RTD2_max = RTD2_list.max()
if RTD1_max >= RTD2_max:
    RTD_max_index = np.argwhere(RTD1_list ==RTD1_max)[-1]
    RTD_max = RTD1_max
else:
    RTD_max_index = np.argwhere(RTD2_list ==RTD2_max)[-1]
    RTD_max = RTD2_max
tip = "max T="+str(RTD_max)+ "℃"+\
"\nx ="+ str(RTD_max_index)+"\nmax Gap =%.2f"%RTD_gap
plt_xy = (Relative_time[int(RTD_max_index)],RTD_max)
plt_xytext = (Relative_time[int(RTD_max_index*2//3)], \
	RTD1_list[0]+(RTD_max-RTD1_list[0])*3//4)

Arg = mdates.DateFormatter("%H:%M:%S")

f,ax = plt.subplots(1,1)
ax.annotate(s = tip, xy = plt_xy, xytext = plt_xytext,\
    arrowprops = dict(facecolor="red",shrink=0.05),)

ax.set(xlabel="time",ylabel="Temperature [℃]",title="Temperature vs Time")
ax.plot(Relative_time,RTD1_list,label="RTD 1")
ax.plot(Relative_time,RTD2_list,label="RTD 2")
ax.plot(Relative_time[int(RTD_max_index)],RTD_max,"ro")

ax.xaxis.set_major_formatter(Arg)
ax.yaxis.grid(True,which="major")
plt.legend()
plt.show()