#! /usr/bin/python3
# -*- coding:UTF-8 -*-
import matplotlib.pyplot as plt
from nptdms import TdmsFile
import numpy as np

filepath = "../M24P4-S-192XYY-M88-000032-A01-010-190322$Windheating.tdms"
tdms = TdmsFile(filepath)
RTD1_list = tdms.object("Meta Data","MA-RTD 1").data
RTD2_list = tdms.object("Meta Data","MA-RTD 2").data
#print(RTD1_list)
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
fig = plt.figure()
plt_xy = (RTD_max_index,RTD_max)
plt_xytext = (RTD_max_index - 400, RTD_max - 30)
plt.annotate(s = tip, xy = plt_xy, xytext = plt_xytext,\
    arrowprops = dict(facecolor="red",shrink=0.05),)
plt.title("Temperature vs Time")
plt.xlabel("Time")
plt.ylabel("Temperatue[℃]")
x_list = range(len(RTD1_list))
plt.plot(x_list,RTD1_list,label="RTD 1")
plt.plot(x_list,RTD2_list,label="RTD 2")
plt.plot(RTD_max_index,RTD_max,"ro")
plt.legend()
plt.show()