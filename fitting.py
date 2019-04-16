#曲线拟合
from nptdms import TdmsFile
import numpy as np
file = "./M24P4-S-192XYY-M88-000032-A01-010-190322$Windheating.tdms"
group = "Meta Data"
channel = "MA-RTD 1"

f = TdmsFile(file)
RTD_list = f.object(group,channel).data

x_list = np.arange(0,len(RTD_list))
