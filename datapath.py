#! /usr/bin/python3
# -*- coding:UTF-8 -*-
import os

class DataPath():# input value set
    def file_path_set(self):
        file_path = os.listdir(".")
        Excel_path_list = []
        for Excel in file_path:
            if ".xlsx" in Excel.lower():
                Excel_path_list.append(Excel)
        num = range(0,len(Excel_path_list))
        file_dict = dict(zip(num,Excel_path_list))
        for Key, Value in file_dict.items():
            print("{0}:{1}".format(Key,Value))
        while True:
            filename_num = input("请选择需要操作的Excel\n")
            filename_num = int(filename_num.strip())
            if filename_num >= 0 and filename_num < len(Excel_path_list):
                filename = file_dict[filename_num]
                break
            else:
                print("请输入 0~%d"%(len(Excel_path_list)-1))
        return filename

    def data_get(self):
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
        choice = input("请选择需要进行数据处理的组[0/1]:\n0.Instantly Data(默认) 1.Meta Data\n")
        if int(choice.strip()) == 0:
            group = "Instantly Data"
        else:
            group ="Meta Data"
        return (filename,group)