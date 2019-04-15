#! /usr/bin/python3
# -*- coding:UTF-8 -*-
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