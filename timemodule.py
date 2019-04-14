# -*- coding:utf-8 -*-
import time
from datetime import datetime
a = datetime.now()
time.sleep(3)
b = datetime.now()
c = b -a
print(a,b,c)
