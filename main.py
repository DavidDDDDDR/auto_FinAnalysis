from spider_func import spider
import pandas as pd
import numpy as np
from analysis import Report,finAnalysis

stkcd = input('输入股票代码（六位数字，如：000001）：')
data = spider(stkcd)

def select():
    print(
    '''
    功能列表：
    1、生成合并报表及报告文本
    2、生成母公司本部报表及报告文本
    3、异常变动、异常财务指标与风险提示
    4、退出
    '''
    )
    func = input('输入所需功能代号：')
    while int(func) not in range(1,5):
        func = input('输入所需功能代号：')
    else: 
        return int(func)

def operation(i):
    if i == 1:
        data.ComData()
        Report(data=data,Own=False)
    elif i == 2:
        data.OwnData()
        Report(data=data,Own=True)
    elif i == 3:
        data.ComData()
        finAnalysis(data=data)

i = select()
while i != 4:
    operation(i)
    i = select()
else:
    exit()
