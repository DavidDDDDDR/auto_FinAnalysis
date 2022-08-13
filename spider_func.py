import os
import sys
if not os.path.exists('./data'):
    os.makedirs('./data')
if not os.path.exists('./母公司报表（从Wind下载）'):
    os.makedirs('./母公司报表（从Wind下载）')

import pandas as pd
import numpy as np
import requests
import json
import time
from bs4 import BeautifulSoup

account_list = ['流动资产','非流动资产','流动负债','非流动负债','所有者权益','资产','负债','负债及股东权益',
                '一、经营活动产生的现金流量','二、投资活动产生的现金流量','三、筹资活动产生的现金流量']

def get_name(stkcd):
    headers = {'user-agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/103.0.0.0 Safari/537.36'}
    url=f'http://vip.stock.finance.sina.com.cn/corp/go.php/vFD_BalanceSheet/stockid/{stkcd}/ctrl/part/displaytype/4.phtml'
    data = requests.get(url,headers=headers)
    cookies = data.cookies
    # headers = data.headers

    soup = BeautifulSoup(data.text,features="html.parser")
    name = soup.find('div',id='center').find('a').text
    print('股票名称：',name)

    k = pd.read_html(data.text)
    years = k[12].values[0][0].split(' ')[2:]
    return name,years,cookies,headers

class spider:
    '''
    获得所有历史三大报表数据
    '''
    def __init__(self,stkcd):
        self.stkcd = stkcd
        self.name,self.year,self.cookies,self.headers = get_name(stkcd)
        self.year = self.year[:5]
        self.Balence = 0
        self.url = 'https://vip.stock.finance.sina.com.cn/corp/go.php/'

    def BalanceSheet(self):
        '''
        爬取该公司全部历史合并资产负债表
        '''
        datas = []
        for i in self.year:
            content = requests.get(self.url + f'vFD_BalanceSheet/stockid/{self.stkcd}/ctrl/{i}/displaytype/4.phtml',
                                   headers=self.headers,cookies=self.cookies)
            self.cookies = content.cookies
            datas.append(pd.read_html(content.text)[13].replace('--',np.nan))
            # time.sleep(2)
        datas = pd.concat(datas,axis=1)
        datas[datas.isin(account_list)] = np.nan
        datas = datas.dropna(how='all').dropna(how='all',axis=1)
        datas.columns = datas.iloc[0]
        datas.index = datas.iloc[:,0]
        datas = datas[[i for i in datas.columns if i != '报表日期']].iloc[1:]
        # print(datas)
        datas = datas[~(datas.iloc[:,0].isna())].astype(float)
        datas.to_excel(f'./data/{self.name}-{self.stkcd}-资产负债表.xlsx',encoding='utf_8_sig')

        self.Balence = datas
        print('资产负债表爬取完毕')
    
    def ProfitSheet(self):
        '''
        爬取该公司全部历史合并利润表
        '''
        datas = []
        for i in self.year:
            content = requests.get(self.url + f'vFD_ProfitStatement/stockid/{self.stkcd}/ctrl/{i}/displaytype/4.phtml',
                                   headers=self.headers,cookies=self.cookies)
            self.cookies = content.cookies
            datas.append(pd.read_html(content.text)[13].replace('--',np.nan))
            # time.sleep(2)
        datas = pd.concat(datas,axis=1)
        datas[datas.isin(account_list)] = np.nan
        datas = datas.dropna(how='all').dropna(how='all',axis=1)
        datas.columns = datas.iloc[0]
        datas.index = datas.iloc[:,0]
        datas = datas[[i for i in datas.columns if i != '报表日期']].iloc[1:]
        datas = datas[~(datas.iloc[:,0].isna())].astype(float)
        datas.to_excel(f'./data/{self.name}-{self.stkcd}-利润表.xlsx',encoding='utf_8_sig')

        self.Profit = datas
        print('利润表爬取完毕')
    
    def CashFlowSheet(self):
        '''
        爬取该公司全部历史合并现金流量表
        '''
        datas = []
        for i in self.year:
            content = requests.get(self.url + f'vFD_CashFlow/stockid/{self.stkcd}/ctrl/{i}/displaytype/4.phtml',
                                   headers=self.headers,cookies=self.cookies)
            self.cookies = content.cookies
            datas.append(pd.read_html(content.text)[13].replace('--',np.nan))
            # time.sleep(2)
        datas = pd.concat(datas,axis=1)
        
        datas[datas.isin(account_list)] = np.nan
        datas = datas.dropna(how='all').dropna(how='all',axis=1)
        datas.index = datas.iloc[:,0]
        datas = datas.loc[:'附注'].dropna(how='all').dropna(how='all',axis=1)
        datas = datas.loc[[i for i in datas.index if i != '附注']]
        datas.columns = datas.iloc[0]
        datas = datas[[i for i in datas.columns if i not in ['报表日期',np.nan]]].iloc[1:]
        datas = datas[~(datas.iloc[:,0].isna())].astype(float)
        datas.to_excel(f'./data/{self.name}-{self.stkcd}-现金流量表.xlsx',encoding='utf_8_sig')

        self.CashFlow = datas
        print('现金流量表爬取完毕')
        
    def OwnData(self):
        '''
        获取Wind下载的母公司报表
        '''
        a = requests.get(self.url+f'vCB_Bulletin/stockid/{self.stkcd}/page_type/ndbg.phtml', headers=self.headers, cookies=self.cookies)
        self.cookies = a.cookies
        c = BeautifulSoup(a.text,features="html.parser")
        self.stkcd1 = c.find('h1').span.text[1:-1]
        a = f'./母公司报表（从Wind下载）/{self.stkcd1}-利润表.xlsx'
        b = f'./母公司报表（从Wind下载）/{self.stkcd1}-资产负债表.xlsx'
        c = f'./母公司报表（从Wind下载）/{self.stkcd1}-现金流量表.xlsx'
        del_ind = ['报告期', '报表类型', '显示币种', '原始币种', '转换汇率', '利率类型', '税率', '税率说明', 
                   '审计意见(境内)', '审计意见(境外)', '调整原因', '调整说明', '公告日期', '数据来源']
        print('请从Wind（F9-深度资料）下载“母公司财务报表”，放入以下位置：\n\t'+a+'\n\t'+b+'\n\t'+c)
        y = input('若已放入对应位置，请输入y并回车：')
        if y in ['y','Y']:
            if os.path.exists(a) and os.path.exists(b) and os.path.exists(c):
                self.OwnProfit = pd.read_excel(a,skipfooter=3)
                self.OwnProfit.index = [i.replace(' ','') for i in self.OwnProfit.iloc[:,0]]
                self.OwnProfit = self.OwnProfit.drop(index=del_ind).iloc[:,1:].dropna()
                self.OwnBalence = pd.read_excel(b,skipfooter=3)
                self.OwnBalence.index = [i.replace(' ','') for i in self.OwnBalence.iloc[:,0]]
                self.OwnBalence = self.OwnBalence.drop(index=del_ind).iloc[:,1:].dropna()
                self.OwnCashFlow = pd.read_excel(c,skipfooter=3)
                self.OwnCashFlow.index = [i.replace(' ','') for i in self.OwnCashFlow.iloc[:,0]]
                self.OwnCashFlow = self.OwnCashFlow.drop(index=del_ind).iloc[:,1:].dropna().loc[:'筹资活动产生的现金流量净额']
            else:
                raise  ValueError("未检测到文件，请检查文件名和路径")
        else:
            exit()
        
    def ComData(self):
        self.BalanceSheet()
        self.ProfitSheet()
        self.CashFlowSheet()


if __name__=='__main__':
    
    i = spider('000002')
    i.ComData()
    i.OwnData()
    
    for data in [i.Profit,i.Balence,i.CashFlow,i.OwnProfit,i.OwnBalence,i.OwnCashFlow,]:
        data = data
        print(data)
        