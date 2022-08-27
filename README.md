# auto_FinAnalysis
Automated Financial Statement Analysis of Chinese Stocks

# 一个自动爬取A股公司财务报表并计算相关指标的小程序

## 功能实现：

1. 生成合并报表及报告文本
2. 生成母公司本部报表及报告文本
3. 异常变动、异常财务指标与风险提示
<img src="[./xxx.png](https://user-images.githubusercontent.com/111165917/187031517-7afffe45-ace1-4bb0-84ce-153aae45fbc0.png)"
     width = "208.6" height = "121.8" alt="财务比率" align=center />


## 数据来源：
* 新浪财经(合并报表)
    * 资产负债表:  
    https://money.finance.sina.com.cn/corp/go.php/vFD_BalanceSheet/stockid/{股票代码}/ctrl/part/displaytype/4.phtml
    * 利润表:  
    https://money.finance.sina.com.cn/corp/go.php/vFD_ProfitStatement/stockid/{股票代码}/ctrl/part/displaytype/4.phtml
    * 现金流量表:  
    https://money.finance.sina.com.cn/corp/go.php/vFD_CashFlow/stockid/{股票代码}/ctrl/part/displaytype/4.phtml

* Wind( F9 - 母公司报表) -- 人工

## 结构：
|功能|代码|
|---|---|
|主程序|main.py|
|爬虫程序|spider_func.py|
|分析程序|analysis.py|

## 环境与依赖
* Python 3.8
* 需要的包：
   bs4==0.0.1
   lxml==4.9.1
   matplotlib==3.5.3
   numpy==1.23.1
   openpyxl==3.0.10
   pandas==1.1.0
   requests==2.28.1
   xlrd==1.2.0
