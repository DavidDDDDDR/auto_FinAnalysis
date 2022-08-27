# auto_FinAnalysis
Automated Financial Statement Analysis of Chinese Stocks

# 一个自动爬取A股公司财务报表并计算相关指标的小程序

## 功能实现：

1. 生成合并报表及报告文本
2. 生成母公司本部报表及报告文本
3. 异常变动、异常财务指标与风险提示

| 类型   | 指标名称           | 计算方法                |
|------|----------------|---------------------|
| 偿债能力 | 流动比率           | 流动资产/流动负债           |
|      | 速动比率           | (流动资产-存货)/流动负债      |
|      | 现金比率           | (流动资产-存货-应收账款)/流动负债 |
|      | 资产负债率          | 总负债/总资产             |
| 盈利能力 | 销售净利率(%)       | 净利润/营业收入            |
|      | 销售毛利率(%)       | 营业利润/营业收入           |
|      | ROA(%)         | 净利润/总资产             |
|      | ROE(%)         | 净利润/所有者权益           |
| 经营能力 | 主营业务收入同比增长率(%) | 主营业务收入同比增长率         |
|      | 应收账款周转率（次数）    | 营业收入/应收账款           |
|      | 存货周转率（次数）      | 营业成本/存货             |
| 现金流  | 净现比            | 经营性现金流净额/净利润        |
|      | 收现比            | 经营性现金流入/营业收入        |

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
