# auto_FinAnalysis
Automated Financial Statement Analysis of Chinese Stocks

# 一个自动爬取A股公司财务报表并计算相关指标的小程序

## 功能实现：

1. 生成合并报表及报告文本
2. 生成母公司本部报表及报告文本
3. 异常变动、异常财务指标与风险提示

## 数据来源：
* 新浪财经(合并报表)
    * 资产负债表
    https://money.finance.sina.com.cn/corp/go.php/vFD_BalanceSheet/stockid/{股票代码}/ctrl/part/displaytype/4.phtml
    * 利润表
    https://money.finance.sina.com.cn/corp/go.php/vFD_ProfitStatement/stockid/{股票代码}/ctrl/part/displaytype/4.phtml
    * 现金流量表
    https://money.finance.sina.com.cn/corp/go.php/vFD_CashFlow/stockid/{股票代码}/ctrl/part/displaytype/4.phtml

* Wind( F9 - 母公司报表) -- 人工

## 结构：
|功能|代码|
|---|---|
|主程序|main.py|
|爬虫程序|spider_func.py|
|分析程序|analysis.py|
