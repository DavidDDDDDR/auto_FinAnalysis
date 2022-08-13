import pandas as pd
import numpy as np
import webbrowser
import matplotlib.pyplot as plt
plt.rcParams["font.sans-serif"]=["SimHei"] #设置字体
plt.rcParams["axes.unicode_minus"]=False #该语句解决图像中的“-”负号的乱码问题
import os
import sys
if not os.path.exists('./images'):
    os.makedirs('./images')

html = '''
<!DOCTYPE html>
<html lang="cn">
<head>
    <meta charset="UTF-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{title}</title>
</head>
<h1>{title}</h1>
<body>
{body}
</body>
</html>'''

red = '<font color="#FF0000">{}</font>'
img = '<center><img src="{}" alt="Pulpit rock" width="800" height="380"></center>'
div = '<div style="width:{}px; height:{}px; overflow:scroll;">{}</div>'

def html_out(string,type='p'):
    return f'<{type}>'+string+f'</{type}>\n'

def Report(data,Own=False):
    
    sum_keys = ['流动资产合计', '非流动资产合计', '资产总计','流动负债合计', '非流动负债合计', '负债合计', 
                '归属于母公司股东权益合计','所有者权益(或股东权益)合计', '负债和所有者权益(或股东权益)总计',
                '归属于母公司所有者权益合计','所有者权益合计','股东权益合计']
    
    if Own:
        CashFlow = data.OwnCashFlow
        Profit = data.OwnProfit
        Balence = data.OwnBalence
        sign = '母公司本部'
        string = html_out(f'（二）{sign}财务报表分析','h3')
    else:
        CashFlow = data.CashFlow
        Profit = data.Profit
        Balence = data.Balence 
        sign = '合并'
        string = html_out('四、公司财务情况','h2')+html_out(f'（一）{sign}财务报表分析','h3')
    
    cash_list = ['经营活动产生的现金流量净额', '投资活动产生的现金流量净额', '筹资活动产生的现金流量净额']
    cash_dict = {}
    for i in CashFlow.index:
        if len(set(i.split('、'))&set(cash_list))>0:
            cash_dict[i] = list(set(i.split('、'))&set(cash_list))[0]

    profit_list = ['营业收入','营业支出','营业成本','净利润']
    profit_dict = {}
    for i in Profit.index:
        if len(set(i.split('、'))&set(profit_list))>0:
            profit_dict[i] = list(set(i.split('、'))&set(profit_list))[0]

    sheet1 = pd.concat([Balence,Profit.loc[profit_dict.keys()].rename(index=profit_dict),CashFlow.loc[cash_dict.keys()].rename(index=cash_dict)])
    sheet1 = sheet1.loc[:,[sheet1.columns[0]] + [i for i in sheet1.columns[1:] if i[-5:] == '12-31']].rename(index={'营业支出':'营业成本'})

    
    date=sheet1.columns.to_list()
    date.sort()
    sheet1 = sheet1[date]
    
    date1 = [i for i in sheet1.columns if i[-5:]=='12-31']
    date2 = date1[-3:] + ['变化额','变化率']
    if sheet1.columns[-1][-5:]!='12-31': 
        date2.append(sheet1.columns[-1])
    last = sheet1.columns[-1]
    
    sheet1['变化额'] = sheet1[date1[-1]]-sheet1[date1[-2]]
    sheet1['变化率'] = sheet1[date1[-1]]/sheet1[date1[-2]]-1
    sheet1 = sheet1[date2]
    
    out_sheet1 = sheet1.copy()
    out_sheet1['变化率'] = [f'{i*100:.2f}%'.replace('nan%','') for i in out_sheet1['变化率']]
    # out_sheet1.to_excel(f'{data.name}+{sign}.xlsx',encoding='utf_8_sig')

    
    def Balence_str(part,df,No):
        A = df.copy()
        Asset = A[date1[-1]].iloc[-1]
        Asset_r = A['变化额'].iloc[-1]
        Asset_r1 = A['变化率'].iloc[-1]
        string1 = html_out((No+part+'情况：'), 'h4')
        if Asset_r>0:
            r_list = [i for i in A[A.变化额/Asset_r>0.1].sort_values('变化额',ascending=False).index.tolist() if i not in sum_keys]
            string1 += html_out(
                f'{data.name}公司{date1[-1][:4]}年{part}总额为{Asset:.0f}万元，较上年新增{Asset_r:.0f}万元，同比增长{Asset_r1*100:.2f}%，增长部分主要体现在：'
                +'，'.join(r_list)+'上。')
        else:
            r_list = [i for i in A[A.变化额/Asset_r<-0.1].sort_values('变化额',ascending=True).index.tolist() if i not in sum_keys]
            string1 += html_out(f'{data.name}公司{date1[-1][:4]}年{part}总额为{Asset:.0f}万元，较上年下降{-Asset_r:.0f}万元，同比下降{-Asset_r1*100:.2f}%，下降部分主要体现在：'
                             +'，'.join(r_list)+'上。')
        string1 += '\n'
        for i in r_list:
            num = A.loc[i][date1[-1]]
            change = A.loc[i]['变化率']
            if change>0:
                string1 += html_out(f'{i}为{num:.0f}万元，较上年同比增长{change*100:.2f}%，原因为：')
            else:
                string1 += html_out(f'{i}为{num:.0f}万元，较上年同比下降{-change*100:.2f}%，原因为：') 
        return string1

    
    A = sheet1.loc[:'资产总计']
    B = sheet1.loc['资产总计':'负债合计'].iloc[1:]
    E = sheet1.loc['负债合计':list(set(['所有者权益(或股东权益)合计','股东权益合计','所有者权益合计'])&set(sheet1.index))[0]].iloc[1:]

    
    string += html_out(f'{data.name}提供了近{min(3,len(date1)):.0f}年由XXXX会计师事务所审计无保留意见的审计报告及{last[:4]}年{last[5:7]}月企业自编{sign}财务部报表，现将详细内容列表如下：')
    string += out_sheet1.to_html(justify='center' ,float_format=lambda x:f'{x:.2f}'.replace('nan',''))+'\n'
    string += Balence_str('资产',A,'1、')+Balence_str('负债',B,'2、')+Balence_str('所有者权益',E,'3、')

    
    string += html_out('4、营收及利润情况：','h4')
    for i in list(set(['营业收入', '营业成本', '净利润'])&set(sheet1.index)):
        num = sheet1.loc[i][date1[-1]]
        change = sheet1.loc[i]['变化率']
        if change>0:
            string += html_out(f'{i}为{num:.0f}万元，较上年同比增长{change*100:.2f}%，其中：')
        else:
            string += html_out(f'{i}为{num:.0f}万元，较上年同比下降{-change*100:.2f}%，其中：')
    
    string += html_out('5、现金流量情况：\n','h4')
    for i in list(set(['经营活动产生的现金流量净额', '投资活动产生的现金流量净额', '筹资活动产生的现金流量净额'])&set(sheet1.index)):
        num = sheet1.loc[i][date1[-1]]
        change = sheet1.loc[i]['变化率']
        if change>0:
            string += html_out(f'{i}为{num:.0f}万元，较上年同比增长{change*100:.2f}%，其中：')
        else:
            string += html_out(f'{i}为{num:.0f}万元，较上年同比下降{-change*100:.2f}%，其中：')

    
    with open(f'{data.name}-{sign}报表分析.html','w',encoding='utf_8_sig') as txt:
        txt.write(html.format(title=f'{data.name}-{sign}报表分析',body=string))
    print(f'{sign}报表分析已输出至：\t'+f'{data.name}-{sign}报表分析.html')
    
    webbrowser.open(f'{data.name}-{sign}报表分析.html')

def finAnalysis(data):
    sheet_sum = pd.concat([data.Balence,data.CashFlow,data.Profit]).T
    date = list(sheet_sum.index)
    date.sort()
    sheet_sum = sheet_sum.loc[date]
    shift = len(set([i.split('-')[1] for i in date]))
    
    string = html_out('一、异常波动科目：','h2')
    
    list_rise = []
    list_drop = []
    date_1 = [i for i in date[:-1] if i[-5:]=='12-31' ]+[date[-1]]
    for i in sheet_sum.columns:
        i_name = i.split('：')[-1].split('、')[-1]
        if i in data.Balence.index:
            select = abs(sheet_sum[i].mean()/sheet_sum['资产总计'].mean())>=0.1
        elif i in data.Profit.index:
            select = abs(sheet_sum[i].mean()/sheet_sum['营业收入'].mean())>=0.1
        else:
            select = abs(sheet_sum[i].mean()/sheet_sum['六、期末现金及现金等价物余额'].mean())>=0.1
        plot1 = sheet_sum[i]/10000
        plot1.name = i_name
        plot2 = (plot1/plot1.shift(shift)-1)*100*(plot1.shift(shift)/plot1.shift(shift).abs())
        plot2.name = '同比增长（%）'
        sheet_html = pd.concat([plot1,plot2],axis=1).T.dropna(axis=1).to_html(justify='center' ,float_format=lambda x:f'{x:.2f}'.replace('nan',''))
        sheet_html = html_out(div.format(1200,100,sheet_html),'center').replace('<td>','<td nowrap="nowrap">').replace('<th>','<th nowrap="nowrap">')
        if select and (plot2.values[-1]==np.max(plot2.values[-1-shift*2:]) or plot2.values[-1]==np.min(plot2.values[-1-shift*2:])):
            fig,ax1 = plt.subplots(figsize=(15,7))
            ax1.bar(date_1[1:],plot1[date_1[1:]],width=0.3,label=i_name)
            plt.tick_params(labelsize=20)
            plt.ylabel(i_name+'（亿元）',fontsize= 17)
            plt.xlabel('时间',fontsize= 17)
            ax2= ax1.twinx()
            ax2.plot(date_1[1:],plot2[date_1[1:]],'r',label='同比增长（%）')
            fig.legend(fontsize= 15,bbox_to_anchor=(1,1),bbox_transform=ax1.transAxes)
            plt.tick_params(labelsize=20)
            plt.ylabel('同比增长（%）',fontsize= 17)
            plt.savefig('images/'+i_name+f'-{data.name}.png')
            plt.close(fig)
            if plot2.values[-1]==np.max(plot2.values[-1-shift*2:]):
                print(f'最新季度({date_1[-1]}):\t'+i_name+'同比增长率两年内最高')
                st = html_out(i_name+'：','h4')+html_out(f'最新季度({date_1[-1]}):\t'+i_name+'同比增长率两年内最高')+img.format('images/'+i_name+f'-{data.name}.png')
                list_rise.append(st+sheet_html+'\n')
            else:
                print(f'最新季度({date_1[-1]}):\t'+i_name+'同比增长率两年内最低')
                st = html_out(i_name+'：','h4')+html_out(f'最新季度({date_1[-1]}):\t'+i_name+'同比增长率两年内最低')+img.format('images/'+i_name+f'-{data.name}.png')
                list_drop.append(st+sheet_html+'\n')
    if len(list_rise)>0:
        string += html_out(red.format('以下报表科目增长率达到两年内最高：'),'h3')+ ''.join(list_rise)
    else:
        string += html_out(red.format('报表科目增长率均未达到两年内最高，经营情况稳定'),'h3')
    if len(list_drop)>0:
        string += html_out(red.format('以下报表科目增长率降至两年内最低：'),'h3')+ ''.join(list_drop)
    else:
        string += html_out(red.format('报表科目增长率均未降至两年内最低，经营情况稳定'),'h3')        
    
    string += html_out('二、异常财务比率：','h2')
    
    sheet_sum.columns = [i.split('：')[-1].split('、')[-1] for i in sheet_sum.columns]
    fin_rate = sheet_sum[['流动资产合计','流动负债合计']].copy()
    
    info = [
            # ['净营运资本','流动资产-流动负债','偿债能力'],
            ['流动比率','流动资产/流动负债','偿债能力'],
            ['速动比率','(流动资产-存货)/流动负债','偿债能力'],['现金比率','(流动资产-存货-应收账款)/流动负债','偿债能力'],
            ['资产负债率','总负债/总资产','偿债能力'],['销售净利率(%)','净利润/营业收入','盈利能力'],
            ['销售毛利率(%)','营业利润/营业收入','盈利能力'],['ROA(%)','净利润/总资产','盈利能力'],['ROE(%)','净利润/所有者权益','盈利能力'],
            ['主营业务收入同比增长率(%)','主营业务收入同比增长率','经营能力'],['应收账款周转率（次数）','营业收入/应收账款','经营能力'],
            ['存货周转率（次数）','营业成本/存货','经营能力'],['净现比','经营性现金流净额/净利润','现金流情况'],['收现比','经营性现金流入/营业收入','现金流情况']]
    info = pd.DataFrame(info,columns=['指标名称','计算方法','类型']).set_index(['类型','指标名称']).to_html(justify='center')
    info = html_out(info,'center').replace('<td>','<td nowrap="nowrap">').replace('<th>','<th nowrap="nowrap">')
    string += html_out(html_out('财务比率计算方法说明','p'),'center')+info
    
    for i in ['应收账款','存货','财务费用','财务费用','所得税费用']:
        if i not in sheet_sum.columns:
            sheet_sum[i] = 0
    fin_rate['adj'] = (12/fin_rate.index.str[5:7].astype(int)).astype(float)
    sheet_sum['last'] = sheet_sum.index.str[:4].astype(int)
    sheet_sum.loc[sheet_sum.index.str[5:7]!='12-31','last'] -= 1
    def last_data(x):
        x = x.sort_index()
        first = x.iloc[0].tolist()
        for i in range(len(x)):
            x.iloc[:,i] = first[i]
        return x
    last_sum = sheet_sum.groupby('last').apply(last_data).drop(columns='last')
    last_sum.index = last_sum.index.droplevel(0)
    sum2 = (sheet_sum+last_sum)/2
    
    # 偿债能力分析
    # fin_rate['净营运资本'] = sheet_sum['流动资产合计']-sheet_sum['流动负债合计']
    fin_rate['流动比率(%)'] = sheet_sum['流动资产合计']/sheet_sum['流动负债合计']*100
    fin_rate['速动比率(%)'] = (sheet_sum['流动资产合计']-sheet_sum['存货'])/sheet_sum['流动负债合计']*100
    fin_rate['现金比率(%)'] = (sheet_sum['流动资产合计']-sheet_sum['存货']-sheet_sum['应收账款'])/sheet_sum['流动负债合计']*100
    fin_rate['资产负债率(%)'] = sheet_sum['负债合计']/sheet_sum['资产总计']*100
    fin_rate.drop(columns=['流动资产合计','流动负债合计'],inplace=True)
    # print('注：报表中未体现利息费用，利息费用暂时用财务费用代替')
    # fin_rate['利息保障倍数'],fin_rate['现金流量的利息保障倍数'] = np.nan,np.nan
    # sheet_sum_i = sheet_sum[sheet_sum['财务费用']>0]
    # fin_rate.loc[sheet_sum['财务费用']>0,'利息保障倍数'] = (sheet_sum_i['净利润']+sheet_sum_i['财务费用']+sheet_sum_i['所得税费用'])/sheet_sum_i['财务费用']
    # fin_rate.loc[sheet_sum['财务费用']>0,'现金流量的利息保障倍数'] = sheet_sum_i['经营活动产生的现金流量净额']/sheet_sum_i['财务费用']
    # del sheet_sum_i
    
    # 盈利能力分析
    fin_rate['销售净利率(%)'] = sheet_sum['净利润']/sheet_sum['营业收入']*100
    fin_rate['销售毛利率(%)'] = sheet_sum['营业利润']/sheet_sum['营业收入']*100
    fin_rate['ROA(%)'] = sheet_sum['净利润']*fin_rate['adj']/sum2['资产总计']*100
    fin_rate['ROE(%)'] = sheet_sum['净利润']*fin_rate['adj']/sum2['资产总计']*100
    
    # 经营能力
    fin_rate['主营业务收入同比增长率(%)'] = (sheet_sum['营业收入']/sheet_sum['营业收入'].shift(shift)-1)*100
    fin_rate['应收账款周转率（次数）'] = sheet_sum['营业收入']*fin_rate['adj']/sum2['应收账款']
    fin_rate['存货周转率（次数）'] = sheet_sum['营业成本']*fin_rate['adj']/sum2['存货']
    
    # 现金流情况
    fin_rate['净现比'] = sheet_sum['经营活动产生的现金流量净额']/sheet_sum['净利润']
    fin_rate['收现比'] = sheet_sum['经营活动现金流入小计']/sheet_sum['营业收入']
    
    fin_rate = fin_rate.drop(columns='adj')
    
    for i in fin_rate.columns:
        string += html_out(i+'情况：','h4')
        if fin_rate[i].values[-1]==np.max(fin_rate[i].values[-1-shift*2:]) or fin_rate[i].values[-1]==np.min(fin_rate[i].values[-1-shift*2:]):
            color = 'r'
            if fin_rate[i].values[-1]==np.max(fin_rate[i].values[-1-shift*2:]):
                print(f'最新季度({date_1[-1]}):\t'+i+'达到两年内最高')
                string += red.format(html_out(i+'：','h4')+html_out(f'最新季度({date_1[-1]}):\t'+i+'达到两年内最高'))
            else:
                print(f'最新季度({date_1[-1]}):\t'+i+'达到两年内最低')
                string += red.format(html_out(i+'：','h4')+html_out(f'最新季度({date_1[-1]}):\t'+i+'达到两年内最低'))
        else:
            color = 'b'
        
        fin_rate1 = fin_rate.copy()
        fin_rate1.index = pd.to_datetime(fin_rate1.index)
        fig,ax = plt.subplots(figsize=(15,7))
        ax.plot(fin_rate1[i],color,label=i)
        plt.tick_params(labelsize=15)
        plt.xlabel('时间',fontsize= 17)
        plt.ylabel(i,fontsize= 17)
        # plt.ylim(min(fin_rate1[i].min(),0)-(fin_rate1[i].max()-fin_rate1[i].min())/20)
        plt.savefig('images/'+i+f'-{data.name}.png')
        plt.close(fig)
        sheet_html = pd.DataFrame(fin_rate[i]).T.dropna(axis=1).to_html(justify='center' ,float_format=lambda x:f'{x:.2f}'.replace('nan',''))
        sheet_html = html_out(div.format(1200,75,sheet_html),'center').replace('<td>','<td nowrap="nowrap">').replace('<th>','<th nowrap="nowrap">')
        string += img.format('images/'+i+f'-{data.name}.png')
        string += sheet_html+'\n'
    
    import webbrowser
    with open(f'{data.name}-风险与异常分析.html','w',encoding='utf_8_sig') as txt:
        txt.write(html.format(title=f'{data.name}-风险与异常分析',body=string))
    webbrowser.open(f'{data.name}-风险与异常分析.html')