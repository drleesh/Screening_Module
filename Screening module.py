#!/usr/bin/env python
# -*- coding: utf8 -*-
# -*- coding: euc-kr -*-


import win32com.client

excel = win32com.client.Dispatch("Excel.Application")
# excel.Visible = True
wb = excel.Workbooks.Open('E:/screening.xls')
ws = wb.ActiveSheet

######################## YoY Screening #########################
# 매출액 YoY 증가율 0%이상인 종목의 코드와 증가율 출력
Sales_Dic = {}
tmp_Dic = {}
for i in range(2,2000):
    Sales = ws.Cells(i,111).Value
    Comp_Code_Sales = ws.Cells(i, 1).Value
    Comp_Code_Sales = str(Comp_Code_Sales)[1:]
    if Sales > 0:
        tmp_Dic = {Comp_Code_Sales: Sales}
    Sales_Dic.update(tmp_Dic)
# print(Sales_Dic)


# 영업이익 YoY 증가율 0%이상인 종목의 코드와 증가율 출력
Profit_List = []
for j in range(2,2000):
    Profit = ws.Cells(j,115).Value
    Comp_Code_Profit = ws.Cells(j,1).Value
    Comp_Code_Profit = str(Comp_Code_Profit)[1:]
    if type(Profit) == float:
        if Profit > 0:
            Profit_List.append(Comp_Code_Profit)
    else:
        try:
            if Profit.encode('utf-8') == "흑전":
                Profit_List.append(Comp_Code_Profit)
        except AttributeError:
            pass
# print(Profit_List)

for k in Profit_List:
    if k in Sales_Dic:
         # print k, Sales_Dic[k]
        Profit_Dic = {k: Sales_Dic[k] for k in Profit_List if k in Sales_Dic}
# profit dic은 매출액, 영업이익까지 screening한 결과



# 당기순이익 YoY 증가율 0%이상인 종목의 코드와 증가율 출력
NetIncome_List = []
for l in range(2, 2000):
    NetIncome = ws.Cells(l, 121).Value
    Comp_Code_NetIncome = ws.Cells(l,1).Value
    Comp_Code_NetIncome = str(Comp_Code_NetIncome)[1:]
    if type(NetIncome) == float:
        if NetIncome > 0:
            NetIncome_List.append(Comp_Code_NetIncome)
    else:
        try:
            if NetIncome.encode('utf-8') == "흑전":
                NetIncome_List.append(Comp_Code_NetIncome)
        except AttributeError:
            pass
# print(NetIncome_List)


for m in NetIncome_List:
    if m in Profit_Dic:
        # print m, Profit_Dic[m]
        NetIncome_Dic = {m: Profit_Dic[m] for m in NetIncome_List if m in Profit_Dic}
# print(NetIncome_Dic)
# NetIncome_Dic에는 매출액, 당기순이익, 영업이익 YoY Screening 결과 종목코드가 key, 매출액 증가율 YoY가 value로 들어감

# YoY_List = []
# YoY_List = NetIncome_Dic.keys()
# print(YoY_List)


######################## ROE Screening #########################
Roe_Dic = {}
RoeScreening_Dic = {}
for n in range(2, 2000):
    Profit_PQ = ws.Cells(n, 86).Value
    Comp_Code_Profit_PQ = ws.Cells(n,1).Value
    Comp_Code_Profit_PQ = str(Comp_Code_Profit_PQ)[1:]
    Capital_PQ = ws.Cells(n, 161).Value
    Capital_LQ = ws.Cells(n, 162).Value

    try:
        Roe = Profit_PQ / ((Capital_PQ + Capital_LQ)/2)
        if Roe >= 0.04:
            tmp_Dic = {Comp_Code_Profit_PQ: Roe}
            Roe_Dic.update(tmp_Dic)
    except TypeError:
        pass
# Roe_Dic에 모든 종목의 {코드: ROE} 형식으로 저장됨
# print(Roe_Dic)

for key in NetIncome_Dic:
    if key in Roe_Dic:
        # print key, Roe_Dic[key]
        RoeScreening_Dic = {key: Roe_Dic[key] for key in NetIncome_Dic if key in Roe_Dic}
# Roe_Dic에 YoY screening 종목의 코드와 Roe가 {코드: RoE} 형식으로 저장됨
print(RoeScreening_Dic)




######################## PBR Screening #########################
excel.Visible = True
wb = excel.Workbooks.Open('E:/data11.xls')
ws = wb.ActiveSheet

for key in RoeScreening_Dic:
    Comp_Code_Aggregate


excel.Quit()