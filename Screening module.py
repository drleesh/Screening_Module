#!/usr/bin/env python
# -*- coding: utf8 -*-


import win32com.client

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True
wb = excel.Workbooks.Open('E:/screening.xls')
ws = wb.ActiveSheet
wb_Data_PQ = excel.Workbooks.Open('E:/data_PQ.xls')
ws_Data_PQ = wb_Data_PQ.ActiveSheet
wb_Data_LQ = excel.Workbooks.Open('E:/data_LQ.xls')
ws_Data_LQ = wb_Data_LQ.ActiveSheet

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


# Aggregate_Dic에는 이번분기 모든 종목의 코드와 시가총액을 담음
Aggregate_Dic = {}
Aggregate_Screening = {}
for a in range(2, 2500):
    Comp_Code_Aggregate = ws_Data_PQ.Cells(a, 1).Value
    Aggregate= ws_Data_PQ.Cells(a, 8).Value
    tmp_Dic = {Comp_Code_Aggregate: Aggregate}
    Aggregate_Dic.update(tmp_Dic)
# print(Aggregate_Dic)

# Aggregate_Screening에는 YoY, ROE 스크리닝 종목의 코드와 시총을 담음
for key in RoeScreening_Dic:
    if key in Aggregate_Dic:
        Aggregate_Screening = {key: Aggregate_Dic[key] for key in RoeScreening_Dic if key in Aggregate_Dic}
print(Aggregate_Screening)

PBRScreening_Dic = {}
Aggregate_Key = []
# Aggregate_Value = []
for b in range(2, 2000):
    Comp_Code = ws.Cells(b,1)
    Comp_Code = str(Comp_Code)[1:]
    Capital_PQ = ws.Cells(b, 161).Value
    Capital_LQ = ws.Cells(b, 162).Value
    Aggregate_Key = Aggregate_Screening.keys()
    # Aggregate_Value = Aggregate_Screening.values()
    Aggregate_Len = len(Aggregate_Key)
    for c in range(0, Aggregate_Len):
        if Comp_Code == Aggregate_Key[c]:
            Aggregate_Value = Aggregate_Screening.get(Aggregate_Key[c])
            Aggregate_Value = Aggregate_Value.replace(',', '', 10)
            Aggregate_Value = int(Aggregate_Value)
            # print(Aggregate_Value)
            PBR = Aggregate_Value / ((Capital_PQ + Capital_LQ)/2) *0.000001
            if PBR <= 3:
                tmp_Dic = {Comp_Code: PBR}
                PBRScreening_Dic.update(tmp_Dic)
print(PBRScreening_Dic)
# PBRScreening_Dic_Len = len(PBRScreening_Dic.keys())
# print(PBRScreening_Dic_Len)
# PBRScreening_Dic_Key = PBRScreening_Dic.keys()

######################## 12% Screening #########################
StockPrice_LQ_Dic = {}
StockPrice_PQ_Dic = {}
for d in range(2, 2500):
    Comp_Code = ws_Data_LQ.Cells(d, 1).Value
    StockPrice = ws_Data_LQ.Cells(d,3).Value
    tmp_Dic = {Comp_Code: StockPrice}
    StockPrice_LQ_Dic.update(tmp_Dic)

for e in range(2, 2500):
    Comp_Code = ws_Data_PQ.Cells(e, 1).Value
    StockPrice = ws_Data_PQ.Cells(e,3).Value
    tmp_Dic = {Comp_Code: StockPrice}
    StockPrice_PQ_Dic.update(tmp_Dic)

StockPriceScreening_LQ_Dic = {}
StockPriceScreening_PQ_Dic = {}
for key in PBRScreening_Dic:
    if key in StockPrice_LQ_Dic:
        StockPriceScreening_LQ_Dic = {key: StockPrice_LQ_Dic[key] for key in PBRScreening_Dic if key in StockPrice_LQ_Dic}

for key in PBRScreening_Dic:
    if key in StockPrice_PQ_Dic:
        StockPriceScreening_PQ_Dic = {key: StockPrice_PQ_Dic[key] for key in PBRScreening_Dic if key in StockPrice_PQ_Dic}

StockPrice_Value_LQ = StockPriceScreening_LQ_Dic.values()
StockPrice_Value_PQ = StockPriceScreening_PQ_Dic.values()
StockPrice_Key = StockPriceScreening_PQ_Dic.keys()
FinalScreening_Dic ={}
for f in range(0,len(StockPrice_Value_LQ)):
    StockPrice_PQ = float(StockPrice_Value_PQ[f].replace(',', '', 10))
    StockPrice_LQ = float(StockPrice_Value_LQ[f].replace(',', '', 10))
    # print(StockPrice_PQ, StockPrice_LQ)
    StockYield = (StockPrice_PQ - StockPrice_LQ) / StockPrice_LQ
    StockYield = round(StockYield, 4)
    # print(StockYield)
    if StockYield >= 0.12:
        tmp_Dic = {StockPrice_Key[f]: StockYield}
        FinalScreening_Dic.update(tmp_Dic)
print(FinalScreening_Dic)

######################## 종목명 출력 #########################
Comp_Name_Dic = {}
for g in range(2, 2000):
    Comp_Code = ws.Cells(g, 1).Value
    Comp_Code = str(Comp_Code)[1:]
    Comp_Name = ws.Cells(g, 2).Value
    # Comp_Name = str(Comp_Name.decode('utf-8'))
    tmp_Dic = {Comp_Code: Comp_Name}
    Comp_Name_Dic.update(tmp_Dic)
# print(Comp_Name_Dic)

BuyRecommended_Dic = {}
for key in Comp_Name_Dic:
    if key in FinalScreening_Dic:
        BuyRecommended_Dic = {key: Comp_Name_Dic[key] for key in Comp_Name_Dic if key in FinalScreening_Dic}
print(BuyRecommended_Dic)

for h in BuyRecommended_Dic.keys():
    print(h)
    print(BuyRecommended_Dic[h])

# BuyRecommended_Name = BuyRecommended_Dic.values()
# # print(type(BuyRecommended_Name[0]))
# print(BuyRecommended_Name)

excel.Quit()