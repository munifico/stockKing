import win32com.client
instCpStockCode = win32com.client.Dispatch("CpUtil.CpStockCode")
instCpCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr") #각종 코드 정보 및 코드 목록 구현
codeList = instCpCodeMgr.GetStockListByMarket(1) # 시장 구분에 따라 주식 종목을 리스트 형태로 
print(codeList)
# kospi = {}
# for code in codeList:
#     name = instCpCodeMgr.CodeToName(code)
#     kospi[code] = name
# f = open('C:\\Users\\mingz\\Desktop\\kospi.csv', 'w')
# for key, value in kospi.items():
#     f.write("%s,%s\n" % (key, value))
# f.close()
instStockChart = win32com.client.Dispatch("CpSysDib.StockChart")
instStockChart.SetInputValue(0, "A003540") #조회하려는 종목의 코드값
instStockChart.SetInputValue(1, ord('2')) #조회할 기간 기간으로 요청할땐 1 개수로 입력할땐 2 ord('2') 아스키코드값으로 변환
instStockChart.SetInputValue(4, 10) # 요청 개수 최근 거래일로부터 10일치에 해당하는 데이터 
instStockChart.SetInputValue(5, 5) #요청할 데이터 종류
instStockChart.SetInputValue(6, ord('D')) # 차트의 종류
instStockChart.SetInputValue(9, ord('1')) #수정 주가의 반영 여부 
instStockChart.BlockRequest() #서버에 데이터 요청W
numData = instStockChart.GetHeaderValue(3)
for i in range(numData):
    print(instStockChart.GetDataValue(0, i))
# objStockMst = win32com.client.Dispatch("DsCbo1.StockMst")
# objStockMst.SetInputValue(0, 'A005930')   #종목 코드 - 삼성전자
# objStockMst.BlockRequest()

instCpTdUtil = win32com.client.Dispatch("CpTrade.CpTdUtil")
instCpTd0311 = win32com.client.Dispatch("CpTrade.CpTd0311")
instCpTdUtil.TradeInit()
accountNumber = instCpTdUtil.AccountNumber[0]
accFlag = instCpTdUtil.GoodsList(accountNumber, 1)
print('accountNumber:',accountNumber,'   ',accFlag)
instCpTd0311.SetInputValue(0, 2)
instCpTd0311.SetInputValue(1, accountNumber)
instCpTd0311.SetInputValue(2, accFlag[0])
instCpTd0311.SetInputValue(3, 'A130660')
instCpTd0311.SetInputValue(4, 10)
instCpTd0311.SetInputValue(5, 13550)
instCpTd0311.BlockRequest()