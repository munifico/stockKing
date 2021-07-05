import win32com.client
import time
def CheckVolumn(instStockChart, code):
    # SetInputValue
    instStockChart.SetInputValue(0, code)
    instStockChart.SetInputValue(1, ord('2'))
    instStockChart.SetInputValue(4, 60)
    instStockChart.SetInputValue(5, 8)
    instStockChart.SetInputValue(6, ord('D'))
    instStockChart.SetInputValue(9, ord('1'))

    # BlockRequest
    instStockChart.BlockRequest()

    # GetData
    volumes = []
    numData = instStockChart.GetHeaderValue(3)
    for i in range(numData):
        volume = instStockChart.GetDataValue(0, i)
        volumes.append(volume)

    # Calculate average volume
    averageVolume = (sum(volumes) - volumes[0]) / (len(volumes) -1)

    if(volumes[0] > averageVolume * 10):
        return 1
    else:
        return 0

if __name__ == "__main__":
    instStockChart = win32com.client.Dispatch("CpSysDib.StockChart") #거래량 구하는 모듈
    instCpCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")
    inStockMst = win32com.client.Dispatch("dscbo1.StockMst")
    inStockMst.SetInputValue(0, "A000660")   
    #GetStockListByMarket 메서드의 인자로 1을 전달하면 유가증권시장의 종목을 파이썬 튜플 형태로 반환
    codeList = instCpCodeMgr.GetStockListByMarket(1)
    buyList = []
    for code in codeList:
        if CheckVolumn(instStockChart, code) == 1:
            buyList.append(code)
            # 현재가 알아오는 코드
            inStockMst.SetInputValue(0,code)
            inStockMst.BlockRequest() 
            current = inStockMst.GetHeaderValue(11)
            print(code,' 이름: ',instCpCodeMgr.CodeToName(code),' 현재가: ',current)
        time.sleep(0.2)