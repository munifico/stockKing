import sys
from PyQt5.QtWidgets import *
import win32com.client
import ctypes
# 설명: 당일 상승률 상위 200 종목을 가져와 현재가  실시간 조회하는 샘플
# CpEvent: 실시간 현재가 수신 클래스
# CpStockCur : 현재가 실시간 통신 클래스
# Cp7043 : 상승률 상위 종목 통신 서비스 - 연속 조회를 통해 200 종목 가져옴
# CpMarketEye: 복수 종목 조회 서비스 - 200 종목 현재가를 조회 함.
 
# CpEvent: 실시간 이벤트 수신 클래스

## 전역 OBJECT
objCpTrade = win32com.client.Dispatch('CpTrade.CpTdUtil')
g_objCpStatus = win32com.client.Dispatch('CpUtil.CpCybos')
# PLUS 실행 기본 체크 함수
def InitPlusCheck():
    # 프로세스가 관리자 권한으로 실행 여부
    if ctypes.windll.shell32.IsUserAnAdmin():
        print('정상: 관리자권한으로 실행된 프로세스입니다.')
    else:
        print('오류: 일반권한으로 실행됨. 관리자 권한으로 실행해 주세요')
        return False
 
    # 연결 여부 체크
    if (g_objCpStatus.IsConnect == 0):
        print("PLUS가 정상적으로 연결되지 않음. ")
        return False
 
    # 주문 관련 초기화
    if (objCpTrade.TradeInit(0) != 0):
        print("주문 초기화 실패")
        return False
 
    return True
 
 


class CpEvent:
    def set_params(self, client,caller):
        self.client = client
        self.caller = caller
 
    def OnReceived(self):
        item = {}
        codes = {}
        code = self.client.GetHeaderValue(0)  # 초
        name = self.client.GetHeaderValue(1)  # 초
        timess = self.client.GetHeaderValue(18)  # 초
        exFlag = self.client.GetHeaderValue(19)  # 예상체결 플래그
        cprice = self.client.GetHeaderValue(13)  # 현재가
        diff = self.client.GetHeaderValue(2)  # 대비
        cVol = self.client.GetHeaderValue(17)  # 순간체결수량
        vol = self.client.GetHeaderValue(9)  # 거래량
        item['현재가'] = cprice
        item['종목코드'] = code
        codes[code] = item
        if (exFlag == ord('1')):  # 동시호가 시간 (예상체결)
            print("실시간(예상체결)", name, timess, "*", cprice, "대비", diff, "체결량", cVol, "거래량", vol)
        elif (exFlag == ord('2')):  # 장중(체결)
            print("실시간(장중 체결)", name, timess, cprice, "대비", diff, "체결량", cVol, "거래량", vol)
 
# CpStockCur: 실시간 현재가 요청 클래스
class CpStockCur:
    def Subscribe(self, code,caller):
        self.objStockCur = win32com.client.Dispatch("DsCbo1.StockCur")
        handler = win32com.client.WithEvents(self.objStockCur, CpEvent)
        self.objStockCur.SetInputValue(0, code)
        handler.set_params(self.objStockCur,caller)
        self.objStockCur.Subscribe()
 
    def Unsubscribe(self):
        self.objStockCur.Unsubscribe()
 
 
# Cp7043 상승률 상위 요청 클래스
class Cp7043:
    def __init__(self):
        # 통신 OBJECT 기본 세팅
        self.objRq = win32com.client.Dispatch("CpSysDib.CpSvrNew7043")
        self.objRq.SetInputValue(0, ord('0')) # 거래소 + 코스닥
        self.objRq.SetInputValue(1, ord('2'))  # 상승
        self.objRq.SetInputValue(2, ord('1'))  # 당일
        self.objRq.SetInputValue(3, 21)  # 전일 대비 상위 순
        self.objRq.SetInputValue(4, ord('1'))  # 관리 종목 제외
        self.objRq.SetInputValue(5, ord('0'))  # 거래량 전체
        self.objRq.SetInputValue(6, ord('0'))  # '표시 항목 선택 - '0': 시가대비
        self.objRq.SetInputValue(7, 0)  #  등락율 시작
        self.objRq.SetInputValue(8, 30)  # 등락율 끝
 
    # 실제적인 7043 통신 처리
    def rq7043(self, retcode):
        self.objRq.BlockRequest()
        # 현재가 통신 및 통신 에러 처리
        rqStatus = self.objRq.GetDibStatus()
        rqRet = self.objRq.GetDibMsg1()
        print("통신상태", rqStatus, rqRet)
        if rqStatus != 0:
            return False
 
        cnt = self.objRq.GetHeaderValue(0)
        cntTotal  = self.objRq.GetHeaderValue(1)
        print(cnt, cntTotal)
        #retcode.append('A372290')
        for i in range(cnt):
            code = self.objRq.GetDataValue(0, i)  # 코드
            retcode.append(code)
            if len(retcode) >  10:       # 최대 200 종목만,
                break
            name = self.objRq.GetDataValue(1, i)  # 종목명
            diffflag = self.objRq.GetDataValue(3, i)
            diff = self.objRq.GetDataValue(4, i)
            vol = self.objRq.GetDataValue(6, i)  # 거래량
            print(code, name, diffflag, diff, vol)
 
    def Request(self, retCode):
        self.rq7043(retCode)
 
        # 연속 데이터 조회 - 50 개까지만.
        while self.objRq.Continue:
            self.rq7043(retCode)
            print(len(retCode))
            if len(retCode) > 10:
                break
 
        # #7043 상승하락 서비스를 통해 받은 상승률 상위 200 종목
        # size = len(retCode)
        # for i in range(size):
        #     print(retCode[i])
        return True
 
 
# CpMarketEye : 복수종목 현재가 통신 서비스
class CpMarketEye:
    def Request(self, codes, rqField,codeItem):
        # 연결 여부 체크
        objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
        bConnect = objCpCybos.IsConnect
        if (bConnect == 0):
            print("PLUS가 정상적으로 연결되지 않음. ")
            return False
 
        # 관심종목 객체 구하기
        objRq = win32com.client.Dispatch("CpSysDib.MarketEye")
        # 요청 필드 세팅 - 종목코드, 종목명, 시간, 대비부호, 대비, 현재가, 거래량
        # rqField = [0,17, 1,2,3,4,10]
        objRq.SetInputValue(0, rqField) # 요청 필드
        objRq.SetInputValue(1, codes)  # 종목코드 or 종목코드 리스트
        objRq.BlockRequest()
 
 
        # 현재가 통신 및 통신 에러 처리
        rqStatus = objRq.GetDibStatus()
        rqRet = objRq.GetDibMsg1()
        print("CpMarketEye 통신상태", rqStatus, rqRet)
        if rqStatus != 0:
            return False
 
        cnt  = objRq.GetHeaderValue(2)
        # 요청 필드 배열 - 종목코드, 시간, 대비부호 대비, 현재가, 거래량, 종목명
        #rqField = [0, 1, 2, 3, 4, 10, 17]  #요청 필드
        for i in range(cnt):
            item = {}
            rpCode = objRq.GetDataValue(0, i)  # 코드
            rpTime = objRq.GetDataValue(1, i)  # 시간
            rpDiff= objRq.GetDataValue(2, i)  # 대비
            rpDiffFlag = objRq.GetDataValue(3, i)  # 대비부호
            rpCur = objRq.GetDataValue(4, i)  # 현재가
            rpVol = objRq.GetDataValue(5, i)  # 거래량
            rpName = objRq.GetDataValue(6, i)  # 종목명
            item['종목코드'] =rpCode
            item['종목명'] =rpName
            item['현재가'] =rpCur
            item['거래량'] =rpVol
            key = rpCode
            codeItem[key] = item
            # print('아아')
            # print(rpCode, rpName, rpTime,  rpDiffFlag, rpDiff, rpCur, rpVol)
 
        return True
class CpTrade:
    def Request(self, codes,flag):
        # 주문 관련 초기화
        if (objCpTrade.TradeInit(0) != 0):
            print("주문 초기화 실패")
            return False
        # 계좌번호 조회
        acc = objCpTrade.AccountNumber[0]  # 계좌번호
        accFlag = objCpTrade.GoodsList(acc, 1)  # 주식상품 구분
        print(acc, accFlag[0])

        objStockOrder = win32com.client.Dispatch("CpTrade.CpTd0311")
        ##당일 상위 종목 50개 받아와서 매수
        ## 현재가가 5만원 이하이때만 현재가로 10주 매수        
        for i in codes:
            print('띠용: ',i,' ',codes[i],' ',codes[i]['현재가'])
            #if(codes[i]['현재가'] < 50000):
            objStockOrder.SetInputValue(0, flag)   # 2: 매수
            objStockOrder.SetInputValue(1, acc )   #  계좌번호
            objStockOrder.SetInputValue(2, accFlag[0])   # 상품구분 - 주식 상품 중 첫번째
            objStockOrder.SetInputValue(3, i)   # 종목코드 - A003540 - 대신증권 종목
            objStockOrder.SetInputValue(4, 10)   # 매수수량 10주
            objStockOrder.SetInputValue(5, codes[i]['현재가'] )   # 주문단가  - 14,100원
            objStockOrder.SetInputValue(7, "0")   # 주문 조건 구분 코드, 0: 기본 1: IOC 2:FOK
            objStockOrder.SetInputValue(8, "01")   # 주문호가 구분코드 - 01: 보통
            objStockOrder.BlockRequest()
# Cp6033 : 주식 잔고 조회
class Cp6033:
    def __init__(self):
        objCpTrade.TradeInit()
        acc = objCpTrade.AccountNumber[0]  # 계좌번호
        accFlag = objCpTrade.GoodsList(acc, 1)  # 주식상품 구분
        print(acc, accFlag[0])
 
        self.objRq = win32com.client.Dispatch("CpTrade.CpTd6033")
        self.objRq.SetInputValue(0, acc)  # 계좌번호
        self.objRq.SetInputValue(1, accFlag[0])  # 상품구분 - 주식 상품 중 첫번째
        self.objRq.SetInputValue(2, 50)  # 요청 건수(최대 50)
        self.dicflag1 = {ord(' '): '현금',
                         ord('Y'): '융자',
                         ord('D'): '대주',
                         ord('B'): '담보',
                         ord('M'): '매입담보',
                         ord('P'): '플러스론',
                         ord('I'): '자기융자',
                         }
 
 
    # 실제적인 6033 통신 처리
    def requestJango(self, caller):
        while True:
            self.objRq.BlockRequest()
            # 통신 및 통신 에러 처리
            rqStatus = self.objRq.GetDibStatus()
            rqRet = self.objRq.GetDibMsg1()
            print("Cp6033 통신상태", rqStatus, rqRet)
            if rqStatus != 0:
                return False
 
            cnt = self.objRq.GetHeaderValue(7)
            print('아아아아: ',cnt)
 
 
            for i in range(cnt):
                item = {}
                code = self.objRq.GetDataValue(12, i)  # 종목코드
                item['종목코드'] = code
                item['종목명'] = self.objRq.GetDataValue(0, i)  # 종목명
                # item['현금신용'] = self.dicflag1[self.objRq.GetDataValue(1,i)] # 신용구분
                # print(code, '현금신용', item['현금신용'])
                item['대출일'] = self.objRq.GetDataValue(2, i)  # 대출일
                item['잔고수량'] = self.objRq.GetDataValue(7, i)  # 체결잔고수량
                item['매도가능'] = self.objRq.GetDataValue(15, i)
                item['장부가'] = self.objRq.GetDataValue(17, i)  # 체결장부단가
                #item['평가금액'] = self.objRq.GetDataValue(9, i)  # 평가금액(천원미만은 절사 됨)
                #item['평가손익'] = self.objRq.GetDataValue(11, i)  # 평가손익(천원미만은 절사 됨)
                # 매입금액 = 장부가 * 잔고수량
                item['매입금액'] = item['장부가'] * item['잔고수량']
                item['현재가'] = 0
                item['대비'] = 0
                item['거래량'] = 0
 
                # 잔고 추가
#                key = (code, item['현금신용'],item['대출일'] )
                key = code
                caller.jangoData[key] = item
 
                if len(caller.jangoData) >= 10:  # 최대 200 종목만,
                    break
 
            if len(caller.jangoData) >= 10:
                break
            if (self.objRq.Continue == False):
                break
        return True
 
class MyWindow(QMainWindow):
 
    def __init__(self):
        super().__init__()

        # plus 상태 체크
        if InitPlusCheck() == False:
            exit()

        self.setWindowTitle("PLUS API TEST")
        self.setGeometry(300, 300, 300, 180)
        self.isSB = False
        self.objCur = []
        self.jangoData = {} # 잔고 object
        btnStart = QPushButton("요청 시작", self)
        btnStart.move(20, 20)
        btnStart.clicked.connect(self.btnStart_clicked)
 
        btnStop = QPushButton("요청 종료", self)
        btnStop.move(20, 70)
        btnStop.clicked.connect(self.btnStop_clicked)
 
        btnExit = QPushButton("종료", self)
        btnExit.move(20, 120)
        btnExit.clicked.connect(self.btnExit_clicked)
 
    def StopSubscribe(self):
        if self.isSB:
            cnt = len(self.objCur)
            for i in range(cnt):
                self.objCur[i].Unsubscribe()
            print(cnt, "종목 실시간 해지되었음")
        self.isSB = False
 
        self.objCur = []
 
    def btnStart_clicked(self):
        self.StopSubscribe()
        codes = []
        codeItem = {}
        obj7043 = Cp7043()
        if obj7043.Request(codes) == False:
            return
        print("상승종목 개수:", len(codes))
        print('상승 종목: ',codes)

        # 요청 필드 배열 - 종목코드, 시간, 대비부호 대비, 현재가, 거래량, 종목명
        rqField = [0, 1, 2, 3, 4, 10, 17]  #요청 필드
        objMarkeyeye = CpMarketEye()
        if (objMarkeyeye.Request(codes, rqField,codeItem) == False):
            exit()
        #매수코드
        objTrade = CpTrade()
        objTrade.Request(codeItem,"2")
        
        cnt = len(codes)
        for i in range(cnt):
            self.objCur.append(CpStockCur())
            self.objCur[i].Subscribe(codes[i],self)

        self.obj6033 = Cp6033()
        if self.obj6033.requestJango(self) == False:
            return
        print('잔고: ',self.jangoData)
        print("빼기빼기================-")
        print(cnt , "종목 실시간 현재가 요청 시작")
        self.isSB = True
 
    def btnStop_clicked(self):
        self.StopSubscribe()
 
 
    def btnExit_clicked(self):
        self.StopSubscribe()
        exit()
 
if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()
