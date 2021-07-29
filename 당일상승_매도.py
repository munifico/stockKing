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
    def set_params(self, client,name,caller):
        self.client = client
        self.name = name
        self.caller = caller
 
    def OnReceived(self):
        if self.name == 'stockcur':
            code = self.client.GetHeaderValue(0)  # 초
            name = self.client.GetHeaderValue(1)  # 초
            timess = self.client.GetHeaderValue(18)  # 초
            exFlag = self.client.GetHeaderValue(19)  # 예상체결 플래그
            cprice = self.client.GetHeaderValue(13)  # 현재가
            diff = self.client.GetHeaderValue(2)  # 대비
            cVol = self.client.GetHeaderValue(17)  # 순간체결수량
            vol = self.client.GetHeaderValue(9)  # 거래량
 
 
            item = {}
            item['code'] = code
            #rpName = self.objRq.GetDataValue(1, i)  # 종목명
            #rpDiffFlag = self.objRq.GetDataValue(3, i)  # 대비부호
            item['diff'] = diff
            item['cur'] = cprice
            item['vol'] = vol
 
            # 현재가 업데이트
            self.caller.updateJangoCurPBData(item)
        # 실시간 처리 - 주문체결
        elif self.name == 'conclution':
            # 주문 체결 실시간 업데이트
            conc = {}
 
            # 체결 플래그
            conc['체결플래그'] = self.dicflag14[self.client.GetHeaderValue(14)]
 
            conc['주문번호'] = self.client.GetHeaderValue(5)  # 주문번호
            conc['주문수량'] = self.client.GetHeaderValue(3)  # 주문/체결 수량
            conc['주문가격'] = self.client.GetHeaderValue(4)  # 주문/체결 가격
            conc['원주문'] = self.client.GetHeaderValue(6)
            conc['종목코드'] = self.client.GetHeaderValue(9)  # 종목코드
            conc['종목명'] = g_objCodeMgr.CodeToName(conc['종목코드'])
 
            conc['매수매도'] = self.dicflag12[self.client.GetHeaderValue(12)]
 
            flag15  = self.client.GetHeaderValue(15) # 신용대출구분코드
            if (flag15 in self.dicflag15):
                conc['신용대출'] = self.dicflag15[flag15]
            else:
                conc['신용대출'] = '기타'
 
            conc['정정취소'] = self.dicflag16[self.client.GetHeaderValue(16)]
            conc['현금신용'] = self.dicflag17[self.client.GetHeaderValue(17)]
            conc['주문조건'] = self.dicflag19[self.client.GetHeaderValue(19)]
 
            conc['체결기준잔고수량'] = self.client.GetHeaderValue(23)
            loandate = self.client.GetHeaderValue(20)
            if (loandate == 0):
                conc['대출일'] = ''
            else:
                conc['대출일'] = str(loandate)
            flag18 = self.client.GetHeaderValue(18)
            if (flag18 in self.dicflag18):
                conc['주문호가구분'] = self.dicflag18[flag18]
            else:
                conc['주문호가구분'] = '기타'
 
            conc['장부가'] = self.client.GetHeaderValue(21)
            conc['매도가능수량'] = self.client.GetHeaderValue(22)
 
            print(conc)
            self.caller.updateJangoCont(conc)
 
            return
        # item = {}
        # codes = {}
        # code = self.client.GetHeaderValue(0)  # 초
        # name = self.client.GetHeaderValue(1)  # 초
        # timess = self.client.GetHeaderValue(18)  # 초
        # exFlag = self.client.GetHeaderValue(19)  # 예상체결 플래그
        # cprice = self.client.GetHeaderValue(13)  # 현재가
        # diff = self.client.GetHeaderValue(2)  # 대비
        # cVol = self.client.GetHeaderValue(17)  # 순간체결수량
        # vol = self.client.GetHeaderValue(9)  # 거래량
        # item['현재가'] = cprice
        # item['종목코드'] = code
        # codes[code] = item
        # objTrade = CpTrade()
        # #and self.caller.isSell is False
        # if(len(self.caller.jangoData) !=0):
        #     print('*************** 매도: ',self.caller.jangoData[code]['종목명'],'  장부가: ',self.caller.jangoData[code]['장부가'],'   currentPrice:',cprice)
        #     if(self.caller.jangoData[code]['장부가']* 1.03 < cprice):
        #         # 현재가랑 종목 코드 필요
        #         # 3% 올랐을때 매도
        #         objTrade.Request(codes,"1")
        #     if(self.caller.jangoData[code]['장부가']* 0.97 > cprice):
        #         # 현재가랑 종목 코드 필요
        #         # 3% 떨어졌을때 매도 
        #         # 내일 다시 수정 필요!
        #         objTrade.Request(codes,"1")
        # #    self.caller.isSell = True            
        # if (exFlag == ord('1')):  # 동시호가 시간 (예상체결)
        #     print("실시간(예상체결)", name, timess, "*", cprice, "대비", diff, "체결량", cVol, "거래량", vol)
        # elif (exFlag == ord('2')):  # 장중(체결)
        #     print("실시간(장중 체결)", name, timess, cprice, "대비", diff, "체결량", cVol, "거래량", vol)
################################################
# plus 실시간 수신 base 클래스
class CpPublish:
    def __init__(self, name, serviceID):
        self.name = name
        self.obj = win32com.client.Dispatch(serviceID)
        self.bIsSB = False
 
    def Subscribe(self, var, caller):
        if self.bIsSB:
            self.Unsubscribe()
 
        if (len(var) > 0):
            self.obj.SetInputValue(0, var)
 
        handler = win32com.client.WithEvents(self.obj, CpEvent)
        handler.set_params(self.obj, self.name, caller)
        self.obj.Subscribe()
        self.bIsSB = True
 
    def Unsubscribe(self):
        if self.bIsSB:
            self.obj.Unsubscribe()
        self.bIsSB = False
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
 
        self.objRq = win32com.client.Dispatch(" ")
        self.objRq.SetInputValue(0, acc)  # 계좌번호
        self.objRq.SetInputValue(1, accFlag[0])  # 상품구분 - 주식 상품 중 첫번째
        self.objRq.SetInputValue(2, 10)  # 요청 건수(최대 50)
        self.dicflag1 = {ord(' '): '현금',
                         ord('Y'): '융자',
                         ord('D'): '대주',
                         ord('B'): '담보',
                         ord('M'): '매입담보',
                         ord('P'): '플러스론',
                         ord('I'): '자기융자',
                         }
 
 
    # 실제적인 6033 통신 처리
    def requestJango(self,caller):
        
        while True:
            self.objRq.BlockRequest()
            # 통신 및 통신 에러 처리
            rqStatus = self.objRq.GetDibStatus()
            rqRet = self.objRq.GetDibMsg1()
            print("Cp6033 통신상태", rqStatus, rqRet)
            if rqStatus != 0:
                return False
 
            cnt = self.objRq.GetHeaderValue(7)
            money = self.objRq.GetHeaderValue(8)
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
                item['평가금액'] = self.objRq.GetDataValue(9, i)  # 평가금액(천원미만은 절사 됨)
                item['평가손익'] = self.objRq.GetDataValue(11, i)  # 평가손익(천원미만은 절사 됨)
                # 매입금액 = 장부가 * 잔고수량
                item['매입금액'] = item['장부가'] * item['잔고수량']
                item['현재가'] = 0
                item['대비'] = 0
                item['거래량'] = 0
 
                # 잔고 추가
#                key = (code, item['현금신용'],item['대출일'] )
                key = code
                caller.jangoData[key] = item
                caller.codes.append(code)
                print('값: ',item)
                if len(caller.jangoData) >= 10:  # 최대 200 종목만,
                    break
 
            if len(caller.jangoData) >= 10:
                break
            if (self.objRq.Continue == False):
                break
        return True
# CpPBStockCur: 실시간 현재가 요청 클래스
class CpPBStockCur(CpPublish):
    def __init__(self):
        super().__init__('stockcur', 'DsCbo1.StockCur')
class CpPBConclusion(CpPublish):
    def __init__(self):
        super().__init__('conclution', 'DsCbo1.CpConclusion')
# 현재가 - 한종목 통신
class CpRPCurrentPrice:
    def __init__(self):
        self.objStockMst = win32com.client.Dispatch('DsCbo1.StockMst')
        return
 
    def Request(self, code, caller):
        self.objStockMst.SetInputValue(0, code)
        ret = self.objStockMst.BlockRequest()
        if self.objStockMst.GetDibStatus() != 0:
            print('통신상태', self.objStockMst.GetDibStatus(), self.objStockMst.GetDibMsg1())
            return False
 
        item = {}
        item['code'] = code
        #caller.curData['종목명'] = g_objCodeMgr.CodeToName(code)
        item['cur'] = self.objStockMst.GetHeaderValue(11)  # 종가
        item['diff'] = self.objStockMst.GetHeaderValue(12)  # 전일대비
        item['vol'] = self.objStockMst.GetHeaderValue(18)  # 거래량
        caller.curDatas[code] = item
        '''
        caller.curData['기준가'] = self.objStockMst.GetHeaderValue(27)  # 기준가
        caller.curData['예상플래그'] = self.objStockMst.GetHeaderValue(58)  # 예상플래그
        caller.curData['예상체결가'] = self.objStockMst.GetHeaderValue(55)  # 예상체결가
        caller.curData['예상대비'] = self.objStockMst.GetHeaderValue(56)  # 예상체결대비
        # 10차호가
        for i in range(10):
            key1 = '매도호가%d' % (i + 1)
            key2 = '매수호가%d' % (i + 1)
            caller.curData[key1] = (self.objStockMst.GetDataValue(0, i))  # 매도호가
            caller.curData[key2] = (self.objStockMst.GetDataValue(1, i))  # 매수호가
        '''
 
 
 
        return True
 
################################################
# CpMarketEye : 복수종목 현재가 통신 서비스
class CpMarketEye:
    def __init__(self):
        # 요청 필드 배열 - 종목코드, 시간, 대비부호 대비, 현재가, 거래량, 종목명
        self.rqField = [0, 1, 2, 3, 4, 10, 17]  # 요청 필드
 
        # 관심종목 객체 구하기
        self.objRq = win32com.client.Dispatch("CpSysDib.MarketEye")
 
    def Request(self, codes,caller):
        # 요청 필드 세팅 - 종목코드, 종목명, 시간, 대비부호, 대비, 현재가, 거래량
        self.objRq.SetInputValue(0, self.rqField)  # 요청 필드
        self.objRq.SetInputValue(1, codes)  # 종목코드 or 종목코드 리스트
        self.objRq.BlockRequest()
 
        # 현재가 통신 및 통신 에러 처리
        rqStatus = self.objRq.GetDibStatus()
        rqRet = self.objRq.GetDibMsg1()
        print("통신상태", rqStatus, rqRet)
        if rqStatus != 0:
            return False
 
        cnt = self.objRq.GetHeaderValue(2)
 
        for i in range(cnt):
            item = {}
            item['code'] = self.objRq.GetDataValue(0, i)  # 코드
            #rpName = self.objRq.GetDataValue(1, i)  # 종목명
            #rpDiffFlag = self.objRq.GetDataValue(3, i)  # 대비부호
            item['diff'] = self.objRq.GetDataValue(3, i)  # 대비
            item['cur'] = self.objRq.GetDataValue(4, i)  # 현재가
            item['vol'] = self.objRq.GetDataValue(5, i)  # 거래량
 
            caller.curDatas[item['code']] =item
 
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
        self.isSell = False
        self.objCur = []
        self.jangoData = {} # 잔고 object
        self.codes = []
        # 현재가 정보
        self.curDatas = {}
        self.objRPCur = CpRPCurrentPrice()
        self.objConclusion = CpPBConclusion()
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
    def requestJango(self):
        self.StopSubscribe()
        codeItem = {}
        self.obj6033 = Cp6033()
        if self.obj6033.requestJango(self) == False:
            return
        print('잔고: ',self.codes)

        # 잔고 현재가 통신
        codes = set()
        for code, value in self.jangoData.items():
            codes.add(code)
 
        objMarkeyeye = CpMarketEye()
        codelist = list(codes)
        if (objMarkeyeye.Request(codelist, self) == False):
            exit()
        print('아아아: ',codelist)
        # 실시간 현재가  요청
        cnt = len(codelist)
        for i in range(cnt):
            code = codelist[i]
            self.objCur[code] = CpPBStockCur()
            self.objCur[code].Subscribe(code, self)
        self.isSB = True
 
        # 실시간 주문 체결 요청
        self.objConclusion.Subscribe('', self)
    def btnStart_clicked(self):
        #잔고요청
        self.requestJango()
        # self.StopSubscribe()
        # codeItem = {}
        # self.obj6033 = Cp6033()
        # if self.obj6033.requestJango(self) == False:
        #     return
        # print('잔고: ',self.codes)
        # self.isSB = True
        # objTrade = CpTrade()
        # # if(len(self.caller.jangoData) !=0):
        # #     print('*************** 매도: ',self.caller.jangoData[code]['종목명'],'  잔고금액: ',self.caller.jangoData[code]['장부가'],'   currentPrice:',cprice)
        # #     cnt = len(self.codes)
        # #     for i in range(cnt):
        # #         if(jangoData[code]['장부가']* 1.03 < cprice):
        # #             # 현재가랑 종목 코드 필요
        # #             # 3% 올랐을때 매도
        # #             objTrade.Request(codes,"1")
        # #         if(jangoData[code]['장부가']* 0.97 > cprice):
        # #             # 현재가랑 종목 코드 필요
        # #             # 3% 떨어졌을때 매도 
        # #             # 내일 다시 수정 필요!
        # #             objTrade.Request(codes,"1")
        # cnt = len(self.codes)
        # for i in range(cnt):
        #     self.objCur.append(CpStockCur())
        #     self.objCur[i].Subscribe(self.codes[i],self)
    def btnStop_clicked(self):
        self.StopSubscribe()
 
 
    def btnExit_clicked(self):
        self.StopSubscribe()
        exit()
     # 실시간 주문 체결 처리 로직
    def updateJangoCont(self, pbCont):
        # 주문 체결에서 들어온 신용 구분 값 ==> 잔고 구분값으로 치환
        dicBorrow = {
               '현금': ord(' '),
               '유통융자': ord('Y'),
               '자기융자': ord('Y'),
               '주식담보대출': ord('B'),
               '채권담보대출': ord('B'),
               '매입담보대출': ord('M'),
               '플러스론': ord('P'),
               '자기대용융자': ord('I'),
               '유통대용융자': ord('I'),
               '기타' : ord('Z')
               }
 
        # 잔고 리스트 map 의 key 값
        # key = (pbCont['종목코드'], dicBorrow[pbCont['현금신용']], pbCont['대출일'])
        #key = pbCont['종목코드']
        code = pbCont['종목코드']
 
        # 접수, 거부, 확인 등은 매도 가능 수량만 업데이트 한다.
        if pbCont['체결플래그'] == '접수' or pbCont['체결플래그'] == '거부' or pbCont['체결플래그'] == '확인' :
            if (code not in self.jangoData) :
                return
            self.jangoData[code]['매도가능'] = pbCont['매도가능수량']
            return
 
        if (pbCont['체결플래그'] == '체결'):
            if (code not in self.jangoData) : # 신규 잔고 추가
                if (pbCont['체결기준잔고수량'] == 0) :
                    return
                print('신규 잔고 추가', code)
                # 신규 잔고 추가
                item = {}
                item['종목코드'] = pbCont['종목코드']
                item['종목명'] = pbCont['종목명']
                item['현금신용'] = dicBorrow[pbCont['현금신용']]
                item['대출일'] = pbCont['대출일']
                item['잔고수량'] = pbCont['체결기준잔고수량']
                item['매도가능'] = pbCont['매도가능수량']
                item['장부가'] = pbCont['장부가']
                # 매입금액 = 장부가 * 잔고수량
                item['매입금액'] = item['장부가'] * item['잔고수량']
 
                print('신규 현재가 요청', code)
                self.objRPCur.Request(code, self)
                self.objCur[code] = CpPBStockCur()
                self.objCur[code].Subscribe(code, self)
 
                item['현재가'] = self.curDatas[code]['cur']
                item['대비'] = self.curDatas[code]['diff']
                item['거래량'] = self.curDatas[code]['vol']
 
                self.jangoData[code] = item
 
            else:
                # 기존 잔고 업데이트
                item =self.jangoData[code]
                item['종목코드'] = pbCont['종목코드']
                item['종목명'] = pbCont['종목명']
                item['현금신용'] = dicBorrow[pbCont['현금신용']]
                item['대출일'] = pbCont['대출일']
                item['잔고수량'] = pbCont['체결기준잔고수량']
                item['매도가능'] = pbCont['매도가능수량']
                item['장부가'] = pbCont['장부가']
                # 매입금액 = 장부가 * 잔고수량
                item['매입금액'] = item['장부가'] * item['잔고수량']
 
                # 잔고 수량이 0 이면 잔고 제거
                if item['잔고수량'] == 0:
                    del self.jangoData[code]
                    self.objCur[code].Unsubscribe()
                    del self.objCur[code]
 
        return
 
    # 실시간 현재가 처리 로직
    def updateJangoCurPBData(self, curData):
        code = curData['code']
        self.curDatas[code] = curData
        self.upjangoCurData(code)
 
    def upjangoCurData(self, code):
        # 잔고에 동일 종목을 찾아 업데이트 하자 - 현재가/대비/거래량/평가금액/평가손익
        curData = self.curDatas[code]
        item = self.jangoData[code]
        item['현재가'] = curData['cur']
        item['대비'] = curData['diff']
        item['거래량'] = curData['vol']
 
if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()
