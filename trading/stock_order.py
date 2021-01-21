
# # coding: utf-8

# # In[1]:


# # 대신증권 연결 확인
# import win32com.client

# instCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
# print(instCpCybos.IsConnect)


# # In[3]:


# import win32com.client

# instCpTdUtil = win32com.client.Dispatch("CpTrade.CpTdUtil")
# instCpTd0311 = win32com.client.Dispatch("CpTrade.CpTd0311")

# instCpTdUtil.TradeInit()


# # 총평가금액 / 총평가손익 / 잔고수량 / 현재가 받아오기

# In[30]:


import datetime
import time
now = datetime.datetime.now()
time = now
now = now.strftime('%H%M%S')
now


# In[31]:


import sys
import win32com.client
 
# 설명: 주식 계좌잔고 종목(최대 200개)을 가져와 현재가  실시간 조회하는 샘플
# CpEvent: 실시간 현재가 수신 클래스
# CpStockCur : 현재가 실시간 통신 클래스
# Cp6033 : 주식 잔고 조회
# CpMarketEye: 복수 종목 조회 서비스 - 200 종목 현재가를 조회 함.
 
# CpEvent: 실시간 이벤트 수신 클래스
class CpEvent:
    def set_params(self, client):
        self.client = client
 
    def OnReceived(self):
        code = self.client.GetHeaderValue(0)  # 초
        name = self.client.GetHeaderValue(1)  # 초
        timess = self.client.GetHeaderValue(18)  # 초
        exFlag = self.client.GetHeaderValue(19)  # 예상체결 플래그
        cprice = self.client.GetHeaderValue(13)  # 현재가
        diff = self.client.GetHeaderValue(2)  # 대비
        cVol = self.client.GetHeaderValue(17)  # 순간체결수량
        vol = self.client.GetHeaderValue(9)  # 거래량
 
        if (exFlag == ord('1')):  # 동시호가 시간 (예상체결)
            print("실시간(예상체결)", name, timess, "*", cprice, "대비", diff, "체결량", cVol, "거래량", vol)
        elif (exFlag == ord('2')):  # 장중(체결)
            print("실시간(장중 체결)", name, timess, cprice, "대비", diff, "체결량", cVol, "거래량", vol)

# CpStockCur: 실시간 현재가 요청 클래스
class CpStockCur:
    def Subscribe(self, code):
        self.objStockCur = win32com.client.Dispatch("DsCbo1.StockCur")
        handler = win32com.client.WithEvents(self.objStockCur, CpEvent)
        self.objStockCur.SetInputValue(0, code)
        handler.set_params(self.objStockCur)
        self.objStockCur.Subscribe()
 
    def Unsubscribe(self):
        self.objStockCur.Unsubscribe()
 
 
# Cp6033 : 주식 잔고 조회
class Cp6033:
    def __init__(self):
        # 통신 OBJECT 기본 세팅
        self.objTrade = win32com.client.Dispatch("CpTrade.CpTdUtil")
        initCheck = self.objTrade.TradeInit(0)
        if (initCheck != 0):
            print("주문 초기화 실패")
            return
 
        # 
        acc = self.objTrade.AccountNumber[0]  # 계좌번호
        accFlag = self.objTrade.GoodsList(acc, 1)  # 주식상품 구분
#         print(acc, accFlag[0])
 
        self.objRq = win32com.client.Dispatch("CpTrade.CpTd6033")
        self.objRq.SetInputValue(0, acc)  # 계좌번호
        self.objRq.SetInputValue(1, accFlag[0])  # 상품구분 - 주식 상품 중 첫번째
        self.objRq.SetInputValue(2, 50)  #  요청 건수(최대 50)
 
    # 실제적인 6033 통신 처리
    def rq6033(self, retcode):
        self.objRq.BlockRequest()
 
        # 통신 및 통신 에러 처리
        rqStatus = self.objRq.GetDibStatus()
        rqRet = self.objRq.GetDibMsg1()
        print("통신상태", rqStatus, rqRet)
        if rqStatus != 0:
            return False
 
        cnt = self.objRq.GetHeaderValue(7)
#         print(cnt)
 
#         print("종목코드 종목명 신용구분 체결잔고수량 체결장부단가 평가금액 평가손익 손익단가")
        
        temp_amount = []
        
        for i in range(cnt):
            code = self.objRq.GetDataValue(12, i)  # 종목코드
            name = self.objRq.GetDataValue(0, i)  # 종목명
            
            retcode.append(code)
            if len(retcode) >=  200:       # 최대 200 종목만,
                break
#             cashFlag = self.objRq.GetDataValue(1, i)  # 신용구분
#             date = self.objRq.GetDataValue(2, i)  # 대출일
            amount = self.objRq.GetDataValue(7, i) # 체결잔고수량
#             buyPrice = self.objRq.GetDataValue(17, i) # 체결장부단가
#             evalValue = self.objRq.GetDataValue(9, i) # 평가금액(천원미만은 절사 됨)
#             evalPerc = self.objRq.GetDataValue(10, i) # 평가손익
#             evalbep = self.objRq.GetDataValue(18, i)  # 손익단가
#             print(code,'/',name,'/', cashFlag,'/', amount,'/', buyPrice,'/', evalValue,'/', evalPerc,'/',evalbep)
            temp_amount.append(amount)
#         print('총평가금액 : ', self.objRq.GetHeaderValue(11),'총평가손익 :',self.objRq.GetHeaderValue(4)) #잔고 총 평가금액
        return  self.objRq.GetHeaderValue(11) ,self.objRq.GetHeaderValue(4), temp_amount
        
    def Request(self, retCode):
        a,b,c= self.rq6033(retCode)
 
        # 연속 데이터 조회 - 200 개까지만.
        while self.objRq.Continue:
            self.rq6033(retCode)
#             print(len(retCode))
            if len(retCode) >= 200:
                break
        # for debug
        size = len(retCode)
#         for i in range(size):
#             print(retCode[i])
        return a,b,c
 
 
# CpMarketEye : 복수종목 현재가 통신 서비스
class CpMarketEye:
    def Request(self, codes, rqField):
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
        print("통신상태", rqStatus, rqRet)
        if rqStatus != 0:
            return False
 
        cnt  = objRq.GetHeaderValue(2)
        
        name = []
        price = []
        for i in range(cnt):
#             rpCode = objRq.GetDataValue(0, i)  # 코드
#             rpName = objRq.GetDataValue(1, i)  # 종목명
#             rpTime= objRq.GetDataValue(2, i)  # 시간
#             rpDiffFlag = objRq.GetDataValue(3, i)  # 대비부호
            rpDiff = objRq.GetDataValue(4, i)  # 대비
#             rpCur = objRq.GetDataValue(5, i)  # 현재가
            rpVol = objRq.GetDataValue(6, i)  # 거래량
#             print(rpCode, rpName, rpTime,  rpDiffFlag, rpDiff, rpCur, rpVol)
            price.append(rpDiff)
            name.append(rpVol)
        return name, price

# if __name__ == "__main__":
def get_info():
    codes = []
    obj6033 = Cp6033()
    total_eval,total_profit,amount = obj6033.Request(codes) 
     
#     print("잔고 종목 개수:", len(codes))

    # 요청 필드 배열 - 종목코드, 시간, 대비부호 대비, 현재가, 거래량, 종목명
    rqField = [0,17, 1, 2, 3, 4, 10]  #요청 필드
    objMarkeyeye = CpMarketEye()
    name , price = objMarkeyeye.Request(codes, rqField) 
    
    return total_eval, total_profit, name , amount, price


#########################################################################################################################################
def sell(code,num):
# 주식 매수 주문
    # 주문 초기화
    objTrade =  win32com.client.Dispatch("CpTrade.CpTdUtil")
    initCheck = objTrade.TradeInit(0)
    if (initCheck != 0):
        print("주문 초기화 실패")
        exit()

    acc = objTrade.AccountNumber[0] #계좌번호
    accFlag = objTrade.GoodsList(acc, 1)  # 주식상품 구분
    print(acc, accFlag[0])
    objStockOrder = win32com.client.Dispatch("CpTrade.CpTd0311")
    objStockOrder.SetInputValue(0, "1")   # 1: 매도
    objStockOrder.SetInputValue(1, acc )   #  계좌번호
    objStockOrder.SetInputValue(2, accFlag[0])   # 상품구분 - 주식 상품 중 첫번째
    objStockOrder.SetInputValue(3, code)   # 종목코드 
    objStockOrder.SetInputValue(4, num)   # 매수수량 
#         objStockOrder.SetInputValue(5, price)   # 주문단가  
    objStockOrder.SetInputValue(7, "0")   # 주문 조건 구분 코드, 0: 기본 1: IOC 2:FOK
    objStockOrder.SetInputValue(8, "03")   # 주문호가 구분코드 - 01: 보통 , 03:시장가

    # 매수 주문 요청
    objStockOrder.BlockRequest()

    rqStatus = objStockOrder.GetDibStatus()
    rqRet = objStockOrder.GetDibMsg1()
    print("통신상태", rqStatus, rqRet)
    if rqStatus != 0:
        exit()

def buy(code,num):
    # 주식 매수 주문
    # 주문 초기화
    objTrade =  win32com.client.Dispatch("CpTrade.CpTdUtil")
    initCheck = objTrade.TradeInit(0)
    if (initCheck != 0):
        print("주문 초기화 실패")
        exit()        

    acc = objTrade.AccountNumber[0] #계좌번호
    accFlag = objTrade.GoodsList(acc, 1)  # 주식상품 구분
    print(acc, accFlag[0])
    objStockOrder = win32com.client.Dispatch("CpTrade.CpTd0311")
    objStockOrder.SetInputValue(0, "2")   # 2: 매수
    objStockOrder.SetInputValue(1, acc )   #  계좌번호
    objStockOrder.SetInputValue(2, accFlag[0])   # 상품구분 - 주식 상품 중 첫번째
    objStockOrder.SetInputValue(3, code)   # 종목코드 
    objStockOrder.SetInputValue(4, num)   # 매수수량 
#         objStockOrder.SetInputValue(5, price)   # 주문단가  
    objStockOrder.SetInputValue(7, "0")   # 주문 조건 구분 코드, 0: 기본 1: IOC 2:FOK
    objStockOrder.SetInputValue(8, "03")   # 주문호가 구분코드 - 01: 보통 , 03:시장가

    # 매수 주문 요청
    objStockOrder.BlockRequest()

    rqStatus = objStockOrder.GetDibStatus()
    rqRet = objStockOrder.GetDibMsg1()
    print("통신상태", rqStatus, rqRet)
    if rqStatus != 0:
        exit()
        
def present_price(code):

    import win32com.client

    # 연결 여부 체크
    objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
    bConnect = objCpCybos.IsConnect
    if (bConnect == 0):
        print("PLUS가 정상적으로 연결되지 않음. ")
        exit()

    # 현재가 객체 구하기
    objStockMst = win32com.client.Dispatch("DsCbo1.StockMst")
    objStockMst.SetInputValue(0, code)   
    objStockMst.BlockRequest()

    # 현재가 통신 및 통신 에러 처리 
    rqStatus = objStockMst.GetDibStatus()
    rqRet = objStockMst.GetDibMsg1()
    print("통신상태", rqStatus, rqRet)
    if rqStatus != 0:
        exit()

    # 현재가 정보 조회    
    return objStockMst.GetHeaderValue(11) # 종가