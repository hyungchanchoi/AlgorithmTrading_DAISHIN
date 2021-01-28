# 대신증권 연결 확인
import win32com.client
import pandas as pd
import numpy as np
import math
import time
from datetime import datetime,timedelta


class Real_bidask():

    def __init__(self):      

        instCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
        print(instCpCybos.IsConnect)

        self.today = datetime.today().strftime("%Y%m%d") 

        ########### get code_to_name ###########
        instCpCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")
        codeList = instCpCodeMgr.GetStockListByMarket(1)

        self.code_to_name = {}
        self.name_to_code = {}

        for code in codeList:
            name = instCpCodeMgr.CodeToName(code)
            self.code_to_name[code] = name
            self.name_to_code[name] = code

        ########### set dictionary ###########
        self.kodex200 = {}
        self.tiger200 = {}
        self.kodex_inv = {}
        self.tiger_inv = {}
        self.kodex_active = {}
        self.tiger_active = {}
        self.samsung_group = {}
        self.samsung_value = {}
        self.kodex200_TR = {}
        self.kodex_msci = {}


    ########### get bidask function ###########
    def get_bidask(self,codes):  # 종목, 기간, 오늘, 시점, 분, 시간간격

        now = datetime.now()

        for code in codes:
            instStockChart = win32com.client.Dispatch("CpSysDib.StockChart")
            instStockChart.SetInputValue(0, self.name_to_code[code] )
            instStockChart.SetInputValue(1, ord('1'))
            instStockChart.SetInputValue(2, self.today)
            instStockChart.SetInputValue(3, self.today)
            instStockChart.SetInputValue(5, (0,1))
            instStockChart.SetInputValue(6, ord('m'))  # 'm' : 분, 'T' : 틱
            instStockChart.SetInputValue(7, 1)      # 데이터 주기
            instStockChart.SetInputValue(9, ord('1'))
            instStockChart.SetInputValue(10, 3)
            instStockChart.BlockRequest()

            bid = instStockChart.GetHeaderValue(11)
            ask = instStockChart.GetHeaderValue(12) 

            if code == 'KODEX 200':
                self.kodex200[now.strftime('%H%M%S')] = [bid,ask]
            if code == 'TIGER 200':
                self.tiger200[now.strftime('%H%M%S')] = [bid,ask]
            if code == 'KODEX 인버스':
                self.kodex_inv[now.strftime('%H%M%S')] = [bid,ask]
            if code == 'TIGER 인버스':
                self.tiger_inv[now.strftime('%H%M%S')] = [bid,ask]
            if code == 'KODEX 혁신기술테마액티브':
                self.kodex_active[now.strftime('%H%M%S')] = [bid,ask]
            if code == 'TIGER AI코리아그로스액티브':
                self.tiger_active[now.strftime('%H%M%S')] = [bid,ask]
            if code == 'KODEX 삼성그룹':
                self.samsung_group[now.strftime('%H%M%S')] = [bid,ask]
            if code == 'KODEX 삼성그룹밸류':
                self.samsung_value[now.strftime('%H%M%S')] = [bid,ask]
            if code == 'KODEX 200TR':
                self.kodex200_TR[now.strftime('%H%M%S')] = [bid,ask]
            if code == 'KODEX MSCI Korea TR':
                self.kodex_msci[now.strftime('%H%M%S')] = [bid,ask]

        time.sleep(0.5)



############# main #################       # codes = ['KODEX 200','TIGER 200','KODEX 인버스','TIGER 인버스','KODEX 혁신기술테마액티브','TIGER AI코리아그로스액티브',
                                            #         'KODEX 삼성그룹','KODEX 삼성그룹밸류']

if __name__ == '__main__':
    
    real  = Real_bidask()
    #### current time ####
    while True :    

        now = datetime.now()
        if int(now.strftime('%H%M%S')) > 152000:
            break
        else:
            print('---',now.strftime('%H%M%S'),'---')

        real.get_bidask(['KODEX 200','TIGER 200'])
        real.get_bidask(['KODEX 인버스','TIGER 인버스'])
        real.get_bidask(['KODEX 혁신기술테마액티브','TIGER AI코리아그로스액티브'])
        real.get_bidask(['KODEX 삼성그룹','KODEX 삼성그룹밸류'])
        real.get_bidask(['KODEX MSCI Korea TR','KODEX 200TR'])
    
    df = pd.DataFrame(real.kodex200).transpose()      
    df.index.names = ['time']
    df.columns = ['bid','ask']
    df.to_pickle('real_bidask_data/KODEX 200_'+str(real.today))

    df = pd.DataFrame(real.tiger200).transpose()      
    df.index.names = ['time']
    df.columns = ['bid','ask']
    df.to_pickle('real_bidask_data/TIGER 200_'+str(real.today))

    df = pd.DataFrame(real.kodex_inv).transpose()      
    df.index.names = ['time']
    df.columns = ['bid','ask']
    df.to_pickle('real_bidask_data/KODEX 인버스_'+str(real.today))

    df = pd.DataFrame(real.tiger_inv).transpose()      
    df.index.names = ['time']
    df.columns = ['bid','ask']
    df.to_pickle('real_bidask_data/TIGER 인버스_'+str(real.today))

    df = pd.DataFrame(real.kodex_active).transpose()      
    df.index.names = ['time']
    df.columns = ['bid','ask']
    df.to_pickle('real_bidask_data/KODEX 혁신기술테마액티브_'+str(real.today))

    df = pd.DataFrame(real.tiger_active).transpose()      
    df.index.names = ['time']
    df.columns = ['bid','ask']
    df.to_pickle('real_bidask_data/TIGER AI코리아그로스액티브_'+str(real.today))

    df = pd.DataFrame(real.samsung_group).transpose()      
    df.index.names = ['time']
    df.columns = ['bid','ask']
    df.to_pickle('real_bidask_data/KODEX 삼성그룹_'+str(real.today))

    df = pd.DataFrame(real.samsung_value).transpose()      
    df.index.names = ['time']
    df.columns = ['bid','ask']
    df.to_pickle('real_bidask_data/KODEX 삼성그룹밸류_'+str(real.today))
    
    df = pd.DataFrame(real.kodex200_TR).transpose()      
    df.index.names = ['time']
    df.columns = ['bid','ask']
    df.to_pickle('real_bidask_data/KODEX 200TR_'+str(real.today))

    df = pd.DataFrame(real.kodex_msci).transpose()      
    df.index.names = ['time']
    df.columns = ['bid','ask']
    df.to_pickle('real_bidask_data/KODEX MSCI Korea TR'+str(real.today))

    print('finish')