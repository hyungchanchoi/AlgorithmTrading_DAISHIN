# 대신증권 연결 확인
import win32com.client

instCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
print(instCpCybos.IsConnect)

import pandas as pd
import numpy as np
import math
import time
from datetime import datetime,timedelta

today = datetime.today().strftime("%Y%m%d") 

########### get code_to_name ###########
instCpCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")
codeList = instCpCodeMgr.GetStockListByMarket(1)

code_to_name = {}
name_to_code = {}

for code in codeList:
    name = instCpCodeMgr.CodeToName(code)
    code_to_name[code] = name
    name_to_code[name] = code



########### get strength function ###########
def get_strength(code,code_to_name):  # 종목, 기간, 오늘, 시점, 분, 시간간격
    
    temp = {}
    
    time_= 1530
    
    while time_ > 899:
        
        if len(str(time_)) == 3:
            time_ = '0'+str(time_)
        
        if 60 <= int(str(time_)[-2:]) <=99:
            time_ = int(time_) - 1
            continue

        instStockChart = win32com.client.Dispatch("Dscbo1.StockBid")
        instStockChart.SetInputValue(0, code)
        instStockChart.SetInputValue(2, 80)
        instStockChart.SetInputValue(3, ord('H'))
        instStockChart.SetInputValue(4, str(time_))
        instStockChart.BlockRequest()

        numData = instStockChart.GetHeaderValue(2)
        for i in range(numData):
            temp[instStockChart.GetDataValue(9,i)] = [instStockChart.GetDataValue(4,i),instStockChart.GetDataValue(1,i),
                                                     instStockChart.GetDataValue(5,i) ,instStockChart.GetDataValue(6,i),instStockChart.GetDataValue(8,i)  ]

        time_ = int(time_) - 1
        print(code_to_name[code],time_)
        time.sleep(0.3)

    df = pd.DataFrame(temp).transpose()
    df.index.names = ['time']
    df.columns = ['현재가','전일대비','거래량','순간체결량','체결강도']
        
    df.to_pickle('strength_data/'+ code_to_name[code]+'_'+str(today))
    print(code_to_name[code],'finished')

############# main #################   
if __name__ == '__main__':
    
    codes = ['삼성전자','SK하이닉스','LG화학','삼성전자우','NAVER','삼성바이오로직스','현대차','삼성SDI','셀트리온','기아차','TIGER TOP10']

    for code in codes:
        get_strength(name_to_code[code],code_to_name)
