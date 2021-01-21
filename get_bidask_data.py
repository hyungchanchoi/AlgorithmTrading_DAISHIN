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



########### get bidask function ###########
def get_bidask(code,code_to_name):  
    dfs = []
    
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

        temp = {}
        for i in range(numData):
            temp[instStockChart.GetDataValue(9,i)] = [instStockChart.GetDataValue(2,i),instStockChart.GetDataValue(3,i) ]
        df = pd.DataFrame(temp).transpose()
        df.index.names = ['time']
        df.columns = ['bid','ask']
        dfs.append(df)

        time_ = int(time_) - 1
        print(code_to_name[code],time_)
        time.sleep(0.3)
        
    final = pd.concat(dfs)

    final.to_pickle('bidask_data/'+ code_to_name[code]+'_'+str(today))
    print(code_to_name[code],'finished')

############# main #################

if __name__ == '__main__':
    
    codes = ['KODEX 200','TIGER 200','KODEX 인버스','TIGER 인버스','KODEX 혁신기술테마액티브','TIGER AI코리아그로스액티브',
            'KODEX 코스닥 150','TIGER 코스닥150','KODEX 삼성그룹','KODEX 삼성그룹밸류']
    
    for code in codes:
        get_bidask(name_to_code[code],code_to_name)
