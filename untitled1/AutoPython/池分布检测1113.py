# -*- coding: utf-8 -*-
"""
Created on Tue Nov 12 13:27:52 2019

@author: HUAWEI
"""

#%%



def open_file(data):
    
    columns=data['DistributionType']
    if data['PaymentPeriodID'] is None==True and data['PaymentPeriodID'].sum() !=0:
        A1='PaymentPeriodID-不能为空,且只能为0'
        write1(A1)
        print('PaymentPeriodID-不能为空,且只能为0')


    if data['BucketSequenceNo'].is_unique ==False:
        A2="BucketSequenceNo-有重复"
        write1(A2)
        print('BucketSequenceNo-有重复')
        
        
    data['Count']=data['Count'].astype(float)
    data['Amount']=data['Amount'].astype(float)
            
    AmountPercentage_c=data['AmountPercentage'].sum()
    CountPercentage_c=data['CountPercentage'].sum()
    Count=data['Count'].sum()
    Amount=data['Amount'].sum()
            
    if AmountPercentage_c>=101 or AmountPercentage_c<=99 and CountPercentage_c!=0 and AmountPercentage_c!=0:
        A3="AmountPercentage-列数值设置有误-每个分布总计要等于100(忽略精度)"
        write1(A3)
        print('AmountPercentage-列数值设置有误-每个分布总计要等于100(忽略精度)')

    if CountPercentage_c>=101 or CountPercentage_c<=99 and CountPercentage_c!=0 and AmountPercentage_c!=0:
        A4="CountPercentage-列数值设置有误-每个分布总计要等于100(忽略精度)"
        write1(A4)
        print('--CountPercentage-列数值设置有误-每个分布总计要等于100(忽略精度)')

    if Count>Amount:
        A5="Count-不能大于-Amount"
        print('Count-不能大于-Amount')

def write1(x):
    with open(r'C:\Users\HUAWEI\Desktop\池分布表检测反馈1114.xlsx','a') as f:
        f.write(x)
        f.write('\n')
        f.close()




    
if __name__=="__main__":
    import pandas as pd
    import numpy as np
    import os
    import time
    import datetime
    path=r'\\172.16.7.114\已整理受托报告\实习生-北京\2019年9月份开始新增的产品\10月份新增总\说明书\池分布'
    for p_file in os.listdir(path):
        filepath=os.path.join(path,p_file)
        
        if '池分布' in p_file:
            print(filepath)
            write1(filepath)
            data=pd.read_excel(filepath)
            try:
                data.columns=['PaymentPeriodID','DistributionType','DatabaseItem','资产池分布类型','BucketSequenceNo','Bucket','Amount','AmountPercentage','Count','CountPercentage']
            
                data.dropna(subset=['Amount','Bucket','AmountPercentage'],inplace=True)
#                print(data)
                data.groupby('DistributionType').apply(open_file)
            except:
                print('未知错误-列名不能加中文')
                pass
            
                
                
                
























