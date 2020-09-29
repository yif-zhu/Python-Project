# -*- coding: utf-8 -*-
"""
Created on Tue Oct  8 11:36:01 2019

@author: HUAWEI
"""

#%%
def open_file(path):
    
    Trust=pd.read_excel(path,sheet_name='Trust')
    TrustBond=pd.read_excel(path,sheet_name='TrustBond')
    TrustBond=TrustBond[['TrustBondID','ItemID','ItemCode','ItemValue']]
    TrustExtension=pd.read_excel(path,sheet_name='TrustExtension')
    TrustExtension=TrustExtension[['ItemCode','ItemValue']]
    PrincipalSchedule=pd.read_excel(path,sheet_name='PrincipalSchedule')
    
    huanben=PrincipalSchedule.loc[1][2]
#    print(huanben)
    
    TrustBond.dropna(subset=['ItemCode'],inplace=True)
    TrustBond.dropna(subset=['ItemID'],inplace=True)
    write1(file_path)
    try:
    #产品设立日
        Creation_Date=Trust.iloc[0][12]#12
        #法定到期日
        Maturity_date=Trust.iloc[0][13]#13
        Creation_Date=str(Creation_Date)
        Maturity_date=str(Maturity_date)
        Maturity_date=datetime.datetime.strptime(Maturity_date,"%Y-%m-%d %H:%M:%S")
        Creation_Date=datetime.datetime.strptime(Creation_Date,"%Y-%m-%d %H:%M:%S")
    except:
        print('Trust_检测出错！')
        pass
    
    
    
    try:
        
        TrusrBond_(TrustBond,Trust,huanben)
    except:
        write1('TrusrBond_检测出错！')
        print('TrusrBond_检测出错！')
        pass
    try:
        TrustExtension_(TrustExtension,Creation_Date,Maturity_date)
    except:
        print('TrustExtension_检测出错！')
        write1('TrustExtension_检测出错！')
        pass
    
    
def write1(x):
    with open(r'C:\PyCharm\pdf-docx\source\Example\1\基础表检测反馈1114.xlsx','a') as f:
        f.write(x)
        f.write('\n')
        f.close()


def TrusrBond_(TrustBond,Trust,huanben):
    write1('TrusrBond')
#TrustBond校验
    AmountC=0
    for i in TrustBond.index:
        TrustBondId=TrustBond.loc[i][0]
        ItemId=int(TrustBond.loc[i][1])
        ItemCode=TrustBond.loc[i][2]
        ItemValue=TrustBond.loc[i][3]
#        print(TrustBondId,ItemId,ItemCode,ItemValue)
        
        
        if ' ' in (ItemCode or ItemValue):
            A1="{},'单元格中不能有空格'".format(p_file)
            write1(A1)
            print(p_file,'单元格中不能有空格')
            
        if ItemCode=='SecurityExchangeCode' and ('.' or 'IB') in str(ItemValue):
            A2="{},{},'不能包含英文字符'".format(TrustBondId,ItemCode)
            write1(A2)
            print(TrustBondId,ItemCode,'不能包含英文字符')
            
        if ItemCode=='InterestPaymentType' and ItemValue not in (1,3,6,12):
            A3="{},'付息频率填写错误!'".format(TrustBondId)
            write1(A3)
            print(TrustBondId,'付息频率填写错误!')
            
            
        if ItemCode=='PaymentConvention' and ItemValue not in ('固定摊还','一次性还本付息','到期一次性还本','过手摊还','按期等额本金','按期等额本息'):
            A4="{},'还本付息方式填写错误!'".format(TrustBondId)
            write1(A4)
            print(TrustBondId,'还本付息方式填写错误!')
        
        if ItemCode=='PaymentConvention' and ItemValue=='固定摊还' and pd.isnull(huanben) :
            A5="{},'固定摊还应有还本计划'".format(TrustBondId)
            write1(A5)
            print(TrustBondId,'固定摊还应有还本计划')
            
        
        if ItemCode=='CouponPaymentReference' and ItemValue not in ('浮动利率','固定利率'):
            A6="{},'利率形式填写错误!',{},'--应该填写固定利率或浮动利率'".format(TrustBondId,ItemValue)
            write1(A6)
            print(TrustBondId,'利率形式填写错误!',{ItemValue},'--应该填写固定利率或浮动利率')
            
        if ItemCode=='InterestRateCalculation' and ItemValue not in ('按天','按月','按半年','按年'):
            A7="{},'计息方式填写错误!'".format(TrustBondId)
            write1(A7)
            print(TrustBondId,'计息方式填写错误!')
            
        if TrustBond.TrustBondID.drop_duplicates().count()==1 and ItemCode=='ClassType' and ItemValue !='FirstClass':
            A8="{},债券层级填写错误!".format(TrustBondId)
            write1(A8)
            print(TrustBondId,'债券层级填写错误')
            
        if TrustBond.TrustBondID.drop_duplicates().count()==2 and ItemCode=='ClassType' and ItemValue not in ('FirstClass','EquityClass'):
            A9="{},'债券层级填写错误'".format(TrustBondId)
            write1(A9)
            
            print(TrustBondId,'债券层级填写错误')
            
            
        if ItemCode=='OfferAmount':
            Amount=ItemValue
            AmountC=AmountC+Amount
    if  Trust.iloc[0][7]-AmountC>10 or Trust.iloc[0][7]-AmountC<-10:  #7
        A10="发行金额{}与层级金额{}相加不符".format(Trust.iloc[0][7],AmountC)
        write1(A10)
        print('发行金额{}与层级金额{}相加不符'.format(Trust.iloc[0][7],AmountC))
    
    
    
def TrustExtension_(TrustExtension,Creation_Date,Maturity_date):
    
    write1('TrustExtension')
    
    for j in TrustExtension.index:
        ItemCode=TrustExtension.loc[j][0]
        ItemValue=TrustExtension.loc[j][1]
#        print(ItemCode,ItemValue)
        
        #检测是否有空格
        if ' ' in (ItemCode or ItemValue):
            A11="单元格中不能有空格"
            write1(A11)
            print('单元格中不能有空格')
        
           
        if ItemCode in ('B_InterestCollectionDate','B_CollectionDate','B_PaymentDate','R_CollectionDate','R_InterestCollectionDate','R_PaymentDate') and ItemValue is not np.nan :
            A12="{},{},填写值为 ({}) 错误 -需为空".format(p_file,ItemCode,ItemValue)
            write1(A12)
            print(p_file,ItemCode,'填写值为 ({}) 错误 -'.format(ItemValue),'需为空')
        
        
        #获取第一个计息日
        try:
            if ItemCode in ('B_InterestCollectionDate_FirstDate','B_CollectionDate_FirstDate','B_PaymentDate_FirstDate'):
                dated_date=str(ItemValue)
                
                dated_date=datetime.datetime.strptime(dated_date,"%Y-%m-%d %H:%M:%S")
                
    #            print(dated_date,type(dated_date))
                if dated_date<Creation_Date or dated_date>Maturity_date:
                    A13="{},第一个计息日应大于产品成立日,小于法定到期日".format(ItemCode)
                    write1(A13)
                    
                    print(ItemCode,'第一个计息日应大于产品成立日,小于法定到期日')
        except:
            print('B_InterestCollectionDate_FirstDate,B_CollectionDate_FirstDate,B_PaymentDate_FirstDate -日期格式填写错误,应为‘YYYY-mm-dd’')
            A14="B_InterestCollectionDate_FirstDate,B_CollectionDate_FirstDate,B_PaymentDate_FirstDate -日期格式填写错误,应为‘YYYY-mm-dd’"
            write1(A14)
            pass
                
                
        if ItemCode in ('B_InterestCollectionDate_Condition','B_CollectionDate_Condition','B_PaymentDate_Condition') and ItemValue not in ('True','False'):
            
            A15="{},填写值为 ({}) 错误 -只能填写-True or False'".format(ItemCode,ItemValue)
            write1(A15)
            
            print(ItemCode,'填写值为 ({}) 错误 -'.format(ItemValue)+'只能填写-True or False')
            
        if ItemCode in ('B_InterestCollectionDate_ConditionCalendar','B_CollectionDate_ConditionCalendar','B_PaymentDate_ConditionCalendar') and ItemValue not in ('NaturalDay','WorkingDay','TradingDay'):
            
            A16="{},填写值为 ({}) 错误 - 只能填写-NaturalDay、WorkingDay、TradingDay".format(ItemCode,ItemValue)
            write1(A16)
            
            print(ItemCode,'填写值为 ({}) 错误 -'.format(ItemValue)+'只能填写-NaturalDay、WorkingDay、TradingDay')
            
        if ItemCode in ('B_InterestCollectionDate_Frequency','B_CollectionDate_Frequency','B_PaymentDate_Frequency') and ItemValue not in (1,3,6,12):
            
            A17="{},填写值为 ({}) 错误 -只能填写- 1、3、6、12'".format(ItemCode,ItemValue)
            write1(A17)
            print(ItemCode,'填写值为 ({}) 错误 -'.format(ItemValue)+'只能填写- 1、3、6、12')
            
            
        if ItemCode in ('B_InterestCollectionDate_ConditionTarget','B_CollectionDate_ConditionTarget','B_PaymentDate_ConditionTarget') and ItemValue not in ('BeginingOfMonth','EndOfMonth'):
            A18="{},'填写值为 ({}) 错误 -只能填写- BeginingOfMonth,EndOfMonth".format(ItemCode,ItemValue)
            write1(A18)
            print(ItemCode,'填写值为 ({}) 错误 -'.format(ItemValue)+'只能填写- BeginingOfMonth,EndOfMonth')
            
        if ItemCode in ('IsTopUpAvailable') and ItemValue not in ('True','False'):
            
            A19="填写值为 ({}) 错误 -只能填写- True,False".format(ItemCode,ItemValue)
            write1(A19)
            print(ItemCode,'填写值为 ({}) 错误 -'.format(ItemValue)+'只能填写- True,False')
        
    write1('\n')
            
#        
#        
        
if __name__=="__main__":
    import pandas as pd
    import numpy as np
    import os
    import time
    import datetime
    path=r'\\172.16.7.114\已整理受托报告\实习生-北京\2019年9月份开始新增的产品\10月份新增总\说明书\基础表'


    for p_file in os.listdir(path):
        file_path=os.path.join(path,p_file)
#        print(file_path)
#        os.chdir(file_path)
#        for file in os.listdir(file_path):
    #        print(file)
        if '基础表.xlsx' in file_path:
            
            print(file_path)
#            try:
            open_file(file_path)
            
#            except:
#                print('未知错误')
#                pass


        
        
        
        
        
        
        
        
        
        
       
        
        
        
        