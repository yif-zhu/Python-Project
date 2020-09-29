# _*_ coding:utf-8 _*_

import sys
import os,re
import os.path
import datetime
import xml.etree.ElementTree as XETree
import pdfplumber
import pyodbc
import calendar
import pandas as pd


logtxtFilePath = None
errtxtFilePath = None
dbConnectionStr = None

##### Helper Methods #####
def writeLog(msg):
    if not os.path.exists(logtxtFilePath):
        f = open(logtxtFilePath, "w")
    print(msg)
    with open(logtxtFilePath, "a") as f:
        ts = datetime.datetime.now().strftime('[%H:%M:%S]')
        f.write('{0}:  {1}\n'.format(ts, msg))

def execSQLCmd(sql):
    # print(sql)
    cnxn = pyodbc.connect(dbConnectionStr)
    try:
        cursor = cnxn.cursor()
        cursor.execute(sql)
        cnxn.commit()
    except Exception as ex:
        writeLog(str(ex))
        print(str(ex))
        # raise ex
    finally:
        cnxn.close()
def deleteOldData(fileName):
    sql = "delete from  PortfolioManagement.DvImport.StaticPoolData where FileNames = N'{0}'".format(fileName)
    execSQLCmd(sql)

def getDate(value):
    model1 = r"(\d{4}年\d{1,2}月(\d{1,2}日)?)"
    model2 = r"(\d{4}[-/]\d{1,2}[^a-z]([-/]\d{1,2}[^a-z])?)|((\d{1,2}[^a-z][-/])?\d{1,2}[^a-z][-/]\d{4})"
    model3 = r"^(\d{5,6})$"
    if re.search(model1, value) is not None:
        value = re.sub('\D$', '', value)
        value = re.sub(r'\D', r'-', value)
        value = pd.to_datetime(value)
    elif re.search(model3, value) is not None:
        if re.search(model2, value) is not None:
            value = pd.to_datetime(value)
        else:
            value = pd.to_datetime(value, format=('%Y%m'))
    else:
        value = pd.to_datetime(value)
    year, month = str(value).split('-')[0], str(value).split('-')[1]  # 获取月末日期
    end = calendar.monthrange(int(year), int(month))[1]
    value = pd.to_datetime('%s-%s-%s' % (year, month, end))
    return value.strftime('%Y%m%d')

def getNumber(Value,Type):
    cvalue = str(Value).replace(' ', '').replace(',', '').replace('.00', '').replace('\t', '').replace('，', '')
    if cvalue == 'NA' or cvalue == '-' or cvalue == '':
        return 0
    if Type == 'Int':
        return int(cvalue)
    elif Type == 'Float':
        return '%.2f' %float(cvalue)

def main(configFilePath, dateId):
    global logtxtFilePath
    global errtxtFilePath
    global dbConnectionStr
    global trustID
    global paymentPeriodID

    mappingTree = XETree.parse(configFilePath)
    cfgRoot = mappingTree.getroot()
    sql = cfgRoot.attrib['sql']
    filePath = cfgRoot.attrib['filePath']
    beginPage = int(cfgRoot.attrib['beginPage'])
    endPage = int(cfgRoot.attrib['endPage'])
    columns = cfgRoot.attrib['columns']
    # ExcelColumns = cfgRoot.attrib['ExcelColumns']
    dbConnectionStr = 'DRIVER={SQL Server};SERVER=172.16.6.143\MSSQL;DATABASE=PortfolioManagement;UID=sa;PWD=PasswordGS2017'

    fileName = os.path.basename(filePath)
    scriptFolderPath = os.path.dirname(os.path.abspath(__file__))
    log_Path = os.path.join(scriptFolderPath, "Logs")
    if not os.path.exists(log_Path):
        os.mkdir(log_Path)
    logtxtFilePath = os.path.join(scriptFolderPath, 'Logs', '{0}.txt'.format(dateId))
    Types = []      #用来存储每列对应的类型
    deleteOldData(fileName)     #删除同名文件的旧数据
    for column in columns.split(','):
        if column == '贷款发放时间' or column =='报告月份月末':
            Types.append('Date')
        elif '元' in column:
            Types.append('Float')
        else:
            Types.append('Int')

    with pdfplumber.open(filePath) as pdf:
        money = 1
        page = pdf.pages[beginPage]
        text = page.extract_text()
        #table = page.extract_tables()[-1][0]
        #columnNum = []  #获取对应PDF对应的列
        if '万元' in text:  # 确定金额的单位
            money = 10000
        elif '亿元' in text:
            money = 100000000
        # for excelColumn in ExcelColumns.split(','):
        #     for index, row in enumerate(table):
        #         if row == excelColumn:
        #             columnNum.append(index)
        #             break
        cellValue = ''
        for pageNum in range(beginPage, endPage+1):
            execSql = sql
            page = pdf.pages[pageNum]
            tables = page.extract_tables()
            tablesNum = 0
            if pageNum == beginPage:   #取当前页对应的表格，静态池的表格在首页时，一定是最后面那个表格
                tablesNum = len(tables)-1
                cellValue = tables[tablesNum][0][0]
            for index, Row in enumerate(tables[tablesNum]):
                if index == 0 and Row[0] == cellValue:
                    continue
                if Row[0] is None and Row[1] is None:
                    continue
                execSql = execSql + "( N'"+ fileName + "',"
                for i in range(len(Types)):
                    if Types[i] == 'Date':
                        execSql = execSql + getDate(Row[i]) + ','
                    if Types[i] == 'Float':
                        execSql = execSql + str(getNumber(Row[i], 'Float')*money) + ','
                    if Types[i] == 'Int':
                        execSql = execSql + str(getNumber(Row[i], 'Int')*money) + ','
                execSql = execSql.rstrip(',') + '),'
            execSql = execSql.rstrip(',')
            writeLog('当前页数据已上传，当前页数为：'+ str(pageNum))
            execSQLCmd(execSql)
            # for pageNum in range(beginPage, endPage):     #考虑录入pdf字段时的方案
            #     execSql = sql
            #     page = pdf.pages[pageNum]
            #     tables = page.extract_tables()
            #     tablesNum = 0
            #     if pageNum == beginPage:  # 取当前页对应的表格，静态池的表格在首页时，一定是最后面那个表格
            #         tablesNum = len(tables) - 1
            #         cellValue = tables[tablesNum][0][0]
            #     for index, Row in tables[tablesNum]:
            #         if index == 0 and Row[0] == cellValue:
            #             continue
            #         execSql = execSql + "( N'" + fileName + "',"
            #         for i in columnNum:
            #             if Types[i] == 'Date':
            #                 execSql = execSql + getDate(Row[i]) + ','
            #             if Types[i] == 'Float':
            #                 execSql = execSql + getNumber(Row[i], 'Float') * money + ','
            #             if Types[i] == 'Int':
            #                 execSql = execSql + getNumber(Row[i], 'Int') * money + ','
            #         execSql = execSql.rstrip(',') + '),'
            #     execSql = execSql.rstrip(',')
            #     execSQLCmd(execSql)




