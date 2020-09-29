import pdfplumber
import pyodbc
import pandas as pd
import re,os
import calendar

path = r'C:\Users\DELL\Desktop\关于未来两年“兴晴”系列个人消费类贷款资产支持证券注册申请报告(1).pdf'
cpath = r'C:\PyCharm\untitled1\Request\pdf_download\“飞驰建普”系列微小企业贷款支持证券注册申请报告.pdf'
dbConnectionStr = 'DRIVER={SQL Server};SERVER=172.16.6.143\MSSQL;DATABASE=PortfolioManagement;UID=sa;PWD=PasswordGS2017'
def execSQLCmd(sql):
    # print(sql)
    cnxn = pyodbc.connect(dbConnectionStr)
    try:
        cursor = cnxn.cursor()
        cursor.execute(sql)
        cnxn.commit()
    except Exception as ex:
        # writeLog(str(ex))
        # writeErr(str(ex))
        print(str(ex))
    finally:
        cnxn.close()

def execSQLCmdFetchAll(sql):
    # print(sql)
    cnxn = pyodbc.connect(dbConnectionStr)
    try:
        cursor = cnxn.cursor()
        rows = cursor.execute(sql).fetchall()
        return rows
    except Exception as ex:
        #writeLog(str(ex))
        raise ex
    finally:
        cnxn.close()

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

def getBeginPage(pdfPath, CompareChar):

    pass

def insertDatabase(pdfPath, begin, end):
    with pdfplumber.open(pdfPath) as pdf:
        money = 1
        page = pdf.pages[begin]
        text = page.extract_text()
        tables = page.extract_tables()

        fileName = os.path.basename(pdfPath)
        if '万元' in text:        #确定金额的单位
            money = 10000
        elif '亿元' in text:
            money = 100000000
        for column in page.extract_tables():
            pass
        for pageNum in range(begin, end):
            Sql = "insert into PortfolioManagement.DvImport.StaticPoolData(FileNames,LoanDate,ReportDate) values "
            page = pdf.pages[pageNum]
            tables = page.extract_tables()
            tablesNum = 0
            if pageNum == begin:   #取当前页对应的表格，静态池的表格在首页时，一定是最后面那个表格
                tablesNum = len(tables)-1
            for index, Row in enumerate(tables[tablesNum]):
                if index > 0:
                    Sql = Sql + "(N'"+ fileName + "',"+ getDate(Row[1])+","+getDate(Row[4])+"),"
            Sql = Sql.rstrip(',')
            print('当前页数据已上传，当前页数为：'+ str(pageNum))
            execSQLCmd(Sql)


if __name__ == '__main__':
    sql = insertDatabase(path, 84, 120)
    #print(sql)
    #execSQLCmd(sql)
    # with pdfplumber.open(cpath) as pdf:
    #     page = pdf.pages[73]
    #     text = page.extract_text()
    #     #print(text)
    #     tables = page.extract_tables()
    #     #print(len(tables))
    #     #table是一个list 每行是一条数据
    #     for index, t in enumerate(tables):
    #         print('t.size:'+str(len(t)))
    #         # for row in t:
    #         #     #打印每行的第二列0]
    #         #     print(row[1])
    #         #得到的table是嵌套list类型，转化成DataFrame更加方便查看和分析
    #         df = pd.DataFrame(t[1:], columns=t[0])
    #         print(df)