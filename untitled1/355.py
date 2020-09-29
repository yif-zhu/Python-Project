# encoding:utf8
import sys

import pymssql
import requests
import json


def GetData():
    url = 'http://www.chinamoney.com.cn/r/cms/www/chinamoney/data/currency/bk-lpr.json'
    res = requests.get(url).text
    data = json.loads(res)
    pub_date = data['data']['showDateCN'].split()[0]
    val = []
    for i in data['records']:
        if i['termCode'] == '1Y':
            val.append(('短期贷款', '1Years(LPR)', '一年(LPR)', float(i['shibor']), 'loan', pub_date
                        ))
        elif i['termCode'] == '5Y':
            val.append(('中长期贷款', 'MoreThan5Years(LPR)', '五年以上(LPR)', float(i['shibor']), 'loan', pub_date
                        ))
    return val


def InsertData(val):
    sql = "INSERT INTO TrustManagement.tblPBCRates" \
          "( Category, SubCategoryCode,SubCategory, BaseRate, type, PubDate) " \
          "VALUES ( %s, %s, %s, %d, %s, %s)"
    cursor.executemany(sql, val)
    conn.commit()


if __name__ == '__main__':
    server = "172.16.6.143\mssql"
    user = "sa"
    password = "PasswordGS2017"
    database = "PortfolioManagement"
    conn = pymssql.connect(server, user, password, database, charset='utf8')
    cursor = conn.cursor()

    # 获取插入数据
    val = GetData()

    InsertData(val)
    conn.close()
