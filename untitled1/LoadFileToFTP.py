# _*_ coding:utf-8 _*_

import sys
import os
import os.path
import datetime
import pyodbc
from ftplib import error_perm, FTP
from posixpath import dirname

dbConnectionStr = 'DRIVER={SQL Server};SERVER=192.168.1.149\MSSQL;DATABASE=TaskCollection;UID=sa;PWD=PasswordGS15'
##### Helper Methods #####
def writeLog(msg):
    if not os.path.exists(logtxtFilePath):
        f = open(logtxtFilePath, "w")
    print(msg)
    with open(logtxtFilePath, "a") as f:
        ts = datetime.datetime.now().strftime('[%H:%M:%S]')
        f.write('{0}:  {1}\n'.format(ts, msg))

def writeErr(msg):
    if not os.path.exists(errtxtFilePath):
        f = open(errtxtFilePath, "w")
    # print(msg)
    with open(errtxtFilePath, "a") as f:
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

def execSQLCmdFetchOne(sql):
    # print(sql)
    cnxn = pyodbc.connect(dbConnectionStr)
    try:
        cursor = cnxn.cursor()
        row = cursor.execute(sql).fetchone()
        return row
    except Exception as ex:
        writeLog(str(ex))
        raise ex
    finally:
        cnxn.close()

def writeToTable(orderId, OrderFileType, fileName):
    dt = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    sql = "insert into [TaskCollection].[dbo].[OrderAttachment] values('{0}','{1}',N'{2}','{3}')".format(orderId, OrderFileType, fileName, dt)
    execSQLCmd(sql)

def getFtpPath(LinkUrl):
    sql = "SELECT [FolderPath] + '/TargetFile' as FtpPath FROM [TaskCollection].[dbo].[OrderFolder] WHERE LinkURL = '{0}'".format(LinkUrl)
    try:
        FtpPath = execSQLCmdFetchOne(sql).FtpPath
        return FtpPath
    except Exception as ex:
        writeErr(str(ex))
        return 0

def getOrderID(LinkUrl):
    sql = "SELECT OrderID  FROM [TaskCollection].[dbo].[OrderFolder] WHERE LinkURL = '{0}'".format(LinkUrl)
    try:
        OrderID = execSQLCmdFetchOne(sql).OrderID
        return OrderID
    except Exception as ex:
        writeErr(str(ex))
        return 0

def ftp_makedirs_cwd(ftp, path, first_call=True):
    """设置“FTP”中给出的FTP连接的当前目录
参数(如ftplib)。)，不存在创建所有父目录
。ftplib对象必须已经连接并登录
    """
    try:
        ftp.cwd(path)
    except error_perm:
        ftp_makedirs_cwd(ftp, dirname(path), False)
        ftp.mkd(path)
        if first_call:
            ftp.cwd(path)

def logFTP(host, user, passwd, ftpPath, localPath, filename):
    ftp = FTP(host)
    try:
        ftp.encoding = 'GB2312'
        ftp.login(user, passwd)
        ftp_makedirs_cwd(ftp, ftpPath, True)
        ftp.cwd(ftpPath)

        fp = open(localPath, 'rb')  # 处理之后的文件上传
        ftp.storbinary('STOR ' + filename, fp)
        fp.close()

    finally:
        ftp.quit()


scriptFolder = r"C:\Users\DELL\Desktop\excel交付\excel交付"
errtxtFilePath = os.path.join(scriptFolder, 'Errors', 'Log_ExtractInsert_{0}.txt'.format(datetime.datetime.now().strftime('%m-%d %H%M%S')))
logtxtFilePath = os.path.join(scriptFolder, 'Logs', 'Log_ExtractInsert_{0}.txt'.format(datetime.datetime.now().strftime('%m-%d %H%M%S')))

for parent, dirnames, filenames in os.walk(scriptFolder, followlinks=True):
    for filename in filenames:
        webPath = 'http://quickdealservice.com/Community/#/question'
        file = filename.split('_')[0]
        LinkUrl = webPath + '/' + file
        localPath = os.path.join(parent, filename)
        ftpPath = getFtpPath(LinkUrl)
        orderId = getOrderID(LinkUrl)
        print('LinkUrl:'+ LinkUrl)
        print('localPath:'+ localPath)
        print('ftpPath:' + ftpPath)

        # LinkUrl = ''
        # localPath=r'C:\Users\DELL\Desktop\excel交付\excel交付\1727_178.xlsx'
        #ftpPath = '/Test/006B85F3-B4F9-4C88-ABDB-781BD4A635B125/TargetFile'
        host = '192.168.1.211'
        user = 'gsuser'
        passwd = 'Password01'

        writeToTable(orderId, 'TargetFile', filename)
        #logFTP(host, user, passwd, ftpPath, localPath, filename) #上传FTP文件




