import urllib.request
import time
import json
import re
import os
import requests
#13位时间戳转换时间格式
def getDateFormat(timeInt):
    timeStamp = float(timeInt / 1000)
    timeArray = time.localtime(timeStamp)
    otherStyleTime = time.strftime("%Y-%m-%d %H:%M:%S", timeArray)
    return otherStyleTime
#获取文件
def getFile(url,fileName):
    #file_name = url.split('/')[-1]
    u = urllib.request.urlopen(url)
    f = open(fileName, 'wb')

    block_sz = 8192
    while True:
        buffer = u.read(block_sz)
        if not buffer:
            break

        f.write(buffer)
    f.close()
    print("Sucessful to download" + " " + fileName)

root_url = 'http://www.chinamoney.com.cn/chinese/qwjsn/?searchValue=%25E6%25B3%25A8%25E5%2586%258C%25E7%2594%25B3%25E8%25AF%25B7%25E6%258A%25A5%25E5%2591%258A'
#下载地址中查询出来的Pdf部分
raw_url = 'http://www.chinamoney.com.cn'

data ={
    'sort':"relevance",
    'text':"注册申请报告",
    'date':"all",
    'field': "title",
    'pageIndex': 1,
    'pageSize': 500
}

headers={
    'User-Agent':"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/78.0.3904.108 Safari/537.36",
    'Referer':"http://www.chinamoney.com.cn/chinese/qwjsn/?searchValue=%25E6%25B3%25A8%25E5%2586%258C%25E7%2594%25B3%25E8%25AF%25B7%25E6%258A%25A5%25E5%2591%258A",
    'Accept':'*/*',
    'Accept-Encoding': "gzip, deflate",
    'Accept-Language': "zh-CN,zh;q=0.9"
}
requestUrl = 'http://www.chinamoney.com.cn/ses/rest/cm-u-notice-ses-cn/query?sort=relevance&text=%E6%B3%A8%E5%86%8C%E7%94%B3%E8%AF%B7%E6%8A%A5%E5%91%8A&date=all&field=title&start=&end=&pageIndex=1&pageSize=500&public=false&infoLevel=YrsR7YRg8k2F/oR/Go0wW5lCaidTipCNqRZjDIs4h5NLrjMyYk3XBcVGjrOMqUYIwEA%2BvHLfrMCA%0ArG79NjEARD5fVkLrdv9VSGKa/9RUwGGmpsv89DuZcDE09s7wFNe7WdmMu5rhVhyAW%2BwcNIvfCGQu%0A7Kh//4zJrgc81lQv9v0=%0A&sign=Q/d8solfMh3GOoMI5WmGUaZA1ukiCpO5sMwap9ByMZnt4tsJZeSkX6Wq1v3lRrKsnQLcWdAPun00%0ALsYa5AtcTZpCs2CvuKf8xTKL5JKkAphGIIEbpsADAhjeg2dCZIBVMUOFd2LaiLvRLJLML9AfJTc/%0AI44XV2MvFkyyEBuTLsA=%0A&channelIdStr=2496,2556,2632,2663,2589,2850,3300,&nodeLevel=1'
#urls='https://www.lagou.com/jobs/list_python/p-city_252?px=default&gx=%E5%85%A8%E8%81%8C&gj=&xl=%E6%9C%AC%E7%A7%91&isSchoolJob=1#filterBox'
s = requests.Session()
s.get(raw_url, headers=headers, timeout=3)
cookie = s.cookies
response = s.post(requestUrl,data=data,headers=headers, cookies=cookie,timeout=5)
#print(response.text)
os.mkdir('pdf_download')
os.chdir(os.path.join(os.getcwd(), 'pdf_download'))
for Item in json.loads(response.text)['data']['result']['pageItems']:
    FileName = Item['title']
    links = Item['paths']
    timeInt = Item['releaseDate']
    if len(links) == 0:
        continue
    url = raw_url + links[0]    #获取完整路径
    PublishDate = getDateFormat(timeInt) #时间戳转日期格式
    strinfo = re.compile('<font color=\'red\'>|</font>')
    FullName = strinfo.sub('', FileName)+'.pdf'
    getFile(url, FullName)
    #print("{0}, Url: {1} , 发布日期：{2}".format(FullName, url, PublishDate))

