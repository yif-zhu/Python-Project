import urllib.request

response = urllib.request.urlopen('http://www.chinamoney.com.cn/chinese/qwjsn/?searchValue=%25E6%25B3%25A8%25E5%2586%258C%25E7%2594%25B3%25E8%25AF%25B7%25E6%258A%25A5%25E5%2591%258A')
print(response.read().decode('utf-8'))

# import urllib.parse
# import urllib.request
# from http import cookiejar
# url = "http://quickdealservice.com/DealViewer/#/Product/Detail/8942/ProductInformation"
# response1 = urllib.request.urlopen(url)
# print("第一种方法")
# #获取状态码，200表示成功
# print(response1.getcode())
# #获取网页内容的长度
# print(str(response1.read()))
# print(len(response1.read()))
# print("第二种方法")
# request = urllib.request.Request(url)
# #模拟Mozilla浏览器进行爬虫
# request.add_header("user-agent","Mozilla/5.0")
# response2 = urllib.request.urlopen(request)
# print(response2.getcode())
# print(len(response2.read()))
# print("第三种方法")
# cookie = cookiejar.CookieJar()
# #加入urllib2处理cookie的能力#
# opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(cookie))
# urllib.request.install_opener(opener)
# response3 = urllib.request.urlopen(url)
# print(response3.getcode())
# print(len(response3.read()))
# print(cookie)