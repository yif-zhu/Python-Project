import os
import os.path

def createTask(pyFile, funName, xml):
    try:
        import time
        pyName = pyFile.split('\\')[-1][0:-3]
        dir_path = os.path.dirname(os.path.abspath(__file__)) + '\\Task\\'
        now = time.strftime('%Y%m%d_%H%M%S', time.localtime())  #将指定格式的当前时间以字符串输出
        suffix = ".py"
        newfile= dir_path + now + suffix  #新建任务的绝对路径
        if not os.path.exists(newfile):
            Task = open(newfile, 'w+')
            Task.write('import Action.' + pyName + ' as PyPath \n')
            Task.write('xml = r"' + xml + '" \n')
            Task.write('dateId = r"' + now + '" \n')
            Task.write('PyPath.' + funName + '(xml, dateId) \n')
            Task.close()
            print(newfile + " created.")
        else:
            print(newfile + " already existed.")
            Task = open(newfile, 'w+')
            Task.write('import Action.' + pyName + ' as PyPath \n')
            Task.write('xml = r"' + xml + '" \n')
            Task.write('dateId = r"' + now + '" \n')
            Task.write('PyPath.' + funName + '(xml, dateId) \n')
            Task.close()
        with open(newfile, encoding="UTF-8") as f:
            exec(f.read())
        pwd = os.getcwd()
        mappingPath = os.path.join(os.path.abspath(os.path.dirname(pwd) + os.path.sep + "."),
                                   "FileTranslator", 'MappingXml', '{0}.xml'.format(now))
        return mappingPath
    except ValueError as e:
        print(e.args)
        return e.args




pyName = 'C:\\PyCharm\\untitled1\\Action\\example.py'
funName = 'main'
xml = 'C:\\PyCharm\\xml\\444.xml'

pyName1 = 'C:\\PyCharm\\untitled1\\Action\\NonFileFormatCheck.py'
funName1 = 'main'
xml1 = 'C:\\PyCharm\\pdf-docx\\FileFormatCheck_FTP.xml'

new = createTask(pyName, funName, xml)
#new = createTask(pyName1, funName1, xml1)
print(new)