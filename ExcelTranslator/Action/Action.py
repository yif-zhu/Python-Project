import os
import os.path
import xml.etree.ElementTree as XE


class Action:
    configfilepath = "./Config/Actions.xml"

    def __init__(self, actionCode, kwargs):
        self.actionCode = actionCode
        tree = XE.parse(self.configfilepath)
        root = tree.getroot()
        actionConfig = root.find("Action[@ActionCode = '{0}']".format(actionCode))
        for variable in actionConfig.iter(tag="Variable"):
            parameterName = variable.attrib["ParameterName"]
            if parameterName in kwargs:
                variable.set("Value", kwargs[parameterName])
        tree.write(self.configfilepath, encoding='utf-8', xml_declaration=True)
        self.createAction()

    def createAction(self):
        tree = XE.parse(self.configfilepath)
        root = tree.getroot()
        actionConfig = root.find("Action[@ActionCode = '{0}']".format(self.actionCode))
        self.xmlFilePath = actionConfig.find("*[@VariableName = '{0}']".format("XmlFile")).attrib["Value"]
        self.pyFile = actionConfig.find("*[@VariableName = '{0}']".format("PythonFile")).attrib["Value"]
        self.methodName = actionConfig.find("*[@VariableName = '{0}']".format("MethodName")).attrib["Value"]

        tree = XE.parse(self.xmlFilePath)
        root = tree.getroot()
        for variable in actionConfig.findall("*[@IsParameter = 'True']"):
            root.set(variable.attrib["VariableName"], variable.attrib["Value"])
        tree.write(self.xmlFilePath, encoding='utf-8', xml_declaration=True)

    def runAciton(self):
        pyFile = self.pyFile
        funName = self.methodName
        xml = self.xmlFilePath
        try:
            pyName = pyFile.split('\\')[-1][0:-3]
            import time
            dir_path = os.path.dirname(os.path.abspath(__file__)) + '\\Task\\'
            now = time.strftime('%Y%m%d_%H%M%S', time.localtime())  # 将指定格式的当前时间以字符串输出
            suffix = ".py"
            newfile = dir_path + now + suffix  # 新建任务的绝对路径
            if not os.path.exists(dir_path):
                os.mkdir(dir_path)
            if not os.path.exists(newfile):
                Task = open(newfile, 'w+', encoding="utf-8")
                Task.write('import FileTranslator.' + pyName + ' as PyPath \n')
                Task.write('xml = r"' + xml + '" \n')
                Task.write('dateId = r"' + now + '" \n')
                Task.write('PyPath.' + funName + '(xml, dateId) \n')
                Task.close()
                print(newfile + " created.")
            else:
                print(newfile + " already existed.")
                Task = open(newfile, 'w+')
                Task.write('import FileTranslator.' + pyName + ' as PyPath \n')
                Task.write('xml = r"' + xml + '" \n')
                Task.write('dateId = r"' + now + '" \n')
                Task.write('PyPath.' + funName + '(xml, dateId) \n')
                Task.close()
            with open(newfile, encoding="UTF-8") as f:
                exec(f.read())
            pwd = os.getcwd()
            mappingPath = os.path.join(pwd, "FileTranslator", 'MappingXml', '{0}.xml'.format(now))
            return mappingPath
        except ValueError as e:
            print(e.args)
            return e.args
if __name__ == "__main__":
    dir = r"C:\Users\HHH\Documents\HHH\橙易2016年第一期持证抵押贷款资产支持证券 - 副本"

    parameters = dict()
    parameters["inputFile"] = dir + r"\橙易2016年第一期持证抵押贷款证券化信托2016年5月受托机构月度报告（第3期）.xlsx"
    parameters["XmlFile"] =  dir + r"\222.xml"
    parameters["outputFile"] = r"C:\Users\HHH\Documents\HHH\result"
    parameters["templateFile"] = r"C:\Users\HHH\Documents\HHH\受托报告导入模板-新.xlsx"

    action = Action("FillData", parameters)
    print(action.runAciton())

