

import os
import xml.etree.ElementTree as XETree

def indent(elem, level=0):   #给xml增加换行符
    i = "\n" + level * "\t"
    if len(elem):
        if not elem.text or not elem.text.strip():
            elem.text = i + "\t"
        if not elem.tail or not elem.tail.strip():
            elem.tail = i
        for elem in elem:
            indent(elem, level + 1)
        if not elem.tail or not elem.tail.strip():
            elem.tail = i
    else:
        if level and (not elem.tail or not elem.tail.strip()):
            elem.tail = i

input = "C:\\PyCharm\\pdf-docx\\source\\task1"
output= "C:\\PyCharm\\pdf-docx\\dest\\task1"

MapPath = "C:\\PyCharm\\untitled1\\Action\\MappingXml\\20190716_175704_1.xml"
tree = XETree.parse(MapPath)
root = tree.getroot()
node = root.findall("Mapping")[0]
MapPath1 = XETree.Element('map')  # 创建节点,单个文件的mapping文件
MapPath1.set("sourcecell", '')
MapPath1.set("destCell", '')
MapPath1.set("comment", '')
node.append(MapPath1)
indent(node)
tree.write(MapPath, encoding='utf-8', xml_declaration=True)
