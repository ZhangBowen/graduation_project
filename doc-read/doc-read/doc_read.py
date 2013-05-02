#coding=utf-8
#cp936
import os,win32com.client,sys
from xml.dom.minidom import Document
reload(sys)
sys.setdefaultencoding('utf-8')
conf = {}

conffile = open(r'E:\doc read\doc-read\doc-read\conf\shenqingshu.ini')
for line in conffile.readlines():
    temp = line.split('=')
    value=temp[0]
    for key in temp[1].strip().split('|'):
        conf[key]=value
conffile.close()

word = win32com.client.Dispatch('Word.Application')

doc = word.Documents.Open(r'E:\aaa.doc')

xmlfile = Document()
information = xmlfile.createElement('information')

for i in range(doc.paragraphs.count):
    infor = str(doc.paragraphs[i]).strip().split('：')
    for j in range(len(infor)):
        infor[j]=infor[j].strip()
    if infor[0].decode('utf-8').encode('cp936') in conf.keys():
        name=conf[infor[0].decode('utf-8').encode('cp936')]
        value=infor[1]
        inforitem = xmlfile.createElement(name)

        fa=doc.paragraphs[i].range.font
        inforitem.setAttribute('size',str(fa.size))
        inforitem.setAttribute('isBold',str(fa.bold))
        inforitem.setAttribute('isItalic',str(fa.italic))
        inforitem.setAttribute('hasUnderline',str(fa.underline))
        if fa.name !="" :
            inforitem.setAttribute('TT',fa.name)
        else :
            inforitem.setAttribute('TT',"MIXTT")
        if value != "" :
            textnode = xmlfile.createTextNode(value)
        else:
            textnode = xmlfile.createTextNode("NULL")
        inforitem.appendChild(textnode)
        information.appendChild(inforitem)

xmlfile.appendChild(information)
f=open('output.xml','w')
xmlfile.writexml(f, "\t", "\t", "\n", "utf-8")
f.close()

#doc.Close()
#word.Quit()
os.system('pause')