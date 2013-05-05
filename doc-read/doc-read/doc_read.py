#coding=utf-8
#cp936

import os,win32com.client,sys
from xml.dom.minidom import Document

no_limit_flag='N/A'

class judge_infor:
    dest=""     # hit field
    text=[]     
    size=[]
    type=[]
    UL=[]
    Italic=[]
    bold=[]
    
    def __init__ (self,i):
        self.dest=i['dest']
        self.text=i['text']
        self.size=i['size']
        self.type=i['type']
        self.UL=i['UL']
        self.Italic=i['Italic']
        self.bold=i['bold']
    
    def judge_hit(self,te,s,ty,u,i,b):
        if no_limit_flag not in self.text and te.decode('utf-8').encode('cp936') not in self.text:
            return False
        if no_limit_flag not in self.size and s not in self.size :
            return False
        if no_limit_flag not in self.type and ty.decode('utf-8').encode('cp936') not in self.type:
            return False
        if no_limit_flag not in self.UL and u not in self.UL:
            return False
        if no_limit_flag not in self.Italic and i not in self.Italic:
            return False
        if no_limit_flag not in self.bold and b not in self.bold:
            return False
        return True

    def __str__(self):
        return ','.join([self.dest,str(self.text),str(self.size),str(self.type),str(self.UL),str(self.Italic),str(self.bold)])


def is_hit(te,s,ty,u,i,b):
    for i in conf:
        if i.judge_hit(te,s,ty,u,i,b):
            return i
    return None

reload(sys)
sys.setdefaultencoding('utf-8')
conf = []

conffile = open(r'conf\shenqingshu.ini')

for line in conffile.readlines():
    temp = line.split('=')

    #init judge element
    je={}
    je['dest']=temp[0]
    je['text']=[no_limit_flag]
    je['size']=[no_limit_flag]
    je['type']=[no_limit_flag]
    je['UL']=[no_limit_flag]
    je['Italic']=[no_limit_flag]
    je['bold']=[no_limit_flag]

    for des_infor in temp[1].strip().split(';'):
        elements=des_infor.split(':')
        je[elements[0]]=elements[1].split('|')
    conf.append(judge_infor(je))
conffile.close()


word = win32com.client.Dispatch('Word.Application')

doc = word.Documents.Open(r'E:\graduation_project\doc-read\aaa.doc')


xmlfile = Document()
information = xmlfile.createElement('information')


for i in range(doc.paragraphs.count):
    infor = str(doc.paragraphs[i]).strip().split('：')
    for j in range(len(infor)):
        infor[j]=infor[j].strip()
    
    fa=doc.Range(doc.paragraphs[i].Range.Start,doc.paragraphs[i].Range.Start+len(unicode(infor[0]))).font
#    print str(doc.Range(doc.paragraphs[i].Range.Start,doc.paragraphs[i].Range.Start+len(unicode(infor[0])))).decode('utf-8').encode('cp936')
    te=infor[0]
    s=str(fa.size)
    ty=fa.name
    u=str(fa.underline)
    it=str(fa.italic)
    b=str(fa.bold)

    ans=is_hit(te,s,ty,u,it,b)
    if ans!= None:
        name=ans.dest
        value=infor[1]
        inforitem = xmlfile.createElement(name)

        inforitem.setAttribute('size',s)
        inforitem.setAttribute('isBold',b)
        inforitem.setAttribute('isItalic',it)
        inforitem.setAttribute('hasUnderline',u)

        if ty !="" :
            inforitem.setAttribute('TT',ty)
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
