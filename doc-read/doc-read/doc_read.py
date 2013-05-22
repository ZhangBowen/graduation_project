#coding=utf-8
#cp936

import os
import win32com.client
import sys
import getopt
from xml.dom.minidom import Document
import re


####################################################################################


no_limit_flag='N/A'
INF = 1e9
tab_const=18


####################################################################################



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

####################################################################################

class XMLE:
    name=''
    data=None
    def __init__ (self,n,d):
        self.name=n
        self.data=d


####################################################################################

def is_hit(te,s,ty,u,i,b):
    for i in conf:
        if i.judge_hit(te,s,ty,u,i,b):
            return i
    return None

reload(sys)
sys.setdefaultencoding('utf-8')
conf = []
elements=[]


####################################################################################


def get_deep(str):
    ans=0
    for char in str:
        if char =='\t' :
            ans+=1
        else:
            return ans


####################################################################################

def serach_E(n):
    for i in elements:
        if i.name==n:
            if i.data.hasChildNodes()==False:
                xmlfile = Document()
                temp1=xmlfile.createElementNS('xs','complexType')
                temp1.setAttribute('mixed','true')
                i.data.appendChild(temp1)
                temp2=xmlfile.createElementNS('xs','sequence')
                temp1.appendChild(temp2)
            return i.data.getElementsByTagName('complexType')[0].getElementsByTagName('sequence')[0]
    return None

####################################################################################

def serach_E_R(n):
    for i in elements:
        if i.name==n:
            if i.data.hasChildNodes()==False:
                xmlfile = Document()
                temp1=xmlfile.createElementNS('xs','complexType')
                temp1.setAttribute('mixed','true')
                i.data.appendChild(temp1)
                temp2=xmlfile.createElementNS('xs','sequence')
                temp1.appendChild(temp2)
                #temp3=xmlfile.createElementNS('xs','element')
                #temp3.setAttribute('name','text')
                #temp2.appendChild(temp3)
            return i.data.getElementsByTagName('complexType')[0].getElementsByTagName('sequence')[0]
    return None



####################################################################################

pattern1 = re.compile(r'^<(/xs:sequence|xs:complexType|xs:sequence|/xs:complexType)') 
pattern2 = re.compile(r'^</')

def is_useless(str):
    if pattern1.match(str.strip()):
        return True
    return False

def get_out(str):
    if pattern2.match(str.strip()):
        return True
    return False

####################################################################################

def get_XSD_r(conf):
    xmlfile = Document()
    
    information = xmlfile.createElementNS('xs','schema')
    information.setAttributeNS('xmlns','xs','http://www.w3.org/2001/XMLSchema')
    information.setAttribute("targetNamespace",'http://www.w3school.com.cn')
    information.setAttribute('xmlns','http://www.w3school.com.cn')
    xmlfile.appendChild(information)
    conffile=open(conf)
    father=None

    while len(elements):
        elements.pop()
    for line in conffile.readlines():
        temp=line.split('=')
        for j in range(len(temp)):
            temp[j]=temp[j].strip()
        te = xmlfile.createElementNS('xs','element')
        te.setAttribute('name',temp[0])
        elements.append(XMLE(temp[0],te))
        if len(temp)>1:
            temp[0]=temp[1].split(';')
            for j in temp[0]:
                j=j.strip()
                if j == '':
                    continue
                j=j.split(':')
                if j[0]=='type':
                    if j[1]=='str':
                        te.setAttribute('type','xs:string')
                    elif j[1]=='int':
                        te.setAttribute('type','xs:positiveInteger')
                elif j[0]=='fix':
                    te.setAttribute('fixed',j[1])
                elif j[0]=='word'or j[0]=='contain' or j[0]=='maxo' or j[0]=='mino':
                    te.setAttribute(j[0],j[1])
                elif j[0]=='fat':
                    father=serach_E(j[1])
                    if father == None:
                        te.setAttribute('error','no father node named :%s' % j[1])
            
            if father == None:
                information.appendChild(te)
            else:
                father.appendChild(te)
                father=None
    conffile.close()
    
    temp = conf.split('\\')            
    f=open('%s.xsdr' % temp[-1].split('.')[0] ,'w')
    xmlfile.writexml(f, "\t", "\t", "\n", "gbk")
    f.close()


####################################################################################


def get_XSD(conf):
    xmlfile = Document()
    
    schema = xmlfile.createElementNS('xs','schema')
    schema.setAttributeNS('xmlns','xs','http://www.w3.org/2001/XMLSchema')
    schema.setAttribute("targetNamespace",'http://www.w3school.com.cn')
    schema.setAttribute('xmlns','http://www.w3school.com.cn')
    schema.setAttribute('elementFormDefault','qualified')
    xmlfile.appendChild(schema)

    root=xmlfile.createElementNS('xs','element')
    root.setAttribute('name','root ')
    complex = xmlfile.createElementNS('xs','complexType')
    schema.appendChild(root)

    root.appendChild(complex)

    information = xmlfile.createElementNS('xs','sequence')
    complex.appendChild(information)

    conffile=open(conf)
    father=None

    while len(elements):
        elements.pop()
        
    for line in conffile.readlines():
        temp=line.split('=')
        for j in range(len(temp)):
            temp[j]=temp[j].strip()
        te = xmlfile.createElementNS('xs','element')
        te.setAttribute('name',temp[0])
        elements.append(XMLE(temp[0],te))
        if len(temp)>1:
            temp[0]=temp[1].split(';')
            for j in temp[0]:
                j=j.strip()
                if j == '':
                    continue
                j=j.split(':')
                if j[0]=='type':
                    if j[1]=='str':
                        te.setAttribute('type','xs:string')
                    elif j[1]=='int':
                        te.setAttribute('type','xs:positiveInteger')
                elif j[0]=='fix':
                    te.setAttribute('fixed',j[1])
                elif j[0]=='maxo' :
                    if j[1]!='INF':
                        te.setAttribute("maxOccurs",j[1])
                    else:
                        te.setAttribute("maxOccurs","unbounded")
                elif j[0]=='mino':
                    te.setAttribute("maxOccurs",j[1])
                elif j[0]=='fat':
                    father=serach_E_R(j[1])
                    if father == None:
                        te.setAttribute('error','no father node named :%s' % j[1])
            
            if father == None:
                information.appendChild(te)
            else:
                father.appendChild(te)
                father=None
    conffile.close()
    temp = conf.split('\\')            
    f=open('%s.xsd' % temp[-1].split('.')[0] ,'w')
    xmlfile.writexml(f, "\t", "\t", "\n", "gbk")
    f.close()

####################################################################################


def word_to_xml(conf,wordpath,output=None):

    word = win32com.client.Dispatch('Word.Application')
    doc = word.Documents.Open(wordpath)
    
    xmlfile = Document()
    docname=wordpath.split('\\')[-1]
    root = xmlfile.createElement("root")
    xmlfile.appendChild(root)

    conffile = open (conf)
    contain=[]
    father =[]
    deep=2
    now_root=root
    combo=0
    last=None
    text=None
    tab=0
    father.append(None)

    conffile.readline()
    conffile.readline()         #忽略XSD文档前两行
    line = conffile.readline()

    para = doc.paragraphs[0]
    while para !=None and line !='':
        if str(para) == '\r':
            while len(contain)> 0 and (contain[-1] == 'NS' or contain[-1]=='tab'):
                contain.pop()
                line=conffile.readline()
                while is_useless(line):
                    line=conffile.readline()
                if get_out(line):
                    now_root=father.pop()
                    line=conffile.readline()
                tab=0
                combo=0
            para=para.Next()
            continue

        while is_useless(line):
            line=conffile.readline()
            combo=0

        while get_out(line):
            now_root=father.pop()
            line=conffile.readline()
            combo=0

        if line == '':
            break

        if len(contain)>0 :
            if contain[-1]=='tab' and para.LeftIndent+para.FirstLineIndent < tab-2:
                contain.pop()
                tab-=tab_const
                line=conffile.readline()
                combo=0
                continue

        d={}
        d['type']=None
        d['name']=None
        d['fixed']=None
        d['word']=None
        d['contain']=None
        d['maxo']=1
        d['mino']=1

        infor = line.strip('\t\n</>').split(' ')
        
        for i in infor:
            j = i.split('=')
            if j[0] in d.keys():
                d[j[0]]=j[1].strip('"')
        
        
        te = xmlfile.createElement(d['name'])
        
        text=str(para).decode('utf-8').encode('gbk').strip('\n\r')

        now_deep = get_deep(line)
        if now_deep > deep:
            father.append(now_root)
            now_root=last

        deep=now_deep

        

        combo+=1

        if d['word']:
            
            tn=xmlfile.createTextNode(text.split(' ')[int(d['word'].strip('"'))-1])
            te.appendChild(tn)
            now_root.appendChild(te)
            last=te
            line=conffile.readline()
            combo=0;
            continue



        tn=xmlfile.createTextNode(text)
        te.appendChild(tn)

        now_root.appendChild(te)
        last=te
        para=para.Next()


        
        if d['contain']!=None:
            contain.append(d['contain'])
            if d['contain']=='tab':
                tab+=tab_const

        if combo >= d['maxo']:
            line=conffile.readline()
            combo=0
            continue    
    if output==None:
        f=open('%s.xml' % docname.split('.')[0] ,'w')
    else:
        f=open(output,'w')
    xmlfile.writexml(f, "\t", "\t", "\n", "gbk")
    f.close()

    doc.Close()
    word.Quit()


####################################################################################


def get_r_xml(confp,wordpath,output=None):
    print confp,wordpath

    conffile = open(confp)

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
            element=des_infor.split(':')
            je[element[0]]=element[1].split('|')
        conf.append(judge_infor(je))
    conffile.close()

    xmlfile = Document()
    information = xmlfile.createElement('root')
    xmlfile.appendChild(information)

    word = win32com.client.Dispatch('Word.Application')
    doc = word.Documents.Open(wordpath)

    for i in range(doc.paragraphs.count):
        infor = str(doc.paragraphs[i]).strip().split('：')
        for j in range(len(infor)):
            infor[j]=infor[j].strip()
    
        fa=doc.Range(doc.paragraphs[i].Range.Start,doc.paragraphs[i].Range.Start+len(unicode(infor[0]))).font
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

    if output ==None:
        f=open('output.xml','w')
    else:
        f=open(output,'w')
    xmlfile.writexml(f, "\t", "\t", "\n", "utf-8")
    f.close()

    doc.Close()
    word.Quit()

####################################################################################

def usage():
    print '''
    -c 检验模式。检验一个Word文档是否符合特定Schema格式要求。需输入待检验Word文档路径，配置文件路径或者已有的xsdr路径。
    -r 提取模式。根据一个配置文件从一个Word文档中提取出感兴趣的文档内容。需输入待检验Word文档路径，配置文件路径。
    -h 帮助信息。
    -i 指定配置文件路径。
    -w 指定word文档路径。
    -x 指定XML Schema 路径。
    -o 指定输出的XML文件路径。
    -a 提取用用户模板语法说明。
    -b 验证用用户模板语法说明。
    '''.decode('utf-8').encode('gbk')
    
def usage1():
    print ''' 
    配置项以行为单位。
    开头是命中后生成的XML element 的名称，后跟‘=’。‘=’后面是对于匹配命中条件的描述。
    命中条件包括:
    text 规定命中文本的文本内容。
    size 规定命中文本的文本字号。填写以磅为单位的字号。
    type 规定命中文本的文本字体。
    UL   规定命中文本是否含有下划线。-1代表含下划线，其他内容代表不含下划线。
    Iitalic 规定命中文本是否是斜体字。-1代表是斜体字，其他内容代表不是斜体字。
    bold   规定命中文本是否是粗体字。-1 代表是粗体字，其他内容代表不是粗体字。
    
    如果不指定字段则代表对该字段没有任何要求。
    各个命中条件之中用’;’相隔。
    
    例如：
    SHENQINGDAIMA=text:申请代码;type:宋体;size:10.5
    
    对于同一字段可以填写多个条件，各个条件之间是或关系，用‘|’连接各个条件。
    
    例如：
    SHENQINGZHE=text:申请者|申 请 者
    '''.decode('utf-8').encode('gbk')

def usage2():
    print '''
    配置项以行为单位。
	开头是内存中读取到配置项后生成的XML element 的名称，后跟‘=’。‘=’后面是对于生成节点的包含关系和属性的描述。
	处于同一层次的节点需要保证配置项与word文档的书写顺序相匹配，否则会得出错误的输出结果。
	描述的关键词包括：
	fix 规定word转化生成的XML文本必须严格符合fix后的文本。默认无限制。
	type规定word转化生成的XML文本必须严格符合type后的类型。后面所填的的类型必须是Schema官方类型中的一个。默认无限制。
	maxo 规定此字段至多出现次数。填数字代表确切次数，如果不对至多出现次数做限制，那么填写INF。默认为1。
	mino规定此字段至少出现次数。填数字代表确切次数，如果不对至少出现次数做限制，那么填写INF。默认为1。
	fat  此配置项的父节点
	contain表示此XML element将包含多个word段落。后面跟随包含的条件。
    contain: tab表示此XML element 将包含下面所有含有相同行首缩进的段落。
    contain：NS表示此XML element 将包含下面所有的段落，直到出现一个空行。
    各个描述信息之间用‘;’相隔。
    '''.decode('utf-8').encode('gbk')

####################################################################################
if __name__ == '__main__':
    mode = None
    conffile =None
    wordfile = None
    xmlfile = None
    outputfile = None
    opts,args = getopt.getopt(sys.argv[1:],'abhcri:w:x:o:')
    for op,value in opts:
        if op == '-c':
            mode =1
        elif op == '-r':
            mode =2
        elif op == '-i':
            conffile = value
        elif op == '-x':
            xmlfile = value
        elif op == '-w':
            wordfile = value
        elif op == '-o':
            outputfile =value
        elif op == '-h':
            usage()
            sys.exit(0)
        elif op == '-a':
            usage1()
            sys.exit(0)
        elif op == '-b':
            usage2()
            sys.exit(0)

    if mode ==None:
        usage()
        sys.exit(1)
    elif mode ==1:
        if conffile and wordfile :
            get_XSD_r(conffile)
            get_XSD(conffile)
            word_to_xml('%s.xsdr' % conffile.split('\\')[-1].split('.')[0],wordfile,outputfile)
        elif xmlfile and wordfile :
            word_to_xml(xmlfile,wordfile,outputfile)
        else:
            usage()
            sys.exit(2)
    elif mode ==2:
        if conffile and wordfile:
            get_r_xml(conffile,wordfile,outputfile)
        else:
            usage()
            sys.exit(3)

####################################################################################
