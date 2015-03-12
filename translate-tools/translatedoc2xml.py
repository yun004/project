import win32com 
import win32com.client
#import Dispatch, constants 
import sys
import string
import codecs
import os
import re
import time
import copy

class easyWord:
    """A utility to make it easier to get at Word.  Remembering
    to save the data is your problem, as is  error handling.
    Operates on one workbook at a time."""

    def __init__(self, filename=None):
        self.xlApp = win32com.client.Dispatch('Word.Application')
        if filename:
            self.filename = filename
            self.xldoc = self.xlApp.Documents.Open(filename)
        else:
            self.xldoc = self.xlApp.Documents.Add()
            self.filename = ''  
    def save(self, newfilename=None):
        if newfilename:
            self.filename = newfilename
            self.xldoc.SaveAs(newfilename)
        else:
            self.xldoc.Save()    

    def close(self):
        self.xldoc.Close(SaveChanges=0)
        del self.xlApp


def usage():
    print("python translatordoc2xml.py doc_file_name xml_name")
    sys.exit()

def set_step(xml,key,value):
    xml.write('''<%s>%s</%s>\n'''%(key,value,key))

cwd = os.getcwd()
sys.path.append(cwd)

if __name__ == "__main__" :
    
    KEYS = {"Test Procedures":"actions","Expected Results":"expectedresults","Reference":"reference","Objective":"objective","Pre-test Conditions":"preconditions","Priority":"importance"}
    KEYS_T = {"TC-":"testcase","Title":"summary"} 
    if len(sys.argv) < 2 :
        usage()

    doc_file = os.path.join(cwd,sys.argv[1])
    word = easyWord(doc_file)
    suite = sys.argv[2]
    xml_file = os.path.join(cwd,"%s.xml"%suite)
    xml = codecs.open(xml_file,encoding='utf-8',mode='w+')

    xml.write('''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<testcases xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">\n''')

    doc2txt = os.path.join(cwd,"%s.txt"%suite) 
    word.xldoc.SaveAs(doc2txt, FileFormat=3)
    f = open(doc2txt,'rU')

    lines = f.readlines()
    tmp_lines = []
    for line in lines :
        if line.strip().startswith("TC-"):
            tmp = re.sub('\t+',' ',line.strip())			    
            if tmp not in tmp_lines:
                tmp_lines.append(tmp)				
    print('line count: %d'%len(tmp_lines)) 
    f.close()
    i = 0
    try : 
        for table in word.xldoc.Tables:
            for row in table.Rows:
                line = row.ConvertToText(Separator='\t').Text.split('\t')
                s = line[0].strip()
                value = ''
                for v in line[1:]:
                    value = value + v + '\r'  
                    
                if s == "Reference":
                    tmp = tmp_lines[i].split(' ',1)
                    if len(tmp) == 2 : 
                        summary=re.sub('(>|<|&)','',tmp[1])
                    else:
                        summary=re.sub('(>|<|&)','',tmp)
                    name=re.sub('(>|<|&|"|\')','',tmp_lines[i].strip())
                    xml.write("""<%s name='"""%'testcase')
                    xml.write(" %s'>\n"%name)
                    set_step(xml,"summary",summary)
                    values = value.split('\r')
                    value = '<![CDATA['
                    for v in values:
                        value = value + '<p>'+ v + '</p>'
                    value = value + ']]>'
                    set_step(xml,KEYS[s],value.strip())
                    i = i + 1
                elif s == "Priority":
                    values = value.split('\r')
                    value = '<![CDATA['
                    for v in values:
                        value = value + '<p>'+ v + '</p>'
                    value = value + ']]>'
                    set_step(xml,KEYS[s],value.strip())
                elif s == "Pre-test Conditions":
                    values = value.split('\r')
                    value = '<![CDATA['
                    for v in values:
                        value = value + '<p>'+ v + '</p>'
                    value = value + ']]>'
                    set_step(xml,KEYS[s],value.strip())					
                elif s == "Test Procedures":
                    xml.write("<steps>\n")
                    xml.write("<step>\n")
                    set_step(xml,"step_number",1)
                    values = value.split('\r')
                    value = '<![CDATA['
                    for v in values:
                        value = value + '<p>'+ v + '</p>'
                    value = value + ']]>'
                    set_step(xml,KEYS[s],value.strip())
                elif s == "Expected Results":
                    values = value.split('\r')
                    value = '<![CDATA['
                    for v in values:
                        value = value + '<p>'+ v + '</p>'
                    value = value + ']]>'
                    set_step(xml,KEYS[s],value.strip())
                    xml.write("</step>\n")
                    xml.write("</steps>\n") 
                    xml.write("</testcase>\n")

        xml.write("</testcases>\n")
        xml.close()
        word.close()
    except Exception as e:
        xml.write("</testcases>\n")
        print(e)
        xml.close()
        word.close()


