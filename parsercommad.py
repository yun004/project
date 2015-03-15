#!/usr/bin/env python
from pyparsing import *
import sys

lines = '''#ADD-XX:A='1',B='2';
ADD-XX:A='1',B='2';
'''

class Parameter:
    def __init__(self,name,value):
        self.name = name
        self.value = value

class MO:
    def __init__(self,name):
        self.action = ""
        self.name = name
        self.parameters = []
        self.cur_parameter = None

    def setaction(self,action):
        self.action = action
        
    def setparameter(self,name,value):
        self.cur_parameter = Parameter(name,value)

    def setparameters(self):
        self.parameters.append(self.cur_parameter)

class Parsecommand:
    def __init__(self):
        pass
        
    def parsecommand(self,line):
        result = ''
        COLON  = Suppress(":")
        SEMICOLON = Suppress(";")
        COMMA = Suppress(",")
        EQUAL = Suppress("=")
        MINUS = Suppress("-")
        DOUBLEQUOTE = Suppress('"')
        QUOTE = Suppress("'")
        COMMENT = Suppress("#")
        value = alphanums + "+-._/:*@"
        AF_IP = Keyword('A_IP')
        CF_IP = Keyword('C_IP')
        DF_IP = Keyword('D_IP')
        MF_IP = Keyword('M_IP')
        TM_IP = Keyword('T_IP')
        IPDU_IP = Keyword('I_IP')
        STOPDATE = Keyword('SDATE')
        STOPTIME = Keyword('STIME')

        linedef = Word(alphas).setResultsName('action')+ MINUS + Word(alphanums).setResultsName('name') + COLON + ZeroOrMore(Group(ZeroOrMore(COMMA)+ \
                  Word(alphas) + EQUAL + ZeroOrMore(DOUBLEQUOTE) +  ZeroOrMore(QUOTE) + (A_IP|C_IP|D_IP|M_IP|T_IP|I_IP|SDATE|STIME|Word(value)) \
                  + ZeroOrMore(DOUBLEQUOTE) + ZeroOrMore(QUOTE))).setResultsName('parameters') + SEMICOLON
        pythoncomments = COMMENT + restOfLine
        linedef.ignore(pythoncomments)
        try:
            result = linedef.parseString(line)
        except ParseException,e :
            print "Parsing failed:"
            print line
            print "%s^" % (' '*(e.col-1))
            print e.msg
        return result

if __name__ == '__main__':
    MOS =[]
    parse =Parsecommand()
    for line in lines.split('\n'):
        if line.strip() != '':
            result = parse.parsecommand(line.strip())
            if result != '' :
                mo = MO(result.name)
                mo.setaction(result.action)
                for para in result.parameters:
                    mo.setparameter(para[0],para[1])
                    mo.setparameters()
                MOS.append(mo)
            if MOS != [] :
                for i in  range(len(MOS)):
                    print("command: "+ MOS[i].action +"-"+ MOS[i].name+":")
                    for j in range(len(MOS[i].parameters)):
                        print(MOS[i].parameters[j].name+"="+MOS[i].parameters[j].value)
