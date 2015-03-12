#!/usr/bin/env python
import sys

class event:
  UNKNOWN = 0
  CASEID = 1
  TITLE  = 2
  STEP   = 3
  EXPE_RESULT = 4
  TEST_RESULT = 5
  key_words=['Test Case ID','Title','Test Procedures','Expected Results', 'Test Results']
    
  def __init__(self,text):
    self.id = event.UNKNOWN
    self.text=text.strip()
   
    for i in range(len(event.key_words)):
      if text.startswith(event.key_words[i]):
        self.id = i+1
        self.text = text[len(event.key_words[i]):].strip()
        break

class state_base:
   def __init__(self, _reactions):
     self.reactions = _reactions

   def default_handler(self, _event):
     return self
   
   def react(self, _event):
     return self.reactions.setdefault(_event.id, self.default_handler)(_event)

class init_state(state_base):
   def __init__(self):
     state_base.__init__(self, {event.CASEID : lambda _event: name_state(_event.text)})
     
class name_state(state_base):
   case_map = {}
   def __init__(self, text):
     name_state.case_map[text] = name_state.case_map.setdefault(text, 0) + 1
     state_base.__init__(self, {event.TITLE : lambda _event : summary_state(_event.text)})
     sys.stdout.write('\n  <testcase name="%s ' %text.strip())
     #print '\n  <testcase name="%s' %text.strip()

class summary_state(state_base):
    def __init__(self, text):
     state_base.__init__(self,{event.STEP : lambda _event : step_state(_event.text)})
     print '%s">' %text.strip()
     #sys.stdout.write('%s">' %text.strip())
     print '\t<summary>%s</summary>' %text.strip()

class step_state(state_base):
   def __init__(self, text):
     state_base.__init__(self, {event.EXPE_RESULT : self.handle_result_event})
     print '\t<steps>\n\t  '
     print '\t<step>\n\t   '
     print '\t<step_number>%s</step_number>\n\t'%'1' 
     print '\t<actions><![CDATA['
     #print text
     for s in text.split('\n'): 
       print '<p>'
       print "%s" %s
       print '</p>'

   def handle_result_event(self, _event):
     print ']]></actions>'
     return result_state(_event.text)

   def default_handler(self, _event):
     if _event.text: print '<p>    %s </p>' %_event.text
     return self  

class result_state(step_state):
   def __init__(self, text):
     state_base.__init__(self, {event.TEST_RESULT : self.handle_init_event})
     print '\t<expectedresults><![CDATA['
     for s in text.split('\r'): 
       print '<p>'
       print "%s" %s
       print '</p>'

   def handle_init_event(self, _event):
     print ']]></expectedresults>\n </step> \n </steps>\n </testcase>'
     return init_state()
     
if __name__ == '__main__':
  file=open(sys.argv[1],'r')
  
  print '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<testcases xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'''

  state = init_state()
  for line in file:
       _event = event(line)
       state=state.react(_event)

  print '</testcases>'
  #check duplicated caseid
  for key in name_state.case_map:
    if name_state.case_map[key] > 1 :
      sys.stderr.write('Error:the caseid %s is duplicated %d times\n' %(key,name_state.case_map[key]))
  
