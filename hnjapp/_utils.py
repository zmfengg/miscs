# coding=utf-8
'''
Created on 2018-04-28

@author: zmFeng
'''

import re,csv,codecs,sys

_joptn_ = re.compile("^[a-zA-Z0-9][0-9]{4,}$")

def fmtjono(jn):
    """ turn a numeric JO# into string one, mainly removing the decimal point """
    
    import numbers
    if(isinstance(jn, numbers.Number)):
        jn = "%d" % jn
    else:
        if(jn):
            jn = jn.strip()
            if(not _joptn_.match(jn)): jn = None
    return jn

''' failed
class Csvwriter(csv.writer):
    _ec_ = codecs.getincrementaldecoder(sys.getfilesystemencoding())
    def writerow(self,row):
        if(row):
            r = [self._ec_.encode(x) if isinstance(x,basestring) else x for x in row]
            super(Csvwriter,self).writerow(r) 
'''