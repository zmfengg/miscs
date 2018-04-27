# coding = utf-8
'''
Created on 2018-04-28
classes to read data from C1's monthly invoices
need to be able to read the 2 kinds of files: C1's original and calculator file
@author: zmFeng
'''

from collections import namedtuple
from utils import xwu
from xlwings import constants
import os,sys

class InvRdr():
    """
        read the monthly invoices from both C1 version and CC version
    """
    
    import logging
    logger = logging.getLogger(__name__)
    C1InvItem  = namedtuple("C1InvItem","source,jono,labour,settings,remarks,stones,parts")    
    
    def read(self,fldr):
        """
        perform the read action 
        @param fldr: the folder contains the invoice files
        @return: a list of C1InvItem
        """
         
        if(not os.path.exists(fldr)): return
        root = fldr + os.path.sep if fldr[len(fldr) - 1] <> os.path.sep else ""
        fns = [root + unicode(x,sys.getfilesystemencoding()) for x in \
            os.listdir(fldr) if(x.lower().find("_f") > 0)]
        if(len(fns) == 0): return
        killxw,app = xwu.app(False);wb = None
        try:
            cnsc1 = u"工单号,镶工,胚底/件,备注".split(",")
            cnscc = u"镶石费$,胚底费$,工单,参数,备注".split(",")
            for fn in fns:
                wb = app.workbooks().open(fn)
                items = list()
                for sht in wb.sheets():
                    rngs = list()
                    for s0 in cnsc1:
                        rng = xwu.find(sht, s0, lookat = constants.LookAt.xlWhole)
                        if(rng): rngs.append(rng)
                    if(len(cnsc1) == len(rngs)):
                        items.append(self._readc1(sht))
                    else:
                        for s0 in cnscc:
                            rng = xwu.find(sht, s0, lookat = constants.LookAt.xlWhole)
                            if(rng): rngs.append(rng)
                        if(len(cnsc1) == len(rngs)): items.append(self._readcalc(sht))                           
        finally:            
            if(killxw): app.quit()
    
    def _readc1(self,sht):
        """
        read c1 invoice file
        @param   sht: the sheet that is verified to be the C1 format
        @return: a list of C1InvItem with source = "C1"
        """
        pass
    
    def _readcalc(self,sht):
        """
        read cc file
        @param   sht: the sheet that is verified to be the CC format
        @return: a list of C1InvItem with source = "CC"
        """
        cns = u"镶石费$,胚底费$,工单,参数,配件,笔电,链尾,分色,电咪,其它,银夹金,石料,形状,尺寸,粒数,重量,镶法,备注".split(",")
        rng = xwu.find(sht,cns[0],lookat= constants.LookAt.xlWhole)
        x = xwu.usedrange(sht)
        rng = sht.range((rng.row,x.columns.count),(x.last_cell().row,x.last_cell().column))
        vvs = rng.value


def dirx(obj,args = None):
    if(not obj): return None
    if(not args): return dir(obj)
    s0 = set()
    xall = dir(obj)
    for s in args.split(","):
        s0.add([x for x in xall if(x.find(s) > 0)])
    return list(s0)