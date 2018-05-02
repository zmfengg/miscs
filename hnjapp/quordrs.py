# coding=utf-8
'''
Created on Apr 23, 2018
module try to read data from quo catalog
@author: zmFeng
'''

from collections import namedtuple
from models.utils import JOElement
import logging
from os import path
import re

def readagq(fn):
    """
    read AGQ reference prices
    @param fn: the file to read data from  
    """
    
    if(not path.exists(fn)): return
    
    from utils import xwu
    import numbers
    
    kxl,app = xwu.app(False)    
    wb = app.books.open(fn)
    try:
        rng = xwu.usedrange(wb.sheets(r'Running lines'))
        cidxs = list()
        vals = rng.value
        Point = namedtuple("Point","x,y")
        val = "styno,idx,type,mps,uprice,remarks"
        Item =namedtuple("Item",val)
        pt = re.compile("cost\s*:", re.IGNORECASE)        
        items = list()
        items.append(Item._make(val.split(",")))        
        hts = list()
                
        ccnt = 3
        for ridx in range(len(vals)):
            tr = vals[ridx]
            if(len(cidxs) < ccnt):
                for cidx in range(len(tr)):
                    val = tr[cidx]
                    if(isinstance(val, basestring) and pt.match(val)):
                        if(len(cidxs) < ccnt and not (cidx in cidxs)):
                            cidxs.append(cidx)
            if(len(cidxs) < ccnt): continue
            val = tr[cidxs[0]] if isinstance(tr[cidxs[0]], basestring) else None
            if(not(val and  pt.match(val))): continue
            for ii in range(0,ccnt):
                hts.append(Point(ridx,cidxs[ii]))
        
        #hardcode, 4 prices, in the 16th columns
        mpss = [vals[x][16] for x in range(4)]
        for pt in hts:
            stynos = list()
            #RG + 5% is special case, treat it as a new item
            rgridx = 0
            #10 rows up, discard if not found
            for x in range(1,10):
                ridx = pt.x - x
                if(ridx < 0): break
                val = vals[ridx][pt.y]
                if(isinstance(val,basestring)):
                    if(val.lower().find("style#") == 0):
                        for x in val[len("style#"):].split(","):
                            je = JOElement(x.strip())
                            if(len(je.alpha) == 1 and je.digit > 0): stynos.append(str(je))
                        break
                    else:
                        if(len(val) < 5): continue
                        if(val.lower()[:2] == 'rg'): rgridx = ridx                            
                        for x in val.split(","):
                            je = JOElement(x.strip())
                            if(len(je.alpha) == 1 and je.digit > 0): stynos.append(str(je))
            if(not stynos):
                logging.getLogger(__name__).debug("failed to get sty# for pt %s" % (pt,))                
            else:
                #4 rows down, must have
                rxs = [x + pt.x for x in range(1,5)]
                if(rgridx): rxs.append(rgridx)
                for x in rxs:
                    v0 = vals[pt.x][pt.y+2]
                    v0 = "" if not v0 else v0.lower()
                    #some items with stone, extend the columns if necessary
                    ccnt = 2 if v0 == "labour" else 3
                    tr = vals[x]
                    for jj in range(1,ccnt):
                        val = tr[pt.y + jj]
                        if(not isinstance(val,numbers.Number)): continue
                        #remark/type
                        rmk = tr[pt.y]
                        tp = "SS" if(rmk.lower() == "silver") else "RG+" if x == rgridx else "S+"                        
                        v0 = vals[pt.x][pt.y + jj]
                        if(v0) : rmk += ";" + v0
                        if(x == rgridx):
                            mpsidx = 1
                        else:
                            mpsidx = (x - pt.x - 1) % 2
                        mps = "S=%3.2f;G=%3.2f" % (mpss[mpsidx + 2],mpss[mpsidx])                        
                        for s0 in stynos:
                            items.append(Item(s0,mpsidx if x <> rgridx else 2 ,\
                                tp,mps,round(val,2),rmk.upper()))
        wb1 = app.books.add()
        sht = wb1.sheets[0]
        vals = list(items)
        v0 = sht.range((1,1),(len(items),len(items[0])))
        v0.value = vals
    finally:
        wb.close()
        if(kxl): app.books[0].close()
        if(not wb1 and kxl):
            app.quit()
        else:
            if(wb1): app.visible = True

if __name__ == "__main__":
    for x in (r'd:\temp\1200&15.xls',r'd:\temp\1300&20.xls'):
        readagq(x)
