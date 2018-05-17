# coding=utf-8
'''
Created on 2018-05-03

class to implement the PAJ Cost calculating

all cost related field are USD based
all weight related field are GM based

@author: zmFeng
'''

from collections import namedtuple
import logging

# the fineness map for this calculation
_fineness = {8:33.3, 81:33.3, 88:33.3, 9:37.5, 91:37.5, 98:37.5, 10:41.7, 101:41.7, 108:41.7, 14:58.5, 141:58.5, 148:58.5, \
    18:75.0, 181:75.0, 188:75.0, 200:100, "925PAJ":95.0, "925":92.5, 9925:0.95}

MPSINVALID = -10000.00

# karat and weight
class WgtInfo(namedtuple("WgtInfo", "karat,wgt")):

    def __new__(_cls, karat, wgt):
        return super(_cls, WgtInfo).__new__(_cls, karat, round(wgt, 2))


# mps string and the corresponding silver/gold value
class MPS():
    
    def __init__(self, mps):
        self._slots = [None, None, None, None]
        self._parse(mps)
        
    def _parse(self, mps):
        mps = mps.strip().upper()
        tarmps = None
        for mp in mps.split(";"):
            ps = mp.split("=")
            if len(ps) == 2:
                idx = 0 if ps[0] == "S" or ps[0] == "SILVER" else 1 if ps[0] == "G" or ps[0] == "GOLD" else -1
                if idx >= 0:
                    self._slots[idx] = float(ps[1])
                    tarmps = tarmps + ";" if tarmps else ""
                    tarmps += ("S" if idx == 0 else "G") + "=" + ps[1]
        self._slots[3] = bool(tarmps)        
        if tarmps: self._slots[2] = tarmps

    @property
    def isvalid(self):
        return self._slots[3]
    
    @property
    def gold(self):
        return self._slots[1] if self._slots[1] else 0
    
    @property
    def silver(self):
        return self._slots[0] if self._slots[0] else 0
    
    @property
    def value(self):
        return self._slots[2]
    
    def __str__(self):
        return self._slots[2] if self._slots[2] else ""
    
    def __repr__(self):
        return repr(self._slots)


Increment = namedtuple("Increment", "wgts,silver,gold,lossrate")


# the china cost related data    
class PajChina(namedtuple("PajChina", "china,increment,mps,discount,metalcost")):
    
    @property
    def lossrate(self):
        return self.increment.lossrate
    
    def othercost(self):
        return self.china - self.metalcost


# product weight, of mainpart/auxpart/parts
class PrdWgt(namedtuple("PrdWgt", "main,aux,part")):
    __slots__ = ()
    
    # pydev会报错,但是实际上是OK的
    def __new__(_cls, main, aux=None, part=None):
        return super(_cls, PrdWgt).__new__(_cls, main, aux, part)
        '''
        kws = (main, aux, part)
        hix = [ii for ii in range(len(kws)) if kws[ii] and kws[ii].wgt > 0]
        lr = lossrate if lossrate > 0 else max([1.1 if kws[x].karat == 925 else 1.06 for x in hix])
        dc = 0.9 if kws[0].karat == 925 or (kws[1] and kws[1].karat == 925) else 0.85
        inst = super(_cls, PrdWgt).__new__(_cls, main, aux, part, lr, dc)        
        return inst
        '''

    @property
    def wgts(self):
        return (self.main, self.aux, self.part)
    

# constants
PAJCHINAMPS = MPS("S=30;G=1500")


def newchina(cn, wgt):
    """while knowing the china/wgts, return a PajChina instance
    @param cn: the known China cost
    @param wgt: A PrdWgt instance
    """
    cc = PajCalc()
    return PajChina(cn, cc.calcincrement(wgt), PAJCHINAMPS, cc.calcdiscount(wgt), \
        cc.calcmtlcost(wgt, PAJCHINAMPS))


class PajCalc():
    """the PAJ related calculations"""

    @classmethod
    def getfiness(self, karat, vendor="PAJ"):
        """return the finenss of the given karat
        @param karat: the karat in numeric form, for example, 8 or 81
        @param vendor: PAJ or Non-PAJ or None  
        """
        
        lst = (karat, "%d%s" % (karat, vendor) if vendor else str(karat))
        rr = [_fineness[x] for x in lst if _fineness.has_key(x)]
        if rr: return rr[0] / 100.0

    @classmethod
    def calclossrate(self, prdwgt):
        rts = [1.1 if x.karat == 925 else 1.06 for x in [x for x in (prdwgt.main, prdwgt.aux) if x and x.wgt > 0]]
        return max(rts) if rts else 1.06
    
    @classmethod
    def calcmtlcost(self, prdwgt, mps, lossrate=None, vendor="PAJ"):
        """
        calculate the metal cost      
            
        """
        kws = prdwgt.wgts
        hix = [ii for ii in range(len(kws)) if kws[ii] and kws[ii].wgt > 0]
        lr0 = lossrate if lossrate else self.calclossrate(prdwgt) ; r0 = 0
        for idx in hix:
            x = kws[idx]
            mp = mps.silver if x.karat == 925 else 0 if x.karat == 200 else mps.gold
            if not mp and x.karat != 200:
                r0 = MPSINVALID
                break
            r0 += (x.wgt * self.getfiness(x.karat, vendor) * (lr0 if idx < 2 else 1.0) *
                mp / 31.1035)
        return round(r0, 2)
    
    @classmethod
    def calcdiscount(self, prdwgt):
        """ the discount rate of PAJ"""
        kws = prdwgt.wgts
        return 0.9 if kws[0].karat == 925 or (kws[1] and kws[1].karat == 925) else 0.85

    @classmethod
    def calcincrement(self, prdwgt, lossrate=None, vendor=None):
        """ calculate the increment based on the product weight provided
            @param prdwgt:  weight of the product
            @param vendor: PAJ or Non-PAJ or None
        """
        kws = prdwgt.wgts
        s = g = 0;
        hix = [ii for ii in range(len(kws)) if kws[ii] and kws[ii].wgt > 0]
        
        if not lossrate:
            lossrate = self.calclossrate(prdwgt)        
        for idx in hix:
            kw = kws[idx]
            # parts does not have loss
            r0 = kw.wgt * self.getfiness(kw.karat, vendor) * (lossrate if(idx < 2) else 1.0) / 31.1035;
            if(kw.karat == 925):                
                s += r0;                    
            elif(kw.karat == 200):
                pass
            else:
                g += r0;
    
        return Increment(prdwgt, s, g, lossrate)
    
    @classmethod
    def _checkargs(cls,incr,refmps,tarmps):
        """check if incr/mpss is valid"""
        if not incr: return True
        return (not incr.gold or incr.gold and tarmps.gold and refmps.gold) and \
            (not incr.silver or incr.silver and tarmps.silver and refmps.silver)

    @classmethod
    def calchina(self, prdwgt, refup, refmps, tarmps=None, lossrate=None):
        """ calculate the China cost based on the provided arguments
        @param prdwgt: weights of the product
        @param refup: the reference unit price
        @param refmps: the reference mps of the @refup
        @param tarmps: the target MPS the need to be calculated. PAJ's china MPS is S=30;G=1500  
        """
        if not all((prdwgt, refup, refmps)): return None
        if not tarmps: tarmps = PAJCHINAMPS
        if not (refup > 0 and refmps.isvalid and tarmps.isvalid): return None 
        
        # the discount ratio, when there is silver, follow silver, silver = 0.9 while gold = 0.85
        incr = self.calcincrement(prdwgt, None, "PAJ");
        dc = self.calcdiscount(prdwgt)
        if not self._checkargs(incr, refmps, tarmps):
            cn = MPSINVALID
            mc = MPSINVALID
            logging.debug("MPS(%s) not enough for calculating increment(%s)" % (tarmps.value,str(incr)))
        else:
            cn = refup / dc + incr.gold * (tarmps.gold - refmps.gold) * 1.25 \
                + incr.silver * (tarmps.silver - refmps.silver) * 1.25
            mc = self.calcmtlcost(prdwgt, tarmps, incr.lossrate, "PAJ")
        return PajChina(round(cn, 2), incr, tarmps, dc, mc)
    
    @classmethod
    def calctarget(self, cn, tarmps):
        """calculate the target unit price based on the data _NotProvided
        @param cn: the PAJChina cost
        @param tarmps: the target MPS      
        @return: a PajChina object, the china is the current value
        """
        
        if isinstance(tarmps,basestring): tarmps = MPS(tarmps)
        incr = cn.increment
        if not self._checkargs(incr, cn.mps, tarmps):
            r0 = MPSINVALID
            mc = MPSINVALID
            logging.debug("MPS(%s) not enough for calculating increment(%s)" % (tarmps.value,str(incr)))
        else:
            r0 = cn.china + (tarmps.gold - cn.mps.gold) * incr.gold * 1.25 \
                + (tarmps.silver - cn.mps.silver) * incr.silver * 1.25
            r0 = round(r0 * cn.discount, 2)
            mc = self.calcmtlcost(incr.wgts, tarmps, incr.lossrate, "PAJ")
        return PajChina(r0, cn.increment, tarmps, cn.discount, mc)


class P17Decoder():
    """classes to fetch the parts(for example, karat) out from a p17"""
    
    def __init__(self):
        self._cats_ = self._getp17cats()
    
    @classmethod
    def _getp17cats(self):
        """return the categories of all the P17s (from database)
        @return: a map of items containing "catid/cat/digits. This module should not have db code, so hardcode here
        """
        cats = {}
        for x in ((1, "KARAT", "1,11"), (2, "PRODTYPE", "2"), (3, "VERSION", "3-6"), (4, "STONE", "7-8"), \
            (5, "SIZEORPART", "2,9-10"), (6, "SPROCESS", "12-13"), (7, "QCNCHOP", "14-15"), (8, "STLEVEL", "7-8,16-17")):
            cats[x[1]] = x
        return cats
    
    def _getdigits(self, p17, digits):
        """ parse the p17's given code out
        @param p17: the p17 code need to be parse out
        @param digits: the digits, like "1,11"
        """
        rc = ""
        for x in digits.split(","):
            pts = x.split("-")
            rc += p17[int(x) - 1] if len(pts) == 1 else p17[int(pts[0]) - 1:(int(pts[1]))]
        return rc
    
    def _getpart(self, cat, code):
        """fetch the cat + code from database"""
        # todo:: no database now, try from csv or other or sqlitedb 
        # "select description from uv_p17dc where category = '%(cat)s' and codec = '%(code)s'"
        return code
    
    def decode(self, p17, parts=None, div=","):
        """parse a p17's parts out
        @param p17: the p17 code
        @param parts: the combination of the parts name delimited with ",". None to fetch all 
        """
        ns = parts.split(",") if parts else self._cats_.keys();ss = []        
        for x in ns:
            ss.append("%s=%s" % (x, self._getpart(x, self._getdigits(p17, self._cats_[x][2]))))
        return div.join(ss)
