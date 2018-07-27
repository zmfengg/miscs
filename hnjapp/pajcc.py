# coding=utf-8
'''
Created on 2018-05-03

class to implement the PAJ Cost calculating

all cost related field are USD based
all weight related field are GM based

@author: zmFeng
'''

from collections import namedtuple
from decimal import Decimal
from hnjcore import karatsvc,Karat
from .common import _logger as logger
from os import path
import csv
from csv import DictReader, Dialect
from utilz import trimu


# the fineness map for this calculation
MPSINVALID = -10000.00

def _tofloat(val,precious = 4):
    """ convenient method for (maybe) decimal to float """
    try:
        return round(float(val),precious)
    except Exception as e:
        print(e)
        return -1
    

# karat and weight
class WgtInfo(namedtuple("WgtInfo", "karat,wgt")):

    def __new__(_cls, karat, wgt):
        return super(_cls, WgtInfo).__new__(_cls, karat, _tofloat(wgt,2))


# mps string and the corresponding silver/gold value
class MPS():

    def __init__(self, mps):
        self._slots = [None, None, None, None]        
        self._parse(mps)

    def _parse(self, mps):
        if not mps: return
        mps = mps.strip().upper()
        tarmps = None
        for mp in mps.split(";"):
            ps = mp.split("=")
            if len(ps) == 2:
                idx = 0 if ps[0] == "S" or ps[0] == "SILVER" else 1 if ps[0] == "G" or ps[0] == "GOLD" else -1
                if idx >= 0:
                    self._slots[idx] = float(ps[1])
            ps = self._slots[:2]
        if any(ps):
            ttls = ("S","G")
            for idx in range(len(ps)):
                if not ps[idx]: continue
                tarmps = tarmps + ";"  if tarmps else ""
                tarmps += "%s=%4.2f" % (ttls[idx],ps[idx])
        self._slots[3] = bool(tarmps)
        if tarmps: self._slots[2] = tarmps
    
    @classmethod
    def _floateq(self,flt0,flt1):
        return abs(flt0 - flt1) < 0.0001

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

    def __eq__(self, other):
        return MPS._floateq(self.silver,other.silver) and  MPS._floateq(self.gold,other.gold)

    def __hash__(self):        
        return hash((int(self.silver * 10000),int(self.gold * 10000)))

    def __str__(self):
        return self._slots[2] if self._slots[2] else ""

    def __repr__(self):
        return repr(self._slots)


Increment = namedtuple("Increment", "wgts,silver,gold,lossrate")


# the china cost related data
class PajChina(namedtuple("PajChina", "china,increment,mps,discount,metalcost")):

    def __new__(_cls, china,increment,mps,discount,metalcost):
        return super(_cls, PajChina).__new__(_cls, _tofloat(china,4), \
        increment,mps,_tofloat(discount,4),_tofloat(metalcost,4))

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
        lr = lossrate if lossrate > 0 else max(
            [1.1 if kws[x].karat == 925 else 1.06 for x in hix])
        dc = 0.9 if kws[0].karat == 925 or (
            kws[1] and kws[1].karat == 925) else 0.85
        inst = super(_cls, PrdWgt).__new__(_cls, main, aux, part, lr, dc)
        return inst
        '''

    @property
    def wgts(self):
        return (self.main, self.aux, self.part)
    
    def __str__(self):
        d = {"main":self.main, "sub":self.aux, "part":self.part}
        return ";".join(["%s(%s=%s)" % (kw[0],kw[1].karat,kw[1].wgt) \
            for kw in d.items() if kw[1]])


# constants
PAJCHINAMPS = MPS("S=30;G=1500")


def newchina(cn, wgts):
    """while knowing the china/wgts, return a PajChina instance
    @param cn: the known China cost
    @param wgt: A PrdWgt instance
    """
    cc = PajCalc()
    return PajChina(cn, cc.calcincrement(wgts), PAJCHINAMPS, cc.calcdiscount(wgts), \
        cc.calcmtlcost(wgts, PAJCHINAMPS))


class PajCalc():
    """the PAJ related calculations"""

    @classmethod
    def getfineness(self, karat, vendor="PAJ"):
        #only PAJ's 925 item has different fineness
        if vendor and karat == 925: karat = "925PAJ"
        kt = karatsvc.getkarat(karat)
        return kt.fineness if kt else 0

    @classmethod
    def calclossrate(self, prdwgt):
        rts = [1.1 if x.karat == 925 else 1.06 for x in [
            x for x in (prdwgt.main, prdwgt.aux) if x and x.wgt > 0]]
        return max(rts) if rts else 1.06

    @classmethod
    def calcmtlcost(self, prdwgt, mps, lossrate=None, vendor="PAJ",oz2gm = 31.1035):
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
            r0 += (x.wgt * self.getfineness(x.karat, vendor) * (lr0 if vendor != "PAJ" or idx < 2 else 1.0) *
                mp / oz2gm)
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
        s = 0; g = 0
        hix = [ii for ii in range(len(kws)) if kws[ii] and kws[ii].wgt > 0]

        if not lossrate:
            lossrate = self.calclossrate(prdwgt)
        for idx in hix:
            kw = kws[idx]
            # parts does not have loss
            r0 = kw.wgt * \
                self.getfineness(kw.karat, vendor) * \
                               (lossrate if(idx < 2) else 1.0) / 31.1035
            if(kw.karat == 925):
                s += r0
            elif(kw.karat == 200):
                pass
            else:
                g += r0

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
        @param tarmps: the target MPS the need to be calculated. PAJ's china MPS is S = 30; G = 1500
        return: a PajChina Object
        """
        if not all((prdwgt, refup, refmps)): return None
        if not tarmps: tarmps = PAJCHINAMPS
        if isinstance(tarmps,str): tarmps = MPS(tarmps)
        if isinstance(refmps,str): refmps = MPS(refmps)
        if isinstance(refup,Decimal): refup = float(refup)
        if not (refup > 0 and refmps.isvalid and tarmps.isvalid): return None

        # the discount ratio, when there is silver, follow silver, silver = 0.9 while gold = 0.85
        incr = self.calcincrement(prdwgt, None, "PAJ")
        dc = self.calcdiscount(prdwgt)
        if not self._checkargs(incr, refmps, tarmps):
            cn = MPSINVALID
            mc = MPSINVALID
            logger.debug("MPS(%s) not enough for calculating increment(%s)" % (
                tarmps.value,str(incr)))
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

        if isinstance(tarmps,str): tarmps = MPS(tarmps)
        incr = cn.increment
        if not self._checkargs(incr, cn.mps, tarmps):
            r0 = MPSINVALID
            mc = MPSINVALID
            logger.debug("MPS(%s) not enough for calculating increment(%s)" % (
                tarmps.value,str(incr)))
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

    def _getp17cats(self):
        """return the categories of all the P17s(from database)
        @return: a map of items containing "catid/cat/digits. This module should not have db code, so hardcode here
        """
        
        with open(path.join(path.dirname(__file__),"res","pcat.csv"),"r") as fh:
            rdr = DictReader(fh, dialect='excel-tab')
            return dict([(trimu(x["name"]),(x["cat"],x["digits"])) for x in rdr])
        cats = {}

    def _getdigits(self, p17, digits):
        """ parse the p17's given code out
        @param p17: the p17 code need to be parse out
        @param digits: the digits, like "1,11"
        """
        rc = ""
        for x in digits.split(","):
            pts = x.split("-")
            rc += p17[int(x) - 1] if len(pts) == 1 else p17[int(pts[0]
                          ) - 1:(int(pts[1]))]
        return rc

    def _getpart(self, cat, code):
        """fetch the cat + code from database"""
        # todo:: no database now, try from csv or other or sqlitedb
        # "select description from uv_p17dc where category = '%(cat)s' and codec = '%(code)s'"
        if not hasattr(self,"_ppart"):
            with open(path.join(path.dirname(__file__),"res","ppart.csv"),"r") as fh:
                rdr = DictReader(fh)
                self._ppart = dict([(x["catid"]+x["codec"],x["description"].strip()) for x in rdr])
        code = self._ppart.get(self._cats_[cat][0] + code,code)
        if not isinstance(code,str):
            code = code["description"]
        return code

    def decode(self, p17, parts=None):
        """parse a p17's parts out
        @param p17: the p17 code
        @param parts: the combination of the parts name delimited with ",". None to fetch all
        @return the actual value if only one parts, else return a dict with part as key
        """
        ns, ss = [trimu(x) for x in parts.split(",")] if parts else self._cats_.keys(), []
        [ss.append((x, self._getpart(x, self._getdigits(p17, self._cats_[x][1])))) for x in ns]
        if len(ns) <= 1:
            return ss[0][1]
        else:
            return dict(ss)