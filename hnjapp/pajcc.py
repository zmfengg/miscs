# coding=utf-8
'''
Created on 2018-05-03

class to implement the PAJ Cost calculating

all cost related field are USD based
all weight related field are GM based

@author: zmFeng
'''

from collections import namedtuple
from csv import DictReader
from decimal import Decimal
from os import path

from hnjcore import karatsvc
from utilz import getvalue as gv
from utilz import trimu

from .common import _logger as logger

# the fineness map for this calculation
MPSINVALID = -10000.00


def _tofloat(val, precious=4):
    """ convenient method for (maybe) decimal to float """
    try:
        return round(float(val), precious)
    except:
        return -1


def addwgt(prdwgt, wi, isparts=False, autoswap=True):
    """ add wgt to target prdwgt """
    if not wi:
        return
    if not prdwgt:
        prdwgt = PrdWgt(wi)
    else:
        #act table: 0 -> main; 1 -> +main, 10 -> aux; 11 -> +aux; 20 -> part; 21 -> +part
        if isparts:
            act = 21 if prdwgt.part and prdwgt.part.wgt else 20
        elif prdwgt.main:
            act = 0 if not prdwgt.main.wgt else 1 if prdwgt.main.karat == wi.karat else 11 if prdwgt.aux and prdwgt.aux.wgt else 10
        else:
            act = 0
        if act == 0:
            prdwgt = prdwgt._replace(main=wi)
        elif act == 1:
            prdwgt = prdwgt._replace(
                main=WgtInfo(wi.karat, wi.wgt + prdwgt.main.wgt))
        elif act == 10:
            prdwgt = prdwgt._replace(aux=wi)
        elif act == 11:
            prdwgt = prdwgt._replace(
                aux=WgtInfo(wi.karat, wi.wgt + prdwgt.aux.wgt))
        elif act == 20:
            prdwgt = prdwgt._replace(part=wi)
        else:
            prdwgt = prdwgt._replace(
                part=WgtInfo(wi.karat, wi.wgt + prdwgt.part.wgt))
    if autoswap and prdwgt:
        if prdwgt.main and prdwgt.aux and prdwgt.main.wgt < prdwgt.aux.wgt:
            prdwgt = prdwgt._replace(main=prdwgt.aux, aux=prdwgt.main)

    return prdwgt


def cmpwgt(expected, actual, tor=5, strictkt=False):
    """
    tor is the toralent, positive stands for percentage, negative for actual wgt
    """
    if not (expected and actual):
        return None
    flag = True
    if not tor:
        tor = 5
    if tor > 1:
        tor = tor / 100.0
    for exp, act in zip(expected.wgts, actual.wgts):
        if bool(exp) ^ bool(act):
            return False
        if exp:
            flag = exp.karat and act.karat
            if flag:
                flag = exp.karat == act.karat if strictkt else karatsvc.getfamily(
                    exp.karat).karat == karatsvc.getfamily(act.karat).karat
            if flag:
                if tor > 0:
                    #xw = min(exp.wgt, act.wgt)
                    xw = exp.wgt
                    flag = round(abs(
                        (exp.wgt - act.wgt) / xw), 2) <= tor if xw else False
                else:
                    flag = round(abs(exp.wgt - act.wgt), 2) <= -tor
            if not flag:
                break
    return flag


class WgtInfo(namedtuple("WgtInfo", "karat,wgt")):
    """
    Product weight with karat and weight
    """

    @staticmethod
    def __new__(cls, karat, wgt, precious=2):
        if karat and wgt:
            return super().__new__(cls, karat, _tofloat(wgt, precious))
        return super().__new__(cls, 0, 0)
    
    def __str__(self):
        return "0" if not self.wgt else "%d=%4.2f" % (self.karat, self.wgt)


# mps string and the corresponding silver/gold value
class MPS():
    """
    class to hold gold/silver price
    you can construct by:
        .MPS("S=1;G=2") or MPS("G=2;S=1")
        .MPS(s=1, g=2) or MPS(silver=1, gold=2)
    """

    def __init__(self, mps=None, **kwds):
        self._slots = [None, ]*4
        flag = True
        if kwds:
            self._slots[0] = gv(kwds, "s") or gv(kwds, "silver")
            self._slots[1] = gv(kwds, "g") or gv(kwds, "gold")
            if any(self._slots[:2]):
                self._fmt()
        if flag:
            self._parse(mps)

    def _parse(self, mps):
        if not mps:
            return
        mps = mps.strip().upper()
        for mp in mps.split(";"):
            ps = mp.split("=")
            if len(ps) != 2:
                continue
            idx = 0 if ps[0] == "S" or ps[0] == "SILVER" else 1 if ps[
                0] == "G" or ps[0] == "GOLD" else -1
            if idx >= 0:
                self._slots[idx] = max(0, _tofloat(ps[1]))
        self._fmt()

    def _fmt(self):
        """
        create other slots from slot[:2]
        """
        ps = self._slots[:2]
        self._slots[3] = any(ps)
        if self._slots[3]:
            tarmps = []
            for t, p in zip(("S", "G"), ps):
                if not p:
                    continue
                tarmps.append("%s=%4.2f" % (t, p))
            self._slots[2] = ";".join(tarmps)
        else:
            self._slots[2] = None


    @classmethod
    def _floateq(cls, flt0, flt1):
        return abs(flt0 - flt1) < 0.0001

    @property
    def isvalid(self):
        """
        does this object contain valid data
        """
        return self._slots[3]

    @property
    def gold(self):
        """
        gold price
        """
        return self._slots[1] if self._slots[1] else 0

    @property
    def silver(self):
        """
        silver price
        """
        return self._slots[0] if self._slots[0] else 0

    @property
    def value(self):
        """
        a well-formatted MPS, for example, S=12.30;G=1234.50
        """

        return self._slots[2]

    def __eq__(self, other):
        return MPS._floateq(self.silver, other.silver) and MPS._floateq(
            self.gold, other.gold)

    def __hash__(self):
        return hash((int(self.silver * 10000), int(self.gold * 10000)))

    def __str__(self):
        return self._slots[2] if self._slots[2] else ""

    def __repr__(self):
        return repr(self._slots)


Increment = namedtuple("Increment", "wgts,silver,gold,lossrate")


# the china cost related data
class PajChina(
        namedtuple("PajChina", "china,increment,mps,discount,metalcost")):
    """ class to hold pajChina related data:
        china,increment,mps,discount,metalcost
    """

    @staticmethod
    def __new__(cls, china, increment, mps, discount, metalcost):
        return super().__new__(cls, _tofloat(china, 4)\
        , increment, mps, _tofloat(discount, 4), _tofloat(metalcost, 4))

    @property
    def lossrate(self):
        """
        lossrate of this calculation
        """
        return self.increment.lossrate

    def othercost(self):
        """
        cost except metal cost
        """
        return self.china - self.metalcost


class PrdWgt(namedtuple("PrdWgt", "main,aux,part,netwgt")):
    """
        product weight, of mainpart/auxpart/parts
    """
    __slots__ = ()

    @staticmethod
    def __new__(cls, main, aux=None, part=None, netwgt=0):
        return super(cls, PrdWgt).__new__(cls, main, aux, part, netwgt)

    @property
    def wgts(self):
        """
        return a tuple of WgtInfo presenting main/aux/part data
        """
        return (self.main, self.aux, self.part)

    def __str__(self):
        d = {"main": self.main, "sub": self.aux, "part": self.part}
        return ";".join(["%s(%s=%s)" % (kw[0], kw[1].karat, kw[1].wgt) \
            for kw in d.items() if kw[1]])

    @property
    def metal(self):
        """
        return a list of WgtInfo Object(s)
        """
        return tuple(x for x in (self.main, self.aux) if x and x.wgt > 0)

    @property
    def chain(self):
        """ WgtInfo type of chain weight """
        wi = self.part
        if not (wi and wi.wgt):
            return None
        return WgtInfo(wi.karat, wi.wgt if wi.wgt > 0 else -wi.wgt / 100)

    @property
    def metal_stone(self):
        """ metal & stone weight without chain """
        return self.netwgt - (self.chain.wgt if self.chain else 0)

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
    def getfineness(cls, karat, vendor="PAJ"):
        """
        only PAJ's 925 item has different fineness
        """
        if vendor and karat == 925:
            karat = "925PAJ"
        kt = karatsvc.getkarat(karat)
        return kt.fineness if kt else 0

    @classmethod
    def calclossrate(cls, prdwgt):
        """
        calcualte the lossrate based on the prdwgt data
        """
        rts = [
            1.1 if x.karat == 925 else 1.06
            for x in [x for x in (prdwgt.main, prdwgt.aux) if x and x.wgt > 0]
        ]
        return max(rts) if rts else 1.06

    @classmethod
    def calcmtlcost(cls,
                    prdwgt,
                    mps, **kwds):
        """
        calculate the metal cost

        """

        lossrate, vendor, oz2gm = tuple(kwds.get(x) for x in ("lossrate vendor oz2gm".split()))
        if not vendor:
            vendor = "PAJ"
        if not oz2gm:
            oz2gm = 31.1035
        kws = prdwgt.wgts
        hix = [ii for ii in range(len(kws)) if kws[ii] and kws[ii].wgt > 0]
        lr0 = lossrate if lossrate else cls.calclossrate(prdwgt)
        r0 = 0
        if isinstance(mps, str):
            mps = MPS(mps)
        for idx in hix:
            x = kws[idx]
            mp = mps.silver if x.karat == 925 else 0 if x.karat == 200 else mps.gold
            if not mp and x.karat != 200:
                r0 = MPSINVALID
                break
            r0 += (x.wgt * cls.getfineness(x.karat, vendor) *
                   (lr0 if vendor != "PAJ" or idx < 2 else 1.0) * mp / oz2gm)
        return round(r0, 2)

    @classmethod
    def calcdiscount(cls, prdwgt):
        """ the discount rate of PAJ"""
        kws = prdwgt.wgts
        return 0.9 if kws[0].karat == 925 or (kws[1] and
                                              kws[1].karat == 925) else 0.85

    @classmethod
    def calcincrement(cls, prdwgt, lossrate=None, vendor=None):
        """ calculate the increment based on the product weight provided
            @param prdwgt:  weight of the product
            @param vendor: PAJ or Non-PAJ or None
        """
        kws, s, g = prdwgt.wgts, 0, 0
        hix = [ii for ii in range(len(kws)) if kws[ii] and kws[ii].wgt > 0]

        if not lossrate:
            lossrate = cls.calclossrate(prdwgt)
        for idx in hix:
            kw = kws[idx]
            # parts does not have loss
            r0 = kw.wgt * cls.getfineness(kw.karat, vendor) * \
                    (lossrate if idx < 2 else 1.0) / 31.1035
            if kw.karat == 925:
                s += r0
            elif kw.karat == 200:
                pass
            else:
                g += r0

        return Increment(prdwgt, s, g, lossrate)

    @classmethod
    def _checkargs(cls, incr, refmps, tarmps):
        """check if incr/mpss is valid"""
        if not incr:
            return True
        return (not incr.gold or incr.gold and tarmps.gold and refmps.gold) and \
            (not incr.silver or incr.silver and tarmps.silver and refmps.silver)

    @classmethod
    def calchina(cls, prdwgt, refup, refmps, tarmps=None):
        """ calculate the China cost based on the provided arguments
        @param prdwgt: weights of the product
        @param refup: the reference unit price
        @param refmps: the reference mps of the @refup
        @param tarmps: the target MPS the need to be calculated. PAJ's china MPS is S = 30; G = 1500
        return: a PajChina Object
        """
        if not all((prdwgt, refup, refmps)):
            return None
        if not tarmps:
            tarmps = PAJCHINAMPS
        if isinstance(tarmps, str):
            tarmps = MPS(tarmps)
        if isinstance(refmps, str):
            refmps = MPS(refmps)
        if isinstance(refup, Decimal):
            refup = float(refup)
        if not (refup > 0 and refmps.isvalid and tarmps.isvalid):
            return None

        # the discount ratio, when there is silver, follow silver, silver = 0.9 while gold = 0.85
        incr = cls.calcincrement(prdwgt, None, "PAJ")
        dc = cls.calcdiscount(prdwgt)
        if not cls._checkargs(incr, refmps, tarmps):
            cn = MPSINVALID
            mc = MPSINVALID
            logger.debug("MPS(%s) not enough for calculating increment(%s)" %
                         (tarmps.value, str(incr)))
        else:
            cn = refup / dc + incr.gold * (tarmps.gold - refmps.gold) * 1.25 \
                + incr.silver * (tarmps.silver - refmps.silver) * 1.25
            mc = cls.calcmtlcost(prdwgt, tarmps, lossrate=incr.lossrate, vendor="PAJ")
        return PajChina(round(cn, 2), incr, tarmps, dc, mc)

    @classmethod
    def calctarget(cls, cn, tarmps):
        """calculate the target unit price based on the data _NotProvided
        @param cn: the PAJChina cost
        @param tarmps: the target MPS
        @return: a PajChina object, the china is the current value
        """

        if isinstance(tarmps, str):
            tarmps = MPS(tarmps)
        incr = cn.increment
        if not cls._checkargs(incr, cn.mps, tarmps):
            r0 = MPSINVALID
            mc = MPSINVALID
            logger.debug("MPS(%s) not enough for calculating increment(%s)" %
                         (tarmps.value, str(incr)))
        else:
            r0 = cn.china + (tarmps.gold - cn.mps.gold) * incr.gold * 1.25 \
                + (tarmps.silver - cn.mps.silver) * incr.silver * 1.25
            r0 = round(r0 * cn.discount, 2)
            mc = cls.calcmtlcost(incr.wgts, tarmps, lossrate=incr.lossrate, vendor="PAJ")
        return PajChina(r0, cn.increment, tarmps, cn.discount, mc)


class P17Decoder():
    """
    classeto fetch the parts(for example, karat) out from a p17
    """

    def __init__(self):
        self._cats_ = self._getp17cats()
        self._ppart = None

    @classmethod
    def _getp17cats(cls):
        """return the categories of all the P17s(from database)
        @return: a map of items containing "catid/cat/digits. This module should not have db code, so hardcode here
        """
        rdr = path.join(path.dirname(__file__), "res", "pcat.csv")
        with open(rdr, "r") as fh:
            rdr = DictReader(fh, dialect='excel-tab')
            return {trimu(x["name"]): (x["cat"], x["digits"]) for x in rdr}

    @classmethod
    def _getdigits(cls, p17, digits):
        """ parse the p17's given code out
        @param p17: the p17 code need to be parse out
        @param digits: the digits, like "1,11"
        """
        rc = ""
        for x in digits.split(","):
            pts = x.split("-")
            rc += p17[int(x) - 1] if len(pts) == 1 else p17[int(pts[0]) - 1:(
                int(pts[1]))]
        return rc

    def _getpart(self, cat, code):
        """fetch the cat + code from database"""
        # todo:: no database now, try from csv or other or sqlitedb
        # "select description from uv_p17dc where category = '%(cat)s' and codec = '%(code)s'"
        if not self._ppart:
            fn = path.join(path.dirname(__file__), "res", "ppart.csv")
            with open(fn, "r") as fh:
                rdr = DictReader(fh)
                self._ppart = {x["catid"] + x["codec"]: x["description"].strip()\
                    for x in rdr}
        code = self._ppart.get(self._cats_[cat][0] + code, code)
        if not isinstance(code, str):
            code = code["description"]
        return code

    def decode(self, p17, parts=None):
        """parse a p17's parts out
        @param p17: the p17 code
        @param parts: the combination of the parts name delimited with ",". None to fetch all
        @return the actual value if only one parts, else return a dict with part as key
        """
        ns, ss = tuple(trimu(x) for x in parts.split(",")) if parts\
            else self._cats_.keys(), []
        for x in ns:
            ss.append((x,
                       self._getpart(x, self._getdigits(p17,
                                                        self._cats_[x][1]))))
        return ss[0][1] if len(ns) <= 1 else dict(ss)
