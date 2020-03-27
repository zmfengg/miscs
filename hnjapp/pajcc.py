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
from datetime import datetime

from hnjcore import karatsvc
from utilz import getvalue as gv

from .common import _logger as logger, config

# the fineness map for this calculation
MPSINVALID = -10000.00


def _tofloat(val, precious=4):
    """ convenient method for (maybe) decimal to float """
    try:
        return round(float(val), precious)
    except:
        return -1

def addwgt(prdwgt, wi, isparts=False, autoswap=False):
    """ add wgt to target prdwgt """
    if not wi:
        return
    if not prdwgt:
        prdwgt = PrdWgt(wi)
    else:
        #act table: 0 -> main; 1 -> +main, 10 -> aux; 11 -> +aux; 20 -> part; 21 -> +part
        # don't use below rule because it will block some aux-for-parts adding to the parts
        if False:
            if isparts and (wi.karat != 925 and wi.wgt < 0.3 or wi.karat == 925 and wi.wgt < 1.4):
                isparts = False
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
    def _raw_cmp(w0, w1):
        if tor > 0:
            xw = w0 or 1
            flag = round(abs(
                (_adj_wgt(w0) - _adj_wgt(w1)) / xw), 2) <= tor if xw else False
        else:
            flag = round(abs(_adj_wgt(w0) - _adj_wgt(w1)), 2) <= -tor
        return flag
    flag = _raw_cmp(expected.netwgt or 0, actual.netwgt or 0)
    if not flag:
        return flag
    for exp, act in zip(expected.wgts, actual.wgts):
        if bool(exp) ^ bool(act):
            return False
        if exp:
            flag = exp.karat and act.karat
            if flag:
                flag = exp.karat == act.karat if strictkt else karatsvc.getfamily(
                    exp.karat).karat == karatsvc.getfamily(act.karat).karat
            if flag:
                flag = _raw_cmp(exp.wgt, act.wgt)
            if not flag:
                break
    return flag

def _adj_wgt(wgt):
    ''' for the wgt < 0, turn it to actual value '''
    if not wgt or wgt > 0:
        return wgt
    return wgt / -100.0

class WgtInfo(namedtuple("WgtInfo", "karat,wgt")):
    """
    Product weight with karat and weight
    """

    @staticmethod
    def __new__(cls, karat, wgt, precious=2):
        # if karat and wgt:
        if karat:
            return super().__new__(cls, karat, _tofloat(wgt, precious))
        return super().__new__(cls, 0, 0)

    def __str__(self):
        return "0" if not self.wgt else "%d=%4.2f" % (self.karat, self.wgt)


# mps string and the corresponding silver/gold value
class MPS(object):
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
    def __new__(cls, main=None, aux=None, part=None, netwgt=0):
        return super(cls, PrdWgt).__new__(cls, main, aux, part, netwgt)

    @property
    def wgts(self):
        """
        return a tuple of WgtInfo presenting main/aux/part data
        """
        return (self.main, self.aux, self.part)

    def __str__(self):
        d = {"main": self.main, "sub": self.aux, "part": self.part}
        d0 = ";".join(["%s(%s=%s)" % (kw[0], kw[1].karat, kw[1].wgt) \
            for kw in d.items() if kw[1]])
        if self.netwgt:
            d0 += ";net=%s" % self.netwgt
        return d0

    def terms(self, div='-'*15):
        '''
        return an string in the terms of metal/stone/chain
        '''
        _fmt_wgtinfo = lambda wi, tn: "%s%s:%4.2fgm" % (karatsvc.getkarat(wi.karat).name, tn, wi.wgt)
        aio = []
        var = [_fmt_wgtinfo(x, '') for x in self.metal]
        aio.extend(var)
        # show netwgt only when there is stone
        var = self.metal_stone
        if var:
            if div:
                aio.append(div)
            aio.append("w/st:%4.2fgm" % var)
        var = self.chain
        if var:
            if div:
                aio.append(div)
            aio.append(_fmt_wgtinfo(var, "chain"))
        return "\n".join(aio)

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
        """
        metal & stone weight without chain, if no stone, this should return 0
        """
        st = self.netwgt - (self.chain.wgt if self.chain else 0)
        return 0 if abs(sum(x.wgt for x in self.metal) - st) < 0.008 else st

    @property
    def metal_jc(self):
        """ the metal weight for jocost, that is the weight of all metals """
        return sum(x.wgt if x.wgt > 0 else -x.wgt / 100 for x in self.wgts if x)

    def follows(self, mKarat):
        '''
        stupid method, if main karat differs from mKarat, swap aux and main
        simple and stupid
        '''
        wi = self.main
        if karatsvc.getfamily(wi.karat).karat != karatsvc.getfamily(mKarat).karat and self.aux is not None:
            return PrdWgt(self.aux, self.main, self.part, self.netwgt)
        return self


# constants
PAJCHINAMPS = MPS(config.get('pajcc.consts')['paj.ref.mps'])

class DiscountHis(object):
    ''' manage the discount histories
    '''
    def __init__(self):
        mp = config.get('pajcc.consts')
        ttl = 'frm,thru,silver,other'
        nc, ttl = namedtuple('DiscountDef', ttl), ttl.split(',')
        lst = [[x.get(y) for y in ttl] for x in mp['discounts']]
        for x in lst:
            for idx in range(2):
                if not x[idx]:
                    continue
                x[idx] = datetime.strptime(x[idx], '%Y/%m/%d').date()
        self._dscs = sorted([nc(*x) for x in lst], key=lambda x: x.frm)
        self._first, self._latest = self._dscs[0], self._dscs[-1]

    def discount(self, d0=None):
        ''' return the discount of given date as tuple
        Args:
            d0=None:  the date for that discount, None for the initial date
        '''
        inst = None
        if not d0:
            inst = self._first
        else:
            if isinstance(d0, datetime):
                d0 = d0.date()
            idxf, idxt = 0, len(self._dscs) - 1
            while idxf <= idxt:
                idx = (idxf + idxt) // 2
                inst0 = self._dscs[idx]
                if not inst0.thru or (inst0.frm <= d0 < inst0.thru):
                    inst = inst0
                    break
                if inst0.frm > d0:
                    idxt = idx - 1
                else:
                    idxf = idx + 1
        return (inst.silver, inst.other) if inst else None

class PajCalc(object):
    """the PAJ related calculations"""

    _markup = _lr_silver = _lr_other = None
    _disc_his = DiscountHis()

    @classmethod
    def _args(cls):
        if cls._lr_silver:
            return
        mp = config.get('pajcc.consts')
        cls._lr_silver, cls._lr_other, cls._markup = (mp[x] for x in ('lossrate.silver', 'lossrate.other', 'paj.ref.markup'))

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
        cls._args()
        rts = [
            cls._lr_silver if x.karat == 925 else cls._lr_other for x in [x for x in (prdwgt.main, prdwgt.aux) if x and x.wgt > 0]
        ]
        return max(rts) if rts else cls._lr_other

    @classmethod
    def calcmtlcost(cls, prdwgt, mps, **kwds):
        """
        calculate the metal cost

        Args:

            prdwgt: PrdWgt instance

            mps:    MPS instance

            lossrate=None:  lossrate for the prdwgt

            vendor=PAJ: the vendor ID

            oz2gm=31.1035:  onze to gm conversion


        """

        lossrate, vendor, oz2gm = (kwds.get(x) for x in "lossrate vendor oz2gm".split())
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
    def calcdiscount(cls, prdwgt, affdate=None):
        """ calculate the discount rate of given wgts

        Args:

            prdwgt: PrdWgt instance

            affdate=None: the date that the discount blongs to, when omitted, it's 2017/03/01

        """
        dsc = cls._disc_his.discount(affdate)
        kws = prdwgt.wgts
        #return 0.9 if kws[0].karat == 925 or (kws[1] and kws[1].karat == 925) else 0.85
        idx = 0 if kws[0].karat == 925 or (kws[1] and kws[1].karat == 925) else 1
        return dsc[idx]

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
            elif kw.karat != 200:
                g += r0

        return Increment(prdwgt, s, g, lossrate)

    @classmethod
    def newchina(cls, cn, wgts, affdate=None):
        ''' new an PajChina instance based on the cn and wgts provided

        Args:

            cn:     the china cost

            wgts:   the PrdWgt data

            affdate=None: affected date of this calculation

        '''
        return PajChina(cn, cls.calcincrement(wgts), PAJCHINAMPS,
            cls.calcdiscount(wgts, affdate), cls.calcmtlcost(wgts, PAJCHINAMPS))

    @classmethod
    def _checkargs(cls, incr, refmps, tarmps):
        """check if incr/mpss is valid"""
        if not incr:
            return True
        return (not incr.gold or incr.gold and tarmps.gold and refmps.gold) and \
            (not incr.silver or incr.silver and tarmps.silver and refmps.silver)

    @classmethod
    def calchina(cls, prdwgt, refup, refmps, affdate=None, tarmps=None):
        """ calculate the China cost based on the provided arguments

        Args:

            prdwgt: weights of the product

            refup: the reference unit price

            refmps: the reference mps of the @refup

            tarmps=None: the target MPS the need to be calculated. PAJ's china MPS is S = 30; G = 1500

            affdate=None: date the refup belongs to

        Returns:

            a PajChina Object
        """
        if isinstance(refmps, str):
            refmps = MPS(refmps)
        if not (all((prdwgt, refup, refmps)) and (refup > 0 and refmps.isvalid)):
            return None
        if not tarmps:
            tarmps = PAJCHINAMPS
        for x in (tarmps, refmps):
            if isinstance(x, str):
                x = MPS(x)
        if isinstance(refup, Decimal):
            refup = float(refup)

        # the discount ratio, when there is silver, follow silver, silver = 0.9 while gold = 0.85
        incr = cls.calcincrement(prdwgt, None, "PAJ")
        dc = cls.calcdiscount(prdwgt, affdate)
        if not cls._checkargs(incr, refmps, tarmps):
            cn = mc = MPSINVALID
            logger.debug("MPS(%s) not enough for calculating increment(%s)" %
                         (tarmps.value, str(incr)))
        else:
            mc = 0
            for cat in ('gold', 'silver'):
                wgt = getattr(incr, cat)
                if not wgt:
                    continue
                wgt *= cls._markup
                mc += wgt * (getattr(refmps, cat) - getattr(tarmps, cat))
            cn = refup / dc - mc
            mc = cls.calcmtlcost(prdwgt, tarmps, lossrate=incr.lossrate, vendor="PAJ")
        return PajChina(round(cn, 2), incr, tarmps, dc, mc)

    @classmethod
    def calctarget(cls, cn, tarmps, affdate=None):
        """calculate the target unit price based on the data _NotProvided

        Args:

            cn: an PajChina instance

            tarmps: the target MPS

            affdate=None: date for this calculation

        Returns:

            an PajChina object whose china is the target cost
        """

        if isinstance(tarmps, str):
            tarmps = MPS(tarmps)
        incr = cn.increment
        if not cls._checkargs(incr, cn.mps, tarmps):
            r0 = mc = MPSINVALID
            logger.debug("MPS(%s) not enough for calculating increment(%s)" %
                         (tarmps.value, str(incr)))
        else:
            r0 = cn.china + (tarmps.gold - cn.mps.gold) * incr.gold * cls._markup \
                + (tarmps.silver - cn.mps.silver) * incr.silver * cls._markup
            dc = cls.calcdiscount(incr.wgts, affdate)
            r0 = round(r0 * dc, 2)
            mc = cls.calcmtlcost(incr.wgts, tarmps, lossrate=incr.lossrate, vendor="PAJ")
        return PajChina(r0, cn.increment, tarmps, dc, mc)
