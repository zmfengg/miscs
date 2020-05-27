#! coding=utf-8
'''
* @Author: zmFeng
* @Date: 2018-07-01 11:38:33
* @Last Modified by:   zmFeng
* @Last Modified time: 2018-07-01 11:38:33
'''
import json
from collections import namedtuple
from numbers import Number
from operator import attrgetter
from os import path
from threading import RLock

from .miscs import splitarray, trimu
from .common import thispath

__all__ = ["Karat", "KaratSvc", "RingSizeSvc", "stsizefmt", "UnitCvtSvc"]


class _SzFmt(object):
    ''' helper class for stsizefmt function '''
    _don_touch = {"0000": 0.69, "000": 0.79, "00": 0.89, "0": 0.99}
    def format(self, sz, shortform=False):
        ''' @refer to stsizefmt function '''
        if not sz:
            return None
        if isinstance(sz, Number):
            sz = str(sz)
        if sz.find("~") > 0:
            sz = sz.replace("~", "-")

        sz = trimu(sz) if isinstance(sz, str) else str(sz)
        segs, parts, idx, rng = [""], [], 0, False
        for x in sz:
            if x.isdigit() or x == ".":
                segs[idx] += x
            elif x == "-":
                idx = self._inc(segs)
                rng = True
            elif x in ("X", "*"):
                idx = self._inc(segs)
                if rng:
                    break
            elif rng:
                break
        if not any(segs):
            return sz
        for x in (x for x in (self._fmtpart(x, shortform) for x in segs) if x):
            if isinstance(x, str):
                parts.append(x)
            else:
                parts.extend(x)
        return ("-" if rng else "X").join(sorted(
            parts, key=self._part_power, reverse=True)) if parts else None

    @classmethod
    def _inc(cls, segs):
        segs.append("")
        return len(segs) - 1

    @classmethod
    def _fmt_no_digit(cls, val):
        if int(val) == val:
            return "%d" % val
        return "%r" % val

    @classmethod
    def _fmtpart(cls, s0, shortform):
        if not s0 or s0 in cls._don_touch:
            return s0
        try:
            ln = len(s0)
            if ln < 4 or s0.find(".") >= 0:
                s0 = float(s0) * 100
                s0 = cls._fmt_no_digit(s0 / 100) if shortform else "%04d" % s0
            else:
                s0 = splitarray(s0, 4)
                if shortform:
                    for ii, it in enumerate(s0):
                        s0[ii] = cls._fmt_no_digit(int(it) / 100)
        except:
            s0 = None
        return s0

    @classmethod
    def _part_power(cls, s0):
        return cls._don_touch.get(s0) or float(s0)


_st_fmtr = _SzFmt()


def stsizefmt(sz, shortform=False):
    """ format a stone size into long or short form, with big -> small sorting, some examples are
    @param sz: the string to format
    @param shortform: return a short format
        "3x4x5mm" -> "0500X0400X0300"
        "3x4x5" -> "0500X0400X0300"
        "3.5x4.0x5.3" -> "0530X0400X0350"
        "4" -> "0400"
        "053004000350" -> "0530X0400X0350"
        "040005300350" -> "0530X0400X0350"
        "0400X0530X0350" -> "0530X0400X0350"
        "4m" -> "0400"
        "4m-3.5m" -> "0400-0350"
        "3x4x5", False, True -> "5X4X3"
        "0500X0400X0300" -> "5X4X3"
        "0300X0500X0400" -> "5X4X3"
        "0000" -> "0000"
        "000" -> "000"
    """
    return _st_fmtr.format(sz, shortform)


Karat = namedtuple("Karat", "karat,fineness,density,name,category,color")


class KaratSvc(object):
    """
    Karat service for id/name resolving and fineness and so on
    also support querying by below names:
    CATEGORY_GOLD/CATEGORY_SILVER
    COLOR_YELLOW
    """
    def __init__(self, fn=None):
        if not fn or not path.exists(fn):
            fn = path.join(thispath, "res", "settings.json")
        settings = None
        with open(fn, "rt") as fh:
            settings = json.load(fh)
        if not settings:
            return
        settings, fn = settings["KaratSvc"], []
        for x in settings["karats"]:
            byid = [x.get(y) for y in "karat fineness density".split()]
            if byid[1] > 1.0:
                byid[1] = byid[1] / 100
            byid.extend((x.get(y).strip() for y in "name category color".split()))
            fn.append(Karat(*byid))
        self._priorities = settings["_priorities"]
        self._byid, self._byname, fingrp, self._byfamily, self._cats, self._colors = ({} for x in range(6))
        for x in fn:
            self._byid[x.karat] = self._byname[x.name] = x
            fin = x.fineness
            if fin < 100.0 and x.category == "GOLD":
                fingrp.setdefault(fin, []).append(x)
            self._cats[x.category] = x.category
            self._colors[x.color] = x.color
        for mp, x in zip((self._cats, self._colors), tuple(settings.get(x) for x in "category_alias color_alias".split())):
            if not x:
                continue
            mp.update(x)
        for x in fingrp.values():
            y = sorted(x, key=attrgetter("category", "karat"))
            for it in y[1:]:
                self._byfamily[it.karat] = y[0]
        self._byfineness = None

    @property
    def all(self):
        """ return all the karat instances """
        return self._byid.values()

    @property
    def COLORS(self):
        '''
        all the available colors inside me
        '''
        return set(self._colors.values())

    @property
    def CATEGORIES(self):
        '''
        all the categories inside me
        '''
        return set(self._cats.values())

    def __getitem__(self, key):
        return self.getkarat(key)

    def getkarat(self, karat):
        """
            return the karat object by id or by name
            for example, getkarat(8) or getkarat("8K")
        """
        if isinstance(karat, str):
            if karat.isdigit():
                karat = int(karat)
            else:
                karat = trimu(karat).replace("KY", "K")
            if karat == 'COPPER':
                karat = 'BRONZE'
        for x in (self._byid, self._byname):
            if karat in x:
                return x[karat]
        return None

    def getbyfineness(self, fineness):
        """ fineness must be an integer, the actual fineness * 1000, if not, I do it for you """
        if isinstance(fineness, Number):
            if fineness < 0:
                fineness = int(fineness * 1000)
            if not self._byfineness:
                self._byfineness = {x.fineness * 1000: x for x in self.all}
            if fineness in self._byfineness:
                return self.getfamily(self._byfineness[fineness])
        return None

    def getfamily(self, karat):
        """ the legacy karat issue: 9 -> 91 -> 98 10 -> 101 -> 108 ... """
        if karat and not isinstance(karat, Karat):
            karat = self.getkarat(karat)
        if not karat:
            return None
        if karat.karat in self._byfamily:
            karat = self._byfamily[karat.karat]
        return karat

    def issamecategory(self, k0, k1):
        """
        check if given karats are of the same karat
        @param k0: a Karat instance or int/str type of karat
        """
        kx = [x if isinstance(x, Karat) else self[x] for x in (k0, k1)]
        return kx[0].category == kx[1].category if all(kx) else None

    def convert(self, k0, wgt0, k1):
        """
        convert the wgt0 in k0 to k1
        Args:
            k0: source Karat object or a name of karat
            wgt0: source weight
            k1: target Karat object or a name of karat
        Returns:
            the target weight if conversion is OK, else 0 or None
        """
        k0, k1 = [x if isinstance(x, Karat) else self.getkarat(x) for x in (k0, k1)]
        if not all((k0, k1)):
            return None
        wgt1 = wgt0 * k1.density / k0.density if k0.density else None
        return round(wgt1, 2) if wgt1 else None

    def compare(self, k0, k1):
        """ check if 2 given karat are the same, Only same category items can be compared
        return:
            1 if k0.finess > k1.finess
            -1 if k0.finess < k1.finess
            0 if k0 is k1
        """
        if k0 is k1:
            return 0
        tcs = (k0, k1)
        cps = [self._priorities[x.category] for x in tcs]
        rc = cps[0] - cps[1]
        if rc == 0:
            rc = tcs[0].fineness - tcs[1].fineness
            if rc == 0:
                rc = k0.karat - k1.karat
        rc = 1 if rc > 0 else -1 if rc < 0 else 0

        return rc

    def __getattribute__(self, name):
        if name.startswith("CATEGORY_"):
            return self._cats[name[9:]]
        if name.startswith("COLOR_"):
            return self._colors[name[6:]]
        return super().__getattribute__(name)

class RingSizeSvc(object):
    """ ring size converting between different standards """
    _szcht, _szgrp = None, None
    _rlck = RLock()

    @classmethod
    def _loadrgcht(cls):
        #if the file with BOM as first character, use utf-8-sig to open it
        '''
        with open(
                path.join(thispath, "res", "rszcht.csv"),
                "r+t",
                encoding="utf-8-sig") as fh:
            rdr = DictReader(fh)
            lst = list(rdr)
        '''
        mp = None
        with open(path.join(thispath, "res", "settings.json")) as fh:
            mp = json.load(fh)["RingSizeSvc"]
        if not mp:
            return None
        d0 = {}
        for k, x in [(k, x) for x in mp["sizes"] for k in x]:
            d0.setdefault(k, {})[x[k]] = x
        mp = {x: k for k, v in mp["groups"].items() for x in v.split(",")}
        return d0, mp

    def _getgrp(self, cn):
        self._rlck.acquire()
        try:
            if not self._szcht:
                self._szcht, self._szgrp = RingSizeSvc._loadrgcht()
        except:
            pass
        finally:
            self._rlck.release()
        cn = trimu(cn)
        return cn if cn in self._szcht else None if cn not in self._szgrp else self._szgrp[
            cn]

    def _getitem(self, cn0, sz0):
        cn0 = self._getgrp(cn0)
        if not cn0:
            return None
        d0, sz0 = self._szcht[cn0], trimu(sz0)
        if sz0 not in d0:
            return None
        return d0[sz0]

    def convert(self, cn0, sz0, cn1):
        """ convert ring size between different standards
        @param cn0: the country name sth. like "US","HK"
        @param sz0,sz1: the size code
        """
        it = self._getitem(cn0, sz0)
        if not it:
            return None
        cn1 = self._getgrp(cn1)
        if not cn1:
            return None
        sz1 = it[cn1]
        if sz1 == "-":
            sz1 = None
        return sz1

    def getcirc(self, cn, sz):
        """
        return the ring's circumference of a ring size. EU's size is the circumference in mm
        @param cn: the country code, sth. like "US","EU","CN","HK"
        @param sz: the ring size
        """
        sz1 = self.convert(cn, sz, "EU")
        if not sz1:
            return None
        return float(sz1)

class UnitCvtSvc(object):
    '''
    do unit conversion, mainly for weight
    '''

    def __init__(self):
        super().__init__()
        self._cats = self._cvtrs = None


    def _load_data(self):
        with open(path.join(thispath, "res", "settings.json")) as fh:
            mp0 = json.load(fh)["UnitCvtSvc"]
            # use override stagery
            if self._cats is None:
                self._cats = {}
                self._cvtrs = {}
            mp = self._cvtrs
            n2c = self._cats
            for n, cnp in mp0.items():
                n = self._nrl(n)
                cnp[0] = self._nrl(cnp[0])
                mp[self._key(cnp[0], n)] = cnp[1]
                if cnp[1] == 1:
                    mp[self._key(cnp[0], '__ORIGINE')] = n
                n2c[n] = cnp[0]

    def _key(self, category, name):
        return category + '.' + name

    def _nrl(self, n):
        return trimu(n)

    def add(self, name, category, weight):
        '''
        add def
        Args:
            name(String): name of the unit
            category(String): the category of that name
            weight(float): the weight to standard value
        '''
        if self._cats is None:
            self._load_data()
        name = self._nrl(name)
        category = self._nrl(category)
        self._cats[name] = category
        self._cvtrs[self._key(category, name)] = weight


    def convert(self, val, uFrm, uThru):
        '''
        convert given value from uFrm to uThru
        Args:
            uFrm: the unit from
            uThru: the unit to
        '''
        if self._cvtrs is None:
            self._load_data()
        fnt = [self._nrl(x) for x in (uFrm, uThru)]
        if fnt[0] == fnt[1]:
            return val
        cats = [self._cats.get(x) for x in fnt]
        if not all(cats):
            raise OverflowError('unit(%s) not defined by config file' % ",".join(fnt[i] for i in range(len(cats)) if not cats[i]))
        if cats[0] != cats[1]:
            raise TypeError('category of %s is %s, not %s of %s' % (uFrm, cats[0], cats[1], uThru))
        return val * self._cvtrs['%s.%s' % (cats[0], fnt[0])] / self._cvtrs['%s.%s' % (cats[0], fnt[1])]
