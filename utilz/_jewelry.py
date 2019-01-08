#! coding=utf-8
'''
* @Author: zmFeng
* @Date: 2018-07-01 11:38:33
* @Last Modified by:   zmFeng
* @Last Modified time: 2018-07-01 11:38:33
'''
from collections import namedtuple
from csv import DictReader
from numbers import Number
from operator import attrgetter
from os import path
from threading import RLock

from ._miscs import splitarray, trimu
from .common import thispath

__all__ = ["Karat", "KaratSvc", "RingSizeSvc", "stsizefmt"]


class _SzFmt(object):
    ''' helper class for stsizefmt function '''
    _don_touch = {"0000": 0.69, "000": 0.79, "00": 0.89, "0": 0.99}
    def format(self, sz, shortform=False):
        ''' @refer to stsizefmt function '''
        if not sz:
            return None
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
        for x in [x for x in segs if x]:
            x = self._fmtpart(x, shortform)
            if not x:
                continue
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


Karat = namedtuple("Karat", "karat,name,fineness,category,color")


class KaratSvc(object):
    """ Karat service for id/name resolving and fineness and so on """
    CATEGORY_GOLD = "GOLD"
    CATEGORY_SILVER = "SILVER"
    CATEGORY_BRONZE = "BRONZE"
    CATEGORY_BONDEDGOLD = "BG"

    COLOR_WHITE = "WHITE"
    COLOR_YELLOW = "YELLOW"
    COLOR_ROSE = "ROSE"
    COLOR_BLACK = "BLACK"
    COLOR_BLUE = "BLUE"

    _priorities = {
        CATEGORY_BRONZE: -100,
        CATEGORY_SILVER: -50,
        CATEGORY_BONDEDGOLD: -10,
        CATEGORY_GOLD: 10
    }
    """ class help to solve karat related issues """

    def __init__(self, fn=None):
        if not fn or not path.exists(fn):
            fn = path.join(thispath, "res", "karats.csv")
        lst = []
        with open(fn, "rt") as fh:
            rdr = DictReader(fh)
            for x in rdr:
                kt = x["karat"]
                if kt.isdigit():
                    kt = int(kt)
                fin = float(x["fineness"])
                if fin > 1.0:
                    fin = fin / 100
                lst.append(
                    Karat(kt, x["name"].strip(), fin, x["category"].strip(),
                          x["color"].strip()))
        byid, byname, fingrp, fml = {}, {}, {}, {}
        for x in lst:
            byid[x.karat] = byname[x.name] = x
            fin = x.fineness
            if fin < 100.0 and x.category == "GOLD":
                fingrp.setdefault(fin, []).append(x)

        for x in fingrp.values():
            y = sorted(x, key=attrgetter("category", "karat"))
            for it in y[1:]:
                fml[it.karat] = y[0]

        self._byid, self._byname, self._byfamily, self._byfineness = byid, byname, fml, None

    @property
    def all(self):
        """ return all the karat instances """
        return self._byid.values()

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


class RingSizeSvc(object):
    """ ring size converting between different standards """
    _szcht, _szgrp = None, None
    _rlck = RLock()

    @classmethod
    def _loadrgcht(cls):
        #if the file with BOM as first character, use utf-8-sig to open it
        with open(
                path.join(thispath, "res", "rszcht.csv"),
                "r+t",
                encoding="utf-8-sig") as fh:
            rdr = DictReader(fh)
            lst = list(rdr)
        #use a 2 layer dict to index the size chart
        d0 = {}
        for x in lst:
            for k in x:
                d0.setdefault(k, {})[x[k]] = x
        dg0 = {}
        with open(path.join(thispath, "res", "rszgrp.csv")) as fh:
            for x in fh.readlines():
                if x.startswith("#"):
                    continue
                ss = trimu(x).split("=")
                for yy in ss[1].split(","):
                    dg0[yy] = ss[0]
        return d0, dg0

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
