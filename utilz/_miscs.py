#! coding=utf-8
'''
* @Author: zmFeng 
* @Date: 2018-06-16 15:44:32 
* @Last Modified by:   zmFeng 
* @Last Modified time: 2018-06-16 15:44:32 
'''

from collections import OrderedDict
from collections.abc import Iterator, Sequence
from math import ceil
from numbers import Integral
from os import listdir, path
from random import random
import re
from sys import getfilesystemencoding, version_info

from sqlalchemy.orm import Session

from .common import _logger as logger

__all__ = ["NamedList", "NamedLists", "appathsep", "daterange", "deepget", "getfiles", "isnumeric", "list2dict", "na", "splitarray", "stsizefmt", "triml", "trimu", "removews"]

na = "N/A"

def splitarray(arr, logsize=100):
    """split an array into arrays whose len is less or equal than logsize
    @param arr: the sequence object that need to split
    @param logsize: len of each sub-array's size  
    """
    if not arr: return
    if not (isinstance(arr,tuple) or isinstance(arr,list) or isinstance(arr,str)):arr = tuple(arr)
    if not logsize:
        logsize = 100
    return [arr[x * logsize:(x + 1) * logsize] for x in range(int(ceil(1.0 * len(arr) / logsize)))]


def isnumeric(val):
    flag = True
    try:
        float(val)
    except:
        flag = False
    return flag


def appathsep(fldr):
    """append a path sep into given path if there is not"""
    return fldr + path.sep if fldr[len(fldr) - 1:] != path.sep else fldr


def list2dict(lst, trmap=None, dupdiv="", bname=None):
    """ turn a list into zero-id based, name -> id lookup map 
    @param lst: the list or one-dim array containing the strings that need to do the name-> pos map
    @param trmap: An translation map, make the description -> name translation, if ommitted, description become name
                  if the description is not sure, split them with candidates, for example, "jono":"Job,JS"
    @param dupdiv: when duplicated item found, a count will be generated, dupdiv will be
        placed between the original and count
    @param bname: default name for the blank item
    @return: a dict with name -> id map   
    """
    if not lst:
        return None, None

    lstl = [triml(x) for x in lst]
    mp = {}
    for ii in range(len(lstl)):
        x = lstl[ii]
        if not x and bname:
            lstl[ii] = bname
        if x in mp:
            mp[x] += 1
            if dupdiv == None:
                dupdiv = ""
            if lstl[ii]:
                lstl[ii] += dupdiv + str(mp[x])
            else:
                lstl[ii] = dupdiv + str(mp[x])
        else:
            mp[x] = 0
    if not trmap:
        trmap = {}
    else:
        trmap = dict([(triml(x[1]),x[0]) for x in trmap.items()])
        for x in [x for x in trmap.keys() if(x.find(",") >= 0)]:
            for y in x.split(","):
                if not y:
                    continue
                y = y.lower()
                cnds = [x0 for x0 in range(len(lstl)) if lstl[x0] and lstl[x0].find(y) >= 0]
                if(len(cnds) > 0):
                    s0 = str(random())
                    lstl[cnds[0]] = s0
                    trmap[s0] = triml(trmap[x])
                    break
    return OrderedDict(zip([trmap[x] if(x in trmap) else x for x in lstl], range(len(lstl))))


def deepget(obj, names):
    """ get deeply from the object """
    #gtr, rc = object.__getattribute__ if version_info.major >= 3 else object.__getattr__, obj
    gtr, rc =  getattr, obj
    for k in names.split("."):
        rc = gtr(rc,k)
    return rc


def getfiles(fldr, part=None, nameonly=False):
    """ return files under given folder """
    """ @param nameonly : don't return the full-path """

    if fldr:
        fldr = appathsep(fldr)
        if part:
            part = part.lower()
            fns = [x if version_info.major >= 3 else str(x, getfilesystemencoding())
                   for x in listdir(fldr) if x.lower().find(part) >= 0]
        else:
            fns = [x if version_info.major >= 3 else str(x, getfilesystemencoding())
                   for x in listdir(fldr)]
        if not nameonly:
            fns = [fldr + x for x in fns]
    return fns


def daterange(year, month, day=1):
    """ make a from,thru tuple for the given month, thru is the first date of next month """
    import datetime as dtm
    df = dtm.date(year, month, day if day > 0 else 1)
    month += 1
    if month > 12:
        year += 1
        month = 1
    dt = dtm.date(year, month, 1)
    del dtm
    return df, dt


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
    """
    def _inc(segs):
        segs.append("")
        return len(segs) - 1

    def _fmtpart(s0, shortform):
        ln = len(s0)
        if ln < 4 or s0.find(".") >= 0:
            s0 = "%04d" % (float(s0) * 100)
            if shortform:
                s0 = "%d" % (int(s0) / 100)
        else:
            s0 = splitarray(s0, 4)
            if shortform:
                for ii in range(len(s0)):
                    s0[ii] = "%d" % (int(s0[ii]) / 100)
        return s0

    sz = sz.strip().upper()
    segs, parts, idx, rng = [""], [], 0, False
    for x in sz:
        if x.isdigit() or x == ".":
            segs[idx] += x
        elif x == "-":
            idx = _inc(segs)
            rng = True
        elif x in ("X", "*"):
            idx = _inc(segs)
            if rng:
                break
        elif rng:
            break
    for x in segs:
        x = _fmtpart(x, shortform)
        if isinstance(x, str):
            parts.append(x)
        else:
            parts.extend(x)
    return ("-" if rng else "X").join(sorted(parts, reverse=True))

def removews(s0):
    """ remove the invalid white space """
    if not s0: return
    return re.sub(r"\s{2,}", " ", s0.strip())


def trimu(s0, removewsps = True):
    """ trim/strip and upper case """
    if s0 and isinstance(s0, str):
        s0 = s0.strip().upper()
        if removewsps: s0 = removews(s0)
    return s0


def triml(s0, removewsps = True):
    """ trim and lower case """
    if s0 and isinstance(s0, str):
        s0 = s0.strip().lower()
        if removewsps: s0 = removews(s0)
    return s0


class NamedList(object):
    """
    the wrapper of the list/tuple that make it operatable by .name or [name] or [i]
    self._dtype:
        0 -> not data set
        1 -> data is list
        2 -> data is dict
        10 -> data is object
        ... your turn to extend me
    """


    def __init__(self, nmap, data=None):
        if isinstance(nmap,tuple) or isinstance(nmap,list):
            nmap = list2dict(nmap)
        elif isinstance(nmap,str):
            nmap = list2dict(nmap.split(","))
        elif isinstance(nmap,dict):
            nmap = dict([(triml(x[0]),x[1]) for x in nmap.items()])
        self._nmap, self._idmap = nmap, None
        self._dtype = 0
        if data:
            self.setdata(data)

    def clone(self, data = None):
        """ create a clone with the same definination as me, but not the same data set """
        return NamedList(self._nmap, data)

    def setdata(self, data):
        if data:
            if isinstance(data, Sequence) and len(self._nmap) != len(data):
                data = None
        if not data:
            self._dtype = 0
        else:
            self._dtype = 1 if isinstance(data, Sequence) else 2 if isinstance(data, dict) else 10
        self._data = data
        return self

    def _checkarg(self, name):
        if not (self._dtype and (self._dtype != 1 or name in self._nmap)):
            raise AttributeError("no attribute(%s) found or data not set" % name)

    def __getattr__(self, name):
        name = triml(name)
        if self._dtype == 1:
            return self._data[self._nmap[name]]
        elif name in self._nmap:
            name = self._nmap[name]
        return self._data[name] if self._dtype == 2 else getattr(self._data, name)

    def __setattr__(self, name, val):
        #self._checkarg(name)
        if name.startswith("_"):            
            object.__setattr__(self, name, val)
        else:
            name = triml(name)
            self._checkarg(name)
            if self._dtype == 1:
                self._data[self._nmap[name]] = val
            else:
                if name in self._nmap:
                    name = self._nmap[name]
                if self._dtype == 2:
                    self._data[name] = val
                else:
                    setattr(self._data, name, val)

    def __getitem__(self, key):
        if isinstance(key, slice) or isinstance(key, Integral):
            return self._data[key]
        return self.__getattr__(key)

    def __setitem__(self, key, val):
        if isinstance(key, Integral):
            self._data[key] = val
        else:
            self.__setattr__(key, val)

    def _mkidmap(self):
        if not self._idmap:
            self._idmap = dict([x[1],x[0]] for x in self._nmap.items())

    def get(self,kon,default = None):
        """ simulate the dict's get function, for easy life only """
        rc = default
        try:
            rc = self[kon]
        except:
            pass
        return rc

    def getcol(self, nameorid):
        """
        return colname ->  colid or colid -> colname
        """
        if isinstance(nameorid,str):
            rc = self._nmap.get(triml(nameorid),None)
        else:
            self._mkidmap()
            rc = self._idmap.get(nameorid,None)
        return rc

    @property
    def _colnames(self):
        return tuple(self._nmap.keys())

    @property
    def _colids(self):
        self._mkidmap()
        return tuple(self._idmap.keys())

    @property
    def data(self):
        return self._data


class NamedLists(Iterator):
    """ 
    make a list of list(2d array) accessable by name, for example, you read data from a csv
    lsts = (("id","name","price"),(1,"Jan",23.45),(2,"Pet",30.25)), you don't want to get id by
        lsts[0][0] 
        or 
        nmap = dict([(lsts[0][idx],idx) for x in range(len(lsts[0]))])
        lsts[0][nmap["id"]]

    Use this as:
        its = NamedLists(lsts):
        for x in its:
            id = x.id...

    """

    def __init__(self, lsts, trmap=None, newinst=True):
        """ 
        init one named list instance
        @param lsts: the list(or tuple) of a list(or tuple, but when it's a tuple, you can not assigned value)
            always send the title rows to the first item
        @param trmap: nmap translation map. used when nmap == None and you want to do some name tranlation
                    @refer to list2dict for more info.
        @param newinst: set this to False if you use "for" loop to save memory
            set it to True if you use lst = [x for x in nl] or lst = list(nl).
            for safe reason, it's True by default
        """
        super(NamedLists, self).__init__()
        nmap = list2dict(lsts[0], trmap)
        lsts = lsts[1:]
        self._lsts, self._nmap, self._ptr, self._ubnd, self._newinst = lsts, nmap, \
            -1, len(lsts), newinst
        if not newinst:
            self._wrpr = NamedList(nmap)

    def __iter__(self):
        return self

    def __next__(self):
        self._ptr += 1
        if not self._lsts or self._ptr >= self._ubnd:
            raise StopIteration()
        if self._newinst:
            return NamedList(self._nmap, self._lsts[self._ptr])
        else:
            self._wrpr.setdata(self._lsts[self._ptr])
            return self._wrpr

    @property
    def namemap(self):
        return self._nmap