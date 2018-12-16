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
import imghdr
import struct
from sys import getfilesystemencoding, version_info

import tkinter as tk

__all__ = ["NamedList", "NamedLists", "appathsep", "daterange", "deepget", "getfiles", "getvalue", "isnumeric",
           "imagesize", "list2dict", "lvst_dist", "na", "splitarray", "triml", "trimu", "updateopts", "removews", "easydialog", "easymsgbox"]

na = "N/A"

_jpgsof = {192, 193, 194, 195, 197, 198, 199, 201, 202, 203, 205, 206, 207}

def _norm_exp(the_mode):
    if not the_mode:
        the_mode = lambda x: x.strip() if x is not None else x
    elif the_mode.lower().find("low") >= 0:
        the_mode = triml
    else:
        the_mode = trimu
    return the_mode

def splitarray(arr, logsize=100):
    """split an array into arrays whose len is less or equal than logsize
    @param arr: the sequence object that need to split
    @param logsize: len of each sub-array's size
    """
    if not arr:
        return []
    if not isinstance(arr, (tuple, list, str)):
        arr = tuple(arr)
    if not logsize:
        logsize = 100
    return [arr[x * logsize:(x + 1) * logsize] for x in range(int(ceil(1.0 * len(arr) / logsize)))]


def getvalue(dct, key, def_val=None):
    """
    @param key: the keyname you want to get from the dict, can not contain ,. When , is found, will be treated as 2 keywords
    get the dict value by below seq:
        normal -> trimu -> triml
    """
    for kw in key.split(","):
        i = 0
        while i < 3:
            if kw in dct:
                return dct.get(kw)
            i += 1
            kw = trimu(kw) if kw[0] == kw[0].lower() else triml(kw)
    return def_val

def isnumeric(val):
    """
    check if given val is a numeric
    """
    flag = True
    try:
        float(val)
    except:
        flag = False
    return flag


def appathsep(fldr):
    """
    append a path sep into given path if there is not
    """
    return fldr + path.sep if fldr[len(fldr) - 1:] != path.sep else fldr

def updateopts(defaults, kwds):
    """
    return a dict which merge kwds and defaults's value, if neither, the item value is None
    @param defaults: dict, an example, {"name": ("alias1,alias2", SomeValue)}
    @param kwds: dict, always put the one you accepted from your function
    """
    if not any((defaults, kwds)):
        return None
    if not defaults:
        return kwds
    if not kwds:
        kwds = {}
    for knw in defaults.items():
        its = [x for x in knw[1][0].split(",") if x in kwds]
        if not its:
            kwds[knw[0]] = knw[1][1]
        else:
            kwds[knw[0]] = kwds[its[0]]
            del kwds[its[0]]
    return kwds


def list2dict(lst, **kwds):
    """ turn a list into zero-id based, name -> id lookup map
    @param lst: the list or one-dim array containing the strings that need to do the name-> pos map
    @param trmap: An translation map, make the description -> name translation, if ommitted, description become name
                  if the description is not sure, split them with candidates, for example, "jono":"Job,JS"
    @param dupdiv: when duplicated item found, a count will be generated, dupdiv will be placed between the original and count
    @param bname: default name for the blank item
    @param normalize: can be one of upper/lower/(blank or None), do keyword normalization using lower/upper or no normalization, default is lower
    @return: a dict with name -> id map
    """
    if not lst:
        return None
    if isinstance(lst, str):
        lst = lst.split(",")
    else:
        lst = tuple(str(x) if x is not None else "" for x in lst)
    mp = updateopts({"dupdiv": ("dupdiv,div,dup_div", ""), "trmap": ("name_map,trmap,alias", None), "bname": ("bname,blank_name", None), "normalize": ("normalize,", "lower")}, kwds)
    dupdiv, bname, trmap, _norm = getvalue(mp, "dupdiv,div,dup_div"), mp.get("bname"), getvalue(mp, "trmap,alias"), getvalue(mp, "normalize,norm")
    if dupdiv is None:
        dupdiv = ""
    lst_lower, ctr = [], {}
    _norm = _norm_exp(_norm)
    for x in (_norm(x) or bname or "" for x in lst):
        if x in ctr:
            ctr[x] += 1
            x += dupdiv + str(ctr[x])
        else:
            ctr[x] = 0
        lst_lower.append(x)
    trmap = {_norm(x[1]): _norm(x[0]) for x in trmap.items()} if trmap else {}
    ctr = list(range(len(lst_lower)))
    if trmap:
        for x in [x for x in trmap if x.find(",") > 0]:
            for y in [x1 for x1 in x.split(",") if x1]:
                cnds = [x for x in ctr if lst_lower[x].find(y) >= 0]
                if not cnds:
                    continue
                s0 = str(random())
                lst_lower[cnds[0]] = s0
                trmap[s0] = trmap[x]
                break
    return OrderedDict(zip([trmap.get(x, x) for x in lst_lower], ctr))


def deepget(obj, names):
    """ get deeply from the object """
    #gtr, rc = object.__getattribute__ if version_info.major >= 3 else object.__getattr__, obj
    gtr, rc = getattr, obj
    for k in names.split("."):
        rc = gtr(rc, k)
    return rc


def imagesize(fn):
    '''detemine jpeg/png/gif/bmp's dimension
    code from https://stackoverflow.com/questions/8032642/how-to-obtain-image-size-using-standard-python-class-without-using-external-lib'''
    with open(fn, 'rb') as fhandle:
        head = fhandle.read(26)
        if len(head) != 26:
            return None
        itp = imghdr.what(fn)
        if not itp:
            return None

        def _sz_png():
            return struct.unpack('>ii', head[16:24]) if struct.unpack('>i', head[4:8])[0] == 0x0d0a1a0a else None

        def _sz_gif():
            return struct.unpack('<HH', head[6:10])

        def _sz_bmp():
            sig = head[:2].decode("ascii")
            if sig == "BM":  # Microsoft
                sig = struct.unpack("<II", head[18:26])
            else:  # IBM
                sig = struct.unpack("<HH", head[18:22])
            return sig

        def _sz_jpeg():
            try:
                ftype = 0
                fhandle.seek(0, 0)
                trunksz = 4096
                brs, ptr, offset = fhandle.read(trunksz), 0, 2
                while ftype not in _jpgsof:
                    ptr += offset
                    offset = ptr - len(brs)
                    if offset >= 0:
                        fhandle.seek(offset, 1)
                        brs, ptr = fhandle.read(trunksz), 0
                    while brs[ptr] == 0xff:
                        ptr += 1
                    ftype = brs[ptr]
                    offset = struct.unpack('>H', brs[ptr + 1:ptr + 3])[0] + 1
                rc = struct.unpack('>HH', brs[ptr+4: ptr+8])
                rc = (rc[1], rc[0])
            except:  # IGNORE:W0703
                rc = None
            return rc

        ops = {"png": _sz_png, "gif": _sz_gif, "bmp": _sz_bmp, "jpeg": _sz_jpeg}
        rc = ops.get(itp)
        if not rc:
            return None
        return rc()

def getfiles(fldr, part=None, nameonly=False):
    """
    return files under given folder
    @param nameonly : don't return the full-path
    """

    if fldr and path.exists(fldr):
        if part:
            part = part.lower()
            fns = [x if version_info.major >= 3 else str(x, getfilesystemencoding())
                   for x in listdir(fldr) if x.lower().find(part) >= 0]
        else:
            fns = [x if version_info.major >= 3 else str(x, getfilesystemencoding())
                   for x in listdir(fldr)]
        if not nameonly:
            fns = [path.join(fldr, x) for x in fns]
        return fns
    return None


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


def removews(s0):
    """
    remove the white space
    """
    return re.sub(r"\s{2,}", " ", s0.strip()) if s0 else None


def trimu(s0, removewsps=True):
    """ trim/strip and upper case """
    if s0 and isinstance(s0, str):
        s0 = s0.strip().upper()
        if removewsps:
            s0 = removews(s0)
    return s0


def triml(s0, removewsps=True):
    """ trim and lower case """
    if s0 and isinstance(s0, str):
        s0 = s0.strip().lower()
        if removewsps:
            s0 = removews(s0)
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

    def __init__(self, nmap, data=None, **kwds):
        mp = updateopts({"normalize": ("normalize,", "lower")}, kwds)
        self._nrl = _norm_exp(mp.get("normalize"))
        if isinstance(nmap, (tuple, list)):
            nmap = list2dict(nmap, **kwds)
        elif isinstance(nmap, str):
            nmap = list2dict(nmap.split(","), **kwds)
        elif isinstance(nmap, dict):
            nmap = {self._nrl(x[0]): x[1] for x in nmap.items()}
        self._nmap, self._idmap, self._kwds, self._dtype = nmap, None, None, 0
        if data:
            self.setdata(data)

    def clone(self, data=None):
        """ create a clone with the same definination as me, but not the same data set """
        return NamedList(self._nmap, data)

    def _replace(self, trmap, data=None):
        """ do name replacing, return a new instance
        trmap has the same meaning of list2dict
        """
        th = tuple(zip(*[(x[0], x[1]) for x in self._nmap.items()]))
        th = (list(th[0]), th[1])
        for x in trmap.items():
            ss = x[1].split(",")
            if len(ss) > 1:
                for sx in ss:
                    hit = False
                    for ii in range(len(th[0])):
                        if th[0][ii].find(sx) >= 0:
                            th[0][ii] = x[0]
                            hit = True
                            break
                    if hit:
                        break
            else:
                for ii in range(len(th[0])):
                    if th[0][ii] == x[1]:
                        th[0][ii] = x[0]
                        break
        return NamedList(dict(zip(*th)), data if data else self._data)

    def setdata(self, val):
        """
        set the internal data(should be of tuple/list). always return myself
        so that you can use continous action
        """
        if val:
            if isinstance(val, Sequence) and len(self._nmap) != len(val):
                val = None
        if not val:
            self._dtype = 0
        else:
            self._dtype = 1 if isinstance(val, Sequence) else 2 if isinstance(val, dict) else 10
        self._data = val
        return self

    def _checkarg(self, name):
        if not (self._dtype and (self._dtype != 1 or name in self._nmap)):
            raise AttributeError("no attribute(%s) found or data not set" % name)

    def __getattr__(self, name):
        name = self._nrl(name)
        if self._dtype == 1:
            return self._data[self._nmap[name]]
        if name in self._nmap:
            name = self._nmap[name]
        return self._data[name] if self._dtype == 2 else getattr(self._data, name)

    def __setattr__(self, name, val):
        # self._checkarg(name)
        if name.startswith("_"):
            object.__setattr__(self, name, val)
        else:
            name = self._nrl(name)
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
        if isinstance(key, (slice, Integral)):
            return self._data[key]
        return getattr(self, key)

    def __setitem__(self, key, val):
        if isinstance(key, Integral):
            self._data[key] = val
        else:
            setattr(self, key, val)

    def _mkidmap(self):
        if not self._idmap:
            self._idmap = dict([x[1], x[0]] for x in self._nmap.items())

    def __contains__(self, key):
        return self.getcol(key) is not None

    def get(self, kon, default=None):
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
        if isinstance(nameorid, str):
            rc = self._nmap.get(self._nrl(nameorid))
        else:
            self._mkidmap()
            rc = self._idmap.get(nameorid, None)
        return rc

    @property
    def colnames(self):
        """
        return the column names(that you can use it to access me)
        """
        return tuple(self._nmap.keys())

    @property
    def colids(self):
        """
        return a tuple if int:column ids(that you can use to access me)
        """
        self._mkidmap()
        return tuple(self._idmap.keys())

    @property
    def data(self):
        """
        return the internal list/tuple
        """
        return self._data


    def __str__(self):
        return self.__repr__()

    def __repr__(self):
        if not self._data:
            return None
        return repr(dict(zip(self.colnames, self._data)))


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

    def __init__(self, lsts, trmap=None, newinst=True, **kwds):
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
        nmap = list2dict(lsts[0], alias=trmap, **kwds)
        lsts = lsts[1:]
        self._lsts, self._nmap, self._ptr, self._ubnd, self._newinst = lsts, nmap, \
            -1, len(lsts), newinst
        if not newinst:
            self._wrpr = NamedList(nmap, **kwds)
        self._kwds = kwds

    def __iter__(self):
        return self

    def __next__(self):
        self._ptr += 1
        if not self._lsts or self._ptr >= self._ubnd:
            raise StopIteration()
        if self._newinst:
            return NamedList(self._nmap, self._lsts[self._ptr], **self._kwds)
        self._wrpr.setdata(self._lsts[self._ptr])
        return self._wrpr

    @property
    def namemap(self):
        """
        return the translate naming map
        """
        return self._nmap

    def __str__(self):
        return self._lsts.__repr__() if self._lsts else None

    def __repr__(self):
        return self._lsts.__repr__() if self._lsts else None


def easydialog(dlg):
    """
    open a tk dialog and return sth. easily
    use dlg.show() works, but sometimes there is a background windows there
    so, use for better looking
    """
    rt = tk.Tk()
    rt.withdraw()
    dlg.master = rt
    rc = dlg.show()
    # non of quit()/destroy() can kill tk while shown in excel, mainloop() even make it non-reponsible
    rt.quit()
    # rt.mainloop()
    # rt.destroy()
    return rc

def easymsgbox(box, *args):
    r"""
    show a messagebox with provided arguments, common snippets, the only usage
    is to hide the master window:
    from tkinter import messagebox as mb
    rc = easymsgbox(mb.showinfo, "hello", "you")
    or
    rc = easymsgbox(mb.askyesno, "attention", "need to delete sth?")
    """
    rt = tk.Tk()
    rt.withdraw()
    rc = box(*args, master=rt)
    rt.quit()
    return rc

def lvst_dist(s, t):
    """
    calculate the minimum movement steps(LevenshteinDistance) from string s to string t
    """
    if not s:
        return t
    if not t:
        return s
    n, m = len(s), len(t)
    p = [x + 1 for x in range(n + 1)] #'previous' cost array, horizontally
    d = [0] * (n + 1) # cost array, horizontally

    for j in range(1, m + 1):
        d[0], t_j = j + 1, t[j - 1]
        for i in range(1, n + 1):
            d[i] = min(min(d[i - 1], p[i]) + 1, p[i - 1] + (0 if s[i - 1] == t_j else 1))
        p, d = d, p
    return p[n] - 1
