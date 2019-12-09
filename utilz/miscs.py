#! coding=utf-8
'''
* @Author: zmFeng
* @Date: 2018-06-16 15:44:32
* @Last Modified by:   zmFeng
* @Last Modified time: 2018-06-16 15:44:32
'''

import tkinter as tk
from collections import OrderedDict
from collections.abc import Iterator, Sequence
from datetime import date
from imghdr import what
from inspect import currentframe, getfile
from math import ceil
from numbers import Integral
from os import listdir, path, remove
from random import random
from re import sub
from struct import unpack
from sys import getfilesystemencoding, version_info
from copy import copy, deepcopy
from ._miscclz import Config, Salt, Number2Word, Literalize, Segments, NumericRange

_sh = _se = None

__all__ = [
    "NamedList", "NamedLists", 'config', "appathsep", "daterange", "deepget",
    "easydialog", "easymsgbox", "getfiles", "getvalue", "iswritable", "isnumeric",
    "imagesize", "list2dict", "lvst_dist", "monthadd", "na", "NA", "removews",
    "Config", "Salt", "shellopen", "Number2Word", "Literalize", "splitarray", "tofloat", "triml", "trimu", "updateopts"
]

NA = "N/A"
na = "n/a"

_jpg_offsets = {192, 193, 194, 195, 197, 198, 199, 201, 202, 203, 205, 206, 207}
# max date of a month
_max_dom = (31, 0, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31,)


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
    return [
        arr[x * logsize:(x + 1) * logsize]
        for x in range(int(ceil(1.0 * len(arr) / logsize)))
    ]

def getpath():
    '''
    Return the path of the given object. If no m is provided, return the caller's path

    path.dirname(getabsfile(method_name)) not so good, because it need a method and getabsfile lower the path. so I use getfile(frame)

    Args:
        m=None: One of non-built-in class/method/module/package
    '''
    return  path.abspath(path.dirname(getfile(currentframe().f_back)))

def newline(ln):
    ''' find the newline position(as negative or None)
    while reading from a text file, the CRLF/LF may need this
    '''
    if not ln:
        return None
    return -2 if ln[-2] == '\r' else -1 if ln[-1] == '\n' else None

def getmodule(pk_name):
    '''
    get the module of given package name. A example to use:
        getmodule(__package__)
    @pk_name:
        name of the package, simply feed by __package__
    '''
    from sys import modules
    try:
        return modules[pk_name or __package__]
    except:
        return None

def getvalue(dct, key, def_val=None):
    """
    @param key: the keyname you want to get from the dict, can not contain ,. When , is found, will be treated as 2 keywords
    get the dict value by below seq:
        normal -> trimu -> triml
    """
    dev = ',' if key.find(',') > 0 else ' '
    for kw in key.split(dev):
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

def shellopen(fns, exec=False):
    '''
    open given files in the explorer(works only under windows)
    @param fns:
        A list of files to be show in the explorer. Files should be in the same folder, or only files in the first folder will be shown
    @param exec:
        True to execute it(for example, if it's an excel file, using excel to open it)
    @return:
        -1 if platform not support
        0 if not argument not valid
        1 if shown
        2 if executed
    '''
    if not fns:
        return 0
    global _sh, _se
    if _sh is None:
        try:
            from win32com.shell import shell as _sh
            from win32api import ShellExecute as _se
        except:
            _sh, rc = 0, -1
    if _sh:
        rt = rt0 = None
        fids = []
        for fn in fns:
            rt = path.dirname(fn)
            if rt != rt0 and rt0:
                continue
            rt0 = rt
            idx = _sh.SHILCreateFromPath(fn, 0)[0]
            if idx:
                fids.append(idx)
        pid = _sh.SHILCreateFromPath(rt0, 0)[0]
        if pid:
            _sh.SHOpenFolderAndSelectItems(pid, fids)
            rc = 1
        if exec:
            for fn in fns:
                _se(0, None, fn, None, None, 0)
            rc = 2
    return rc

def iswritable(fn):
    '''
    check if the given file is writable. According to article in the pydoc, using os.access() might sometimes lead to security hold, use the try/except case (EAFP pattern)
    '''
    fp, flag, rv, fnx = (None, ) * 4
    try:
        if path.isdir(fn):
            fnx = path.join(fn, str(random()))
            fp, rv = open(fnx, "w"), True
        else:
            # open(fn, 'w') might create a file in os.getcwd() if it's not already there, so need special care
            if not path.exists(fn):
                rv, fnx = True, fn
            fp = open(fn, "w")
        flag = True
    except:
        pass
    finally:
        if fp:
            fp.close()
        if rv and fnx and path.exists(fnx):
            remove(fnx)
    return flag

def tofloat(val):
    ''' often try to convert a str to float, but using isnumeric() can not
    let those with "." go, so use this directly. This function is handy
    '''
    try:
        return float(val)
    except:
        return 0


def appathsep(fldr):
    """
    append a path sep into given path if there is not
    """
    return fldr + path.sep if fldr[len(fldr) - 1:] != path.sep else fldr


def updateopts(defaults, kwds):
    """
    return a dict which merge kwds and defaults's value, if neither, the item value is None
    @param defaults: dict, an example, {"name": ("alias1,alias2", SomeValue)} or
        {"name": value}
    @param kwds: dict, always put the one you accepted from your function
    """
    if not any((defaults, kwds)):
        return None
    if not defaults:
        return kwds
    if not kwds:
        kwds = {}
    for knw in defaults.items():
        tp = isinstance(knw[1], (tuple, list))
        its = [x for x in knw[1][0].split(",") if x in kwds] if tp else (knw[0], ) if knw[0] in kwds else None
        if not its:
            its = knw[1][1] if tp else knw[1]
            kwds[knw[0]] = its
        elif knw[0] != its[0]:
            # changed the name
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
    example:
        nl = list2dict('a,b,ck', trmap={'jono': 'c,k'})
        assertTrue(2, nl['jono'])
    """
    if not lst:
        return None
    if isinstance(lst, str):
        lst = lst.split("," if lst.find(',') > 0 else ' ')
    else:
        lst = tuple(str(x) if x is not None else "" for x in lst)
    mp = updateopts({
        "dupdiv": ("dupdiv,div,dup_div", ""),
        "trmap": ("name_map,trmap,alias", None),
        "bname": ("bname,blank_name", None),
        "normalize": ("normalize,", "lower")
    }, kwds)
    dupdiv, bname, trmap, _norm = (getvalue(mp, x) for x in ("dupdiv,div,dup_div", "bname", "trmap,alias", "normalize,norm"))
    if dupdiv is None:
        dupdiv = ""
    lst_nrm, idxr = [], {}
    _norm = _norm_exp(_norm)
    # generate serial# for those duplicated colname or the blank colname
    for x in (_norm(x) or bname or "" for x in lst):
        if x in idxr:
            idxr[x] += 1
            x += dupdiv + str(idxr[x])
        else:
            idxr[x] = 0
        lst_nrm.append(x)
    if not trmap:
        return OrderedDict((x, idx) for idx, x in enumerate(lst_nrm))
    return _tr_l2d(_norm, trmap, lst_nrm)


def _tr_l2d(_norm, trmap, lst_nrm):
    trmap = {_norm(x[1]): _norm(x[0]) for x in trmap.items()}
    ht = set()
    for cands in [x for x in trmap if x.find(",") > 0]:
        flag = False
        for elm in (x for x in cands.split(",") if x):
            for idx, s0 in enumerate(lst_nrm):
                if idx in ht or s0.find(elm) < 0:
                    continue
                lst_nrm[idx] = str(random())
                trmap[lst_nrm[idx]], flag = trmap[cands], True
                ht.add(idx)
                break
            if flag:
                break
    return OrderedDict((trmap.get(x, x), idx) for idx, x in enumerate(lst_nrm))


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
        itp = what(fn)
        if not itp:
            return None

        def _sz_png():
            return unpack('>ii', head[16:24]) if unpack(
                '>i', head[4:8])[0] == 0x0d0a1a0a else None

        def _sz_gif():
            return unpack('<HH', head[6:10])

        def _sz_bmp():
            sig = head[:2].decode("ascii")
            if sig == "BM":  # Microsoft
                sig = unpack("<II", head[18:26])
            else:  # IBM
                sig = unpack("<HH", head[18:22])
            return sig

        def _sz_jpeg():
            try:
                ftype = 0
                fhandle.seek(0, 0)
                trunksz = 4096
                brs, ptr, offset = fhandle.read(trunksz), 0, 2
                while ftype not in _jpg_offsets:
                    ptr += offset
                    offset = ptr - len(brs)
                    if offset >= 0:
                        fhandle.seek(offset, 1)
                        brs, ptr = fhandle.read(trunksz), 0
                    while brs[ptr] == 0xff:
                        ptr += 1
                    ftype = brs[ptr]
                    offset = unpack('>H', brs[ptr + 1:ptr + 3])[0] + 1
                rc = unpack('>HH', brs[ptr + 4:ptr + 8])
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

    Args:

    part=None:  part of the file name
    
    nameonly=False: True to return name only, else name with full-path
    """

    if not (fldr and path.exists(fldr)):
        return None
    _fn = lambda fn: fn if version_info.major >= 3 else\
            str(fn, getfilesystemencoding())
    fns = [_fn(x) for x in listdir(fldr)]
    if part and fns:
        part = part.lower()
        fns = [x for x in fns if x.lower().find(part) >= 0]
    if not nameonly:
        fns = [path.join(fldr, x) for x in fns]
    return fns


def daterange(year, month, day=1):
    """ make a from,thru tuple for the given month, thru is the first date of next month """
    df = date(year, month, day if day > 0 else 1)
    month += 1
    if month > 12:
        year += 1
        month = 1
    dt = date(year, month, 1)
    return df, dt


def _isleap(year):
    """ check if given year is a leap year
    based on https://en.wikipedia.org/wiki/Leap_year
    """
    lp = False
    if year % 4:
        return False
    if year % 100:
        lp = True
    elif not year % 400:
        lp = True
    return lp


def monthadd(d0, months):
    """
    add months(negative/postive) to given date(d0), attention, add 1 month
    to dates like 2018/01/31 will return 2018/02/31, which is not a valid date
    Here I follow VBA's result, return 2018/02/28
    """
    m = d0.month - 1 + months
    y, m = d0.year + m // 12, m % 12 + 1
    d = _max_dom[m - 1]
    if not d:
        d = 29 if _isleap(y) else 28
    d = min(d, d0.day)
    if isinstance(d0, date):
        return date(y, m, d)
    return d0.replace(year=y, month=m, day=d)


def removews(s0):
    """
    remove the white space
    """
    return sub(r"\s{2,}", " ", s0) if s0 else ""


def trimu(s0, removewsps=False):
    """ trim/strip and upper case """
    if not s0: #and isinstance(s0, str):
        return None
    s0 = s0.strip().upper()
    if removewsps:
        s0 = removews(s0)
    return s0


def triml(s0, removewsps=False):
    """ trim and lower case """
    if s0 and isinstance(s0, str):
        s0 = s0.strip().lower()
        if removewsps:
            s0 = removews(s0)
    return s0

def sign(val):
    return 0 if not val else 1 if val > 0 else -1

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
        return len(t) if t else 0
    if not t:
        return len(s) if s else 0
    n, m = len(s), len(t)
    p = [x + 1 for x in range(n + 1)]  #'previous' cost array, horizontally
    d = [0] * (n + 1)  # cost array, horizontally

    for j in range(1, m + 1):
        d[0], t_j = j + 1, t[j - 1]
        for i in range(1, n + 1):
            d[i] = min(
                min(d[i - 1], p[i]) + 1,
                p[i - 1] + (0 if s[i - 1] == t_j else 1))
        p, d = d, p
    return p[n] - 1

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
        self._nmap, self._idmap, self._kwds = [None,] * 3
        self._dtype = 0

        if not nmap:
            return
        if isinstance(nmap, (tuple, list)):
            nmap = list2dict(nmap, **kwds)
        elif isinstance(nmap, str):
            div = ',' if nmap.find(',') > 0 else ' '
            nmap = list2dict(nmap.split(div), **kwds)
        elif isinstance(nmap, dict):
            nmap = {self._nrl(x[0]): x[1] for x in nmap.items()}
        self._nmap = nmap
        if data:
            self.setdata(data)

    def clone(self, struct_only=True):
        """ clone me

        Args:

            struct_only=True:    only clone the structure, no data will be cloned

        """
        self._mkidmap()
        nl = NamedList(None)
        nl._nmap, nl._idmap, nl._kwds, nl._nrl = self._nmap, self._idmap, self._kwds, self._nrl # save memory, share map
        if not struct_only:
            nl.setdata(deepcopy(self._data))
        return nl

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
        """set the internal data(should be of tuple/list)
        Args:
            val: the array that need to be feed to me
        Returns:
            myself
        """
        if val:
            if isinstance(val, Sequence) and len(self._nmap) != len(val):
                val = None
        if val is None:
            self._dtype = 0
        else:
            self._dtype = 1 if isinstance(
                val, Sequence) else 2 if isinstance(val, dict) else 10
        self._data = val
        return self

    def newdata(self, setme=True):
        """
        new an list of None that can be handle by me.
        @param setme: send the data to myself
        """
        rc = [
            None,
        ] * len(self._nmap)
        if setme:
            self.setdata(rc)
        return rc

    def _checkarg(self, name):
        if not (self._dtype and (self._dtype != 1 or name in self._nmap)):
            raise AttributeError(
                "no attribute(%s) found or data not set" % name)

    def __getattr__(self, name):
        name = self._nrl(name)
        if self._dtype == 1:
            return self._data[self._nmap[name]]
        if name in self._nmap:
            name = self._nmap[name]
        return None if self._dtype == 0 else self._data[name] if self._dtype == 2 else getattr(
            self._data, name)

    def __setattr__(self, name, val):
        # self._checkarg(name)
        if name.startswith("_"):
            object.__setattr__(self, name, val)
        else:
            if self._dtype == 0:
                self.newdata()
                self._dtype = 1
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

    @property
    def isblank(self):
        '''
        this object contains no data or all data is None
        '''
        if self._dtype == 0:
            return True
        elif self._dtype in (1, 2):
            return not [x for x in self._data if x]
        elif self._dtype == 10:
            return not [x for x in self._data.value() if x]
        return True

    def __str__(self):
        return self.__repr__()

    def __repr__(self):
        if not self._data:
            return None
        return repr(dict(zip(self.colnames, self._data)))


class NamedLists(Iterator):
    """
    make a list of list(2d array) accessable by name,

    for example, you read data from a csv

    lsts = (("id","name","price"),(1,"Jan",23.45),(2,"Pet",30.25)),

    you don't want to get id by
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
        """ init one named list instance

        Args:
            lsts: A list(or tuple) of a list(or tuple, but when it's a tuple, you can not assigned value),

            always put the title row at first

            trmap:  nmap translation map. used when nmap == None and you want to do some name tranlation

                    @refer to list2dict for more info.

            newinst=True:
                False to save memory in "for" loop

                set it to True if you use lst = [x for x in nl] or lst = list(nl).

                for safe reason, it's True by default
        """
        super(NamedLists, self).__init__()
        nmap = list2dict(lsts[0], alias=trmap, **kwds)
        lsts = lsts[1:]
        self._lsts, self._nmap, self._ptr, self._ubnd, self._newinst = lsts, nmap, \
            -1, len(lsts), newinst
        self._wrpr = NamedList(nmap, **kwds)
        self._kwds = kwds

    def __iter__(self):
        return self

    def __next__(self):
        self._ptr += 1
        if not self._lsts or self._ptr >= self._ubnd:
            raise StopIteration()
        if self._newinst:
            return self._wrpr.clone().setdata(self._lsts[self._ptr])
        return self._wrpr.setdata(self._lsts[self._ptr])

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

config = Config()
