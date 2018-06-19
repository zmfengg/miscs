#! coding=utf-8 
'''
* @Author: zmFeng 
* @Date: 2018-06-16 15:44:32 
* @Last Modified by:   zmFeng 
* @Last Modified time: 2018-06-16 15:44:32 
'''

from os import path
import math
import sys
import os
import threading
import inspect
from sqlalchemy.orm import Session
from .common import _logger as logger

__all__ = ["splitarray","appathsep","deepget","getfiles", \
    "stsizefmt", "daterange"]

def splitarray(arr, logsize = 100):
    """split an array into arrays whose len is less or equal than logsize
    @param arr: the sequence object that need to split
    @param logsize: len of each sub-array's size  
    """    
    if not arr: return
    if not logsize: logsize = 100 
    return [arr[x * logsize:(x + 1) * logsize] for x in range(int(math.ceil(1.0 * len(arr) / logsize)))]

def appathsep(fldr):
    """append a path sep into given path if there is not"""
    return fldr + path.sep if fldr[len(fldr) - 1:] != path.sep else fldr

def deepget(obj,names):
    """ get deeply from the object """
    rc = None
    for k in names.split("."):
        rc = rc.__getattribute__(
            k) if rc else obj.__getattribute__(k)
    return rc

def getfiles(fldr,part = None, nameonly = False):
    """ return files under given folder """
    """ @param nameonly : don't return the full-path """

    if fldr:
        fldr = appathsep(fldr)
        if part:
            part = part.lower()
            fns = [x if sys.version_info.major >= 3 else str(x, sys.getfilesystemencoding()) 
                for x in os.listdir(fldr) if x.lower().find(part) >= 0]
        else:
            fns = [x if sys.version_info.major >= 3 else str(x, sys.getfilesystemencoding()) 
                for x in os.listdir(fldr)]
        if not nameonly:
            fns = [fldr + x for x in fns]
    return fns

def daterange(year,month,day = 1):
    """ make a from,thru tuple for the given month, thru is the first date of next month """
    import datetime as dtm
    df = dtm.date(year,month,day if day > 0 else 1)
    month += 1
    if month > 12:
        year += 1
        month = 1        
    dt = dtm.date(year, month, 1)
    del dtm
    return df, dt

def stsizefmt(sz, shortform = False):
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
            if shortform: s0 = "%d" % (int(s0) / 100)
        else:
            s0 = splitarray(s0,4)
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
        elif x in ("X","*"):
            idx = _inc(segs)
            if rng: break            
        elif rng:
            break
    for x in segs:
        x = _fmtpart(x,shortform)
        if isinstance(x,str):
            parts.append(x)
        else:
            parts.extend(x)                
    return ("-" if rng else "X").join(sorted(parts,reverse = True))