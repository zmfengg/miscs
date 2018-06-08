# coding=utf-8
'''
Created on 2018-05-18

@author: zmFeng
'''

from os import path
import math
import sys
import os
import threading

__all__ = ["splitarray","appathsep","deepget","getfiles","samekarat","ResourceMgr"]

_silveralphas = set(("4", "5"))

def splitarray(arr, logsize = 100):
    """split an array into arrays whose len is less or equal than logsize
    @param arr: the sequence object that need to split
    @param logsize: len of each sub-array's size  
    """    
    if not arr: return
    if not logsize: logsize = 100 
    return (arr[x * logsize:(x + 1) * logsize] for x in range(int(math.ceil(1.0 * len(arr) / logsize))))

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

def samekarat(srcje, tarje):
    """ detect if the given 2 JOElement are of the same karat """
    if not (srcje and tarje): return
    return srcje.alpha == tarje.alpha or (srcje.alpha in _silveralphas and tarje.alpha in _silveralphas)

def getfiles(fldr,ext = None, nameonly = False):
    """ return files under given folder """
    """ @param nameonly : don't return the full-path """

    if fldr:
        fldr = appathsep(fldr)
        if ext:
            ext = ext.lower()
            fns = [unicode(x, sys.getfilesystemencoding()) 
                for x in os.listdir(fldr) if not x.lower().find(ext) >= 0]
        else:
            fns = [unicode(x, sys.getfilesystemencoding()) 
                for x in os.listdir(fldr)]
        if not nameonly:
            fns = [fldr + x for x in fns]
    return fns


class ResourceMgr(object):
    """ a resource manager, mainly for session management, Thread safe
        use acq() to new resource request, ret() to return
        use get() to borrow an existing one, if no existing, won't return anything
        this class act like a stack object, ack -> push, ret -> pop, get -> peak        
    """

    def __init__(self,ctr,dtr):
        """ @param ctr: the construct method for a resource
            @param dtr: the destruct method of a resource, should accept dtr(res) method
        """
        self._tl = threading.local()
        self._tl.stacks = {}
        self._ctr = ctr
        self._dtr = dtr
    
    def _getlist(self):
        id = thread.get_ident()
        return self._tl.stacks.setdefault(id,{})

    def acq(self):
	""" acquire for a new resource and store it, prepare for returning
    """        
        res = self._ctr()
        self._getlist().append(res)
        return res

    def ret(self, **kws):
        """ return the resource to me """
        lst = self._getlist()
        if lst:
            res = lst.pop()
            self._dtr(res)

    def get(self):
        """ get an existing resource, won't return. If no current, return None """
        lst = self._getlist()
        return lst.pop() if lst else None
