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
import inspect
from sqlalchemy.orm import Session
from ._res import _logger as logger

__all__ = ["splitarray","appathsep","deepget","getfiles","samekarat","Resctx","Resmgr", "Sessionmgr"]

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

class Resmgr(object):

    def __init__(self,crtor,dctr):
        """ create a thread-safe resource manager by providng the constructor/destructor
        function which can create/dispose an resource
        @param crtor: the method to create a resource
        @param dctor: the method to dispose a resource
        usage example(direct use Resmgr):
            def _newcnn():
                return pyodbc.connect(...)
            def _dpcnn(cnn):
                cnn.close()
            cnnmgr = Resmgr(_newcnn,_dpcnn)
            cnn,token = cnnmgr.acq()
            try:
                ...
            finally:
                cnnmgr.ret(token)
        
        usage example(using contextmgr):

        """
        self._create = crtor
        self._dispose = dctr
        self._storage = {}

    def _getstorage(self):
        return self._storage.setdefault(threading.get_ident(),([],{}))

    def hasres(self):
        return bool(self._getstorage()[0])
    
    def get(self):
        stk = self._getstorage()[0]
        rc = stk[len(stk) - 1] if stk else None
        logger.debug("existing resource(%s) reused" % (rc) if rc else "No resource available, use acq() to new one")
        return rc
    
    def acq(self):
        """ return an tuple, the resource as [0] while the token as [1]
        you need to provide the token while ret()
        """
        import random
        
        stg = self._getstorage()
        stk, rmap, token = stg[0], stg[1], 0
        while(not token or token in rmap):
            token = random.randint(1,65535)
        res = self._create()
        logger.debug("resource(%s) created by acq() with token(%d)" % (res,token))
        rmap[token] = res
        stk.append(res)
        return res,token
    
    def ret(self,token):
        stg = self._getstorage()
        stk, rmap = stg[0], stg[1]
        
        if not stk:
            raise Exception("Invalid stack status")
        if token not in rmap:
            raise Exception("not returing sth. borrowed from me")
        res = rmap[token]
        del rmap[token]; stk.pop()
        self._dispose(res)
        logger.debug("resource(%s) return and disposed by token(%d)" % (res,token))

class Resctx(object):

    def __init__(self,resmgrs):
        from collections import Iterable
        self._src = list(resmgrs) if isinstance(resmgrs,Iterable) else [resmgrs]

    """ a ball to catch mon(resource)"""
    def __enter__(self):
        self._closes, self._ress = [], []
        for ii in range(len(self._src)):
            x = self._src[ii]
            if x.hasres():
                self._closes.append(False)
                self._ress.append((x.get(),0,ii))
            else:
                self._closes.append(True)
                lst = list(x.acq()); lst.append(ii)
                self._ress.append(lst)
        
        return self._ress[0][0] if len(self._ress) == 1 else [x[0] for x in self._ress]
                
    def __exit__(self, exc_type, exc_value, traceback):
        cnt = len(self._closes)
        for ii in range(cnt - 1,-1,-1):
            if self._closes[ii]:
                self._src[self._ress[ii][2]].ret(self._ress[ii][1])
        return True if not exc_type else False

class Sessionmgr(Resmgr):
    """ a sqlalchemy engine session manager by providing a sqlalchemy engine """
    def __init__(self,engine):
        self._engine = engine
        super(Sessionmgr,self).__init__(self._newsess,self._closesess)

    def _newsess(self):
        return Session(self._engine)
    
    def _closesess(self,sess):
        sess.close()

    def dispose(self):
        logger.debug("sqlachemy engine(%s) disposed" % self._engine)
        self._engine.dispose()
    
    def close(self):
        self.dispose()