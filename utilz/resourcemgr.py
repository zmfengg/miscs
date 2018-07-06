#! coding=utf-8 
'''
* @Author: zmFeng 
* @Date: 2018-06-16 14:20:28 
* @Last Modified by:   zmFeng 
* @Last Modified time: 2018-06-16 14:20:28 
A thread-safe resource manager(ResourceMgr), a Context based Resource consumer(ResourceCtx). An sqlachemy session
resource manager for the ResourceCtx using for cross-method session sharing
'''

import threading
from .common import _logger as logger
from sqlalchemy.orm import Session

__all__ = ["ResourceMgr","ResourceCtx"]

class ResourceMgr(object):
    """ a thread-safe resource manager, for shared objects like session
        usage example(direct use ResourceMgr):
            def _newcnn():
                return pyodbc.connect(...)
            def _dpcnn(cnn):
                cnn.close()
            cnnmgr = ResourceMgr(_newcnn,_dpcnn)
            cnn,token = cnnmgr.acq()
            try:
                ...
            finally:
                cnnmgr.ret(token)
        
        usage example(using contextmgr):

        """

    def __init__(self,crtor,dctr):
        """ create a thread-safe resource manager by providng the constructor/destructor
        function which can create/dispose an resource
        @param crtor: the method to create a resource
        @param dctor: the method to dispose a resource
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

class SessionMgr(ResourceMgr):
    """ a sqlalchemy engine session manager by providing a sqlalchemy engine """
    def __init__(self,engine):
        self._engine = engine
        super(SessionMgr,self).__init__(self._newsess,self._closesess)

    def _newsess(self):
        return Session(self._engine)
    
    def _closesess(self,sess):
        sess.close()
    
    @property
    def engine(self):
        return self._engine

    def dispose(self):
        logger.debug("sqlachemy engine(%s) disposed" % self._engine)
        self._engine.dispose()
    
    def close(self):
        self.dispose()

class ResourceCtx(object):

    def __init__(self,ResourceMgrs):
        from collections import Iterable
        self._src = list(ResourceMgrs) if isinstance(ResourceMgrs,Iterable) else [ResourceMgrs]

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