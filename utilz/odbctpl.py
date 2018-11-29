# coding=utf-8
'''
Created on Apr 19, 2018
help to easily get odbc connection to:
    sybase
@author: zmFeng
'''

import locale as lc
from os import path

from sqlalchemy.engine import create_engine

import pyodbc

from ._miscs import triml, getvalue as gv
from .common import _logger as logger


class _Wrapper(object):
    _drv_mp = {}

    def _getrawcs(self, vendor):
        """ return a raw connection string(including the space holder)
        for different vendors
        """
        vendor = triml(vendor)
        if vendor in self._drv_mp:
            return self._drv_mp[vendor]
        cs, drvs = None, pyodbc.drivers()
        drvs = tuple((x.lower(), x) for x in drvs)
        if vendor == "sybase":
            drvs = [
                x for x in drvs if x[0].find("sybase ase odbc") >= 0 or
                x[0].find("adaptive server enterprise") >= 0
            ]
            if drvs:
                drvs = drvs[0]
                drv = "Driver={%s};" % drvs[1]
                if drvs[0].find("adaptive server") >= 0:
                    cs = "app=%(app)s;server=%(addr)s;port=%(port)s;db=%(db)s;uid=%(uid)s;pwd=%(pwd)s;%(miscs)s"
                elif drvs[0].find("sybase ase odbc") >= 0:
                    cs = "NA=%(addr)s,%(port)s;Database=%(db)s;LogonId=%(uid)s;Password=%(pwd)s;%(miscs)s"
                if cs:
                    cs = drv + cs
        elif vendor == "vfp":
            drvs = [
                x[1] for x in drvs if x[0].find("microsoft visual foxpro") >= 0
            ]
            if drvs:
                cs = "Driver={%s}%s" % (
                    drvs[0],
                    r";SourceType=DBF;SourceDB=%(fldr)s;Exclusive=%(exc)s")
        elif vendor == "msaccess":
            drvs = [
                x[1] for x in drvs if x[0].find("microsoft access driver") >= 0
            ]
            if drvs:
                cs = "Driver={%s}%s" % (
                    drvs[0],
                    r";Dbq=%(fn)s;Uid=%(uid)s;Pwd=%(pwd)s;Exclusive=%(exc)s")
        if not cs:
            logger.debug("No driver is found for vendor(%s)" % vendor)
        self._drv_mp[vendor] = cs
        return cs

    def _makeSybCstr(self, addr, db, uid, pwd, **kwds):
        """
        make a connect string for sybase connection
        if you want a direct connect, call getSybConn() instead
        @param addr:         IP of the server
        @param uid & pwd:    the userId and password
        @param tests:        the key=value pair with ; as delimiter
        @param autoCP:       get the code page automatically
        """

        if not (addr and db):
            return None
        cs = self._getrawcs("sybase")
        if not cs:
            return None
        miscs = []
        if "autoCP" in kwds:
            if "charset" not in kwds:
                miscs.append("charset=%s" % lc.getdefaultlocale()[1])
            del kwds["autoCP"]
        if kwds:
            miscs.extend(["=".join(x) for x in kwds.items()])
        miscs = ";".join(miscs) if miscs else ""
        return cs % {
            "app": gv(kwds, "app", "python"),
            "addr": addr,
            "db": db,
            "uid": uid,
            "pwd": pwd,
            "port": gv(kwds, "port", "5000"),
            "miscs": miscs
        }

    # get a sybase connection with auto detection of the sybase odbc driver
    def getSybConn(self, addr, db, uid, pwd, **kwds):
        """
            return a connection to sybase if your system has one of below driver installed
                .Adaptive Server
                .Sybase ASE ODBC
        """
        cs = self._makeSybCstr(addr, db, uid, pwd, **kwds)
        return pyodbc.connect(cs, timeout=5) if cs else None

    def getAccess(self, fn, exclusive=False, uid=None, pwd=None):
        """ get an access connection """
        if not path.exists(fn):
            return None
        cs = self._getrawcs("msaccess")
        if not cs:
            return None
        cs = cs % {
            "fn": fn,
            "uid": uid or "admin",
            "pwd": pwd or "admin",
            "exc": exclusive
        }
        return pyodbc.connect(cs, autocommit=False)

    def getXBase(self, fldr, exclusive=False):
        """get an xbase connection """
        if not path.exists(fldr):
            return None
        cs = self._getrawcs("vfp")
        if not cs:
            return None
        cs = cs % {"fldr": fldr, "exc": "YES" if exclusive else "NO"}
        return pyodbc.connect(cs)

_wrapper = _Wrapper()

def getSybConn(addr, db, uid, pwd, **kwds):
    """ return a new connection to a sybase server, @refer to _Wrapper.getSybConn() FMI """
    return _wrapper.getSybConn(addr, db, uid, pwd, **kwds)

def getAccess(fn, exclusive=False, uid=None, pwd=None):
    """ return a connection to a MSAccess db, @refer to _Wrapper.getAccess() FMI """
    return _wrapper.getAccess(fn, exclusive, uid, pwd)

def getXBase(fldr, exclusive=False):
    """ return a conneciton to a dbf file or folder, @refer to _Wrapper.getXBase() FMI """
    return _wrapper.getXBase(fldr, exclusive)

def newSybEngine(cnnfunc, **kwds):
    """ create a alchemy enginge based on sybase + existing connection
        @param cnnfunc: the function that will return a pyodbc connection to sybase
    """
    return create_engine(
        "sybase+pyodbc://?driver=xx", creator=cnnfunc, **kwds)
