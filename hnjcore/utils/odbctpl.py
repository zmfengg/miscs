# coding=utf-8
'''
Created on Apr 19, 2018
help to easily get odbc connection to:
    sybase
@author: zmFeng
'''

import pyodbc
import locale as lc
from os import path
from sqlalchemy.engine import create_engine

_sybdrv = None
_vfpdrv = None


def getSybCstr(addr, db, uid, pwd,
           port="5000", app=None, autoCP=True, **kws):
    """
    make a connect string for sybase connection
    if you want a direct connect, call getSybConn() instead
    @param addr:         IP of the server
    @param uid & pwd:    the userId and password
    @param tests:        the key=value pair with ; as delimiter
    @param autoCP:       get the code page automatically       
    """
    
    if(not(addr and db)): return
    if(not(port)):  port = "5000"
    if(not(app)): app = "Python"

    global _sybdrv
    cs = _sybdrv
    if not cs:
        drvs = [x for x in pyodbc.drivers() 
            if x.find("Sybase ASE ODBC") >= 0 or x.find("Adaptive Server Enterprise") >= 0]
        if not drvs: return
        drv = "Driver={%s};" % drvs[0]
        if(drvs[0].find("Adaptive Server") >= 0):
            cs = ("app=%(app)s;server=%(addr)s;"
                "port=%(port)s;db=%(db)s;uid=%(uid)s;pwd=%(pwd)s;%(miscs)s")
        elif(drvs[0].find("Sybase ASE ODBC") >= 0):
            cs = ("NA=%(addr)s,%(port)s;Database=%(db)s;"
                "LogonId=%(uid)s;Password=%(pwd)s;%(miscs)s")
        if cs:
            cs = drv + cs 
            _sybdrv = cs
    if not cs: return
    
    miscs = []
    if autoCP and (not kws or not ("charset" in kws)): 
        miscs.append("charset=%s" % lc.getdefaultlocale()[1])
    if kws:
        miscs.extend([x[0] + "=" + x[1] for x in kws.iteritems()])
    miscs = ";".join(miscs) if miscs else ""
    return cs % {"app" : app, "addr":addr, "db":db,
        "uid":uid, "pwd":pwd , "port":port, "miscs":miscs}
    

# get a sybase connection with auto detection of the sybase odbc driver
def getSybConn(addr, db, uid, pwd,
       port="5000", app=None, autoCP=True, **kws):
    cs = getSybCstr(addr, db, uid, pwd, port, app, autoCP, **kws)
    if(cs): return pyodbc.connect(cs)


def getAccess(fn, exclusive=False, uid=None, pwd=None):
    """ get an access connection """    
    if(not path.exists(fn)): return
    drvs = [x for x in pyodbc.drivers() if x.lower().find("microsoft access driver") >= 0]
    if(len(drvs) <= 0): return
    if(not uid): uid = "admin"
    if(not pwd): pwd = "admin"
    cs = (r"Driver={%(drv)s};Dbq=%(fn)s;Uid=%(uid)s;Pwd=%(pwd)s;Exclusive=%(exc)s" % 
        {"drv" :drvs[0], "fn":fn, "uid":uid, "pwd":pwd, "exc":exclusive})
    return pyodbc.connect(cs, autocommit=False)


def getXBase(fldr, exclusive=False):
    """get an xbase connection """
    if not path.exists(fldr): return
    global _vfpdrv
    cs = _vfpdrv
    if not cs:
        drvs = [x for x in pyodbc.drivers() if x.lower().find("microsoft visual foxpro") >= 0]
        if not drvs: return
        _vfpdrv = "Driver={%s}" % drvs[0] + ";SourceType=DBF;SourceDB=%s;Exclusive=%s"
    cs = _vfpdrv
    cs = cs % (fldr, "YES" if exclusive else "NO")
    return pyodbc.connect(cs)

def newSybEngine(cnnfunc):
    """ create a alchemy enginge based on sybase + existing connection
        @param cnnfunc: the function that will return a pyodbc connection to sybase
    """
    return create_engine("sybase+pyodbc://?driver=xx", creator=cnnfunc)
