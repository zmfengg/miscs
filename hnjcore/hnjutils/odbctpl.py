# coding=utf-8
'''
Created on Apr 19, 2018
help to easily get odbc connection to:
    sybase
@author: zmFeng
'''

import pyodbc
import locale as lc


def getSybCstr(addr, db, uid, pwd, \
           port="5000", app=None, miscs=None, autoCP=True):
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

    drvs = [x for x in pyodbc.drivers() \
        if x.find("Sybase ASE ODBC") >= 0 or x.find("Adaptive Server Enterprise") >= 0]
    
    if(drvs[0].find("Adaptive Server") >= 0):
        cs = "Driver={%(drv)s};app=%(app)s;server=%(addr)s;" \
            + "port=%(port)s;db=%(db)s;uid=%(uid)s;pwd=%(pwd)s;%(tests)s"
    elif(drvs[0].find("Sybase ASE ODBC") >= 0):
        cs = "Driver={%(drv)s};NA=%(addr)s,%(port)s;Database=%(db)s;" \
            + "LogonId=%(uid)s;Password=%(pwd)s;%(tests)s"
    else:
        return
    
    if autoCP and (not(miscs) or miscs.lower().find("charset") < 0):
        cp = "charset=%s" % lc.getdefaultlocale()[1]
        miscs = miscs + ";" + cp if miscs else cp    
    if(cs):
        cs = cs % {"drv" : drvs[0], "app" : app, "addr":addr, "db":db, \
            "uid":uid, "pwd":pwd , "port":port, "tests":miscs}
    return cs


# get a sybase connection with auto detection of the sybase odbc driver
def getSybConn(addr, db, uid, pwd, \
           port="5000", app=None, miscs=None, autoCP=True):
    cs = getSybCstr(addr, db, uid, pwd, port, app, miscs, autoCP)
    if(cs): return pyodbc.connect(cs)


def getAccess(fn, exclusive=False, uid=None, pwd=None):
    """ get an access connection """
    from os import path
    if(not path.exists(fn)): return
    drvs = [x for x in pyodbc.drivers() if x.lower().find("microsoft access driver") >= 0]
    if(len(drvs) <= 0): return
    if(not uid): uid = "admin"
    if(not pwd): pwd = "admin"
    cs = r"Driver={%(drv)s};Dbq=%(fn)s;Uid=%(uid)s;Pwd=%(pwd)s;Exclusive=%(exc)s" % \
        {"drv" :drvs[0], "fn":fn, "uid":uid, "pwd":pwd, "exc":exclusive}
    return pyodbc.connect(cs, autocommit=False)
