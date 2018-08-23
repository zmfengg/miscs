# coding=utf-8    

'''
Created on Apr 19, 2018
Utils for xlwings's not implemented but useful function

@author: zmFeng
'''
import os

import xlwings
import xlwings.constants as const
from xlwings import Range, xlplatform

from ._miscs import NamedLists, list2dict
from .resourcemgr import ResourceMgr

__all__ = ["app","find","fromtemplate","list2dict","usedrange", "safeopen"]
_validappsws = set("visible,enableevents,displayalerts,asktoupdatelinks,screenupdating".split(","))

__crtappmgr = None

class _AppStg(object):
    def __init__(self, sws = None):
        self._sws = sws
        self._swso = None

    def crtr(self):
        self._kxl, self._app = app(False)
        if self._sws: self._swso = appswitch(self._app, self._sws)
        return self._app

    def dctr(self, app):
        if not hasattr(self,"_app"): return
        if not (self._app is app): return
        if self._kxl:
            self._app.quit()
        elif self._swso:
            appswitch(self._app, self._swso)

def app(vis=True,dspalerts = False):    
    """ launch an excel or connect to existing one
    return (flag,app), where flag is True means it's created by me, the caller should
    dispose() it
    """
    
    flag = xlwings.apps.count == 0
    app = xlwings.apps.active if not flag else xlwings.App(visible=vis, add_book=False)
    if app: app.display_alerts = bool(dspalerts)
    return flag, app

def appswitch(app, sws = None):
    """ turn switches on/off, return a string of the original value so that you can restore
    appswitch(app) or appswitch(app, True) to turn all default switch on
    appswitch(app,False) to turn all default switches off
    appswitch(app,{"visible":False,"screenupdate":True})
    remember to hold the result and call this method again to restore the prior state
    """
    if not app: return
    if sws is None:
        sws = dict([(x,True) for x in _validappsws])
    elif isinstance(sws,bool):
        sws = dict([(x,sws) for x in _validappsws])
    mp = {}
    for knv in sws.items():
        if knv[0] not in _validappsws: continue
        ov = eval("app.api.%s" % knv[0])
        if ov == bool(knv[1]): continue
        mp[knv[0]] = ov        
        exec("app.api.%s = %s" % (knv[0],bool(knv[1])))
    return mp

def apirange(rng):
    """ wrap an range object returned by api, for example, rng.api.mergearea
    """
    if not rng: return
    if isinstance(rng, Range): return rng
    if not isinstance(rng,xlplatform.COMRetryObjectWrapper): return
    return Range(impl = xlplatform.Range(rng))

def usedrange(sh):
    """
    find out the used range of the given sheet
    @param sh: the worksheet you want to find used range from. Maybe the same as sht.cells
    """
    return apirange(sh.api.UsedRange)

def find(sh, val, aftr=None, matchCase=False, lookat=const.LookAt.xlPart, \
         lookin=const.FindLookIn.xlValues, so=const.SearchOrder.xlByRows, \
         sd=const.SearchDirection.xlNext,):
    """
    return a range match the find criteria
    the original API does not provide the find function, here is one from the web
    https://gist.github.com/Elijas/2430813d3ad71aebcc0c83dd1f130e33
    respect the author for this
    @param sh: the sheet you want to perform the find on
    """
    if(not sh): return
    if(not val): val = "*"
    aftr = sh.api.Cells(1, 1) if(not aftr) else \
        sh.api.Cells(aftr.row, aftr.column)
    return apirange(sh.api.Cells.Find(What=val, After=aftr, \
                   LookAt=lookat, LookIn=lookin, \
                   SearchOrder=so, SearchDirection=sd, \
                       MatchCase=matchCase))

def contains(sht, vals):
    """ check if the sheet contains all the value in the vals tuple
    """
    if not (isinstance(vals,tuple) or isinstance(vals,list)):
        vals = (vals,)
    for val in vals:
        if not find(sht, val): return
    return True

def range2dict(vvs,trmap=None, dupdiv = "", bname = None):
    """ read a range's values into a list of dict item. vvs[0] should contains headers
    @param vvs: the range's value, an 2d array, use range.value to get it
    @param otherParm" refer to @list2dict
    """
    cmap = list2dict(vvs[0],trmap,dupdiv,bname)
    cns = [x for x in cmap.keys()]
    return list([x for x in [dict(zip(cns,y)) for y in vvs[1:]]])

def fromtemplate(tplfn, app=None):
    """new a workbook based on the tmpfn template
        @param tplfn: the template file
        @param app: the app you want to new workbook on 
    """
    if not os.path.exists(tplfn): return
    if not app:
        app = xlwings.App() if not xlwings.apps else xlwings.apps(0)
    app.api.Application.Workbooks.Add(tplfn)
    return app.books.active

def freeze(rng,restrfocus = True):
    """ freeze the window at given range """
    app = rng.sheet.book.app
    if restrfocus: orng = app.selection
    def _selrng(rg):
        rg.sheet.activate()
        rg.select()
    try:
        _selrng(rng)
        app.api.ActiveWindow.FreezePanes = True
        if restrfocus:
            _selrng(orng)
    except:
        pass

def safeopen(app, fn, updlnk = False, readonly = True):
    if not app or not os.path.exists(fn) : return
    flag = True
    try:
        app.api.workbooks.Open(fn, updlnk, readonly)
    except:
        flag = False
    if flag: return app.books[-1]

def NamedRanges(rng, skipfirstrow = False, nmap = None, scolcnt = 0):
    """ return the data under or include the range as namedlist list
    @param scolcnt: the count of columns to search, default is unlimited
    """
    if not rng: return
    if skipfirstrow: rng = rng.offset(1,0)
    sht, cr, orgcord = rng.sheet, rng.current_region, (rng.row, rng.column)
    ecol = orgcord[1] + scolcnt if scolcnt > 0 else cr.last_cell.column
    tr, rr, mg = sht.range(rng,sht.range(rng.row, ecol)), (65000 ,0), False
    for cell in tr.columns:
        if cell.api.mergecells:
            if not mg: mg = True
            mr = apirange(cell.api.mergearea)
            rr = (min(rr[0],mr.row),max(rr[1],mr.last_cell.row))
    if not mg: rr = (orgcord[0],)*2
    th = sht.range(sht.range(rr[0],orgcord[1]),sht.range(rr[1],ecol))
    if mg:
        if rr[0] == rr[1]:
            ttl = th.value
            for ii in range(len(ttl)):
                if not ttl[ii] and ii > 0: ttl[ii] = ttl[ii - 1]
        else:
            vals = [list(x) for x in th.value]
            for jj in range(len(vals)):
                lst = vals[jj]
                for ii in range(len(lst)):
                    if not lst[ii]:
                        val = lst[ii - 1] if ii > 0 else None
                        if not val: val = vals[jj - 1][ii] if jj > 0 else None
                        if val: lst[ii] = val
            ttl = [".".join(x) for x in zip(*vals)]
    else:
        ttl = ["%s" % x for x in th.value]
    lst = sht.range(sht.range(rr[1]+1,orgcord[1]),  sht.range(cr.last_cell.row,ecol)).value
    lst.insert(0,ttl)
    return NamedLists(lst,nmap)

def appmgr(sws = {"visible":False,"displayalerts":False}):
    global __crtappmgr
    if __crtappmgr: return __crtappmgr
    aps = _AppStg(sws)
    __crtappmgr = ResourceMgr(aps.crtr,aps.dctr)
    return __crtappmgr