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

__all__ = ["app", "appmgr", "find", "fromtemplate", "list2dict", "usedrange", "safeopen"]
_validappsws = set("visible,enableevents,displayalerts,asktoupdatelinks,screenupdating".split(","))


class _AppStg(object):
    def __init__(self, sws=None):
        self._sws = sws
        self._swso, self._kxl, self._app = (None,) * 3

    def crtr(self):
        """
        the app creator
        """
        self._kxl, self._app = app(False)
        if self._sws:
            self._swso = appswitch(self._app, self._sws)
        return self._app

    def dctr(self, app0):
        """
        the app destroyer
        """
        if not hasattr(self, "_app"):
            return
        if not self._app is app0:
            return
        if self._kxl:
            self._app.quit()
            try:
                self._app.version
                # quit() sometime does not work
                # if the app was closed, .version throws exception
                self._app.kill()
            except:
                pass
            self._app = None
        elif self._swso:
            appswitch(self._app, self._swso)


def app(vis=True, dspalerts=False):
    """ launch an excel or connect to existing one
    return (flag,app), where flag is True means it's created by me, the caller should
    dispose() it
    """

    flag = xlwings.apps.count == 0
    app0 = xlwings.apps.active if not flag else xlwings.App(visible=vis, add_book=False)
    if app0:
        app0.display_alerts = bool(dspalerts)
    return flag, app0


def appswitch(app0, sws=None):
    """ turn switches on/off, return a string of the original value so that you can restore
    appswitch(app) or appswitch(app, True) to turn all default switch on
    appswitch(app,False) to turn all default switches off
    appswitch(app,{"visible":False,"screenupdate":True})
    remember to hold the result and call this method again to restore the prior state
    """
    if not app0:
        return None
    if sws is None:
        sws = dict([(x, True) for x in _validappsws])
    elif isinstance(sws, bool):
        sws = dict([(x, sws) for x in _validappsws])
    mp = {}
    for knv in sws.items():
        if knv[0] not in _validappsws:
            continue
        #ov = getattr(app0.api, knv[0])()
        ov = eval("app0.api.%s" % knv[0])
        if ov == bool(knv[1]):
            continue
        mp[knv[0]] = ov
        #getattr(app0.api, knv[0]) = bool(knv[1])
        exec("app0.api.%s = %s" % (knv[0], bool(knv[1])))
    return mp


def apirange(rng):
    """ wrap an range object returned by api, for example, rng.api.mergearea
    """
    if not rng:
        return None
    if isinstance(rng, Range):
        return rng
    if not isinstance(rng, xlplatform.COMRetryObjectWrapper):
        return None
    return Range(impl=xlplatform.Range(rng))


def usedrange(sh):
    """
    find out the used range of the given sheet
    @param sh: the worksheet you want to find used range from. Maybe the same as sht.cells
    """
    return apirange(sh.api.UsedRange)


def find(sh, val, aftr=None, matchCase=False, lookat=const.LookAt.xlPart,
         lookin=const.FindLookIn.xlValues, so=const.SearchOrder.xlByRows,
         sd=const.SearchDirection.xlNext,):
    """
    return a range match the find criteria
    the original API does not provide the find function, here is one from the web
    https://gist.github.com/Elijas/2430813d3ad71aebcc0c83dd1f130e33
    respect the author for this
    @param sh: the sheet you want to perform the find on
    """
    if not sh:
        return None
    if not val:
        val = "*"
    aftr = sh.api.Cells(1, 1) if(not aftr) else \
        sh.api.Cells(aftr.row, aftr.column)
    return apirange(sh.api.Cells.Find(What=val, After=aftr,
                                      LookAt=lookat, LookIn=lookin,
                                      SearchOrder=so, SearchDirection=sd,
                                      MatchCase=matchCase))


def contains(sht, vals):
    """ check if the sheet contains all the value in the vals tuple
    """
    if not isinstance(vals, (tuple, list)):
        vals = (vals,)
    for val in vals:
        if not find(sht, val):
            return None
    return True


def fromtemplate(tplfn, app0=None):
    """new a workbook based on the tmpfn template
        @param tplfn: the template file
        @param app: the app you want to new workbook on
    """
    if not os.path.exists(tplfn):
        return None
    if not app0:
        app0 = appmgr.acq()[0]
    app0.api.Application.Workbooks.Add(tplfn)
    return app0.books.active


def freeze(rng, restrfocus=True):
    """ freeze the window at given range """
    app0 = rng.sheet.book.app
    if restrfocus:
        orng = app0.selection

    def _selrng(rg):
        rg.sheet.activate()
        rg.select()
    try:
        _selrng(rng)
        app0.api.ActiveWindow.FreezePanes = True
        if restrfocus:
            _selrng(orng)
    except:
        pass


def safeopen(appx, fn, updlnk=False, readonly=True):
    """
    open a workbook with the ability to control readonly/updatelink,
    replace the app.books.open(fn)
    """
    flag = appx and os.path.exists(fn)
    if not flag:
        return None
    try:
        appx.api.workbooks.Open(fn, updlnk, readonly)
    except:
        flag = False
    return appx.books[-1] if flag else None


def NamedRanges(rng, skipfirstrow=False, nmap=None, scolcnt=0):
    """ return the data under or include the range as namedlist list
    @param scolcnt: the count of columns to search, default is unlimited
    """
    if not rng:
        return None
    if rng.size > 1:
        rng = rng[0]
    if skipfirstrow:
        rng = rng.offset(1, 0)
    sht, cr, orgcord = rng.sheet, rng.current_region, (rng.row, rng.column)
    ecol = orgcord[1] + scolcnt if scolcnt > 0 else cr.last_cell.column
    tr, rr, mg = sht.range(rng, sht.range(rng.row, ecol)), (65000, 0), False
    for cell in tr.columns:
        if cell.api.mergecells:
            if not mg:
                mg = True
            mr = apirange(cell.api.mergearea)
            rr = (min(rr[0], mr.row), max(rr[1], mr.last_cell.row))
    if not mg:
        rr = (orgcord[0],)*2
    th = sht.range(sht.range(rr[0], orgcord[1]), sht.range(rr[1], ecol))
    if mg:
        if rr[0] == rr[1]:
            ttl = []
            for val in th.value:
                if not val and ttl:
                    val = ttl[-1]
                ttl.append(val)
        else:
            vals = []
            for lst in [tuple(x) for x in th.value]:
                vals.append([])
                for val in lst:
                    if not val and vals[-1]:
                        val = vals[-1][-1]
                    if not val and len(vals) > 1:
                        val = vals[-2][len(vals[-1])]
                    vals[-1].append(val)
            ttl = [".".join(x) for x in zip(*vals)]
    else:
        ttl = ["%s" % x for x in th.value] if th.value else None
    if not ttl:
        return None
    lst = sht.range(sht.range(rr[1]+1, orgcord[1]), sht.range(cr.last_cell.row, ecol)).value
    lst.insert(0, ttl)
    return NamedLists(lst, nmap)


def _newappmgr(sws=None):
    if not sws:
        sws = {"visible": False, "displayalerts": False}
    aps = _AppStg(sws)
    return ResourceMgr(aps.crtr, aps.dctr)


def escapetitle(pg):
    """ when excel's page title has format set, you can not get the raw directly. this function
    help to get rid of the format, return raw data only
    the string format is:
    ' &"fontName,italia"[&size]. Just remove such pair
    """
    ss = []
    for s0 in pg.split('&"'):
        s0 = s0[s0.find('"') + 1:]
        ss.append(s0[s0.find(" ") + 1:] if s0[0] == "&" else s0)
    s0 = "".join(ss)
    return s0


# an appmgr factory, instead of using app(), use appmgr.acq()/appmgr.ret()
appmgr = _newappmgr()
