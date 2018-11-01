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

from ._miscs import NamedLists, list2dict, updateopts
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
        ov = getattr(app0.api, knv[0]) #ov = eval("app0.api.%s" % knv[0])
        if ov == bool(knv[1]):
            continue
        mp[knv[0]] = ov
        setattr(app0.api, knv[0], bool(knv[1])) #exec("app0.api.%s = %s" % (knv[0], bool(knv[1])))
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


def find(sht, val, **kwds):
    """
    return a range match the find criteria
    the original API does not provide the find function, here is one from the web
    https://gist.github.com/Elijas/2430813d3ad71aebcc0c83dd1f130e33
    respect the author for this
    @param sht: the sheet you want to perform the find on
    @param after: Range, after which to perform the search, default is None
    @param match_case(or matchcase): boolean, search case-sensitive, default is False
    @param look_at(or lookat): xlwings.const.LookAt, default is xlPart
    @param look_in(or lookin): xlwings.const.FindLookIn, default is xlValues
    @param order(or searchorder): const.SearchOrder, default is xlByRows
    @param direction: const.SearchDirection.xlNext, default is xlNext
    """
    if not sht:
        return None
    if not val:
        val = "*"
    after = kwds.get("after")
    after = sht.api.Cells(1, 1) if(not after) else sht.api.Cells(after.row, after.column)

    d1 = updateopts({"What": ("LookAt,look_at", const.LookAt.xlPart), "LookIn": ("lookin,look_in", const.FindLookIn.xlValues), "SearchOrder": ("searchorder,search_order,order", const.SearchOrder.xlByRows), "SearchDirection": ("direction", const.SearchDirection.xlNext), "MatchCase": ("match_case,matchcase,case", False)}, kwds)
    d1["What"], d1["After"] = val, after
    return apirange(sht.api.Cells.Find(**d1))

def contains(sht, vals):
    """ check if the sheet contains all the value in the vals tuple
    """
    if not isinstance(vals, (tuple, list)):
        vals = (vals,)
    for val in vals:
        if not find(sht, val):
            return None
    return True

def detectborder(rng0):
    """
    find all the ranges that was surrounded by borders from this range on
    """
    bts = [(getattr(const.BordersIndex,"xlEdge%s" % x[0]), int(x[1]), int(x[2])) for x in [y.split(",") for y in "Top,0,-1;Left,1,-1;Bottom,0,1;Right,1,1".split(";")]]
    sh, maxDtc, orgs, idx, bds = rng0.sheet, 100, [rng0.row, rng0.column], 0, []
    for ptr in bts:
        idx = 1
        while idx < maxDtc:
            nOff = orgs[ptr[1]] + ptr[2] * idx
            if nOff <= 0:
                break #reach the left/top zero point
            rng = sh.range(orgs[0] if ptr[1] else nOff, nOff if ptr[1] else orgs[1])
            if rng.api.borders(ptr[0]).LineStyle != -4142:
                bds.append(rng.column if ptr[1] else rng.row)
                break
            idx += 1
    if not bds or len(bds) != 4:
        return None
    return sh.range(sh.range(bds[0], bds[1]), sh.range(bds[2], bds[3]))

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


def NamedRanges(rng, **kwds):
    """ return the data under or include the range as namedlist list
    @param skip_first_row: boolean, don't process the first row, default is False
    @param name_map: the name->title mapping, see @list2dict FMI, default is None
    @param col_cnt: the count of columns to search, default is 0, that is unlimited
    """
    if not rng:
        return None
    if rng.size > 1:
        rng = rng[0]
    if kwds.get("skip_first_row"):
        rng = rng.offset(1, 0)
    sht, cur_region, org_pt = rng.sheet, rng.current_region, (rng.row, rng.column)
    var = kwds.get("col_cnt", kwds.get("colcnt")) or 0
    e_colidx = org_pt[1] + var if var > 0 else cur_region.last_cell.column
    tt_rows, var = (65000, 0), False

    var = [x for x in sht.range(rng, sht.range(org_pt[0], e_colidx)).columns if x.api.mergecells]
    if var:
        for cell in var:
            mr = apirange(cell.api.mergearea)
            tt_rows = (min(tt_rows[0], mr.row), max(tt_rows[1], mr.last_cell.row))
    else:
        tt_rows = (org_pt[0],)*2
    th = sht.range(sht.range(tt_rows[0], org_pt[1]), sht.range(tt_rows[1], e_colidx))
    if var:
        if tt_rows[0] == tt_rows[1]:
            var = []
            for val in th.value:
                if not val and var:
                    val = var[-1]
                var.append(val)
        else:
            vals = []
            for var in [tuple(x) for x in th.value]:
                vals.append([])
                for val in var:
                    if not val and vals[-1]:
                        val = vals[-1][-1]
                    if not val and len(vals) > 1:
                        val = vals[-2][len(vals[-1])]
                    vals[-1].append(val)
            var = [".".join(x) for x in zip(*vals)]
    else:
        var = ["%s" % x for x in th.value] if th.value else None
    if not var:
        return None
    rng = sht.range(sht.range(tt_rows[1]+1, org_pt[1]), sht.range(cur_region.last_cell.row, e_colidx))
    th = rng.value
    #one row case, xlwings return a 1-dim array only, make it 2D
    if rng.rows.count == 1:
        th = [th]
    th.insert(0, var)
    return NamedLists(th, kwds.get("name_map"))


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
