# coding=utf-8
'''
Created on Apr 19, 2018
Utils for xlwings's not implemented but useful function

@author: zmFeng
'''
from numbers import Number
from os import path, remove
from tempfile import gettempdir

import xlwings.constants as const
from xlwings import App, Range, apps, xlplatform
from xlwings.utils import col_name

from ._miscs import NamedLists, getvalue, list2dict, updateopts, trimu
from .resourcemgr import ResourceMgr

try:
    from PIL import Image
except:
    Image = None

__all__ = [
    "app", "appmgr", "col", "find", "fromtemplate", "hidden", "insertphoto",
    "list2dict", "NamedRanges", "safeopen", "usedrange"
]
_validappsws = set(
    "visible,enableevents,displayalerts,asktoupdatelinks,screenupdating".split(
        ","))
_alignment_mp = {"L": 0, "M": 1, "R": 2, "T": 0, "C": 1, "R": 2}

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

    flag = apps.count
    app0 = apps.active if flag else App(visible=vis, add_book=False)
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
        sws = {x: True for x in _validappsws}
    elif isinstance(sws, bool):
        sws = {x: sws for x in _validappsws}
    mp = {}
    for knv in sws.items():
        if knv[0] not in _validappsws:
            continue
        ov = getattr(app0.api, knv[0])  #ov = eval("app0.api.%s" % knv[0])
        if ov == bool(knv[1]):
            continue
        mp[knv[0]] = ov
        setattr(app0.api, knv[0], bool(
            knv[1]))  #exec("app0.api.%s = %s" % (knv[0], bool(knv[1])))
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

def findsheet(wb, sn):
    sn = trimu(sn)
    for sht in wb.sheets:
        if trimu(sht.name) == sn:
            return sht
    return None


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
    @param find_all(or findall): find all the instances with given criteria
    """
    if not sht:
        return None
    if not val:
        val = "*"
    after = getvalue(kwds, "After")
    after = sht.api.Cells(1, 1) if not after else sht.api.Cells(
        after.row, after.column)

    dfs = {
        "LookAt": ("LookAt,look_at,lookat", const.LookAt.xlPart),
        "LookIn": ("lookin,look_in", const.FindLookIn.xlValues),
        "SearchOrder": ("searchorder,search_order,order",
                        const.SearchOrder.xlByRows),
        "SearchDirection": ("direction", const.SearchDirection.xlNext),
        "MatchCase": ("match_case,matchcase,case", False),
        "After": ("after", after)
    }
    d1 = updateopts(dfs, kwds)
    d1["What"], d1["After"] = val, after
    find_all = getvalue(d1, "findall,find_all")
    #the api only accept valid keywords, so remove other ones
    dfs = [x for x in d1 if x != "What" and x not in dfs]
    if dfs:
        for x in dfs:
            del d1[x]
    rng = apirange(sht.api.Cells.Find(**d1))
    if find_all and rng:
        st = set(rng,)
        while rng:
            d1["After"] = rng.api
            rng = apirange(sht.api.Cells.Find(**d1))
            if not rng or rng in st:
                rng = st
                break
            if rng:
                st.add(rng)
    return rng


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
    bts = [(getattr(const.BordersIndex, "xlEdge%s" % x[0]), int(x[1]),
            int(x[2])) for x in [
                y.split(",")
                for y in "Top,0,-1;Left,1,-1;Bottom,0,1;Right,1,1".split(";")
            ]]
    sh, maxDtc, orgs, idx, bds = rng0.sheet, 100, [rng0.row, rng0.column], 0, []
    for ptr in bts:
        idx = 0
        while idx < maxDtc:
            nOff = orgs[ptr[1]] + ptr[2] * idx
            if nOff <= 0:
                break  #reach the left/top zero point
            rng = sh.range(orgs[0] if ptr[1] else nOff,
                           nOff if ptr[1] else orgs[1])
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
    if not path.exists(tplfn):
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
    flag = appx and path.exists(fn)
    if not flag:
        return None
    try:
        appx.api.workbooks.Open(fn, updlnk, readonly)
    except:
        flag = False
    return appx.books[-1] if flag else None

def _pos(org, ttl, margin, width, align):
    align, rc = _alignment_mp[align[0]], 0
    if align == 0:
        rc = org + margin
    elif align == 1:
        rc = org + (ttl - width) / 2
    else:
        rc = org + (ttl - width) - margin
    return rc

def insertphoto(fn, rng, **kwds):
    """
    insert photo into target range
    @param chop_at: an integer specified the chop/th position to chop at
    @param chop_img: a loaded PIL image, If its a string, I' load it myself
    @param max_size: the maximum photo size that I can insert. If the photo is bigger than that, I'll trim it down
                        to the max_size, default is (800, 600)
    @param margins: a tuple(x, y) to specified the margins to the x/y side, default is (5, 5)
    @param alignment: a mixture of L/C/R and T/M/B, an example is L,T, default is C,M
    """
    mp = updateopts({
        "max_size": ("max_size", tuple(int(x) for x in "800X600".split("X"))),
        "margins": ("margins", (5, 5))
    }, kwds)
    if not Image:
        return None
    img = Image.open(fn)
    img.load()
    sz, save_it = img.size, False
    h_w = sz[1] / sz[0]
    max_sz, chop_at, chop_img, margins = mp.get("max_size"), mp.get(
        "chop_at"), mp.get("chop_img"), mp.get("margins", (5, 5))
    if chop_at:
        if isinstance(chop_img, str) and path.exists(chop_img):
            chop_img = Image.open(chop_img)
            chop_img.load()
        if not chop_img:
            chop_at = None
        else:
            if not isinstance(chop_at, Number):
                chop_at = 3
    if max_sz and sz > max_sz:
        sz = max(max_sz)
        sz = (sz, sz * h_w)
        img.thumbnail(sz)
        if chop_at:
            img = _chop_at(img, chop_img, chop_at)
        save_it = True
    elif chop_at:
        img = _chop_at(img, chop_img, chop_at)
        save_it = True
    try:
        if save_it:
            save_it = path.join(gettempdir(), path.basename(fn))
            img.save(save_it)
            fn = save_it
        pic = rng.sheet.pictures.add(fn)
        sz = (rng.width, rng.height)
        fn = [i[0] - 2 * i[1] for i in zip(sz, margins)]
        fn = min(((fn[0], fn[0] * h_w), (fn[1] / h_w, fn[1])))
        aligns = mp.get("alignment", "C,M").split(",")
        aligns = [_pos(x, sz[idx], margins[idx], fn[idx], aligns[idx]) for idx, x in enumerate((rng.left, rng.top,))]
        fn = (aligns[0], aligns[1], fn[0], fn[1])
        pic.left, pic.top, pic.width, pic.height = fn
    finally:
        if save_it:
            remove(save_it)
    return pic

def _chop_at(orgimg, chop_img, chop_at=3):
    """
    return the cordinal/resized chop file for paste
    """
    osz, chsz = [x / chop_at for x in orgimg.size], chop_img.size
    ch_r = chsz[1] / chsz[0]
    chsz, osz = [
        int(x) for x in min((osz[0], osz[0] * ch_r), (osz[1] / ch_r, osz[1]))
    ], orgimg.size
    osz, chop_img = ([osz[idx] - val for idx, val in enumerate(chsz)],
                     chop_img.resize(chsz))
    orgimg.paste(chop_img, osz, chop_img)
    return orgimg


def NamedRanges(rng, **kwds):
    """ return the data under or include the range as namedlist list
    @param skip_first_row: boolean, don't process the first row, default is False
    @param name_map: the name->title mapping, see @list2dict FMI, default is None
    @param col_cnt: the count of columns to search, default is 0, that is unlimited
    """
    if not rng:
        return None
    cur_region = rng.current_region if rng.rows.count < 2 else rng
    if rng.size > 1:
        rng = rng[0]
    if kwds.get("skip_first_row"):
        rng = rng.offset(1, 0)
    sht, org_pt = rng.sheet, (rng.row, rng.column)
    var = kwds.get("col_cnt", kwds.get("colcnt")) or 0
    e_colidx = org_pt[1] + var if var > 0 else cur_region.last_cell.column
    tt_rows, var = (65000, 0), False

    var = [
        x for x in sht.range(rng, sht.range(org_pt[0], e_colidx)).columns
        if x.api.mergecells
    ]
    if var:
        for cell in var:
            mr = apirange(cell.api.mergearea)
            tt_rows = (min(tt_rows[0], mr.row), max(tt_rows[1],
                                                    mr.last_cell.row))
    else:
        tt_rows = (org_pt[0],) * 2
    th = sht.range(
        sht.range(tt_rows[0], org_pt[1]), sht.range(tt_rows[1], e_colidx))
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
    rng = sht.range(
        sht.range(tt_rows[1] + 1, org_pt[1]),
        sht.range(cur_region.last_cell.row, e_colidx))
    th = rng.value
    #one row case, xlwings return a 1-dim array only, make it 2D
    if rng.rows.count == 1:
        th = [th]
    th.insert(0, var)
    return NamedLists(th, getvalue(kwds, "name_map,alias"))


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


_col_idx = lambda chr: ord(chr) - 64  #ord('A') is 65
_col_pow = (
    1,
    26,
    26**2,
    26**3,
)


def col(c_i):
    """ given a colname or an index, return the related idx or name,
    examples:
        col(1) == 'A'
        col('A') == 1
        col('AA') = 27
    """
    if isinstance(c_i, str):
        if len(c_i) == 1:
            return _col_idx(c_i)
        s = 0
        for idx, ch in enumerate(c_i[::-1]):
            s += _col_pow[idx] * _col_idx(ch)
    else:
        s = col_name(c_i)
    return s


def _a2(addr):
    addr = addr.split("$")[1:]
    return addr[0], int(addr[1])


def _rows_or_cols(addr, row=True):
    """
    return the rows or cols inside the given address or range
    """
    if not isinstance(addr, str):
        addr = addr.address
    lsts, keys = [], set()
    for x in addr.split(","):
        ss = x.split(":")
        var = _a2(ss[0])
        var = var[1] if row else col(var[0])
        if var in keys:
            continue
        keys.add(var)
        if len(ss) == 2:
            if row:
                rx = (_a2(ss[0])[1], _a2(ss[1])[1])
            else:
                rx = (col(_a2(ss[0])[0]), col(_a2(ss[1])[0]))
            lsts.append(rx)
        else:
            lsts.append((var,) * 2)
    return lsts


def hidden(sht, row=True):
    """ return the hidden row/column inside a sheet's used ranged """
    lsts, rng0 = [], sht.used_range
    if not rng0:
        return None
    _rc = lambda rg, row: rg.row if row else rg.column
    idx, midx = 1, _rc(rng0.last_cell, row)
    rng, ridxs = None, _rc(rng0, row)
    # first several rows not in the used-ranged, append them
    if idx < ridxs:
        rng0 = sht.range(sht.cells(1, 1), rng0.last_cell)
    try:
        # specialCells failed when the rng0 at the end or some other criteria
        rng, ridxs = apirange(
            rng0.api.SpecialCells(12)), None  #xlCellTypeVisible
        ridxs = _rows_or_cols(rng.address, row)
    except:
        rng = None

    if rng is None:
        return lsts or [
            (idx, midx),
        ] if idx < midx else None

    for r in ridxs:
        if r[0] > idx:
            lsts.append((idx, r[0] - 1))
        idx = r[1] + 1
    if idx <= midx:
        lsts.append((idx, midx))
    return lsts if lsts else None


# an appmgr factory, instead of using app(), use appmgr.acq()/appmgr.ret()
appmgr = _newappmgr()
