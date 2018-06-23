# coding=utf-8    

'''
Created on Apr 19, 2018
Utils for xlwings's not implemented but useful function

@author: zmFeng
'''
import xlwings.constants as const
import xlwings
import os
from ._miscs import list2dict

__all__ = ["app","find","fromtemplate","list2dict","usedrange"]

def app(vis=True):    
    """ launch an excel or connect to existing one
    return (flag,app), where flag is True means it's created by me, the caller should
    dispose() it
    """
    
    flag = xlwings.apps.count == 0
    app = xlwings.apps.active if not flag else xlwings.App(visible=vis, add_book=False)
    if app: app.api.DisplayAlerts = False
    return flag, app


def usedrange(sh):
    """
    find out the used range of the given sheet
    @param sh: the worksheet you want to find used range from
    """
    ur = sh.api.UsedRange
    rows = (ur.Row, ur.Rows.Count)
    cols = (ur.Column, ur.Columns.count)
    return sh.range(*zip(rows, cols))


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
    apiRng = sh.api.Cells.Find(What=val, After=aftr, \
                   LookAt=lookat, LookIn=lookin, \
                   SearchOrder=so, SearchDirection=sd, \
                       MatchCase=matchCase)
    if(apiRng):
        apiRng = sh.range((apiRng.row, apiRng.column))
    return apiRng


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