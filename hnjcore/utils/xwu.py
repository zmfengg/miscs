# coding=utf-8    

'''
Created on Apr 19, 2018
Utils for xlwings's not implemented but useful function

@author: zmFeng
'''
import xlwings.constants as const
import xlwings
import os
import random

__all__ = ["app","find","fromtemplate","listodict","usedrange"]

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


def listodict(lst, trmap=None, dupdiv = "", bname = None):
    """ turn a list into zero-id based, name -> id lookup map 
    @param lst: the list or one-dim array containing the strings that need to do the name-> pos map
    @param trmap: An translation map, make the description -> name translation, if ommitted, description become name
                  if the description is not sure, split them with candidates, for example, "Job,JS":"jono"
    @param dupdiv: when duplicated item found, a count will be generated, dupdiv will be
        placed between the original and count
    @param blkcn: default name for the blank item
    @return: a dict with name -> id map   
    """
    if not lst: return None, None
    lstl = [x.lower()  if x and isinstance(x,basestring) else "" for x in lst]
    mp = {}
    for ii in range(len(lstl)):
        x = lstl[ii]
        if not x and bname:
            lstl[ii] = bname
        if x in mp:
            mp[x] += 1
            if dupdiv == None: dupdiv = ""
            lstl[ii] += dupdiv + str(mp[x])
        else:
            mp[x] = 0
    if not trmap:
        trmap = {}
    else:
        for x in [x for x in trmap.keys() if(x.find(",") >= 0)]:
            for y in x.split(","):
                if not y: continue
                y = y.lower()
                cnds = [x0 for x0 in range(len(lstl)) if lstl[x0].find(y) >= 0]
                if(len(cnds) > 0):
                    s0 = str(random.random())
                    lstl[cnds[0]] = s0
                    trmap[s0] = trmap[x]                    
                    break
    return dict(zip([trmap[x] if(x in trmap) else x for x in lstl], range(len(lstl))))


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