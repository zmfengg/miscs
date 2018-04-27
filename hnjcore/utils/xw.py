#   coding = utf-8    

'''
Created on Apr 19, 2018
Utils for xlwings's not implemented but useful function

@author: zmFeng
'''
import xlwings.constants as const


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
