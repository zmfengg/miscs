#   coding = utf-8    

'''
Created on Apr 19, 2018
Utils for xlwings's not implemented but useful function

@author: zmFeng
'''
import xlwings.constants as const


def app(vis=True):    
    """ launch an excel or connect to existing one """
    
    import xlwings
    flag = xlwings.apps.count == 0
    return flag, \
        xlwings.apps.active if not flag else xlwings.App(visible=vis)


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


def arrtodict(lst,trmap = None):
    """ turn a list into zero-id based, name -> id lookup map 
    @param lst: the list or one-dim array containing the strings that need to do the name-> pos map
    @param trmap: An translation map, make the description -> name translation, if ommitted, description become name
                  if the description is not sure, split them with candidates, for example, "Job,JS":"jono"
    @return: a dict with name -> id map   
    """
    if(len(lst) == 0): return None,None
    lstl = [x.lower() for x in lst]
    if(not trmap): trmap = {}
    for x in [x for x in trmap.keys() if(x.find(",") >= 0)]:
        for y in x.split(","):
            cnds = [x0 for x0 in lstl if(len(y) >0 and x0.find(y) >= 0)]
            if(len(cnds) > 0):
                trmap[cnds[0]] = trmap[x]
                break
    return dict(zip([trmap[x] if(x in trmap) else x for x in lstl],range(len(lstl))))
    