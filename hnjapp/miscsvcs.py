#! coding=utf-8
'''
* @Author: zmFeng
* @Date: 2019-01-17 10:30:35
* @Last Modified by:   zmFeng
* @Last Modified time: 2019-01-17 10:30:35
* holds the misc services except db ones
'''
from os import path, listdir
from utilz import getvalue, trimu
from datetime import datetime

class StylePhotoSvc(object):
    TYPE_STYNO = "styno"
    TYPE_JONO = "jono"

    def __init__(self, root, level=2):
        self._root =root
        self._min_level = 2
        self._level = max(min(5, level), self._min_level)        
    
    def _build_root(self, styno):
        parts = [styno[:x] for x in range(self._min_level, self._level)]
        return path.join(self._root, *parts)
    
    def getPhotos(self, styno, atype=StylePhotoSvc.TYPE_STYNO, hints=None, **kwds):
        ''' return the valid photos of given style or jo#
        @param atype: argument type, can be one of StylePhotoSvc.TYPE_JONO/TYPE_STYNO
        return a list of files sorted by below criterias:
            .hints hit(DESC)
            .modified date(DESC)
        '''
        jo = jono = None
        if atype == self.TYPE_JONO:
            eng = getvalue(kwds, "engine cache_db")
            if not eng:
                # no helper to convert JO# to sty#
                return None
            jono, jo = styno, self._get_jo(styno)
            styno = jo.style.name.value
            # find the JO# with same SKU#
            hints = (hints + "," + jono).split(",") if hints else (jono, )
        root = self._build_root(styno)
        fns = listdir(path.join(root, "%s" % styno))
        if not fns:
            return None
        styno = trimu(styno)
        ln, lst = len(styno), []
        for fn in fns:
            if fn[:ln] != styno:
                continue
            if '0' <= fn[ln] <= '9':
                continue
            lst.append(fn, self._match(root, fn, ln, hints))
        if not lst:
            return None
        return [path.join(x[0]) for x in sorted(lst, key=lambda x: x[1], reverse=True)]
        
    def _get_jo(self, jono):
        #TODO
        return jono
    
    def _match(self, root, fn, ln, hints):
        ''' if found in hints, result is positive, 
        else return the days to current days as negative
        so that the call can sort the result
        '''
        if hints:
            cand = trimu(fn[ln:fn.rfind('.')])
            if cand:
                for x in hints:
                    if cand.find(trimu(x)) >= 0:
                        return 100
        cand = datetime.fromtimestamp(path.getmtime(path.join(root, fn)))
        return (datetime.today() - cand).day
