#! coding=utf-8
'''
* @Author: zmFeng
* @Date: 2019-01-17 10:30:35
* @Last Modified by:   zmFeng
* @Last Modified time: 2019-01-17 10:30:35
* holds the misc services except db ones
'''
from datetime import datetime
from os import path
try:
    from os import scandir
except:
    from scandir import scandir

from utilz import getvalue, trimu


class StylePhotoSvc(object):
    '''
    service for getting style photo
    '''
    TYPE_STYNO = "styno"
    TYPE_JONO = "jono"

    def __init__(self, root=r"\\172.16.8.91\Jpegs\style", level=3):
        self._root = root
        self._min_level = 2
        self._level = max(min(5, level), self._min_level)

    def _build_root(self, styno):
        parts = [styno[:x] for x in range(self._min_level, self._level + 1)]
        return path.join(self._root, *parts)

    def getPhotos(self, styno, atype="styno", hints=None, **kwds):
        ''' return the valid photos of given style or jo#
        @param atype: argument type, can be one of StylePhotoSvc.TYPE_JONO/TYPE_STYNO
        return a list of files sorted by below criterias:
            .hints hit(DESC)
            .modified date(DESC)
        '''
        if not styno:
            return None
        jo = jono = None
        if atype == self.TYPE_JONO:
            eng = getvalue(kwds, "engine cache_db")
            if not eng:
                # no helper to convert JO# to sty#
                return None
            jono, jo = styno, self._get_jo(styno)
            styno = jo.style.name.value
            # find the JO# with same SKU#
            hints = (hints + "," + jono) if hints else jono
        if hints:
            hints = hints.split(",")
        root = self._build_root(styno)
        if not path.exists(root):
            return None
        styno = trimu(styno)
        ln, lst = len(styno), []
        fns = [x for x in scandir(root) if x.is_file() and trimu(x.name[:ln] == styno)]
        if not fns:
            return None
        if len(fns) > 1:
            for fn in fns:
                styno = fn.name
                if '0' <= styno[ln] <= '9':
                    continue
                lst.append((fn, self._match(root, styno, ln, hints), ))
            if not lst:
                return None
        else:
            return [fns[0].path, ]
        return [x[0].path for x in sorted(lst, key=lambda x: x[1], reverse=True)]

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
        return (cand - datetime.today()).days
