'''
#! coding=utf-8
@Author:        zmFeng
@Created at:    2019-07-09
@Last Modified: 2019-07-09 3:24:14 pm
@Modified by:   zmFeng
the CZsize/From JO and ...
'''

from os import path
from utilz import stsizefmt, NamedLists, trimu
from ._nrlib import thispath
from hnjapp.svcs.db import SvcBase

class _SNS2Wgt(object):
    ''' given stone, shape and size, return weight
    '''
    def __init__(self):
        self._data = None

    def _load(self):
        with open(path.join(thispath, "sztbl.csv"), 'r') as fh:
            nls = [ln.split(",") for ln in fh if ln[0] != "#"]
        nls = NamedLists(nls)
        self._data = {trimu("%s,%s,%s" % (nl.stone, nl.shape, stsizefmt(nl.size, True))): float(nl.unitwgt) for nl in nls}

    def get(self, name, shape, sz):
        '''
        given name,shape and sz, return unit weight(in ct)
        '''
        if not self._data:
            self._load()
        key = trimu("%s,%s,%s" % (name, shape, stsizefmt(sz, True)))
        return self._data.get(key)
