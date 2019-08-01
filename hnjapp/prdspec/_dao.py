'''
#! coding=utf-8
@Author:        zmFeng
@Created at:    2019-07-19
@Last Modified: 2019-07-19 1:19:15 pm
@Modified by:   zmFeng
help to save/retrieve from db and judge if one item need versioning or already exists
'''

from hnjapp.svcs.db import SvcBase
from utilz.resourcemgr import ResourceCtx
from utilz.miscs import Literalize, Segments
from math import ceil
from ..common import config


class PrdSpec(SvcBase):
    '''
    save/retrieve data for a product specifiction form
    '''
    pass

class NameDao(SvcBase):
    ''' controlling the acquiring/returing of name resource
    A common routine:
        dao = NameDao()
        try:
            name = dao.acq(cat)
            ....
            dao.commit(name)
            ....
        except:
            dao.rollback(name)
    '''
    def acq(self, cat):
        ''' return the next item of given cat
        '''
        with self.sessionctx() as sm:
            pass

    def commit(self, name):
        ''' the named is occupied and should never be acquired again
        '''
        pass

    def rollback(self, code):
        ''' return a code acquired before to this controller
        '''
        pass


class NameGtr(object):
    '''
    class for generate new named based on a prior one

    Resources defined:

        charset="0123456789ABCDEFGHJKLMNPQRTUVWXY": the charset to use

        sn.digits=4: the number digit of SN#

        ver.digits=2: the number digit of Ver#

        ring.size.count=20: the count of ring size per version can hold

        ring.headers=['T']: category of ring product

    Naming pattern for version field definations:

    Non_Ring case:

        Have up to len(charset) ** 2 versions

    Ring case:

        for each ring version, there can hold up to $max_rg_sz_cnt$ sizes, using below pattern

        Naming pattern for ring:

            version: 00, size: 00-0K
            version: 10, size: 10-1K
            version: 20, size: 20-2K
            ...
            version: K0, size: K0-KK
            version: L0, size: L0-LK
            ...
            version: Y0, size: Y0-YK

            fragment :., [L-Y]
            version: 0L, size: 0L-0Y + 1L-1Y
            version: 2L, size: 2L-2Y + 3L-3Y
            ...
            version  XL, size: XL-XY + YL-YY

        Naming pattern for ring 1:
            level 0: [0, Y][0, k],
                example: 00,01,02...0K;10,11,12...1K;...;Y0,Y1,Y2...YK
            level 1: [0, Y][L, Y],
                example: 0L,1L,2L...KL;0M,1M,2M...KM;...;0Y,1Y,2Y...KY
            level 2: [L, Y]
                            [L, Y],
                example: LL,LM,LN...
            figure(assume charsetCnt=8, szCnt=5):
            00000111
            00000111
            00000111
            00000111
            00000111
            00000222
            00000222

        call _ring_sample() to generate all name for a ring

    '''
    def __init__(self, **kwds):
        mp = config.get('prodspec.naming')
        _get = lambda x, df=None: kwds.get(x, mp.get(x, df))
        cs = _get('charset')
        self._gtr_sn, self._gtr_ver = (Literalize(cs, digits=_get(x)) for x in ('sn.digits', 'ver.digits'))
        self._gtr_one = Literalize(cs, digits=1)
        self._cache_args = cm = {}
        self._sgmgr = Segments(len(cs), _get('ring.size.count'))

        hdrs = _get('ring.headers')
        cm['_ring_hdrs'] = set(hdrs)
        self._cat_digits = len(next(iter(hdrs)))

    def _is_ring(self, name):
        return name[0] in self._cache_args['_ring_hdrs']

    def _split(self, name):
        idx = self._cat_digits + self._gtr_sn.digits
        return name[:self._cat_digits], name[self._cat_digits: idx], name[idx:]

    def _validate(self, name):
        if not name or not self._gtr_sn.isvalid(name[self._cat_digits:]):
            raise AttributeError('%s Blank or contains character(s) not in %s' % (name, self._gtr_sn.charset))

    def _c2v(self, ch):
        ''' char to value or vise verse
        '''
        if not ch:
            return ch
        if isinstance(ch, str):
            return [self._gtr_one.valueOf(x) for x in ch]
        return ''.join(self._gtr_one.charOf(x) for x in ch)

    def next(self, current, version=True):
        ''' get next name based on current
        Args:

            current:    the current name, can be without sn and version

            version=True:    True for getting new version, else for new ring size based on current version

        throws:

            OverflowError if current is already at the end
        '''
        self._validate(current)
        cat, sn, ver = self._split(current)
        if self._is_ring(current):
            ver = self._sgmgr.next(self._c2v(ver), version)
            ver = self._c2v(ver)
        else:
            ver = self._gtr_ver.next(ver)
        return cat + sn + ver

    def header(self, name):
        ''' return header of given name. For non-ring, return itself
        '''
        self._validate(name)
        if not self._is_ring(name):
            return name
        cat, sn, ver = self._split(name)
        ver = self._sgmgr.range(self._c2v(ver))
        return cat + sn + self._c2v(ver[0])


    def allNamesOf(self, name=None):
        ''' just for demonstrating all the versions available for a given name
        Args:
            name(string):   style name without version, for example P1234. If no name provided, a name of ring will be used
        '''
        gtr = self._gtr_sn
        if name:
            nm = name[:self._cat_digits + gtr.digits]
        else:
            nm = next(iter(self._cache_args['_ring_hdrs'])) + gtr.next(gtr.charOf(0))
        if not self._is_ring(nm):
            # will return too many items
            return None
        nm0 = nm
        lst = self._sgmgr.all()
        nc = len(lst) * len(lst[0])
        tc = self._gtr_ver.radix ** self._gtr_ver.digits
        print('Names for %s, NameCnt:%d, VerCnt=%d, SzPerVer=%d, UseRate=%3.2f%%' % (nm0, nc, len(lst), len(lst[0]), nc / tc * 100))
        i2n = lambda addr: ''.join(self._gtr_one.charOf(x) for x in addr)
        for lst1 in lst:
            print('*' + i2n(lst1[0]) + ': ' + ','.join([i2n(x) for x in lst1[1:]]))
        return lst
