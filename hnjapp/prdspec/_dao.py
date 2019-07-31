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
from utilz.miscs import Literalize
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
    ROW, COL = 0, 1
    def __init__(self, **kwds):
        mp = config.get('prodspec.naming')
        _get = lambda x, df=None: kwds.get(x, mp.get(x, df))
        cs = _get('charset')
        self._gtr_sn, self._gtr_ver = (Literalize(cs, digits=_get(x)) for x in ('sn.digits', 'ver.digits'))
        self._gtr_one = Literalize(cs, digits=self._gtr_ver.digits - 1)
        self._cache_args = cm = {}
        cm['_szcnt'] = m = _get('ring.size.count')
        self._calc = _SpanCalc(self._gtr_one, m)
        # what's left for the fragment
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

    def next(self, current, level=0):
        ''' get next name based on current
        Args:

            current:    the current name, can be without sn and version

            level=0:    0 for getting new version, 1 for new ring size based on current version

        throws:

            OverflowError if current is already at the end
        '''
        self._validate(current)
        cat, sn, ver = self._split(current)
        if self._is_ring(current):
            if level == 0:
                ver = self._spawnVersion(ver)
            else:
                ver = self._spawnSize(ver)
        else:
            ver = self._gtr_ver.next(ver)
        return cat + sn + ver

    def _spawnSize(self, ver):
        if not ver:
            return self._calc.zero * self._gtr_ver.digits
        nxt, calc = self._gtr_one.next, self._calc
        if calc.size == 1:
            raise OverflowError('1 item span should not have sub item')
        lvl = calc.get_level(ver)
        def _eo_span(v):
            idx = self._gtr_one.valueOf(v)
            if idx and (idx + 1) % calc.size == 0:
                raise OverflowError('level 0 ends')
            return v
        if lvl == 0:
            return ver[calc.ROW] + nxt(_eo_span(ver[self.COL]))
        if lvl == 1:
            return nxt(_eo_span(ver[self.ROW])) + ver[calc.COL]
        return calc.next_element(ver)


    def _spawnVersion(self, ver):
        g1, calc = self._gtr_one.next, self._calc
        if not ver:
            # init
            return self._gtr_ver.next()
        radix, mx = calc.radix, calc.client_org - calc.size + 1
        verxi = self._gtr_one.valueOf(ver[1])
        if verxi < mx:
            # size level 0
            try:
                return g1(ver[0]) + ver[1]
            except OverflowError:
                # next span of level 0
                if verxi + calc.size < radix:
                    return calc.zero + g1(ver[1], steps=calc.size)
        verxi = self._gtr_one.valueOf(ver[0])
        if verxi < mx:
            # level 1
            if verxi + calc.size * 2 < radix:
                return g1(ver[0], steps=calc.size) + ver[1]
            try:
                return calc.zero + g1(ver[1])
            except OverflowError:
                # header for level 2
                hgt = calc.client_height
                if hgt * hgt < calc.size:
                    raise OverflowError('no chance to enter level 2')
                return self._gtr_one.charOf(calc.client_org) * 2
        return calc.add_span(ver)

    def getSeries(self, name):
        ''' return series of given name. For non-ring, return itself
        '''
        self._validate(name)
        cat, sn, ver = self._split(name)
        if ver[-1] <= self._cache_args['_msz']:
            ver = ver[0] + self._cache_args['_zero']
        elif ver[0] <= self._cache_args['_msz']:
            ver = self._cache_args['_zero'] + ver[1:]
        else:
            rpg = self._cache_args['_frag_hdr']
            ver = self._gtr_sn.charOf(self._gtr_sn.valueOf(ver[0]) // rpg * rpg) + self._cache_args['_msz1']
        return cat + sn + ver


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
        nm0 = nm
        lst = []
        while True:
            try:
                nm = self.next(nm)
                lst1 = [nm, ]
                sn1 = nm
                while True:
                    try:
                        sn1 = self.next(sn1, 1)
                        lst1.append(sn1)
                    except OverflowError:
                        break
                # print('size of %s: %s' % (nm, lst1))
                if len(lst1) == self._calc.size:
                    lst.append(lst1)
            except OverflowError:
                break
        nc = len(lst) * len(lst[0])
        tc = self._gtr_ver.radix ** self._gtr_ver.digits
        print('Names for %s, NameCnt:%d, VerCnt=%d, SzPerVer=%d, UseRate=%3.2f%%' % (nm0, nc, len(lst), len(lst[0]), nc / tc * 100))
        for lst1 in lst:
            print('*' + lst1[0] + ': ' + ','.join(lst1[1:]))
        return lst

class _SpanCalc(object):
    ROW, COL = 0, 1

    def __init__(self, ltr, spnSz):
        self._size = spnSz
        self._ltr = ltr
        self._cache = mp = {}
        rdx = ltr.radix
        x = rdx // spnSz * spnSz
        mp['_client_org'] = x
        if x < rdx:
            x = ltr.radix - x
            mp['_client_height'] = x
            mp['_client_area'] = x * x
        else:
            mp['_client_height'] = mp['_client_area'] = 0
        mp['_zero'], mp['_radix'] = ltr.next(), rdx

    def _get_cache(self, key, calc):
        if key not in self._cache:
            self._cache[key] = calc()
        return self._cache[key]

    @property
    def zero(self):
        ''' the zero literal in the charset
        '''
        return self._cache['_zero']

    @property
    def size(self):
        ''' span size
        '''
        return self._size

    @property
    def radix(self):
        ''' len of charset
        '''
        return self._cache['_radix']

    @property
    def client_height(self):
        ''' height of the span area
        '''
        return self._cache['_client_height']

    @property
    def client_area(self):
        ''' the size of client area
        '''
        return self._cache['_client_area']

    @property
    def client_org(self):
        ''' original point of the span area
        '''
        return self._cache['_client_org']

    def get_level(self, ver):
        ''' return the level of given address
        '''
        rc, org = [self._ltr.valueOf(x) for x in ver], self.client_org
        if rc[self.COL] < org:
            return 0
        if rc[self.ROW] < org:
            return 1
        return 2

    def next_element(self, ver):
        ''' calc the next element of addr, assume addr is in client_area
        '''
        if isinstance(ver, str):
            ver = self._convert(ver)
        rng = self.get_range(ver)
        r_c = self._add(ver, 1)
        if self._dim_convert(r_c) <= self._dim_convert(rng[1]):
            return ''.join(self._convert(r_c, False))
        raise OverflowError('end of span')

    def get_range(self, ver):
        ''' return the head/tail of given ver
        '''
        if isinstance(ver, str):
            ver = self._convert(ver)
        ln = self._dim_convert(ver)
        sz = self.size
        hdr = ln // sz
        tail = (hdr + 1) * sz - 1
        if tail >= self.client_area:
            raise OverflowError()
        cv = lambda x: ''.join(self._convert(self._dim_convert(x, False)))
        return [cv(hdr * sz), cv(tail)]

    def _dim_convert(self, sz, offset=False):
        ''' convert between one-dim and 2 dim
        '''
        if isinstance(sz, str):
            sz = self._convert(sz, offset=False)
        if isinstance(sz, (tuple, list)):
            org = self.client_org
            if sz[self.ROW] < org:
                org = 0
            return (sz[self.ROW] - org) * self.client_height + sz[self.COL] - org
        hgt = self.client_height
        org = self.client_org if offset else 0
        return [sz // hgt + org, sz % hgt + org]

    def _convert(self, ver, offset=True):
        ''' translate the addr to row/col or verse
        '''
        org = self.client_org if offset else 0
        if isinstance(ver, str):
            return [self._ltr.valueOf(x) - org for x in ver]
        return [self._ltr.charOf(x + org) for x in ver]

    def add_span(self, ver):
        ''' add one span to given ver
        '''
        if not self.client_height:
            raise OverflowError('perfect fit, no level 2 needed')
        if isinstance(ver, str):
            ver = self._convert(ver)
        if ver[self.ROW] * self.client_height + ver[self.COL] + self.size * 2 > self.client_area:
            raise OverflowError('level 2 overflow')
        tar = self._add(ver)
        return ''.join([self._ltr.charOf(x) for x in tar])

    def _add(self, r_c, steps=None, offset=True):
        if not steps:
            steps = self.size
        if isinstance(r_c, str):
            r_c = self._convert(r_c)
        return self._dim_convert(self._dim_convert(r_c) + steps, offset)
