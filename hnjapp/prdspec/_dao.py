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
from utilz.miscs import Literalize, Segments, trimu
from math import ceil
from hnjapp.localstore import Codetable
from sqlalchemy import and_
from ..common import config


class PrdSpec(SvcBase):
    '''
    save/retrieve data for a product specifiction form
    '''
    pass

class NameItem(object):
    ''' using codetable item to store data
    Args:
        cd: a codetable instance
    '''
    def __init__(self, cd=None, cs=None):
        self._data = cd or Codetable(name='_style_name', codec1=cs)

    @property
    def cat(self):
        return self._data.codec0
    
    @cat.setter
    def cat(self, cat):
        self._data.codec0 = trimu(cat)

    @property
    def name(self):
        return self._data.codec1

    @name.setter
    def name(self, name):
        self._data.codec1 = name

    @property
    def tag(self):
        return 'NEW' if self._data.tag == 0 else 'BUFFER'

    @tag.setter
    def tag(self, tag):
        self._data.tag = 0 if trimu(tag) != 'BUFFER' else 100

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
    def __init__(self, trmgr):
        super().__init__(trmgr)
        self._gtr = NameGtr()

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

class CDDao(SvcBase):
    ''' get from dao while keeping data integrity
    '''

    def _get(self, cat):
        with self.sessionctx() as cur:
            return self._get_buffer(cat, cur) or self._get_top(cat, cur)

    def _get_buffer(self, cat, cur):
        # TODO::
        lst = cur.query(Codetable).filter((Codetable.name == '_style_name', Codetable.codec0 == cat, Codetable.tag == 100, )).all()
        if not lst:
            return None
        while True:
            for cd in lst:
                try:
                    cur.delete(cd)
                except:
                    cur.rollback()

    def _get_top(self, cat, cur):
        pass


class NameGtr(object):
    '''
    class for generate new named based on a prior one

    Args:

        charset="0123456789ABCDEFGHJKLMNPQRTUVWXY": the charset to use

        sn.digits=4: the number digit of SN#

        ver.digits=2: the number digit of Ver#

        ring.size.count=20: the count of ring size per version can hold

        ring.row_first=True: use row_first stegory for ring version/size creation

        ring.headers=['T']: category of ring product

    '''
    def __init__(self, **kwds):
        mp = config.get('prodspec.naming')
        _get = lambda x, df=None: kwds.get(x, mp.get(x, df))
        cs = _get('charset')
        self._gtr_sn, self._gtr_ver = (Literalize(cs, digits=_get(x)) for x in ('sn.digits', 'ver.digits'))
        self._gtr_one = Literalize(cs, digits=1)
        self._cache_args = cm = {}
        self._sgmgr = Segments(len(cs), _get('ring.size.count'), _get('ring.row_first', True))

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

    def versions(self, name):
        ''' return the versions available for given name. If name is not of

        ring, return None

        '''
        self._validate(name)
        if not self._is_ring(name):
            return None
        return [name + self._c2v(x) for x in self._sgmgr.segments]

    def sizes(self, name):
        ''' return the sizes available for given name. If name is not of

        ring, return None

        Args:
            name: the style name with version. If no version is provided, use first version

        '''
        self._validate(name)
        if not self._is_ring(name):
            return None
        cat, digits, ver = self._split(name)
        return [cat + digits + self._c2v(x) for x in self._sgmgr.sectors(self._c2v(ver))]

    def all(self, name=None, file=None):
        ''' just for demonstrating all the versions available for a given name
        Args:
            name(string):   style name without version, for example P1234. If no name provided, a name of ring will be used
        '''
        gtr = self._gtr_sn
        if name:
            self._validate(name)
            nm = name[:self._cat_digits + gtr.digits]
        else:
            nm = next(iter(self._cache_args['_ring_hdrs'])) + gtr.next(gtr.charOf(0))
        if not self._is_ring(nm):
            # will return too many items
            return None
        nm0 = nm
        lst = self._sgmgr.all()
        if file:
            nc = len(lst) * len(lst[0])
            tc = self._gtr_ver.radix ** self._gtr_ver.digits
            print('Names for %s, NameCnt:%d, VerCnt=%d, SzPerVer=%d, UseRate=%3.2f%%' % (nm0, nc, len(lst), len(lst[0]), nc / tc * 100))
            i2n = lambda addr: ''.join(self._gtr_one.charOf(x) for x in addr)
            for lst1 in lst:
                print('*' + i2n(lst1[0]) + ': ' + ','.join([i2n(x) for x in lst1[1:]]), file=file)
        return lst
