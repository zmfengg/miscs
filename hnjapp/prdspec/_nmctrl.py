'''
#! coding=utf-8
@Author:        zmFeng
@Created at:    2019-07-19
@Last Modified: 2019-07-19 1:19:15 pm
@Modified by:   zmFeng
name acquire/return controller
'''


import logging
from datetime import datetime
from itertools import chain
from numbers import Number
from os import path, remove

from sqlalchemy import and_, func, or_
from sqlalchemy.orm import Query

from hnjapp.localstore import Codetable
from hnjapp.svcs.db import SvcBase
from utilz import NA, karatsvc, xwu, NamedLists
from utilz._jewelry import stsizefmt, UnitCvtSvc
from utilz.miscs import Literalize, Segments, triml, trimu

from ..common import _logger, config
from ._fromjo import JOFormHandler
from ._tables import Mat, Style, Stymat, Styp, Stypi, Stypidef, Stystset


class _NameItem(object):
    ''' adopt a codetable to named item to avoid remember what field stands for.
    Args:
        cd(Codetable)   : a codetable instance
        name(String)    : name of the style
    '''
    KEY_NAME = '_style_name'

    def __init__(self, cd=None, name=None):
        if not cd or cd.name != _NameItem.KEY_NAME:
            cd = Codetable(name=_NameItem.KEY_NAME)
            cd.coden0 = cd.coden1 = cd.coden2 = 0
            cd.codec0 = cd.codec2 = NA
            cd.description = 'Style naming of category'
            cd.tag = 0
            cd.createdate = cd.lastmodified = cd.coded0 = cd.coded1 = cd.coded2 = datetime.now()
        self._data = cd
        if name:
            self.name = name

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
    def key(self):
        ''' key for querying in the codetable
        '''
        return self._data.name

    @property
    def isbuffer(self):
        return self._data.tag == 100

    @isbuffer.setter
    def isbuffer(self, flag):
        self._data.tag = 100 if flag else 0

    @property
    def rawData(self):
        return self._data

    @staticmethod
    def part(name):
        s0 = triml(name)
        if s0 == 'name':
            return Codetable.codec1
        elif s0 == 'cat':
            return Codetable.codec0
        elif s0 == 'key':
            return Codetable.name
        return None

class NameRepo(object):
    ''' interface for StyleName Repo that can support NameService's actions
    '''

    def __init__(self):
        super().__init__()

    def getmax(self, cat):
        ''' return the maximum item with cat as prefix
        '''
        return None

    def exists(self, name):
        ''' check if given name exists
        '''
        return False


class NameSvc(SvcBase):
    ''' controlling the acquiring/returing of name resource
    Args:
        trMgr(SessionMgr):  a sessionMgr to access database
        gtr(NameGtr):       a NameGtr instance
        styDao(MockDao):    an object that implements MockDao interface
    '''

    def __init__(self, trmgr, gtr=None, styDao=None):
        super().__init__(trmgr)
        self._gtr = gtr or NameGtr()
        self._ctrlDao = CtrlDao(trmgr, self._gtr)
        self._styDao = styDao or NameRepo()

    def get(self, cat, version=True):
        '''
        return the next item of given cat, when len(cat) is longer than gtr's cat, it's a:
        0.  Cat:    return next Cat+SN+00
        1.  Cat+SN  return next Cat+SN+00 or Cat+SN+VER
        2.  Cat+SN+VER  return next Cat+SN+VER
        3.  Cat+SN+VER(Ring)

        I'll take care of such case
        Args:
            cat(String): the cat or cat+sn + cat+sn+ver
            version(optional): new version or sub-version::
                name: create a new style# with initial version
                version: create a new style# with the same Cat+SN as cat and new version
                size: RING only, create new Style# with the same Cat+SN+VER, but new size
        '''
        name = None
        nt = self._gtr.nameType(cat)
        if nt == NameGtr.NT_INVALID:
            return name
        if not self._gtr.subverable(cat) and not version:
            raise Exception('cat(%s) is not sub-versionable' % cat)
        cds = None
        with self.sessionctx():
            while name is None or self._styDao.getmax(self._gtr.partOf(name, NameGtr.NT_CAT_SN)):
                name0 = None
                if nt == NameGtr.NT_CAT:
                    # acq new style, from controller or actual db, then update controller
                    name, cds = self._ctrlDao.get(cat)
                    if not name:
                        name = self._styDao.getmax(cat)
                    else:
                        name0 = name.name
                        name = self._gtr.next(name.name, name.isbuffer)  # CAT always retrun version
                else:
                    if version:
                        name = self._gtr.partOf(cat, NameGtr.NT_CAT_SN)
                        name = self._styDao.getmax(name)
                        if name:
                            name = self._gtr.next(name, version)
                            # versioning does not need confirm, so just return
                            break
                    else:
                        name = self._gtr.next(cat, version)
                if not name:
                    # no matter what, get the next cat+sn+ver of given cat because there is no hint,
                    # that means, it's a initilization
                    name = self._gtr.next(cat, version)
                # because the controller does not keep version/size info. so trim it
                lst = [self._gtr.partOf(x, NameGtr.NT_CAT_SN) if x else None for x in (name, name0)]
                lst.append(cds)
                self._ctrlDao.confirm(*lst)
        return name

    def recycle(self, name):
        ''' return a code acquired before to this controller
        Args:   the name to return
        Returns:
            True if successfully recycled
        '''
        name0 = self._gtr.partOf(name, NameGtr.NT_CAT_SN)
        if not name0:
            return False
        if self._styDao.getmax(name0):
            _logger.debug('Still records of (%s) inside repository, can not recycle' % name)
        # recycle item always as buffer, not top
        return self._ctrlDao.recycle(name0) is not None

class CtrlDao(SvcBase):
    ''' get from dao while keeping data integrity
    '''

    def __init__(self, trmgr, nameGtr):
        super().__init__(trmgr)
        self._gtr = nameGtr or NameGtr()

    def get(self, cat):
        cd = None
        with self.sessionctx() as cur:
            lst = self._get_all(cur, cat)
            cd = lst[-1] if lst else None
        return (_NameItem(cd) if cd else None), lst

    def confirm(self, nname, oname, cds=None):
        ''' confirm the creation of a new name
        Args:
            nname(String):  the new name
            oname(String):  the old new that the new name based on, can be None
            cds(Collection(Codeble)): a collection of codetable queried befored, for performance tuning only, set it to None if called from out of this package
        '''
        tops = []
        rvs = []
        with self.sessionctx() as cur:
            if oname:
                if not cds:
                    cat = self._gtr.partOf(oname, NameGtr.NT_CAT)
                    cds = self._get_all(cur, cat)
                if cds:
                    cds = [_NameItem(x) for x in cds]
                    tops = [x for x in cds if not x.isbuffer]
                    cds = [x for x in cds if x.name == oname]
                    if cds:
                        for x in cds:
                            if x.isbuffer:
                                rvs.append(x)
                        if len(tops) > 1:
                            rvs.extend(tops[1:])
            if rvs:
                for x in rvs:
                    cur.delete(x.rawData)
            tops = tops[0] if tops else _NameItem()
            if oname and tops.name != oname: # this top is malform? should it be top
                var = self._gtr.compare(tops.name, nname)
                if var < 0:
                    # degrade it to buffer, the sessionMgr will handle this changes when commit
                    tops.isbuffer = True
                    tops = _NameItem()
                    tops.isbuffer = False
                else:
                    tops = None
            elif not oname:
                if not cds:
                    cat = self._gtr.partOf(nname, NameGtr.NT_CAT)
                    # only need to know if exists any item
                    cds = cur.query(Codetable).filter(and_(
                        _NameItem.part('key') == _NameItem.KEY_NAME,
                        _NameItem.part('cat') == cat)).count()
                if cds:
                    # this is not an init, maybe by-hand item, don't send it to controller
                    tops = None
                    _logger.debug("Name(%s) maybe by hand, won't be placed to controller" % nname)
            if tops:
                tops.name = nname
                tops.cat = self._gtr.partOf(nname, NameGtr.NT_CAT)
                tops.rawData.lastmodified = datetime.now()
                cur.add(tops.rawData)
            cur.flush()

    @staticmethod
    def _del_or_upd(cur, cd):
        ni = _NameItem(cd)
        rc = None
        # buffer item should be deleted while top item should be kept so codetable's id increases slow down a bit
        if ni.isbuffer:
            cur.delete(cd)
        else:
            # don't touch it because there will be a confirm process
            ni.name = NA
            rc = cd
        return rc

    def _put(self, cat, name):
        with self.sessionctx() as cur:
            lst = self._get_all(cur, cat)
            if lst:
                lst1 = [x for x in lst if _NameItem(x).name == name]
                if lst1:
                    for x in lst1[1:]:
                        cur.delete(x)
                    return lst1[0]
            cd = None
            if lst:
                ni = _NameItem(lst[0])
                if not ni.isbuffer and (not ni.name or ni.name == NA):
                    cd = lst[0]
            if not cd:
                cd = _NameItem(None, name).rawData
            cur.save(cd)
        return cd

    def get_all(self, cat):
        with self.sessionctx() as cur:
            return self._get_all(cur, cat)

    @staticmethod
    def _get_all(cur, cat):
        lst = cur.query(Codetable).filter(
            and_(
                _NameItem.part('key') == _NameItem.KEY_NAME,
                _NameItem.part('cat') == cat)).all()
        if lst and len(lst) > 1:
            lst = sorted(lst, key=lambda x: (x.tag, -x.lastmodified.timestamp()))
        if lst and _logger.isEnabledFor(logging.DEBUG):
            _logger.debug('Names (%s) from controller' % ','.join((_NameItem(x).name for x in lst)))
        return lst

    def _exists(self, name):
        ''' check if given name eixsts
        '''


    def recycle(self, name):
        ''' put an name back to the controller's buffer
        '''
        name0 = self._gtr.partOf(name, NameGtr.NT_CAT_SN)
        if not name0:
            return True
        if self._exists(name0):
            _logger.debug('Name(%s) already inside controller db, will not be put again' % name0)
        ni = _NameItem(None, name0)
        ni.isbuffer = True
        ni.cat = self._gtr.partOf(name0, NameGtr.NT_CAT)
        ni.rawData.lastmodified = datetime.now()
        with self.sessionctx() as cur:
            cur.add(ni.rawData)
            cur.flush()
        return ni


class NameGtr(object):
    '''
    class for generate new named based on a prior one

    Args:

        charset="0123456789ABCDEFGHJKLMNPQRTUVWXY": the charset to use

        sn.digits=4: the number digit of SN#

        ver.digits=2: the number digit of Ver#

        subver.capacity=20: the count of ring size per version can hold

        ring.row_first=True: use row_first stegory for ring version/size creation

        subver.headers=['T']: category of ring product

    '''
    # for nameType detection
    NT_INVALID = '_INVALID'
    NT_CAT = '_category'
    NT_CAT_SN = '_category_sn'
    NT_FULL = '_category_sn_ver'

    def __init__(self, **kwds):
        mp = config.get('prdspec.naming')
        _get = lambda x, df=None: kwds.get(x, mp.get(x, df))
        cs = _get('charset')
        self._gtr_sn, self._gtr_ver = (
            Literalize(cs, digits=_get(x)) for x in ('sn.digits', 'ver.digits'))
        self._gtr_one = Literalize(
            cs, digits=1)  # use for the liternal/index conversion
        self._cache_partTags = cm = {}
        self._sgmgr = Segments(
            len(cs), _get('subver.capacity'), _get('ring.row_first', True))

        hdrs = _get('subver.headers')
        cm['_subver_hdrs'] = set(hdrs)
        self._cat_digits = len(next(iter(hdrs)))

    @property
    def SnDigits(self):
        ''' length of the sn part
        '''
        return self._gtr_sn.digits

    @property
    def VerDigits(self):
        ''' length of a version part
        '''
        return self._gtr_ver.digits

    @property
    def CatDigits(self):
        ''' length of the category
        '''
        return self._cat_digits

    def nameType(self, name):
        '''
        check the name type by the length of it
        '''
        ln = len(name) if name else 0
        sg = self._cache_partTags.get('_type_segs')
        if not sg:
            sg = [self.CatDigits, self.SnDigits, self.VerDigits]
            for idx in range(1, len(sg)):
                sg[idx] += sg[idx - 1]
            rsts = [NameGtr.NT_CAT, NameGtr.NT_CAT_SN, NameGtr.NT_FULL]
            sg = dict(zip(sg, rsts))
            self._cache_partTags['_type_sgs'] = sg
        return sg.get(ln, NameGtr.NT_INVALID)

    def compare(self, n0, n1):
        ''' check if name0 > name0 in the case of SN#, different cat can not be compared
        Args:
        n0(String): name0
        n1(String): name1
        '''
        lsts = [self._split(x) for x in (n0, n1)]
        if lsts[0][0] != lsts[1][0]:
            raise Exception("different category is not comparable")
        sns = [x[1] for x in lsts if x[1]]
        if len(sns) == 2:
            sns = [self._c2v(x) for x in sns]
            return -1 if sns[0] < sns[1] else 0 if sns[0] == sns[1] else 1
        if not sns:
            return 0
        return 1 if lsts[0][1] else 0

    def subverable(self, name):
        ''' check if a given name is sub-versionable
        '''
        return self.partOf(name, NameGtr.NT_CAT) in self._cache_partTags['_subver_hdrs']

    def _split(self, name):
        '''
        split the given name in 3 parts: (cat, sn, ver,). When some parts does not exist, it's part will be None. If length of the name does not any specification, return None instead of tuple
        Args:
            name(string):   the name to split
        '''
        nt = self.nameType(name)
        lst = None
        if nt == self.NT_INVALID:
            return lst
        if nt == self.NT_CAT:
            lst = name[:self.CatDigits], None, None
        else:
            idx = self.CatDigits + self.SnDigits
            if nt == self.NT_CAT_SN:
                lst = name[:self.CatDigits], name[self.CatDigits:idx], None
            else:
                lst = name[:self.CatDigits], name[self.
                                                  CatDigits:idx], name[idx:]
        return lst

    def partOf(self, name, parts='full'):
        ''' return parts of the name
        Args:
            parts:  one of NameGtr.NT_CAT/NameGtr.NT_CAT_SN/other
        '''
        pts = self._split(name)
        if parts == NameGtr.NT_CAT:
            pts = (pts[0], )
        elif parts == NameGtr.NT_CAT_SN:
            pts = pts[:2]
        if all(pts):
            return ''.join(pts)
        return None

    def _validate(self, name):
        if not name or not self._gtr_sn.isvalid(name[self._cat_digits:]):
            raise AttributeError('%s Blank or contains character(s) not in %s' %
                                 (name, self._gtr_sn.charset))

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

            version=True:    True for getting new version, else for new sub-version(ring size) of current version

        throws:

            OverflowError if current is already at the end
        The return value based on arguments as below:
            current	version	Returns
            Cat	    T/F	    Full(init Version of this Cat)
            Name	T	    Full(This init version)
            Name	F	    Full(Next init version)
            Full	T	    Full(This next version)
            Full	F	    Exception(Not sub-versionable)
            Full		    Full(This next sub-version)
        '''
        self._validate(current)
        cat, sn, ver = self._split(current)
        if sn is None:
            sn = self._gtr_sn.next(self._gtr_sn.next(None))
            ver = None
        elif ver is None and not version:
            sn = self._gtr_sn.next(sn)
        if self.subverable(current):
            ver = self._sgmgr.next(self._c2v(ver), version)
            ver = self._c2v(ver)
        else:
            ver = self._gtr_ver.next(ver)
        return cat + sn + ver

    def header(self, name):
        ''' return header of given name. For non-ring, return itself
        '''
        self._validate(name)
        if not self.subverable(name):
            return name
        cat, sn, ver = self._split(name)
        ver = self._sgmgr.range(self._c2v(ver))
        return cat + sn + self._c2v(ver[0])

    def versions(self, name):
        ''' return the versions available for given name. If name is not of

        ring, return None

        '''
        self._validate(name)
        if not self.subverable(name):
            return None
        return [name + self._c2v(x) for x in self._sgmgr.segments]

    def subvers(self, name):
        ''' return the sub-versions(ring sizes) available for given name. If name is not of

        ring, return None

        Args:
            name: the style name with version. If no version is provided, use first version

        '''
        self._validate(name)
        if not self.subverable(name):
            return None
        cat, digits, ver = self._split(name)
        return [
            cat + digits + self._c2v(x)
            for x in self._sgmgr.sectors(self._c2v(ver))
        ]

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
            nm = next(iter(self._cache_partTags['_subver_hdrs'])) + gtr.next(
                gtr.charOf(0))
        if not self.subverable(nm):
            # will return too many items
            return None
        nm0 = nm
        lst = self._sgmgr.all()
        if file:
            nc = len(lst) * len(lst[0])
            tc = self._gtr_ver.radix**self._gtr_ver.digits
            print(
                'Names for %s, NameCnt:%d, VerCnt=%d, SzPerVer=%d, UseRate=%3.2f%%'
                % (nm0, nc, len(lst), len(lst[0]), nc / tc * 100))
            i2n = lambda addr: ''.join(self._gtr_one.charOf(x) for x in addr)
            for lst1 in lst:
                print(
                    '*' + i2n(lst1[0]) + ': ' + ','.join(
                        [i2n(x) for x in lst1[1:]]),
                    file=file)
        return lst

class FormSvc(SvcBase):
    '''
    class to read/write excel form from/to db
    Args:
        xstyno_as_parent(bool): use the old sty# as parent searching condition when
        new name is needed and the parent field is blank
    '''
    def __init__(self, trmgr, **kwds):
        super().__init__(trmgr)
        self._nameSvc = NameSvc(trmgr, styDao=self)
        self._type2cat_mp = None
        self._partTags = {}
        self._stypSvc = StypSvc(trmgr)
        self._xstyno_as_parent = kwds.get('xstyno_as_parent', True)
        mp = config.get('prdspec.parts.tags')
        self._partTags = {x[0]: x[1] for x in mp.items()}


    def _type2cat(self, tn):
        '''
        Pendant to P and so on
        '''
        if not self._type2cat_mp:
            self._type2cat_mp = {x[0]: x[1] for x in config.get('prdspec.type2cat').items()}
        return self._type2cat_mp.get(trimu(tn))


    def persist(self, xlFn):
        '''
        save data of the given xlFn to db
        throws:
            Exceptions if error or warning found
        returns:
            A set or data by FormHandler
        '''
        app, tk = xwu.appmgr.acq()
        sty = None
        try:
            wb = xwu.safeopen(app, xlFn, readonly=False)
            fh = JOFormHandler(wb)
            mp, logs = fh.read()
            if logs:
                pass
            with self.sessionctx() as cur:
                rmp = self._map2ent(mp)
                sty = rmp['style']
                styno = mp.get('styno')
                if not styno or styno[0] == '_':
                    styno = self._new_styno(mp, rmp)
                sty.name = styno
                matIdx = 0
                for key, lst in rmp.items():
                    if not lst or key == 'style':
                        continue
                    # maybe delete then insert is a good idea because there is many sub-records
                    else:
                        for inst in lst:
                            if isinstance(inst, Stymat):
                                matIdx += 1
                                inst.idx = matIdx
                            inst.style = sty
                            cur.add(inst)
                cur.commit()
                sty = rmp['style']
            if styno:
                fh.write({'styno': styno})
                fn = path.splitext(path.basename(xlFn))
                fn = fn[0] + "_" + styno + fn[1]
                fn = path.join(path.dirname(xlFn), fn)
                wb.api.SaveAs(fn)
                wb.close()
                wb = None
                remove(xlFn)
                _logger.debug('file (%s) was saved as (%s)' % (xlFn, fn))
        finally:
            if wb:
                wb.close()
            if tk:
                xwu.appmgr.ret(tk)
        return sty


    def _new_styno(self, mp, rmp):
        parent = mp.get('parent')
        # when parent# is invalid, treate it as an old style#, get one
        if self._xstyno_as_parent and self._isempty(parent):
            parent = [x for x in rmp['property'] if x.prop.pdef.name == 'XSTYN']
            if parent:
                parent = self._stypSvc.fromStypi(parent[0].prop)
                # now parent[-1] holds field
                q = Query(Style).join(Styp).join(Stypi).join(Stypidef).filter(and_(Stypidef.name == 'XSTYN', parent[-1] == parent[1]))
                with self.sessionctx() as cur:
                    parent = q.with_session(cur).limit(1).all()
                    if parent:
                        parent = parent[0].name
        return self._nameSvc.get(parent or self._type2cat(mp['type']))


    def render(self, styno, hideFlds=None, tplFn=None):
        '''
        dump data of the given style to an excel file
        Args:
            hideFlds(Set(String)): A set of fields that you don't want to show in the excel
            tplFn(String): full path to the template against the default template
        '''
        wb = _FormRender(self.sessmgr()).render(styno, hideFlds, tplFn)
        return wb

    def _map2ent(self, mp):
        ''' map to database entity
        '''
        rmp = {}
        mats = {}

        def _getmat(tn, nl):
            if tn == 'metal':
                kt = xwu.esctext(nl.karat)
            elif tn == 'finishing':
                kt = '%s@%s' % (nl.method, nl.spec)
            elif tn == 'parts':
                kt = xwu.esctext(nl.matid)
            elif tn == 'stone':
                kt = nl.matid
            mat = mats.get(kt)
            if not mat:
                with self.sessionctx() as cur:
                    mat = cur.query(Mat).filter(Mat.name == kt).one_or_none()
                if not mat:
                    if tn == 'metal':
                        tmp = {'type': 'METAL', 'name': kt, 'code': kt, 'wgtunit': 'GM'}
                    elif tn == 'finishing':
                        tmp = {'type': 'FINISHING', 'name': kt, 'code': kt, 'wgtunit': 'MM2', 'method': nl.method, 'spec': nl.spec} # mm^2
                    elif tn == 'parts':
                        tmp = {'type': 'PARTS', 'name': kt, 'code': kt, 'description': nl.remarks, 'wgtunit': 'GM'}
                        val = nl.karat
                        if val:
                            tmp['karat'] = xwu.esctext(val)
                    elif tn == 'stone':
                        tmp = {'type': 'STONE', 'name': kt, 'code': kt, 'description': nl.name, 'shape': nl.shape, 'size': nl.size, 'wgtunit': nl.wgtunit}
                    mat = self._map2mat(tmp)
                mats[kt] = mat
            return mat

        rmp['style'] = sty = _Util.default_values(Style())
        for fn in 'docno createdate lastmodified description netwgt increment dim size qclevel'.split():
            setattr(sty, fn, xwu.esctext(mp.get(fn)))
        tn = 'metal'
        rmp[tn] = lst = []
        for nl in mp[tn]:
            ent = _Util.default_values(Stymat())
            ent.mat = _getmat(tn, nl)
            ent.qty = 1
            ent.wgt = nl.wgt or 0
            ent.remarks = nl.remarks
            lst.append(ent)
        tn = 'finishing'
        rmp[tn] = lst = []
        for nl in mp[tn]:
            ent = _Util.default_values(Stymat())
            ent.mat = _getmat(tn, nl)
            ent.qty = 1
            ent.wgt = 0.001 # should this be area? in VX case
            ent.remarks = nl.remarks or NA
            lst.append(ent)
        tn = 'parts'
        rmp[tn] = lst = []
        for nl in mp[tn]:
            ent = _Util.default_values(Stymat())
            ent.mat = _getmat(tn, nl)
            ent.qty = nl.qty
            ent.wgt = nl.wgt or 0
            ent.remarks = ";".join((nl.type, nl.remarks or NA, ))
            if nl.type in self._partTags:
                ent.tag = self._partTags[nl.type]
            lst.append(ent)
        tn = 'stone'
        rmp[tn] = lst = []
        for nl in mp[tn]:
            ent = _Util.default_values(Stymat())
            ent.mat = _getmat(tn, nl)
            ent.qty = nl.qty
            ent.wgt = nl.wgt or 0
            ent.remarks = nl.remarks
            lst.append(ent)
            stone = _Util.default_values(Stystset())
            stone.stymat = ent
            stone.setting = nl.setting
            lst.append(stone)
        tmp = {'CATEGORY': (mp['type'], ), 'CRAFT': (mp['craft'], )}
        if mp['hallmark']:
            tmp['HALLMARK'] = (mp['hallmark'], )
        for nl in chain(*[x for x in (mp.get(x) for x in 'feature feature1'.split()) if x]):
            fn = triml(nl['catid'])
            if fn == 'remarks':
                sty.remarks = nl['value']
            else:
                if fn == 'keywords':
                    # maybe need to split the keywords into records, for better searching
                    pass
                else:
                    pass
                tmp.setdefault(nl['catid'], []).append(nl['value']) # in fact, only string type can have multiple same-name items
        rmp['property'] = lst = []
        stypSvc = StypSvc(self.sessmgr())
        for key, vals in tmp.items():
            # TODO:: lst should have type convertion
            styp = _Util.default_values(Styp())
            styp.prop = stypSvc.toStypi(key, ';'.join(vals))
            lst.append(styp)
        return rmp


    def _map2mat(self, mp):
        ''' convert a map into Mat entity, create or from db
        '''
        mat = None
        _uSvc = _Util()
        with self.sessionctx() as cur:
            mat = cur.query(Mat).filter(Mat.name == mp['name']).one_or_none()
            if not mat:
                mat = _Util.default_values(Mat())
                mat.name = mat.code = mp['name']
                mat.type = mp['type']
                mat.description = mp.get('description', mp['name'])
                mat.spec = _uSvc.enc_mat_spec(mp)
                mat.unit = mp.get('wgtunit', NA)
                cur.add(mat)
                cur.flush()
        return mat


    def _isempty(self, val):
        return not val or val in ('_NONE_', '_N/A_')

    def _ent2mp(self, ents):
        ''' db entities to mp
        '''
        return None

    def forms(self, stynos):
        '''
        return a map of styno -> file, where the file is the excel form of given style
        '''
        return None

    def getmax(self, cat):
        ''' return the maximum item with cat as prefix
        '''
        if cat[-1] != '%':
            cat += '%'
        with self.sessionctx() as cur:
            lst = cur.query(func.max(Style.name)).filter(Style.name.like(cat)).one_or_none()
        return lst[0]

    def exists(self, name):
        ''' check if given name exists
        '''
        with self.sessionctx() as cur:
            inst = cur.query(Style).filter(Style.name == name).fetchone()
        return bool(inst)

class StyleSvc(SvcBase):
    '''
    services for Style
    '''
    def __init__(self, trmgr):
        super().__init__(trmgr)
        self._partTags = {}
        mp = config.get('prdspec.parts.tags')
        self._partTags = {x[0]: x[1] for x in mp.items()}
        self._ucSvc = UnitCvtSvc()

    def calcIncr(self, styno, lossrates='G=1.06;S=1.07', oz2gm=None):
        '''
        calculate the increment of the given sty#
        Args:
            styno
            lossrates(String): formatted as G=XX[;S=XX]
        Returns:
            A map with increments by metal type
        '''
        with self.sessionctx() as cur:
            q = Query((Stymat.qty, Stymat.wgt, Mat.name, Mat.type, Mat.spec, )).join(Style).join(Mat).filter(and_(Style.name == styno, or_(Mat.type == 'METAL', and_(Mat.type == 'PARTS', Stymat.tag == self._partTags['XP']))))
            pps = q.with_session(cur).all()
            if not pps:
                return None
            lrmp = [(x[0], float(x[1])) for x in (x.split('=') for x in lossrates.split(';'))]
            lrmp = dict(lrmp)
            uSvc = _Util()
            incr = {}
            for pp in pps:
                tpl = triml(pp.type)
                kt = None
                if tpl == 'metal':
                    kt = karatsvc.getkarat(pp.name)
                else:
                    mp = uSvc.dec_mat_spec(tpl, pp.spec)
                    kt = mp.get('karat', None)
                if kt:
                    rx = (1 if tpl != 'metal' else lrmp.get(kt.category[0], 1.06))
                    wgt = float(pp.wgt) * kt.fineness
                    wgt = wgt / oz2gm if oz2gm else self._ucSvc.convert(wgt, 'gm', 'oz')
                    rx *= wgt
                    incr[kt.category] = incr.get(kt.category, 0) + round(rx, 6)
        return incr

    def calcWgt(self, styno, metalOnly=False):
        '''
        calculate the netwgt of the given style
        '''
        with self.sessionctx() as cur:
            q = Query((Stymat.qty, Stymat.wgt, Mat.name, Mat.type, Mat.spec, Mat.unit, )).join(Style).join(Mat).filter(and_(Style.name == styno, or_(Mat.type == 'METAL', Mat.type == 'STONE', and_(Mat.type == 'PARTS', Stymat.tag != self._partTags['MP']))))
            pps = q.with_session(cur).all()
            if not pps:
                return None
            rc = 0
            for pp in pps:
                tpl = triml(pp.type)
                if tpl == 'stone':
                    if metalOnly:
                        continue
                rc += float(pp.wgt) * self._ucSvc.convert(1, pp.unit, 'GM')
        return round(rc, 2)

class _Util(object):

    def __init__(self):
        super().__init__()
        self._mat_enc_fld = None
    

    @classmethod
    def default_values(cls, ent):
        '''
        fill the defaults of a styx item
        '''
        if not ent.lastuserid:
            ent.lastuserid = 1
        if not ent.creatorid:
            ent.creatorid = 1
        if not ent.createddate:
            ent.createddate = datetime.today()
        if not ent.modifieddate:
            ent.modifieddate = datetime.today()
        if ent.tag is None:
            ent.tag = 0
        return ent


    def _get_mat_enc_fld(self, tn):
        if not self._mat_enc_fld:
            mp = config.get('prdspec.matspec.enc')
            self._mat_enc_fld = {x[0]: x[1].split() for x in mp.items()}
        return self._mat_enc_fld.get(tn)

    def enc_mat_spec(self, mp):
        '''
        create an spec field based on the provided map
        '''
        def _2str(tp, name, val):
            tp = triml(tp)
            if tp == 'stone':
                if name == 'size':
                    val = stsizefmt(val, False)
            elif tp == 'metal':
                if name == 'fineness':
                    val = '%6.4f' % val
            return val

        tp = mp['type']
        tpl = triml(tp)
        spec = None
        # infact, a material's spec should be maintence by an individual program, this create-on-demand is for emergency only, so the encoding of spec is just draft
        # an existing mat won't pass this route
        flds = self._get_mat_enc_fld(tpl)
        # flds = [x for x in self._get_mat_enc_fld(tpl) if x in mp]
        if flds:
            if tpl == 'parts':
                spec = []
                for fld in flds:
                    if fld == 'karat':
                        if fld in mp:
                            spec.append('KARAT:%s' % mp[fld])
                    else:
                        spec.append(_2str(tpl, fld, mp.get(fld, NA)))
            elif tpl == 'metal':
                kt = karatsvc.getkarat(xwu.esctext(mp['name']))
                spec = [_2str(tp, x, getattr(kt, x)) for x in flds]
                if 'description' not in mp or not mp['description']:
                    mp['description'] = '%s %s' % (kt.fineness, kt.category)
            else:
                spec = [_2str(tp, x, mp.get(x, NA)) for x in flds]
            if spec:
                spec = (x for x in spec if x)
        return ';'.join(spec) if spec else NA


    def dec_mat_spec(self, tpl, spec):
        ''' decode a material's spec field into maps
        '''
        vals = spec.split(';')
        flds = self._get_mat_enc_fld(tpl)
        if len(vals) != len(flds):
            return None
        mp = dict(zip(flds, vals))
        if 'karat' in mp:
            mp['karat'] = karatsvc.getkarat(mp['karat'].split(':')[-1])
        return mp

class StypSvc(SvcBase):
    '''
    class to handle the style property list <-> db transferring
    '''

    def __init__(self, trmgr):
        super().__init__(trmgr)
        self._dfts = None
        self._get_deft_def('')  # load the default maps, or create them on-demand


    def toStypi(self, key, val, create_on_demand=False):
        '''
        transfer given kye/value to db
        '''
        updCnt = 0
        with self.sessionctx() as cur:
            pdef = cur.query(Stypidef).filter(Stypidef.name == key).one_or_none()
            if not pdef:
                if create_on_demand:
                    pdef = self._new_def(key, val)
                    cur.add(pdef)
                    cur.flush() # to get the Id
                    updCnt += 1
                    _logger.debug('definition for property(%s) was created on-demand' % key)
                else:
                    pdef = self._get_deft_def(val)
                    _logger.debug('using default definition for property(%s)' % key)
            tn, val, tors = self._enc_stypi(pdef, val)
            deft_def = self._is_deft_def(pdef)
            q = Query(Stypi).filter(Stypi.pdef == pdef)
            if tn == 'C':
                # using default, the key should be encoded into valuec, by ';' as separator
                if deft_def:
                    val = '%s;%s' % (key, val)
                q = q.filter(Stypi.valuec == val)
            else:
                # float point query, don't use ==, use range instead
                if deft_def:
                    # using default, the key was kept inside valuec
                    q = q.filter(and_(Stypi.valuec == key, Stypi.valuen > tors[0], Stypi.valuen < tors[1]))
                else:
                    q = q.filter(and_(Stypi.valuen > tors[0], Stypi.valuen < tors[1]))
            rc = q.with_session(cur).one_or_none()
            if not rc:
                rc = _Util.default_values(Stypi())
                rc.pdef = pdef
                rc.valuec = ''
                rc.valuen = 0
                if tn == 'C':
                    rc.valuec = val
                else:
                    rc.valuen = val
                    if deft_def:
                        rc.valuec = key
                cur.add(rc)
                updCnt += 1
            if updCnt:
                cur.flush()
            return rc


    def fromStypi(self, stypi):
        '''
        transfer a Styp instance to key/value pair
        Args:
            stypi(Stypi): an Stypi instance
        returns:
            the (key, value, field) tuple
        '''
        return self._dec_stypi(stypi)


    def _enc_stypi(self, pdef, val):
        tn = triml(pdef.type)
        rc, st = val, 'C'
        tors = None
        if tn.find('numeric') == 0:
            prec = [int(x.strip()) for x in tn[len('numeric'):].split(',')]
            if prec[-1] > 0:
                rc = round(val, prec[-1])
                tors = pow(10, -prec[-1])
            else:
                rc = int(val)
                tors = 0.1
            tors = [rc - tors, rc + tors]
            st = 'N'
        elif tn.find('date') == 0:
            # saving in valuen make comparision better, but no so user-friendly
            # in VBA, I created a module timestamp, inside which there is tsFromDate/tsToDate to make this conversion
            rc = val.timestamp()
            st = 'N'
            tors = 0.1
            tors = [rc - tors, rc + tors]
        return st, rc, tors


    def _is_deft_def(self, pdef):
        return pdef.tag == 100


    def _dec_stypi(self, stypi):
        pdef = stypi.pdef
        tn = triml(pdef.type)
        fld = Stypi.valuen
        if tn.find('numeric') == 0:
            rc = stypi.valuen
        elif tn.find('date') == 0:
            rc = datetime.fromtimestamp(stypi.valuen)
        else:
            rc = stypi.valuec
            fld = Stypi.valuec
        if self._is_deft_def(pdef):
            if tn == 'string':
                rc = rc.split(';')
                key = rc[0]
                rc = ';'.join(rc[1:])
            else:
                key = stypi.valuec
        else:
            key = pdef.name
        return key, rc, fld


    def _create_defts(self):
        '''
        create the default stypidef items
        '''
        updCnt = 0
        mp = config.get('prdspec.stypidef.defs')
        with self.sessionctx() as cur:
            exts = cur.query(Stypidef.name.in_(tuple(mp.keys()))).all()
            if exts:
                exts = {x.name: x for x in exts.items()}
            else:
                exts = []
            for key, lst in mp.items():
                if key in exts:
                    continue
                pdef = _Util.default_values(Stypidef())
                pdef.name = key
                pdef.type, pdef.format, pdef.remarks = lst[:3]
                if len(lst) > 3:
                    pdef.tag = lst[3]
                cur.add(pdef)
                updCnt += 1
        return updCnt


    def _val_type(self, val):
        '''
        val type to stypi type
        '''
        if isinstance(val, Number):
            rt = 'NUMERIC(18, 4)'
        elif isinstance(val, datetime):
            rt = 'DATETIME'
        else:
            rt = 'STRING'
        return rt


    def _get_deft_def(self, val):
        if self._dfts is None:
            with self.sessionctx() as cur:
                q = Query(Stypidef).filter(Stypidef.tag == 100)
                lst = q.with_session(cur).all()
                if not lst:
                    if self._create_defts():
                        cur.commit()
                    lst = q.with_session(cur).all()
                self._dfts = {}
                for pdef in lst:
                    self._dfts[pdef.type] = pdef
        return self._dfts.get(self._val_type(val))


    def _new_def(self, key, val):
        pdef = _Util.default_values(Stypidef())
        pdef.format = '%s'
        pdef.type = self._val_type(val)
        pdef.name = key
        return pdef

class _FormRender(SvcBase):

    def __init__(self, sessmgr):
        super().__init__(sessmgr)
        self._uSvc = _Util()

    def render(self, styno, hideFlds=None, tplFn=None):
        '''
        render the given styno to an excel file
        Args:
            tplFn(String):  the template file to spawn from
            hideFlds(Set(String)): the fields that won't be shown
        '''
        with self.sessionctx() as cur:
            sty = cur.query(Style).filter(Style.name == styno).one_or_none()
            if not sty:
                return None
            alias = {'name': 'styno', 'modifieddate': 'lastmodified', 'createddate': 'createdate'}
            _nrl = triml
            def _fmt(fn, val):
                if fn.find('date') >= 0:
                    val = val.strftime('%Y/%m/%d')
                return val
            if not hideFlds:
                hideFlds = set()
            else:
                hideFlds = {_nrl(x) for x in hideFlds}
            mp0 = {alias.get(x, x): _fmt(x, getattr(sty, x)) for x in 'name createddate modifieddate qclevel size dim docno description'.split()}
            mp0['author'] = '_RENDER_' # TODO::
            # Some properties encoded into styp, get them out if there is
            spsvc = StypSvc(self.sessmgr())
            pps = cur.query(Styp).join(Style).filter(Style.name == styno).all()
            pps = dict(spsvc.fromStypi(x.prop)[:-1] for x in pps)
            p2m = trimu('category craft hallmark').split() # maybe get those stypidef.tag = 200
            tmp = {alias.get(x, x): pps[x] for x in p2m if x in pps}
            mp0.update(tmp)
            if 'feature' not in hideFlds and len(pps) > len(tmp):
                lst = [["catid", "value"]]
                lst.extend([x for x in pps.items() if x[0] not in tmp])
                mp0['feature'] = NamedLists(lst)
            pps = {}
            lst = Query((Stymat.id, Mat.type, Stymat.idx, Stymat.qty, Stymat.wgt, Mat.name, Mat.spec, Mat.unit, Stymat.remarks, )).join(Style).join(Mat).filter(Style.name == styno).order_by(and_(Mat.type, Stymat.idx))
            lst = lst.with_session(cur).all()
            for tmp in lst:
                tpl = triml(tmp.type)
                pps.setdefault(tpl, []).append(tmp)
            for tpl, lst in pps.items():
                if tpl in hideFlds:
                    continue
                lst = self._make_nls(tpl, lst, cur)
                if lst:
                    mp0[tpl] = lst
            wb = xwu.fromtemplate(tplFn or config.get('prdspec.template'))
            if hideFlds:
                for fn in hideFlds:
                    if fn in mp0:
                        del mp0[fn]
            tmp = JOFormHandler(wb).write(mp0)
        return wb


    def _make_stone_nls(self, lst, cur):
        mpx = Query((Stystset.id, Stystset.setting, Stystset.ismain, )).filter(Stystset.id.in_([x.id for x in lst])).with_session(cur).all()
        mpx = {x.id: x for x in mpx}
        lst1 = ['setting shape name main size qty unitwgt wgtunit wgt matid remarks'.split(), ]
        for pp in lst:
            mp = {}
            sett = mpx.get(pp.id)
            if sett:
                mp['setting'] = sett.setting
                mp['main'] = 'Y' if sett.ismain > 0 else 'N'
            mp['qty'] = pp.qty
            mp['unitwgt'] = float(pp.wgt / pp.qty)
            mp['wgt'] = float(pp.wgt)
            mp['wgtunit'] = pp.unit
            mp['matid'] = pp.name
            mp['remarks'] = pp.remarks
            mpxx = self._uSvc.dec_mat_spec('stone', pp.spec)
            mp.update(mpxx)
            mp['name'] = mp['description']
            lst1.append([mp.get(x) for x in lst1[0]])
        return lst1


    def _make_nls(self, tpl, lst, cur):
        lst = sorted(lst, key=lambda x: x.idx) # sort by idx
        if tpl == 'stone':
            lst = self._make_stone_nls(lst, cur)
        elif tpl == 'metal':
            lst = [(x.name, float(x.wgt), x.remarks) for x in lst]
            lst.insert(0, ('karat', 'wgt', 'remarks'))
        elif tpl == 'finishing':
            _u = _Util()
            lst = [(tuple(_u.dec_mat_spec(tpl, x.spec).values()), (x.remarks, )) for x in lst]
            lst = [list(chain(*x)) for x in lst]
            lst1 = list(_u._get_mat_enc_fld(tpl))
            lst1.append('remarks')
            lst.insert(0, lst1)
        elif tpl == 'parts':
            lst1 = ['type qty karat wgt matid remarks'.split(), ]
            for pp in lst:
                mp = {}
                rmks = pp.remarks.split(';')
                mp['type'] = rmks[0]
                mp['remarks'] = rmks[-1]
                mp1 = self._uSvc.dec_mat_spec(tpl, pp.spec)
                mp['karat'] = None if not mp1 or 'karat' not in mp1 else mp1['karat'].name
                mp['qty'] = pp.qty
                mp['wgt'] = float(pp.wgt)
                mp['matid'] = pp.name
                lst1.append([mp[x] for x in lst1[0]])
            lst = lst1
        else:
            lst = None
        return NamedLists(lst) if lst else None
