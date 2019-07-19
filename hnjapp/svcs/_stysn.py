'''
#! coding=utf-8
@Author:        zmFeng
@Created at:    2019-07-02
@Last Modified: 2019-07-02 10:23:11 am
@Modified by:   zmFeng
Sty -> sn mapping, for familiar style searching and locket size finding
'''

import dbm
import tempfile
from difflib import SequenceMatcher
from logging import DEBUG
from os import environ, listdir, path, remove
from re import compile as cp
from sqlite3 import connect

from sqlalchemy import create_engine, func, literal
from sqlalchemy.ext.declarative.api import declarative_base
from sqlalchemy.orm import Query, aliased, relationship
from sqlalchemy.schema import Column, ForeignKey, Index
from sqlalchemy.sql import and_, exists, not_, or_
from sqlalchemy.sql.sqltypes import VARCHAR, Integer

from hnjapp.localstore import Stysn
from hnjcore import JOElement
from utilz import NamedList, splitarray, xwu
from utilz.miscs import getpath, newline, triml
from utilz.resourcemgr import ResourceCtx, SessionMgr

from ..common import _logger as logger
from ..common import config
from .db import formatsn


class StyleFinder(object):
    '''
    given a style, find the styles that shares the SN#s
    the public SNs(like "HB411") will be automatically avoided

    Args:

        `store_sm`: The session manager to the storage. A default session manager can be obtained by:

            from python.routs.common import getSessionMgr
            sm = getSessionMgr('storage')

       `max_stycnt`=0: the max count of style of one SN#, when greater than this, the SN# will be treated as public component like BALE. It's SN# result will be discarded.

    '''

    _tags = None

    @classmethod
    def tag(cls, tn):
        ''' return the tag value of given name
        Args:

        tn: name of the tag. the support values are:
            snno: snno tag
            keyword: keyword tag
            parent: parent tag
            text: text tag
        
        more tags are available in conf.json/stysn.tag
        '''
        if not cls._tags:
            cls._tags = config.get('stysn.tag')
        return cls._tags.get(triml(tn))

    def __init__(self, store_sm, **kwargs):
        self._sm = store_sm
        self._max_stycnt = kwargs.get('max_stycnt', 0)
        self._sn_pts = {}
        mp = {self._nrm_sn(y) for x in config.get('snno.translation').values() for y in x.split(',')}
        # this is some parts that is referred by Sty#
        mp.update(config.get('locketszfinder.sn.by_styno'))
        self._sn_pts[0] = mp
        # some manul sns that should be sth. like pts
        self._sn_pts[1] = config.get('locketszfinder.sn.component')
        self._sn_pts[2] = config.get('locketszfinder.sn.charm')

    def _block(self, sn, level):
        ''' should accept given sn under the level
        '''
        if level == 0:
            return False

        for lvl in range(min(level, 3)):
            flag = sn in self._sn_pts[lvl]
            if flag:
                return True
        return False

    @property
    def _sessionctx(self):
        return ResourceCtx(self._sm)

    @staticmethod
    def _nrm_sn(sn):
        je = JOElement(sn)
        je.suffix = ""
        return je.name

    def find_by_book(self, wb):
        ''' find by providing a workbook
        '''
        srcttl = 'stynos excludes snnos limits'.split()
        nls = xwu.NamedRanges(xwu.usedrange(wb.sheets[0]))
        if not nls:
            self._make_wb_tpl(wb, srcttl)
            return None
        vvs, valid, snflag = [('source', 'result', 'snno')], None, False
        for nl in nls:
            if not valid:
                valid = set(nl.colnames)
                valid = [1 for x in srcttl if x in valid]
                if sum(valid) < len(srcttl):
                    self._make_wb_tpl(wb, srcttl)
                    return None
            if len(vvs) > 1:
                vvs.append((None, ) * len(vvs[0]))
            excludes = nl.excludes
            if excludes:
                excludes = [cp(JOElement(x).value) for x in excludes.split(',')]
            if nl.snnos:
                stynos = nl.snnos
                lst = self._get_by_sns([JOElement(x).value for x in stynos.split(',')], excludes)
                snflag = True
            else:
                stynos = nl.stynos
                lst = self.find_intersect(stynos.split(','), excludes)
                snflag = False
            if lst and lst[0]:
                stymp = {}
                for sn, tlst in lst[1].items():
                    for styno in tlst:
                        stymp.setdefault(styno, []).append(sn)
                tlst = [] if snflag else [(stynos, x, '_self_') for x in stynos.split(',')]
                tlst.extend([(stynos, x, ','.join(stymp[x])) for x in lst[0]])
                vvs.extend(tlst)
            else:
                vvs.append((stynos, None, None))
        nl = wb.sheets.add(after=wb.sheets[0])
        nl.cells(1, 1).value = vvs
        nl.autofit('c')
        xwu.freeze(nl.cells(2, 2))
        return nl

    @staticmethod
    def _make_wb_tpl(wb, ttl):
        sht = wb.sheets.add(before=wb.sheets[0])
        sht.cells(1, 1).value = (ttl, )
        sht.autofit('c')
        try:
            sht.name = 'template'
            sht.cells(2, 1).select()
        except:
            pass
        xwu.freeze(sht.cells(2, 2))

    def find(self, styno, excludes=None, level=1, spread=False):
        ''' when find the familiar style
        Args:

            styno: the styno to find.

            excludes: a :class: `tuple` of regexp, link for SN# matches the reg will be blocked

            spread: when True, SN# will be always appended by the found Styno, this might lead to large result set. default is False.

            level: the link block level.
                0 accept link from Any SN#
                1 block link from Bale
                2 block link from Bale, Component
                3 block link from block link from Bale, Component and Charm
                3+ as 3

        Returns:
            A tuple(list(styno), {sn, list(styno)}) with sty# as key and SN -> STY# as value, which helps to solve some Sty->sn conflicts
        '''
        # the quene and pointer
        styq, snq = [[JOElement(styno).name, 0], ], []
        styp = snp = 0
        styset, snset, hints, pncs = set(), set(), {}, set()
        def _add_sty(styns, snno):
            styns = [x for x in styns if x not in styset]
            if not styns:
                return
            hints[snno] = styns
            for x in styns:
                styq.append([x, styq[styp][1] + 1]) #increase the relation level
                styset.add(x)
        while styp < len(styq):
            s0 = styq[styp]
            styset.add(s0[0])
            if spread or s0[0] == styq[0][0] or s0[0] in pncs:
                sns = [x for x in self._get_sns(s0[0], excludes, level) if x not in snset]
                if sns:
                    for x in sns:
                        snset.add(x)
                        snq.append(x)
                if s0[1] < 5: # only search for 5 level of relation
                    styns = self._get_pnc(s0[0], excludes)
                    pncs.update(styns)
                    _add_sty(styns, s0[0])
            while snp < len(snq):
                styns = self._get_styns(snq[snp], excludes, level)
                _add_sty(styns, snq[snp])
                snp += 1
            styp += 1
        sns = styq[1:]
        if sns:
            sns = [x[0] for x in sns]
        if len(sns) > 1:
            sns = sorted(sns)
        return sns, hints

    def find_intersect(self, stynos, excludes=None, level=1):
        ''' find the styles that contains the public SN#s of stynos provided

        Args:
            stynos: a collection of stynos

        return:
            (tuple(styno), set(snno),). the original stynos won't be returned
        '''
        if len(stynos) == 1:
            sns = self._by_feature(iter(stynos).__next__(), excludes, level=level)
            return (tuple(y for x in sns for y in x[1]), dict(x for x in sns)) if sns else None
        isns = None
        for styno in stynos:
            sns = self._get_sns(styno, excludes, level)
            if isns:
                isns = set(sns).intersection(isns)
            else:
                isns = sns
            if not isns:
                return None
        if not isns:
            return None
        return self._get_by_sns(isns, excludes, level, stynos)

    def _simple_get(self, fld, q, tag):
        with self._sessionctx as cur:
            q = cur.query(fld).filter(q).filter(Stysn.tag == tag)
            lst = q.all()
        return sorted((x[0] for x in lst)) if lst else None

    def getstynos(self, snno):
        ''' return the styno by given snno
        '''
        snno = JOElement(snno).value
        return self._simple_get(Stysn.styno, Stysn.snno == snno, self.tag('snno'))

    def getstyx(self, styno, tn):
        ''' return given styno's x property

        Args:

            tn: name of the tag, refer to :func: tag() FMI

        Returns:
            tuple(String)
        '''
        tn = triml(tn)
        lst = self._simple_get(Stysn.snno, Stysn.styno == JOElement(styno).value, self.tag(tn))
        if tn == 'snno' and lst and len(lst) > 1:
            lst = formatsn(",".join(lst), retuple=True) if lst else None
        return lst

    def _get_by_sns(self, isns, excludes, level=1, stynos=None):
        stymp = snmp = {}
        for sn in isns:
            if excludes:
                for reg in excludes:
                    if reg.match(sn):
                        continue
            lst = self._get_styns(sn, excludes, level)
            if not lst:
                continue
            if not stymp:
                stymp = {x: 1 for x in lst}
                snmp = {sn: lst}
            else:
                if sn in snmp:
                    snmp[sn].extend(lst)
                else:
                    snmp[sn] = lst
                for styno in lst:
                    if styno in stymp:
                        stymp[styno] = stymp[styno] + 1
                    else:
                        stymp[styno] = 1
        if stymp:
            min_cnt = len(isns)
            if not stynos:
                stynos = {}
            lst = tuple(x for x in stymp if stymp[x] >= min_cnt and x not in stynos)
            isns = {x: set(snmp[x]) for x in isns}
        else:
            lst = None
        return (lst, isns) if lst else None

    def shared_sns(self, min_sncnt=2):
        ''' find out the stynos that share at least min_sncnt of sn with other sty#.
        The SQL for this query is:

        select a.styno, b.styno, count(0)
        from stysn a join stysn b on a.snno = b.snno and a.styno <> b.styno and exists (
        select 1 from stysn c where c.styno = b.styno and c.snno <> a.snno
        )
        group by a.styno, b.styno
        having count(0) > 1
        '''
        return None
        # don't know how to write with the exists() function. seems sqlalchemy.sql.expression is needed
        a, b, c = aliased(Stysn), aliased(Stysn), aliased(Stysn)
        with self._sessionctx as cur:
            q = Query((a.styno, b.styno, func.count(0),)).filter(and_(a.snno == b.snno, a.styno != b.styno)).filter(not_(exists(Query(literal(1)).filter(and_(c.styno == b.styno, c.snno != a.snno))))).group_by(a.styno, b.styno).having(func.count(0) > min_sncnt)
            lst = q.with_session(cur).all()
            if lst:
                return [(x[0], x[1]) for x in lst]
        return None

    def find_by_feature(self, styno, excludes=None, threshold=20, level=1):
        ''' when find the familiar style
        Args:
            styno: the sty# to find.
            excludes: a tuple of regexp, if SN of the style match, will be excluded
        Returns:
            A tuple(list(styno), {sn, list(styno)}) with sty# as key and SN -> STY# as value, which helps to solve some Sty->sn conflicts
        '''
        hints = self._by_feature(styno, excludes, threshold, level)
        if not hints:
            return None
        return tuple({y for x in hints for y in x[1]})

    def _by_feature(self, styno, excludes=None, threshold=20, level=1):
        ''' when find the familiar style
        Args:
            styno: the sty# to find.
            excludes: a tuple of regexp, if SN of the style match, will be excluded
        Returns:
            tuple((snno, tuple(styno))
        '''
        hints = self.find(styno, excludes, level=level, spread=False)[1]
        if not hints:
            return None
        hints = [x for x in hints.items()]
        if len(hints) > 1:
            hints = sorted(hints, key=lambda x: len(x[1]))
            lst = [x for x in hints if len(x[1]) < threshold]
            # at least keep one if all of them are greater than the thread-hold
            if not lst:
                lst = (hints[0], )
            hints = lst
        return hints

    def _get_sns(self, styno, excludes=None, level=1):
        ''' from db or a dict
        '''
        with self._sessionctx as cur:
            lst = Query(Stysn.snno).filter(and_(Stysn.styno == styno, Stysn.tag == 'S')).with_session(cur).all() or ()
            if lst:
                lst = [x[0] for x in lst if not self._block(x[0], level)]
                if excludes:
                    lst = [x for x in lst if not any(reg.match(x) for reg in excludes)]
        return lst

    def _get_pnc(self, styno, excludes=None):
        ''' return parent or child directly related to me
        '''
        with self._sessionctx as cur:
            lst = Query([Stysn.styno, Stysn.snno]).filter(and_(Stysn.tag == 'P', Stysn.styno != Stysn.snno, or_(Stysn.styno == styno, Stysn.snno == styno))).with_session(cur).all() or ()
            if lst:
                lst = [x[0] if x[1] == styno else x[1] for x in lst]
                lst = self._exclude(lst, excludes)
        return lst

    @staticmethod
    def _exclude(lst, excludes):
        if not excludes:
            return lst
        return [x for x in lst if not any(reg.match(x) for reg in excludes)]

    def _get_styns(self, snno, excludes=None, level=1):
        if self._block(snno, level):
            return ()
        with self._sessionctx as cur:
            lst = Query(Stysn.styno).filter(and_(Stysn.snno == snno, Stysn.tag == 'S')).with_session(cur).all() or ()
            if lst:
                lst = self._exclude([x[0] for x in lst], excludes)
                if self._max_stycnt and len(lst) > self._max_stycnt:
                    logger.debug('SN#(%s) with style count = %d is discarded' % (snno, len(lst)))
                    lst = ()
        return lst

    def delete(self, stysns):
        ''' delete given sty# -> snno pair
        Args:
            stysns: a tuple of (styno, snno)
        '''
        with self._sessionctx as cur:
            for styno, snno in stysns:
                cur.query(Stysn).filter(Stysn.styno == styno).filter(Stysn.snno == snno).delete(synchronize_session=False)
            cur.commit()

class LKSizeFinder(object):
    """ class to return locket size(if it is) based on history records

        Args:

        style_finder:   a `StyleFinder` instance that help to find familiar style(s) by styno

        stynos:         a collection of Style(string) that need to find the size

        hints:          a {Styno(string): LkSize(string)} map

        josn_idxr:      a `JOSnIndex` instance help to calc the styno -> snno ratio

    """
    def __init__(self, style_finder, stynos=None, **kwargs):
        self._style_finder = style_finder
        self._josn_idxr = kwargs.get('josn_idxr')
        self._stynos = stynos
        hints = kwargs.get('hints')
        if not hints:
            hints = self._load_lkszs()
        self._sz_hints = hints

    def _load_lkszs(self):
        fn = path.join(getpath(), 'res', 'lksz.txt')
        if not path.exists(fn):
            return None
        with open(fn, 'rt') as fh:
            lns = fh.readlines()
        if not lns:
            return None
        lf = newline(lns[0]) #CRLF -> \r\n, LF=\n
        mp = dict((x[:lf].split('\t') for x in lns))
        return mp

    def find(self, deep_find=False, stynos=None):
        ''' return the size of the stynos provided by the constructor
        Args:
            deep_find=False:    when True, try all candiates even if it's inside the hints
        Returns:
            A {Styno(string), tuple(size(string))}
        '''
        nl = NamedList('cand styno size sns')
        # ratio limited, set it to non-zero to block result by some low-ratio SN of a given style
        rltd = 0
        mp, dbg = {}, logger.isEnabledFor(DEBUG)
        dups = [] if dbg else None
        for styn_r in stynos or self._stynos:
            if not deep_find:
                sz = self._sz_hints.get(styn_r)
                if sz:
                    mp[styn_r] = (sz, )
                    continue
            lst, hints = self._style_finder.find(styn_r, level=3)
            if not lst:
                lst, hints = [], []
            gaveups = {}
            stysnmp = {y: x for x in hints for y in hints[x]}
            # slot for itself
            lst.insert(0, styn_r)
            stysnmp[styn_r] = 'HBXX'
            szstymp = {}
            for styn in lst:
                sn = stysnmp.get(styn)
                # SN# starts with P means the BALE of Pendant, so this is a style
                # just ignore it, but in casting case, this should not be ignored.
                # after testing, casting case is rare.
                if not sn or sn[0] != 'H':
                    continue
                rto = self._josn_idxr.ratio(styn_r, sn) if self._josn_idxr and sn else 1.0
                if not rltd or rto >= 0.2:
                    sz = self._sz_hints.get(styn)
                    szstymp.setdefault(sz, []).append(styn)
                else:
                    if dbg and sn not in gaveups:
                        gaveups[sn] = print('SN#(%s) in Sty#(%s) is too weak(%4.2f%%), link blocked' % (sn, styn_r, rto * 100))
            if gaveups:
                for x in gaveups.items():
                    logger.debug(x)
            uks = tuple(x for x in szstymp if x)
            if len(uks) > 1:
                self._lksz_solv_dup(styn_r, hints, nl, szstymp, dups)
            mp[styn_r] = uks
        if dbg and dups:
            self._dump_conflicts(dups, nl)
        return mp

    def _dump_conflicts(self, dups, nl):
        if not dups:
            return
        pns = []
        pns.append('---- styles with conflict sizes ----')
        pns.append(nl.colnames)
        flag = None
        for x in dups:
            nl.setdata(x)
            if nl.styno != flag:
                pns.append('')
                flag = nl.styno
            pns.append(x)
        if len(pns) < 1500:
            for x in pns:
                logger.debug(x)

    def _lksz_solv_dup(self, styn_r, hints, nl, szstymp, dups):
        stysnmp = {y: x for x in hints for y in hints[x]}
        lsts, mcnt = [], 0
        for sz in (x for x in szstymp if x):
            ref_stynos = szstymp[sz]
            sns = {}
            for x in ref_stynos:
                sn = stysnmp.get(x, 'self')
                sns.setdefault(sn, []).append(x)
            mcnt = max(mcnt, max(len(x) for x in sns.values()))
            lsts.append(nl.newdata())
            nl.size, nl.styno, nl.sns, nl.cand = sz, styn_r, sns, False
        pendings, goodsns = [], []
        for x in lsts:
            nl.setdata(x)
            sns = nl.sns
            flag = max(len(x) for x in sns.values()) == mcnt
            if flag:
                nl.cand = True
                goodsns.extend([x for x in sns if x != 'self'])
            else:
                pendings.append(x)
        mtr = SequenceMatcher()
        for x in pendings:
            nl.setdata(x)
            sns = {}
            for y in (x for x in nl.sns.items()):
                if y[0] == 'self':
                    continue
                rto = (self._josn_idxr.ratio(styn_r, y[0]) if self._josn_idxr else 1) or 0
                if rto < 0.2:
                    mtr.set_seq1(y[0])
                    sn = None
                    for flag in goodsns:
                        mtr.set_seq2(flag)
                        sn = mtr.ratio()
                        if sn > 0.9:
                            break
                    if sn and sn > 0.9:
                        if sn < 1:
                            sns['replace SN(%s) with (%s) for JOs' % (y[0], flag)] = self._josn_idxr.find(styn_r, y[0])
                        else:
                            sns['SN#(%s) SUP in Styles' % y[0]] = y[1]
                    else:
                        sns['remove SN(%s) from JOs' % y[0]] = self._josn_idxr.find(styn_r, y[0])
                else:
                    sns[y[0]] = y[1]
            nl.sns = sns
        flag = nl.getcol('cand')
        if dups:
            dups.extend(sorted(lsts, key=lambda x: x[flag], reverse=True))

_style_finderbase = declarative_base()
class JOx(_style_finderbase):
    '''  jo + style table
    '''
    __tablename__ = "jo"
    __table_args__ = (
        Index('idx_jo_sty', 'styno'),
        Index('idx_jo_jono', 'jono'),
        Index('idx_jo_styjn', 'styno', 'jono', unique=True),
    )
    id = Column(Integer, primary_key=True, autoincrement=True)
    styno = Column(VARCHAR(10))
    jono = Column(VARCHAR(10))

class JOSn(_style_finderbase):
    ''' JO sn table
    '''
    __tablename__ = "josn"
    __table_args__ = (
        Index('idx_josn_joid', 'joid'),
        Index('idx_josn_snjo', 'snno', 'joid', unique=True),
    )
    id = Column(Integer, primary_key=True, autoincrement=True)
    snno = Column(VARCHAR(20))
    joid = Column(ForeignKey("jo.id"))
    jo = relationship("JOx")

class JOSnIndex(object):
    """
    build sty+sn -> JO index from excel file that extract JO/sty/SN from HK JO system and help to:

        .find sty -> snno from all JO history(so that can build better stysn result)
        .get sty->sn ratio to improve sty->sn

    Args:

        `xls_folder`: the folder that contains excel files created by the sql. The file name pattern should be `snnos_wJOxxxx.xlsx`

        db_file: file name of the sqlite db that save data in xls_folder

        the SQL to extract data from HK JO system is:

        select convert(varchar(10), jo.alpha) + convert(varchar(10), jo.digit) jono, sty.alpha + convert(varchar(10), sty.digit) styno, jo.snno
        from jo join orderma od on jo.orderid = od.orderid and jo.tag >= 0 and jo.deadline >= '1999/01/01' and jo.deadline < '2019/01/01'
        join styma sty on od.styid = sty.styid
    """

    def __init__(self, **kwargs):
        root = environ['USERPROFILE']
        _e_or_n = lambda x: x if path.exists(x) else None
        self._xls_fldr = kwargs.get('xls_folder')
        if not self._xls_fldr:
            self._xls_fldr = _e_or_n(path.join(root, 'josn_fldr'))
        self._db_fn = kwargs.get('db_file', path.join(root, 'josn.db'))
        self._sm = None
        self._ratios = {}

    def _sqlite(self):
        return connect(self._db_fn)

    @property
    def sm(self):
        if not self._sm:
            eng = create_engine("sqlite:///?criver=xx", creator=self._sqlite)
            JOx.metadata.create_all(eng)
            self._sm = SessionMgr(eng)
        return self._sm

    def create_stysn(self, file):
        r''' create the style -> sn files based on the SN files inside xls_folder. The output file is a csv file with \t as delimiter.

            inside sqlite, execute blow commands to do the import(assume output name is stysn.txt)

            sqlite a.db
            CREATE TABLE stysn (id INTEGER NOT NULL, styno VARCHAR(10), snno VARCHAR(20), PRIMARY KEY (id));
            CREATE INDEX idx_stysn_sn ON stysn (snno);
            CREATE INDEX idx_stysn_sty ON stysn (styno);
            CREATE UNIQUE INDEX idx_stysn_stysn ON stysn (styno, snno);
            .mode tabs
            .import stysn.txt stysn

        Args:
            file: the file to save result to
        '''
        # because style# randomly exists, need a map to store the result. but when
        # the data size is too big, memory consumes too much. Use sqlite to solve this
        app, tk = xwu.appmgr.acq()
        dbfn = None
        with dbm.open(path.join(tempfile.gettempdir(), '_stysn'), flag='c') as db:
            dbfn, cnts = [db._datfile, db._dirfile, db._bakfile], [0] * 5
            for fn in self._files:
                wb = xwu.safeopen(app, fn)
                nls = xwu.NamedRanges(xwu.usedrange(wb.sheets[0]))
                wb.close()
                for nl in nls:
                    cnts[0] += 1
                    if cnts[0] and cnts[0] % 5000 == 0:
                        logger.debug('%d JOSN records scanned', cnts[0])
                    snnos = formatsn(nl.snno, 1, retuple=True)
                    if not snnos:
                        continue
                    styno = nl.styno
                    for sn in snnos:
                        key = styno + ',' + sn
                        if key not in db:
                            db[key] = '1'
                            cnts[1] += 1
                    if cnts[1] > 1000:
                        cnts[1] = 0
            with open(file, 'wt') as fh:
                nl = 0
                for key in db:
                    nl += 1
                    key = key.decode('ascii').split(',')
                    print('%d\t%s\t%s' % (nl, key[0], key[1]), file=fh)
        logger.debug('Total %d JOSN records scanned', cnts[0])
        if tk:
            xwu.appmgr.ret(tk)
        for fn in dbfn:
            remove(fn)

    @property
    def _files(self):
        return [path.join(self._xls_fldr, x) for x in listdir(self._xls_fldr) if x.find('snnos_wJO') >= 0]

    def index(self):
        app, tk = xwu.appmgr.acq()
        cnt = 0
        with ResourceCtx(self.sm) as cur:
            for fn in self._files:
                wb = xwu.safeopen(app, fn)
                for nl in xwu.NamedRanges(xwu.usedrange(wb.sheets[0])):
                    snnos = formatsn(nl.snno, 1, retuple=True)
                    if not snnos:
                        continue
                    jo = JOx(jono=nl.jono, styno=nl.styno)
                    cur.add(jo)
                    for sn in snnos:
                        cur.add(JOSn(jo=jo, snno=sn))
                    cnt += 1
                    if cnt % 5000 == 0:
                        cur.commit()
                        print('%d jo record done' % cnt)
                cur.commit()
                wb.close()
                print('file(%s) done' % fn)
        xwu.appmgr.ret(tk)

    def find(self, styno, snno):
        ''' return the jo#s that is of given snno + styno
        '''
        with ResourceCtx(self.sm) as cur:
            lst = Query(JOx.jono).join(JOSn).filter(JOx.styno == styno).filter(JOSn.snno == snno).with_session(cur)
            return [x[0] for x in lst] if lst else None

    def get_low_ratios(self, stynos, ratio=0.1, delete=False):
        ''' return the JOs of low ratio style->sn
        In the case of P26382:
            HB923 occurs mainly before 2017/12, for 3 Hings
            HB1485 after 2017/12, for 1 hinge. HB1485 has a 13.9 rate only, but it's necessary.
        So just keep the ratio to not higher than 10
        Args:
            stynos: A collection or generator of styno(string)
            ratio:     the low bound, return the one whose ratio is less than that. 20% can be 0.2 or 20
            delete:  delete the SN from found JOs
        Returns:
            a tuple of (styno, sn, (jono), )
        '''
        q0 = Query(JOSn.snno).join(JOx)
        if ratio > 1:
            ratio = ratio / 100
        toDels = set() if delete else None
        with ResourceCtx(self.sm) as cur:
            for styn in stynos:
                sns = q0.filter(JOx.styno == styn).distinct().with_session(cur).all()
                if not sns:
                    continue
                sns = [x[0] for x in sns]
                for sn in sns:
                    r = self.ratio(styn, sn)
                    if r < ratio:
                        jns = self.find(styn, sn)
                        yield styn, sn, round(r * 100, 1), jns
                        if delete and len(jns) < 3:
                            toDels.update([(x, sn) for x in jns])
        if toDels:
            self.delete_stysn(toDels)

    def _all_stynos(self, cur):
        ''' return all the Sty# inside current db
        '''
        pass

    def delete_stysn(self, stysns, styfinder=None):
        ''' delete give tuple((jono, snno)) from db.
        Args:
            stysns:  a tuple of (jono, snno)
            styfinder:  a StyleFinder that need to delete sty# -> SN# data
        '''
        with ResourceCtx(self.sm) as cur:
            sns = []
            for styn, sn in stysns:
                lst = cur.query(JOSn.id, JOx.styno).join(JOx).filter(JOx.styno == styn).filter(JOSn.snno == sn).all()
                if lst:
                    sns.extend((x[0] for x in lst))
            styfinder.delete(stysns)
            for lst in splitarray(sns):
                cur.query(JOSn).filter(JOSn.id.in_(lst)).delete(synchronize_session=False)
            cur.commit()

    def ratio(self, styno, snno):
        ''' return the ratio the given snno inside styno
        '''
        key = (styno, snno)
        r = self._ratios.get(key)
        if not r:
            r = self._cnt(styno)
            self._ratios[key] = r = self._cnt(styno, snno=snno) / r if r else 0
        return r

    def _cnt(self, styno, **kwargs):
        snno = kwargs.get('snno')
        with ResourceCtx(self.sm) as cur:
            q = Query(func.count(JOx.id)).filter(JOx.styno == styno)
            if snno:
                q = q.join(JOSn).filter(JOSn.snno == kwargs.get('snno'))
            return q.with_session(cur).one()[0]
