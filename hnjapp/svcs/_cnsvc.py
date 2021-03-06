'''
#! coding=utf-8
@Author:        zmFeng
@Created at:    2019-06-28
@Last Modified: 2019-06-28 2:08:08 pm
@Modified by:   zmFeng
services for hnjchina
'''

import datetime
import re
from collections.abc import Sequence
from functools import cmp_to_key, partial

from sqlalchemy import and_, desc, func
from sqlalchemy.orm import Query

from hnjcore import JOElement
from hnjcore.models.cn import JO as JOcn
from hnjcore.models.cn import MM, Codetable
from hnjcore.models.cn import Customer as Customercn
from hnjcore.models.cn import (MMMa, StoneIn, StoneMaster, StoneOut,
                               StoneOutMaster, StonePk)
from hnjcore.models.cn import Style as Stylecn
from utilz import (NA, NamedList, karatsvc, splitarray, stsizefmt, trimu)

from ..common import Utilz
from ..common import config
from ..pajcc import PrdWgt, WgtInfo
from ._common import (SvcBase, _getjos, formatsn, idset, idsin, jesin, nameset,
                      namesin)

class CNSvc(SvcBase):
    ''' main services for getting data from hnjchina database
    '''

    def getshpforjc(self, df, dt):
        """return py shipment data for PajJCMkr
        @param df: start date(include) a date ot datetime object
        @param dt: end date(exclude) a date or datetime object
        """
        lst = None
        with self.sessionctx() as cur:
            if True:
                q0 = Query([JOcn, MMMa, Customercn, Stylecn, JOcn, MM])
            else:
                q0 = Query([
                    JOcn.id, MMMa.refdate,
                    Customercn.name.label("cstname"),
                    JOcn.name.label("jono"),
                    Stylecn.name.label("styno"), JOcn.running, JOcn.karat,
                    JOcn.description,
                    JOcn.qty.label("joqty"),
                    MM.qty.label("shpqty"),
                    MM.name.label("mmno")
                ])
            q = q0.join(Customercn).join(Stylecn).join(MM).join(MMMa)\
                .filter(and_(MMMa.refdate >= df, MMMa.refdate < dt)).with_session(cur)
            lst = q.all()
        return lst

    def getjcrefid(self, runn):
        """ return the referenceId of given runn#, return tuple
        tuple[0] = the refid, tuple[1] = (runnf,runnt)
        """
        x = None
        with self.sessionctx() as cur:
            q = Query([Codetable.coden0,Codetable.coden1,Codetable.coden2]).filter(and_(Codetable.tblname == "jocostma",Codetable.colname == "costrefid")).\
                filter(and_(Codetable.coden1 <= runn,Codetable.coden2 >= runn))
            x = q.with_session(cur).one_or_none()
        if x:
            return int(x.coden0), (int(x.coden1), int(x.coden2))

    def getjcmps(self, refid):
        """ return the metal ups of given refid as dict """
        x = None
        with self.sessionctx() as cur:
            q = Query(Codetable).filter(and_(Codetable.tblname == "metalma",Codetable.colname == "goldprice")).\
                filter(Codetable.tag == refid)
            lst = q.with_session(cur).all()
            return dict([(int(x.coden0), float(x.coden1)) for x in lst])

    def getstins(self, btnos):
        """
        return the stonein's by provding a list of btchnos or btchid
        btchno or id is determined by the fist item in btchnos
        @param btnos: should be a collection of btchno(str) or btchid(int)
        """
        if not isinstance(btnos, Sequence):
            btnos = tuple(btnos)
        isstr = isinstance(btnos[0], str)
        if isstr:
            qmkr, smkr = namesin, nameset
        else:
            qmkr, smkr = idsin, idset
        return self._getbyids(Query(StoneIn), btnos, qmkr, StoneIn, smkr)

    def getstpks(self, pknos):
        if not isinstance(pknos, Sequence):
            pknos = tuple(pknos)
        isstr = isinstance(pknos[0], str)
        if isstr:
            qmkr, smkr = namesin, nameset
        else:
            qmkr, smkr = idsin, idset
        return self._getbyids(Query(StonePk), pknos, qmkr, StonePk, smkr)

    def getjos(self, jns):
        return _getjos(self, JOcn, Query(JOcn), jns)

    def getjostcosts(self, runns):
        """
        return the stone costs by map, running or je as key and cost as value
        """
        if not runns: return None
        if not isinstance(runns, Sequence): runns = tuple(runns)
        isjn, jnlv = isinstance(runns[0], str) or isinstance(
            runns[0], JOElement), 0
        if isjn and isinstance(runns[0], str):
            runnsx = tuple(runns)
            runns = [JOElement(x) for x in runns]
            jnlv = 1
        else:
            runnsx = runns
        lst, cdmap = [], None
        sign = lambda x: 0 if x == 0 else 1 if x > 0 else -1
        cols = [
            JOcn.name, StoneOutMaster.isout, StonePk.pricen, StonePk.unit,
            func.sum(StoneOut.qty).label("qty"),
            func.sum(StoneOut.wgt).label("wgt")
        ]
        gcols = [JOcn.name, StoneOutMaster.isout, StonePk.pricen, StonePk.unit]
        if not isjn:
            cols[0], gcols[0] = JOcn.running, JOcn.running
        q0 = Query(cols).join(StoneOutMaster).join(StoneOut).join(StoneIn).join(
            StonePk).group_by(*gcols)
        with self.sessionctx() as cur:
            for arr in splitarray(runns, self._querysize):
                try:
                    if isjn:
                        q = q0.filter(jesin(arr, JOcn))
                    else:
                        q = q0.filter(JOcn.running.in_(arr))
                    lst1 = q.with_session(cur).all()
                    if lst1: lst.extend(lst1)
                except:
                    pass
            if lst:
                lst1 = Query([Codetable.coden0, Codetable.tag]).filter(
                    and_(Codetable.tblname == "stone_pkma",
                         Codetable.colname == "unit")).with_session(cur).all()
                cdmap = dict([(int(x.coden0), x.tag) for x in lst1])
        if lst and cdmap:
            costs = dict(zip(runnsx, [0] * len(runns)))
            for x in lst:
                costs[int(x.running) if not isjn else x.name.value if jnlv ==
                      1 else x.name] += round(
                          sign(float(x.isout)) * float(x.pricen) * (float(
                              x.qty) if cdmap[x.unit] == 0 else float(x.wgt)),
                          2)
            return costs

    @property
    def _avg_q(self):
        return Query((
            StoneIn.wgt / StoneIn.qty,
            StonePk.wgtunit * 5,
        )).join(StonePk).filter(StoneIn.qty > 0).order_by(
            desc(StoneIn.filldate))

    def getavgpkwgt(self, pknos, szs=None):
        '''
        get the average unit wgt of a package(in CT). No reliable data, just get the most-recent 10 batches, remove lower/high and make mean for the left.
        Example:
            getavgpkwgt(('DDR00781', 'DDR00105')) => {'DDR00781': 0.xxx, 'DDR00105': 0.xx}
            getavgpkwgt('DDR00781') => 0.xxx
        Args:
            pknos: a tuple of PK#
            szs: a map as PK#->Sz map for size lookup, default is None
        Returns:
            A map as pk#->wgt(float) when the pknos has more than one item or just the wgt(float)
        '''
        if isinstance(pknos, str):
            pknos = (pknos,)
        if isinstance(szs, str):
            szs = {pknos[0]: szs}
        q0 = self._avg_q
        mp = {}
        with self.sessionctx() as cur:
            for pkno in pknos:
                sz = szs.get(pkno) if szs else None
                lst = q0.filter(StonePk.name == pkno)
                if sz:
                    lst = lst.filter(StoneIn.size == sz)
                lst = lst.limit(10).with_session(cur).all()
                if not lst:
                    continue
                lst = sorted([x[0] * x[1] for x in lst])
                if len(lst) > 5:
                    lst = lst[1:-1]
                mp[pkno] = sum(lst) / len(lst)
        return mp if len(mp) > 1 else next(iter(mp.values())) if mp else None

    def getavgstwgt(self, name, shape, sz):
        '''
        get the average unit wgt of a name+shape+size(in CT). No reliable data, just get the most-recent 10 batches, remove lower/high and make mean for the left.
        Example:
            getavgstwgt('ST', 'R', '1') => 0.xxx
        Args:
            name: stone name, just pk#[:2], not the full name
            shape: the shape abbr(for example, R), not full shape(ROUND)
            size: a size formatted by stsizefmt(sz, True)
        Returns:
            a float or None
        '''
        q0 = self._avg_q
        with self.sessionctx() as cur:
            q = q0.filter(StonePk.name.like(name[:2] + shape[0] + '%')).filter(
                StoneIn.size == sz)
            lst = q.limit(10).with_session(cur).all()
            lst = sorted([x[0] * x[1] for x in lst])
            if len(lst) > 5:
                lst = lst[1:-1]
            return sum(lst) / len(lst) if lst else None


class BCSvc(object):
    """a handy Hnjhk dao for data access in this tests
    now it's only bc services
    """
    _querysize = 20  # batch query's batch, don't be too large

    def __init__(self, bcdb=None):
        self._bcdb = bcdb

    def getbcsforjc(self, runns):
        """return running and description from bc with given runnings """
        if not (self._bcdb and runns):
            return None
        runns = [str(x) for x in runns]
        s0 = "select runn,desc,ston from stocks where runn in (%s)"
        lst = []
        cur = self._bcdb.cursor()
        try:
            for x in splitarray(runns, self._querysize):
                cur.execute(s0 % ("'" + "','".join(x) + "'"))
                rows = cur.fetchall()
                if rows:
                    lst.extend(rows)
        except:
            pass
        finally:
            if cur:
                cur.close()
        return self._trim(lst)

    @classmethod
    def _trim(cls, lst):
        if not lst:
            return None
        for x in lst:
            for idx, s0 in enumerate(x):
                if s0 and isinstance(s0, str) and s0[-1] == " ":
                    x[idx] = s0.strip()
        return lst

    def getbcs(self, jnorunn, isstyno=False):
        """ should be jns or runnings
        runnings should be of numeric type
        """
        if not (self._bcdb and jnorunn):
            return None
        if not isinstance(jnorunn, Sequence):
            jnorunn = tuple(jnorunn)
        if isstyno:
            cn = "styn"
        else:
            cn = "jobn" if isinstance(jnorunn[0], str) else "runn"
            if cn == "runn":
                jnorunn = [str(x) for x in jnorunn]

        s0, lst = "select * from stocks where %s in (%%s)" % cn, []
        cur = self._bcdb.cursor()
        try:
            for x in splitarray(jnorunn, self._querysize):
                cur.execute(s0 % ("'" + "','".join(x) + "'"))
                rows = cur.fetchall()
                if rows:
                    lst.extend(rows)
        finally:
            if cur:
                cur.close()
        return self._trim(lst)

    def build_from_jo(self, jn, hksvc, cnsvc, hints=None):
        '''
        create a bc item based on provided jn or provided hints(dict)

        @param jn: A JO#(str)
        @param hksvc: An HKsvc instance(to fetch JO data from)
        @param hints: a dict contains necessary element, where
            '_raw_data' contains a bc item, when this exists, I use it instead of creating a new one.
            '_stone_data' contains a tuple of list, where stone is defined. First item in the tuple is a namedlist to operate on the stone data. The available columns can be found in PajShpHdlr.read_stone_data(). Up to 2019/04/04, the columns are:
                pcode,stone,stshape,stsize,stwgt

        description split into 4 parts:
        1.karat + plating where plating can be from JO.remark or stones
        2.stones
        3.category
        4.suffix
        then data that should also be send to rem1 to rem8 is:
        1.1 aux wgt(if there is)
        1.2 parts wgt(if there is)
        2.1 Any stone data
        99.1 SN#
        '''

        with hksvc.sessionctx():
            jo = hksvc.getjos((jn,))[0][0]
            if not jo:
                return None
            mp = {"jo": jo}
            # karat in hints is not reliable, so at least get karat from JO system
            mp["wgts"] = self._merge_wgts(
                hksvc.getjowgts(jo),
                hints.get("mtlwgt") if hints else None)
            if hints and '_stone_data' in hints:
                var = self._adopt_stone(hints['_stone_data'], Utilz.nl_ji())
                if var:
                    mp["stones"], mp["nl_stone"] = var, Utilz.nl_ji()
            else:
                var = Utilz.extract_jis((jo,), hksvc)
                if var:
                    var = var[jo.id]
                    for idx, x in enumerate("stones nl_stone miscs".split()):
                        mp[x] = var[idx]
            if var:
                var = mp["nl_stone"]
                var = [var.setdata(x).sto for x in mp["stones"]]
            for fn, x in zip((partial(Utilz.extract_vcs, jo, var, mp["wgts"]),
                              partial(Utilz.extract_micron, jo.remark)),
                             ("plating", "micron")):
                mp[x] = fn()
            if hints:
                for key in (
                        '_raw_data',
                        '_snno',
                ):
                    if key in hints:
                        mp[key] = hints[key]
            return _JO2BC(cnsvc).build(mp)

    @classmethod
    def _merge_wgts(cls, jowgts, hintwgts):
        if not (jowgts and hintwgts):
            return None if not (jowgts or hintwgts) else jowgts or hintwgts
        rst = []
        for j, h in zip(jowgts.wgts, hintwgts.wgts):
            if j or h:
                rst.append(
                    WgtInfo(j.karat if j else h.karat, h.wgt if h else j.wgt))
            else:
                rst.append(None)
        return PrdWgt(*rst)

    @classmethod
    def _adopt_stone(cls, sts, nl_ji):
        nl_src, rsts = sts[0], []
        for x in sts[1:]:
            nl_src.setdata(x)
            rsts.append(nl_ji.newdata())
            nl_ji.stone = nl_ji.sto = nl_src.stone
            nl_ji.shape = nl_ji.shpo = nl_src.shape
            nl_ji.stsize, nl_ji.stqty, nl_ji.stwgt = nl_src.size, nl_src.qty, nl_src.wgt
            nl_ji.szcal = stsizefmt(nl_src.size)
        return rsts


class _JO2BC(object):

    def __init__(self, cnsvc=None):
        # pp holds properties except (locket, locket_pic)
        self._pp = self._pp_x = None
        self._v_c2n = {x["color"]: x for x in config.get("vermail.defs")}
        x = config.get("bc.description.stone.categories")
        self._stcats = {z: x for x, y in x.items() for z in y}
        self._cnsvc = cnsvc
        x = config.get("bc.sns")
        if x:
            self._sns_mp = {(x[0], x[1]): x[2] for x in x}
        else:
            self._sns_mp = None
        self._st_bywgt = {x for x in config.get("bc.stone.bywgt")}

    def build(self, pp):
        '''
        Build a bc instance based on the dict provided
        '''
        self._pp = pp
        self._pp_x = {}
        nl_bc = NamedList(config.get("bc.colnames.alias"))
        vx = self._pp.get("_raw_data")
        if vx:
            nl_bc.setdata(vx)
        else:
            nl_bc.newdata()
            self._fill_basic(nl_bc)
        self._sort_st_ms(nl_bc)
        self._make_desc(nl_bc)
        self._make_rmks(nl_bc)
        return nl_bc

    def _fill_basic(self, nl):
        jo, wgts = (self._pp[x] for x in "jo wgts".split())
        nl.styno = nl.location = jo.style.name.value
        nl.qty, nl.wgt, nl.karat = jo.qty, wgts.main.wgt, wgts.main.karat
        td = datetime.date.today()
        nl.date, nl.mp_year, nl.mp_month = td.strftime(
            "%Y%m%d 00:0000"), td.strftime("%y"), td.strftime("%m")
        nl.cstname, nl.jono, nl.descn = trimu(
            jo.orderma.customer.name), jo.name.value, jo.description

    def _sort_st_ms(self, nl_bc):
        '''
        sort the stone, fill the main stone if there is
        '''
        ms = "--"
        vx, nl = [self._pp.get(x) for x in ("stones", "nl_stone")]
        if vx:
            cat = self._pp["jo"]

            def srt_cmp(st0, st1):
                rc = [nl.setdata(x).stsize or '' for x in (st0, st1)]
                rc = 0 if rc[0] == rc[1] else 1 if rc[0] > rc[1] else -1
                if rc == 0:
                    sts = [nl.setdata(x).stone for x in (st1, st0)]
                    rc = 0 if sts[0] == sts[1] else 1 if sts[0] > sts[1] else -1
                return rc

            cat = Utilz.getStyleCategory(cat.style.name.value, cat.description)
            for x in vx:
                nl.setdata(x)
                if Utilz.is_main_stone(cat, nl.stqty, nl.szcal):
                    ms = nl.sto
                    break
            # sort the stone by MS + "SZ(A)+ST(D)"
            sts = sorted(vx, key=cmp_to_key(srt_cmp))
            st_p = {nl.setdata(st).stone: idx for idx, st in enumerate(sts)}
            sts = sorted(
                vx,
                key=lambda x: (st_p[nl.setdata(x).stone], nl.stsize or ''),
                reverse=True)
            self._pp["stones"] = sts
            if ms == '--':
                ms = nl.setdata(sts[0]).sto
        nl_bc.stone = ms

    def _make_desc(self, nl_bc):
        rc = self._make_karat_plating()
        rc += self._make_desc_stones(nl_bc)
        tc = self._make_style_cat()
        rc, tc = rc + " " + tc[0], tc[1]
        #<pendant's chain>
        # ^E customer don't need wo chain except ESO/ET/EJE/ELH
        if tc in ("PENDANT", 'LOCKET'):
            nl, wgts, jo = [
                self._pp.get(x) for x in ('nl_stone wgts jo'.split())
            ]
            chns = config.get("jo.description.chain")
            if wgts.part or \
                [1 for x in chns if jo.description.find(x) >= 0] or \
                [1 for x in self._pp.get("miscs", []) if nl.setdata(x).sto == Utilz.MISC_MIT and [1 for x in chns if nl.setting.find(x) >= 0]]:
                rc += ' ROPE CHAIN'
                # check if wgts contains part, if not, insert one
                if not wgts.part:
                    self._pp["wgts"] = PrdWgt(wgts.main, wgts.aux,
                                              WgtInfo(wgts.main.karat, 0))
            else:
                vx = trimu(jo.orderma.customer.name)
                # rule according to e-mail title('EJE 新樣品 P41001_P41002') sent in Jul 12, 2018
                if vx[0] != 'E' or vx in config.get(
                        'bc.description.wochain.customers'):
                    rc += ' W/O CHAIN'
        #</pendant's chain>
        nl_bc.description = rc

    def _make_karat_plating(self):
        '''
        mixed the gold/color based on rule
        '''
        return self.knc_mix(self._pp.get('wgts'), self._pp.get("plating"))

    def knc_mix(self, wgts, vcs):
        '''
        9K&VW means partial VW
        9K VW measn all VW
        '''
        kts = [x.karat for x in (wgts.main, wgts.aux) if x and x.karat]
        if vcs is None:
            vcs = set()  #don't return directly, below has 200 case
        else:
            vcs = {x[0] for x in vcs}
        tt = len(kts) > 1
        vx = lambda x: self._v_c2n[x]["name"]
        vxx = lambda y: [vx(x) for x in sorted(tuple(y), key=lambda x: self._v_c2n[x]["priority"])]
        kn = lambda x: "&".join((karatsvc.getkarat(y).name for y in x))
        if not tt:
            rc = kn(kts)
            if vcs:
                rc += "%s%s" % (" " if kts[0] in (925, 200) else "&", "&".join(
                    vxx(vcs)))
        else:
            # when there is bronze, follow master's color
            if kts[1] == 200:
                vcs.add(karatsvc.getkarat(kts[0]).color)
            if kts[0] == 925:
                rc = karatsvc.getkarat(kts[0]).name + " " + "&".join(
                    vxx(vcs)) + "&" + karatsvc.getkarat(kts[1]).name
            else:
                # gold + X case
                rc = kn(kts) + " " + "&".join(vxx(vcs))
        return rc

    def _make_style_cat(self):
        jo, vx = self._pp["jo"], None
        tc = Utilz.getStyleCategory(jo.style.name.value, jo.description)
        if tc in ("PENDANT", "LOCKET"):
            pc = re.compile(config.get("pattern.locket.pics"))
            pc = pc.search(jo.description)
            if pc:
                pc = pc.group(1)
            nl = self._pp.get("nl_stone")
            vx = [
                x for x in self._pp.get("miscs", [])
                if nl.setdata(x).sto == Utilz.MISC_LKSZ
            ]
            if not vx:
                vx = jo.description.find("相盒") >= 0 or jo.remark.find("相盒") >= 0
            if not (vx or pc):
                vx = ""
            elif vx and isinstance(vx, (tuple, list)):
                nl.setdata(vx[0])
                self._pp_x["locket"], vx = "%s:%s" % (nl.stone,
                                                      nl.stsize), 'LOCKET '
            else:
                self._pp_x["locket"], vx = "LKSZ:TODO", 'LOCKET '
            if pc:
                self._pp_x['locket_pic'] = "%s PIC" % pc
        elif tc == "EARRING":
            if jo.description.find("啤") > 0 or jo.remark.find("啤") >= 0:
                vx = 'CREOLE '
        return (vx or "") + tc, tc

    def _make_rmks(self, nl_bc):
        lst = []
        wgts = self._pp["wgts"]
        wgts = (wgts.aux, wgts.part)
        if wgts[0]:
            lst.append("%s %4.2f" % (karatsvc.getkarat(wgts[0].karat).name,
                                     wgts[0].wgt))
        if wgts[1]:
            lst.append("*%sPTS %4.2f" % (karatsvc.getkarat(wgts[1].karat).name,
                                         wgts[1].wgt))
        vx = self._pp.get("plating")
        if vx:
            vx = [x for x in vx if not x[1]], [x for x in vx if x[1]]
            lst.append(self.knc_mix(self._pp.get('wgts'), vx[0]))
            if vx[1]:
                # vk with height case
                for x in vx[1]:
                    lst.append(x[0] + (" " + x[1] if x[1] else ""))
        for vx in (self._pp_x.get(x) for x in ("locket", "locket_pic")):
            if vx:
                lst.append(vx)
        vx, nl = [self._pp.get(x) for x in ("stones", "nl_stone")]
        if vx:
            for x in vx:
                lst.append(self._make_rem_stone(nl.setdata(x)))
        vx = self._pp.get("_snno")
        vx = self._pp["jo"].snno + ("," + vx if vx else "")
        if vx:
            vx = self.extract_snno(vx)
        if vx:
            vx[0] = "SN#" + vx[0]
            lst.extend(vx)
        else:
            lst.append('SN#%s' % NA)
        mc = 8
        # merge them to 8 items if there are more
        for idx, vx in enumerate(lst[:mc]):
            setattr(nl_bc, "rem%d" % (idx + 1), vx)
        if len(lst) > mc:
            setattr(nl_bc, "rem%d" % mc, ",".join(lst[mc - 1:]))
        return nl_bc

    def _make_rem_stone(self, nl):
        rc = str(int(nl.stqty)) if nl.stqty > 1 else ""
        rc += self._nrm_sns(nl.shpo, nl.sto)
        if nl.stwgt and nl.stwgt != nl.stqty and self._is_st_bywgt(nl.sto):
            rc += "-%r" % round(nl.stwgt, 3)
        if not self._is_st_bywgt(nl.sto):
            rc += "-" + (nl.stsize or "") + "MM"
        return rc

    def _nrm_sns(self, shp, stname):
        '''
        return RDD if shp == 'R' and stname == 'DD'
        '''
        return self._sns_mp.get((shp, stname), shp + stname)

    def _is_st_bywgt(self, stname):
        return stname in self._st_bywgt
        # following PK's unit is not a good point, use fixed answer from conf.json
        # with self._cnsvc.sessionctx() as cur:
        #    pk = cur.query(StonePk).join(StoneMaster).filter(StoneMaster.name == stname).first()
        #    if not pk:
        #        return False
        #    return pk.unit in (1, 3, 5)

    def _make_desc_stones(self, nl_bc):
        vx, nl = [self._pp.get(x) for x in ("stones", "nl_stone")]
        if not vx:
            nl_bc.stone = "--"
            return ""
        mc = config.get('bc.description.stone.maxcnt', 3)
        # whenever there is DIA, make it out
        cats, rc, sms, ms_cat = {}, [], {}, None
        lmb_sn = lambda pk: pk[:2]
        for x in vx:
            st = lmb_sn(nl.setdata(x).stone)
            cat = self._stcats.get(st, '*MULTI_STONE')
            if st == nl_bc.stone:
                ms_cat = cat = (" MS" if cat[0] == "*" else cat)
            if st not in sms:
                sms[st] = None
            cat = cats.setdefault(cat, [])
            if st not in cat:
                cat.append(st)
        if self._cnsvc:
            with self._cnsvc.sessionctx() as cur:
                lst = cur.query(StoneMaster).filter(
                    StoneMaster.name.in_(sms.keys())).all()
                sms = {x.name: x.edesc for x in lst}
        else:
            sms = {}
        # sort the cats, ms as first, compositive as second, later by stone name
        #keys = sorted(cats, key=lambda x: ' ' if x == nl_bc.stone else '-' if x in self._stcats else x)
        keys = sorted(
            cats,
            key=lambda x: -10 if ms_cat == x else -1 if len(cats[x]) > mc else 1
        )
        for cn in keys:
            sts = cats.get(cn)
            if len(sts) > mc:
                rc.append(cn)
            else:
                for x in sts:
                    rc.append(sms.get(x, x))
        return ' ' + " ".join(rc)

    def extract_snno(self, snno):
        '''
        parse the snno parts out from snno, the chinese will be removed
        and only valid snno will be left
        # HB1729夾片底HB923,中號光金瓜子耳
        # HB2056底HB943夾片HB114夾層HB982,中號光金瓜子耳
        # #
        # HB1762底P27473,中號光金瓜子耳
        # #HB2600底夾片HB1485,PT5012BL
        '''
        return formatsn(snno, 0, True)
