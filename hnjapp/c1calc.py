'''
#! coding=utf-8
@Author:        zmFeng
@Created at:    2018-12-18
@Last Modified: 2018-12-18 4:42:37 pm
@Modified by:   zmFeng

c1 calculation excel data filler and so on...

'''
from datetime import datetime
from decimal import Decimal
from logging import DEBUG
from numbers import Number
from os import listdir, path
from time import time

from sqlalchemy import and_, func
from sqlalchemy.orm import Query
from xlwings.constants import LookAt

from hnjapp.c1rdrs import C1InvRdr
from hnjapp.svcs.db import jesin
from hnjcore import JOElement
from hnjcore.models.cn import JO as JOcn, Plating, StoneMaster, Style
from hnjcore.models.hk import JO, POItem
from utilz import (NamedList, ResourceCtx, getvalue, karatsvc, na, splitarray,
                   stsizefmt, trimu)
from utilz.xwu import NamedRanges, appmgr, find, fromtemplate, hidden, offset

from .common import config, _JO_Extractor, Utilz
from .common import _logger as logger
from .localstore import C1JC, C1JCFeature, C1JCStone, FeaSOrN


class _C1Utilz(object):
    CN_JONO = "工单"
    CN_MIT = "MIT"
    MICRON_MISSING = -1000

    def __init__(self):
        # names started with f_ is the field of feature
        self.alias = {
            "jono": self.CN_JONO,
            "styno": "款号",
            "f_spec": "参数",
            "f_karat0": "成色1",
            "f_wgt0": "金重1",
            "f_karat1": "成色2",
            "f_wgt1": "金重2",
            "f_parts": "配件",
            "f_pen": "笔电",
            "f_chain": "链尾",
            "f_tt": "分色",
            "f_micron": "电咪",
            "f_other": "其它",
            "f_mtl2": "银夹金",
            "stone": "石料",
            "shape": "形状",
            "stsize": "尺寸",
            "stqty": "粒数",
            "stwgt": "重量",
            "setting": "镶法",
            "c1cost": "C1工费"
        }
        self._st_sns_abbr = {
            "RD": ("DD", "R"),
            "TD": ("DD", "T"),
            "RZ": ("CZ", "R")
        }

    def find_sheet(self, wb):
        """ find the data sheet, because find is slow, don't use list comprehensive,
        return a list with below element:
        sheet, range("工单"), range("镶法"), NamedList("工单->镶法"))
        """
        to_finds, flag = (
            self.alias["c1cost"],
            self.CN_JONO,
            "镶法",
        ), False
        for sht in wb.sheets:
            fnds = []
            for x in to_finds:
                flag = find(sht, x, lookat=LookAt.xlWhole)
                if not flag:
                    break
                fnds.append(flag)
            if flag:
                fnds = sorted(fnds, key=lambda x: x.column)
                fnds.append(
                    NamedList(sht.range(fnds[0], fnds[-1]).value, alias=self.alias))
                fnds.insert(0, sht)
                return fnds
        return None

    def read_his(self, wb, read_cost=False):
        """ read the calculated out from existing
        @param wb: a workbook object or the result of find_sheet(a list)
        """
        sht = wb if isinstance(wb, (list, tuple)) else self.find_sheet(wb)
        if not sht:
            return None
        rng, var = sht[1], sht[-2]
        alias = self.alias
        if read_cost:
            rng = find(sht[0], "镶石费$", lookat=LookAt.xlWhole)
            alias, nl = alias.copy(), {
                "setcost": "镶石费$",
                "basecost": "胚底费$",
                "labcost": "总价$"
            }
            alias.update(nl)
        nls = [
            x for x in NamedRanges(
                rng, alias, col_cnt=(var.column - rng.column))
            if any(x.data)
        ]
        # sometimes the sheet hides the header row, this might throws excepts, show them
        var = find(sht[0], "RMB对USD:")
        if not var:
            for x in hidden(sht[0]):
                sht[0].range("A%d:A%d" % x).api.entirerow.hidden = False
                if x[1] > 10:
                    break
            var = find(sht[0], "RMB对USD:")
        var = var.end("right").value
        if var and var != 1.0:
            alias = tuple(x for x in nl)
        else:
            var = 0
        for nl in nls:
            jn = nl.jono
            if isinstance(jn, Number):
                nl.jono = JOElement.tostr(jn)
            if not var:
                continue
            for rng in alias:
                if not nl[rng]:
                    continue
                nl[rng] = round(nl[rng] * var, 2)
        return nls

    def extract_st_sns(self, sns):
        """ extract stone and shape out of the sns(Stone&Shape)
        return a tuple of stone, shape
        """
        if not sns:
            return (None,) * 2
        shp = self._st_sns_abbr.get(sns)
        sns, shp = (shp[0], shp[1]) if shp else (sns[1:], sns[0])
        return sns, shp

    @classmethod
    def fetch_jo_skunos(cls, cur, jes):
        ''' return a {JO#: (sku#, styno,)} map '''
        skmp = {}
        q0 = Query((JO.name, POItem.skuno, )).join(POItem)
        for jn in splitarray(jes):
            q = q0.filter(jesin(jn, JO)).with_session(cur).all()
            if not q:
                continue
            skmp.update({x[0].value: trimu(x[1]) for x in q})
        return skmp or None

    @classmethod
    def fetch_ca_jos(cls, cur, jnskump):
        ''' return {skuno: c1jc} map,
        subquery's group by seems non-reasonable, but don't know how
        my original query need to get argument from master query
        @param cur: the cursor to execute the query
        @param jnskump: {jo#: (styno, skuno, )} map
        '''
        if not (jnskump and cur):
            return None
        skus = {x[1]: x[0] for x in jnskump.items() if x[1] and x[1][1] != na}
        q = Query(C1JC).order_by(C1JC.name)
        sq = Query((C1JC.skuno, C1JC.styno, func.max(C1JC.lastmodified).label("lmd"),)).group_by(C1JC.skuno, C1JC.styno).subquery()
        mp = {}
        mk_key = lambda styno, skuno: styno + "_" + skuno
        for sku in splitarray(tuple(skus.keys())):
            x = q.join(sq, and_(C1JC.skuno == sq.c.skuno, C1JC.lastmodified == sq.c.lmd)).filter(C1JC.skuno.in_([x[1] for x in sku]))
            lst = x.with_session(cur).all()
            if not lst:
                continue
            for x in lst:
                thekey = mk_key(x.styno, x.skuno)
                # use the last one, so skip existing checking
                mp[thekey] = x
        mp = {x[0]: mp.get(mk_key(*x[1])) for x in jnskump.items()}
        return {x[0]: x[1] for x in mp.items() if x[1]}


class Writer(object):
    """
    read JO data from source excel and then create the c1cc file from it
    Args:
        hksvc: the services help to handle HK data
        cnsvc: the services help to handle CN data
        his_engine|engine: the creator function help to create the history engine
        keep_dtl: keep detail as possible
    """

    def __init__(self, hksvc, cnsvc, **kwds):
        self._hksvc, self._cnsvc = hksvc, cnsvc
        self._cache_sm = getvalue(kwds, "cache engine")
        self._nl = None

        # stone vermail color
        self._st_vc = config.get("vermail.stones")
        # the fixed stone
        # AZ is based on JO# 460049
        self._st_dd = None  #DD
        self._st_fixed = {y: x[1] for x in config.get("c1calc.stone_dialect", {}).items() for y in x[0].split()}
        self._utilz = _C1Utilz()
        self._init_stx()
        self._extr = _JO_Extractor(kwds.get("keep_dtl"))

    @property
    def _loc_sess_ctx(self):
        return ResourceCtx(self._cache_sm)

    def _init_stx(self):
        """ load stone data from workflow """
        if not self._hksvc:
            return
        kw = "_INITED_"
        if kw not in self._st_fixed:
            mp = config.get("c1calc.fixed_stone")
            q = Query((
                StoneMaster.name,
                StoneMaster.sttype,
            ))
            with self._cnsvc.sessionctx() as cur:
                lst = q.filter(StoneMaster.sttype.in_(
                    [x for x in mp])).with_session(cur).all()
                self._st_fixed.update({x.name: mp[x.sttype] for x in lst})
                lst = q.filter(
                    StoneMaster.sttype == "DIAMOND").with_session(cur).all()
                self._st_dd = {x.name for x in lst}
        self._st_fixed[kw] = 1

    @classmethod
    def _getf_other(cls, jo):
        """ the _specified items, return _specified result """
        return None

    @classmethod
    def _karat_cvt(cls, karat):
        """ karat to the excel's term """
        return karatsvc.getfamily(karat).name

    def _write(self, lsts, sht):
        """ write the result back to the sheet, copy the formulas if required """
        rng0 = find(sht, "镶石费$")
        eidx = rng0.end("down").row
        # formula copy if required
        if len(lsts) - (eidx - rng0.row) > 0:
            rng = sht.range(
                sht.cells(eidx, rng0.column),
                sht.cells(eidx,
                          rng0.end("right").column))
            rng.api.copy
            rng = rng.last_cell.column
            rng = sht.range(
                sht.cells(eidx + 1, rng0.column),
                sht.cells(rng0.row + len(lsts), rng0.column))
            rng.row_height = rng0.row_height
            rng.select()
            sht.api.paste
        # rng0 = find(sht, self._utilz.CN_JONO, lookat=LookAt.xlWhole)
        rng0 = find(sht, self._utilz.alias["c1cost"], lookat=LookAt.xlWhole)
        offset(rng0, 1, 0).value = lsts
        sht.book.app.api.CutCopyMode = False
        offset(rng0, 1, 0).select()

    def _from_his(self, c1jc):
        """ return a list of list with filled data from history """
        def _dec_2_flt(arr):
            for idx, val in enumerate(arr):
                if isinstance(val, Decimal):
                    arr[idx] = float(val)
        nl = self._nl
        nl.styno = c1jc.styno + "_" + c1jc.docno[2:-2] + "_" + c1jc.name
        with self._loc_sess_ctx as cur:
            lst = cur.query(C1JCFeature).filter(C1JCFeature.jcid == c1jc.id).all()
            if lst:
                for x in lst:
                    nl["f_" + x.name] = x.value.v
            lst = cur.query(C1JCStone).filter(C1JCStone.jcid == c1jc.id).all()
            if lst:
                var = lambda x: (x[0], x[-1])
                x1 = dict(var(x.split(":")) for x in "stone shape stsize:size stqty:qty stwgt:wgt setting".split())
                # stone data does not contains MainStone data, so don't sort them
                var = []
                for x in lst:
                    for y in x1.items():
                        if y[0] == 'stsize':
                            sz = getattr(x, y[1]) or ""
                            if sz:
                                nl[y[0]] = "'" + stsizefmt(sz, True)
                        else:
                            nl[y[0]] = getattr(x, y[1])
                    _dec_2_flt(nl.data)
                    var.append(nl.newdata(True))
                return var[:-1]
            _dec_2_flt(nl.data)
        return []

    def _from_jo(self, jo):
        """
        return a list of list with filled data from JO, don't need to return
        the first list, because it's already in the master.
        Some BLogics will be performed during data filling
        """
        cat = self._hksvc.getjocatetory(jo)
        # here the mstr record is already inside self._nl
        nl, lsts, var = self._nl, [], None
        # sort the details by name before operation, MIT always at the end
        mstr, pt = nl.data if cat == "EARRING" else None, False  #earring's Pin has _special offer
        wgts = self._hksvc.getjowgts(jo)
        nls, nl1, var = self._extract_jis(jo, wgts)
        kt = karatsvc.getkarat(wgts.main.karat)
        if not kt:
            logger.critical("JO(%s) does not have weight data" % jo.name.value)
            return []
        self._mstr_BL(jo, wgts, kt, cat, var)
        var = [x for x in (wgts.main, wgts.aux) if x]
        for x, wgt in enumerate(var):
            nl["f_karat%d" % x] = self._karat_cvt(wgt.karat)
            nl["f_wgt%d" % x] = wgt.wgt
        pt = False
        for var in nls or []:
            nl1.setdata(var)
            if mstr and nl1.setting and nl1.setting.find("耳迫") >= 0:  # big5
                pt = True
            for x in nl1.colnames:
                nl[x] = nl1[x]
            lsts.append(nl.newdata(True))
        if pt:
            nl.setdata(mstr)
            kt = karatsvc.getkarat(kt.karat)
            if kt.karat == 925 or kt.category == karatsvc.CATEGORY_BRONZE:
                nl.f_parts = "银迫"
            else:
                nl.f_parts = "9K迫" if kt.color == karatsvc.COLOR_YELLOW else "9KRW迫"
        return lsts[:-1]

    def _mstr_BL(self, jo, pw, kt, cat, vcs):
        """ the business logic of master record """
        nl = lambda w: karatsvc.getkarat(w.karat).category if w and w.wgt else None
        kcs = [nl(x) for x in (pw.main, pw.aux)]
        nl = self._nl
        if pw.part and pw.part.wgt and kcs[0] == karatsvc.CATEGORY_SILVER:
            nl.f_chain = "是"
        if all(kcs):
            nl.f_mtl2 = 1  #at least one
            # gold/silver mixing need pen plating, else ignore
            mp = {'GOLD': 1, 'SILVER': 2}
            if sum(mp.get(x, 100) for x in kcs) == 3:
                nl.f_pen = "是"
        if cat == "BRACELET":
            nl.f_spec = "手链单20"  # don't know if it's 20 or 21
        elif cat == "NECKLACE":
            nl.f_other = (nl.f_other or 0) + 3  # f_chain cutting cost
        elif cat == "PENDANT" and [
                1 for x in "相 盒".split() if jo.description.find(x) >= 0
        ]:  #big5
            nl.f_spec = "相盒"
        if not vcs:
            vcs = []
        if self._extr._keep_dtl:
            nl.f_micron = ";".join('%s=%s' % vc for vc in vcs if vc[1])
        for vc, thk in vcs:
            if thk:
                continue
            if not nl.f_tt and kt.color != vc:
                nl.f_pen = "是"
                break
        if sum((1 for x in "打 噴 沙 砂".split() if jo.remark.find(x) >= 0)) > 1:
            nl.f_other = (nl.f_other or 0) + 5
        if jo.description.find("套") >= 0:  # big5
            nl.f_other = (
                nl.f_other or
                0) + 3 if cat == "PENDANT" else 5 if cat == "RING" else 0

    def _get_micron(self, jo):
        """ return the micron price of given JO# """
        q = Query((
            JOcn.name,
            Plating.uprice,
            Plating.filldate,
        )).join(Plating).join(Style)
        q = q.filter(Style.name == jo.style.name)
        with self._cnsvc.sessionctx() as cur:
            q = q.with_session(cur).all()
        if not q:
            return self._utilz.MICRON_MISSING
        q = sorted(q, key=lambda x: x.filldate, reverse=True)
        return q[0].uprice

    def _extract_jis(self, jo, wgts):
        """ calc the main/side stone, sort and calc the BL """
        # ms is mainstone sign, can be one of M/S/None
        mp = self._extr.extract_jis((jo, ), self._hksvc)
        jis, nl_ji, miscs = mp[jo.id] if mp else (None, ) * 3
        vx = self._extr.extract_vcs(jo, [nl_ji.setdata(x).sto for x in jis] if jis else None, wgts)
        if not mp:
            return None, None, vx
        ms_cand = []
        cat = Utilz.getStyleCategory(jo.style.name.value, jo.description)
        is_ds = jo.remark.find("碟") >= 0 or jo.style.name.alpha.find("M") >= 0
        _ms_chk = lambda cat, nl_ji: nl_ji.stqty == (1 if cat != "EARRING" else 2) and nl_ji.szcal and nl_ji.szcal >= "0300"
        for ji in jis:
            nl_ji.setdata(ji)
            if nl_ji.stsize:
                nl_ji.stsize = "'" + nl_ji.stsize
            if nl_ji.stone and nl_ji.stone.find('CZ') == 0 and len(nl_ji.stone) > 2:
                nl_ji.stone, nl_ji.shape = nl_ji.sto, nl_ji.shpo = nl_ji.stone[:2], nl_ji.stone[2]
            # main stone detection
            if _ms_chk(cat, nl_ji):
                nl_ji.ms = "M"
                ms_cand.append(nl_ji.data)
            else:
                nl_ji.ms = "S" if nl_ji.stqty else "X"
            # TODO maybe here use the DF to switch the is_ds flag
        if is_ds:
            self._jis_diskset(jis, nl_ji)
        if len(ms_cand) > 1:
            # find out the actual MS, set others to S(ide):
            pk = nl_ji.getcol("szcal")
            ms_cand = sorted(ms_cand, key=lambda x: x[pk], reverse=True)
            for x in ms_cand[1:]:
                nl_ji.setdata(x)["ms"] = "S"
        _ms_chk = False
        for ji in jis:
            nl_ji.setdata(ji)
            pk = self._calc_st_set(cat, nl_ji)
            if pk and not _ms_chk and pk.find("腊") >= 0:
                _ms_chk = (nl_ji.stone, nl_ji.shape, )
            nl_ji.setting = pk or nl_ji.setting
        # when one wax, all should be wax, in the case of DDR/CZR, an example
        # is 463783
        if _ms_chk and len(jis) > 1:
            self._st_waxset_check(cat, jis, nl_ji, _ms_chk)
        for ji in miscs:
            nl_ji.setdata(ji)
            if not nl_ji.sto:
                vx.add((nl_ji.setting, None))
            else:
                if nl_ji.sto == self._extr.MISC_LKSZ:
                    continue
                nl_ji.stqty, nl_ji.stwgt = (None, ) * 2
                jis.append(ji)

        ms_cand = {"M": "0", "S": "1"}
        def srt_key(data):
            nl_ji.setdata(data)
            return "%s,%s,%s,%s" % (ms_cand.get(nl_ji.ms, "Z"), nl_ji.sto,
                                    nl_ji.shpo, nl_ji.szcal)

        jis, vx = sorted(
            jis, key=srt_key), tuple(vx) if vx else []
        pk, ji = [nl_ji.getcol(x) for x in "szcal setting".split()]
        return [x[:pk] for x in jis], NamedList(
            nl_ji.colnames[:ji + 1]), vx


    def _jis_diskset(self, sts, nl_ji):
        """ create a diskset item """
        sts.append(nl_ji.newdata(True))
        var = nl_ji.getcol("sto")
        var, fields = [x for x in sts if x[var] in self._st_dd
                      ], "stone shape sto shpo stsize stqty".split()
        if var:
            var = [var[0][nl_ji.getcol(x)] for x in fields]
            var[-1] = -var[-1]
        else:
            var = ["DF", "R", "DF", "R", None, -1]
        fields.append("setting")
        var.append("碟(无石)")
        for x in enumerate(fields):
            nl_ji[x[1]] = var[x[0]]

    def _calc_st_set(self, cat, nl, hints=None):
        """ calculate the setting by given arguments """
        if not (nl.stqty and nl.stone):
            return None
        if nl.stqty < 0:
            return None
        rc = None
        st, shp, qty, sz = [nl[x] for x in "sto shpo stqty szcal".split()]
        if not sz:
            return None
        if st in self._st_fixed:
            rc = self._st_fixed[st]
        elif st in self._st_dd and shp == "R":
            if self._is_st_microset(cat, qty):
                rc = "手微(圆钻)"
            else:
                rc = ("腊" if self._is_st_waxset(st, shp, qty, hints) else "手") + "爪/钉(圆钻)"
        else:
            if nl.ms == "M":
                # main
                rc = "手爪(主石)7x5mm或下" if sz <= "0700" else "手爪(主石)8x6mm或上"
            elif nl.ms == "S":
                # side
                if st == "CZ" and shp == "R":
                    if self._is_st_microset(cat, qty) and sz <= "0300":
                        rc = "手微(CZ 3mm或下)"
                    else:
                        rc = ("腊爪/钉" if self._is_st_waxset(st, shp, qty, hints) else "手爪")
                        rc += "(CZ 3mm%s)" % ("以上" if sz > "0300" else "或下")
                else:
                    rc = "手爪(副石)3mm或下" if sz <= "0300" else "手爪(副石)6x4mm或下"
        # hard-code here, GCL should be as pearl
        if nl.stone and nl.stone[:3] == "GCL":
            rc = "珠"
        return rc

    def _st_waxset_check(self, cat, sts, nl, hints):
        if not sts:
            return
        for x in sts:
            nl.setdata(x)
            st = nl.setting
            if  st and nl.stone == hints[0] and nl.shape == hints[1] and st.find("腊") < 0 and st.find("碟") < 0:
                nl.setting = self._calc_st_set(cat, nl, True)

    @classmethod
    def _is_st_microset(cls, cat, qty):
        """ is microset ? """
        return qty >= 40 and cat in ("RING", "EARRING", "PENDANT") or (
            qty >= 50 and cat == "BANGLE")

    def _is_st_waxset(self, st, shp, qty, hints=None):
        ''' determine if given stone should be wax-set '''
        if hints:
            return True
        if st in self._st_dd and shp == "R":
            return qty >= 6
        if st == "CZ" and shp == "R":
            return qty >= 6
        return False

    def _fetch_data(self, jns, **kwds):
        """ fetch data from db to an list """
        extra = kwds.get("extra", {})
        dtls = {} if kwds.get('_dtl') else None
        with ResourceCtx((self._hksvc.sessmgr(), self._cache_sm, )) as curs:
            tc = time()
            logger.debug("begin to fetch data from HK JO system, might take quite a long time")
            jos, jerrs = self._hksvc.getjos(jns.keys())
            if dtls is not None:
                dtls = {jo.name.name: {'sn': jo.snno,
                    'rmk': jo.remark,
                    'description': jo.description,
                    'cstname': jo.orderma.customer.name.strip(),
                    'wgt': self._hksvc.getjowgts(jo)} for jo in jos}
            if jos and logger.isEnabledFor(DEBUG):
                logger.debug("using %4.2f seconds to get %d jos from HK JO system" % (time() - tc, len(jos)))
            skmp = self._utilz.fetch_jo_skunos(curs[0], [JOElement(x) for x in jns])
            if skmp:
                # JO#(463625,P37209), JO#(463068,E21215), same customer, different Sty#, same SKU#, return {jo: sku_styno}
                var = {x.name.value: x for x in jos}
                skmp = {x[0]: (var[x[0]].style.name.value, x[1]) for x in skmp.items()}
                skmp = self._utilz.fetch_ca_jos(curs[1], skmp)
            if skmp is None:
                skmp = {}
            logger.debug("%d history records returned from cache" % len(skmp))
            lsts, nl = [], self._nl
            handlers = {True: self._from_his, False: self._from_jo}
            jos = sorted(
                jos,
                key=lambda jo: "%s,%s" % (jo.style.name.value, jo.name.value))
            for jo in jos:
                lsts.append(nl.newdata(True))
                var = jo.name.value
                nl["c1cost"] = float(jns[var] or 0)
                nl.jono, nl.styno = "'" + var, jo.style.name.value
                tc = skmp.get(var)
                var = handlers[bool(tc)](tc or jo)
                if var:
                    if not tc:
                        # the mstr record not in nl any more, fetch it from top of lsts
                        nl.setdata(lsts[-1])
                        self._adj_micron(nl, extra or {})
                    lsts.extend(var)
            for jo in jerrs:
                lsts.extend(nl.newdata(True))
                nl.jono = jo
        if dtls is None:
            return lsts
        return lsts, dtls

    def _adj_micron(self, nl, extra):
        ''' put the micron to result, if there is already one, put together '''
        # remove the prefix '
        tc = extra.get(nl.jono)
        if not tc:
            return
        ec = nl.f_micron or 0
        nl.f_micron = tc if not ec or ec == self._utilz.MICRON_MISSING else "%f->%f" % (ec, tc)

    def find_sheet(self, wb):
        ''' find the target sheet inside the given workbook, for foreignor book checking purpose '''
        return self._utilz.find_sheet(wb)

    def _tpl_file(self):
        return config.get("default.file.c1cc.tpl")

    def run(self, wb):
        """ read JOs and write report back """
        # there is at most one data sheet in the book
        sht, extra = self.find_sheet(wb), None
        if not sht:
            # check if it's a C1, if yes, prepare the calc
            nls = C1InvRdr.getReader(wb).read(wb, "c1cc")
            nls = [((x.labor or 0) + (x.setting or 0), x.jono, x.remarks) for y in nls for x in y[0]]
            if not nls:
                logger.debug("file(%s) is not valid c1cc or c1cost file" % wb.name)
                return
            sht = fromtemplate(self._tpl_file())
            wb.close()
            wb, sht = sht, self.find_sheet(sht)
            #find the micron if there is
            extra = {("'" + JOElement.tostr(x[1])): self._extr.extract_micron(x[2]) for x in nls if x[2] and x[2].find("咪") >= 0}
            offset(find(sht[0], self._utilz.alias["c1cost"]), 1, 0).value = [(x[0], x[1]) for x in nls if x[0] > 0]
        self._nl = sht[-1]

        # read the JO#s
        nls = {
            JOElement.tostr(nl.jono): nl["c1cost"]
            for nl in NamedRanges(sht[1], self._utilz.alias) if nl.jono
        }
        if nls:
            logger.debug("totally %d JOs need calculation" % len(nls))
            nls = self._fetch_data(nls, extra=extra)
            self._write(nls, sht[0])
        return wb

    def for_x(self, jns):
        wb = fromtemplate(self._tpl_file())
        sht = self.find_sheet(wb)
        self._nl = sht[-1]
        mp = {x: 0.0 for x in jns}
        lst = self._fetch_data(mp, _dtl=True)
        wb.close()
        if not lst:
            return None
        lst[0].insert(0, self._nl)
        # this is a tuple and map item
        return lst


class HisMgr(object):
    """
    read/save/query c1 cost history
    """
    def __init__(self, cache_sm, hksvc):
        self._sessmgr, self._hksvc = cache_sm, hksvc
        self._utilz = _C1Utilz()

    @property
    def _sessionctx(self):
        return ResourceCtx(self._sessmgr)

    @classmethod
    def _norm_fn(cls, fn):
        ''' sometimes there is very stupid a.xls.xlsx case
        '''
        xtr = lambda nm: trimu(path.splitext(path.basename(nm))[0])
        x, x1 = None, xtr(fn)
        while x != x1:
            x = x1
            x1 = xtr(x)
        return x

    def persist(self, fldr):
        """ gather cc file from sub-folder of fldr and persist
        them to the db
        """
        _is_cc = lambda fn: fn.find("CC") >= 0 and fn.find("_F") > 0
        for sfldr in [
                fn for fn in listdir(fldr) if path.isdir(path.join(fldr, fn))
        ]:
            fns = [
                fn for fn in listdir(path.join(fldr, sfldr))
                if _is_cc(trimu(fn))
            ]
            if not fns:
                continue
            self._persist_file(path.join(fldr, sfldr, fns[0]))

    def _persist_file(self, fn):
        """ persist the given file, if already persisted, do nothing """
        flag = self._is_persisted(fn)
        if flag == 1:
            logger.debug("file(%s) has already been persisted",
                         path.basename(fn))
            return None
        pps, jcs, feas, sts = self._read_file(fn), [], [], []
        var = lambda arr: (arr[0], arr[-1], )
        mstr_map = dict(
            (var(x.split(",")) for x in
             "labcost;setcost;basecost;name,jono;styno;docno;skuno".split(";")))
        st_map = dict(
            (var(x.split(",")) for x in
             "stone;shape;size,stsize;qty,stqty;wgt,stwgt;setting".split(";")))
        dates = datetime.today(), datetime.fromtimestamp(path.getmtime(fn))
        # remove existing if there is
        with self._sessionctx as cur:
            # self._remove_exists(pps)
            with self._hksvc.sessionctx() as var1:
                var = self._utilz.fetch_jo_skunos(var1, [JOElement(x) for x in pps])
                if var:
                    for var1, pp in var.items():
                        pps[var1]["skuno"] = pp
            for pp in pps.values():
                pp = self._new_entity(pp, mstr_map, st_map, dates)
                jcs.append(pp[0])
                for var in zip(pp[1:], (feas, sts)):
                    if var[0]:
                        var[1].extend(var[0])
            if flag == -1:
                self._clear_expired(fn, cur)
            cur.add_all(jcs)
            cur.flush()
            if feas:
                cur.add_all(feas)
            if sts:
                cur.add_all(sts)
            cur.commit()
        logger.debug("%d of records persisted from file(%s)" % (len(jcs), fn))
        return jcs

    def _new_entity(self, pp, mstr_mp, st_mp, dates):
        """ create new entity c1jc entity, return a
        tuple(jc, tuple(c1jcfea), tuple(c1jcst))
        """
        jc, feas, sts = C1JC(), [], []
        for y in mstr_mp.items():
            setattr(jc, y[0], pp.get(y[1], 0))
        # sometimes the Sty# in the source file contains Sty_Date_RefJO#, so normalize it
        if jc.styno and jc.styno.find("_") > 0:
            jc.styno = jc.styno.split("_")[0]
        if jc.docno.find(".") > 0:
            jc.docno = self._norm_fn(jc.docno)
        jc.createdate, jc.lastmodified = dates
        jc.tag = 0
        if not jc.skuno:
            jc.skuno = na
        for y in (y for y in pp if y[:2] == "f_"):
            fea = C1JCFeature()
            feas.append(fea)
            fea.name, fea.jc = y[2:], jc
            fea.value = FeaSOrN(pp[y], None)
        fea = pp.get("stone", [])
        for y in fea:
            st = C1JCStone()
            sts.append(st)
            st.jc = jc
            for fea in st_mp.items():
                setattr(st, fea[0], y.get(fea[1]))
            self._nrm_jc_st(st)
        return jc, feas, sts

    def _remove_exists(self, mp):
        """ remove the existing JO(in database) from mp """
        jns, lsts = splitarray([x for x in mp]), []
        with self._sessionctx as cur:
            for jn in jns:
                q = Query(C1JC).filter(C1JC.name.in_(jn))
                lst = q.with_session(cur).all()
                if not lst:
                    continue
                lsts.extend([x.name for x in lst])
        if lsts:
            logger.debug("JOs(%s) already found in db" % lsts)
            if False:
                for x in lsts:
                    del mp[x]
        return mp

    def _nrm_jc_st(self, st):
        if not st.shape:
            st.stone, st.shape = self._utilz.extract_st_sns(st.stone)
        if st.wgt is None:
            st.wgt = 0

    def _read_file(self, fn):
        app, tk = appmgr.acq()
        wb, mp = app.books.open(fn), {}
        try:
            var = self._utilz.find_sheet(wb)
            nls = self._utilz.read_his(var, True)
            if not nls:
                return None
            var = nls[0].colnames
            jo = var.index("stone")
            var, var1 = var[:jo], var[jo:]
            n_fn, l_jn, jo = self._norm_fn(fn), None, {}
            # in the early age, there might be duplicated JO# inside a file
            for nl in nls:
                jn = JOElement.tostr(nl.jono) or l_jn
                if jn != l_jn:
                    l_jn = jn
                    # master properties
                    jo = {x: nl[x] for x in var if nl[x]}
                    jo["stone"], jo["docno"] = [], n_fn
                    if jn not in mp:
                        mp[jn] = jo
                if not nl.stone or not nl.setting or nl.stone == self._utilz.CN_MIT:
                    continue
                jo["stone"].append({x: nl[x] for x in var1 if nl[x]})
        finally:
            if wb:
                wb.close()
            appmgr.ret(tk)
        return mp

    def _is_persisted(self, fn):
        """ check if given file has been persisted
        @return: 1 if persited
                 0 if not persisted
                 -1 if expired
        """
        flag, fnx = 0, self._norm_fn(fn)
        with self._sessionctx as cur:
            q = Query(func.max(C1JC.lastmodified)).filter(
                C1JC.docno == fnx).with_session(cur).first()
            if q[0]:
                flag = 1 if q[0] >= datetime.fromtimestamp(
                    path.getmtime(fn)) else -1
        return flag

    def _clear_expired(self, fn, cur):
        fn = self._norm_fn(fn)
        qm = cur.query(C1JC.id).filter(C1JC.docno == fn)
        for clz in (C1JCFeature, C1JCStone, ):
            cur.query(clz).filter(
                clz.jcid.in_(qm)).delete(synchronize_session='fetch')
        qm.delete()
