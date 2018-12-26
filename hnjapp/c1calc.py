'''
#! coding=utf-8
@Author:        zmFeng
@Created at:    2018-12-18
@Last Modified: 2018-12-18 4:42:37 pm
@Modified by:   zmFeng

c1 calculation excel data filler and so on...

'''
from numbers import Number
from decimal import Decimal
from os import listdir, path
from re import compile as cmpl
from datetime import datetime

from sqlalchemy.orm import Query
from sqlalchemy import func
from xlwings.constants import LookAt

from hnjapp.dbsvcs import idset, idsin
from hnjcore.models.cn import JO as JOcn
from hnjcore.models.cn import Plating, StoneMaster, Style
from hnjcore.models.hk import JO
from hnjcore.models.hk import JOItem as JI
from utilz import NamedList, karatsvc, stsizefmt, trimu, na, SessionMgr, ResourceCtx
from utilz.xwu import NamedRanges, appmgr, find

from .common import _logger as logger
from .localstore import C1JC, C1JCFeature, C1JCStone, FeaSOrN

_strfn = lambda jn: "%d" % int(jn) if isinstance(jn, Number) else jn

class Utilz(object):
    CN_JONO = "工单"
    CN_MIT = "MIT"
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
            "setting": "镶法"
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
        to_finds, flag = (self.CN_JONO, "镶法", ), False
        for sht in wb.sheets:
            fnds = []
            for x in to_finds:
                flag = find(sht, x, lookat=LookAt.xlWhole)
                if not flag:
                    break
                fnds.append(flag)
            if flag:
                fnds.insert(0, sht)
                fnds.append(NamedList(sht.range(*fnds[1:]).value, alias=self.alias))
                return fnds
        return None

    def read_his(self, wb, read_cost=False):
        """ read the calculated out from existing
        @param wb: a workbook object or the result of find_sheet(a list)
        """
        sht = wb if isinstance(wb, (list, tuple)) else self.find_sheet(wb)
        if not sht:
            return None
        rng, var = sht[1:3]
        alias = self.alias
        if read_cost:
            rng = find(sht[0], "镶石费$", lookat=LookAt.xlWhole)
            alias, nl = alias.copy(), {"setcost": "镶石费$", "basecost": "胚底费$", "labcost" :"总价$"}
            alias.update(nl)
        nls = [
            x for x in NamedRanges(
                rng, col_cnt=(var.column - rng.column), alias=alias)
            if any(x.data)
        ]
        # sometimes the sheet hides the header row, this might throws excepts, show them
        var = find(sht[0], "RMB对USD:").end("right").value
        if var and var != 1.0:
            alias = tuple(x for x in nl)
        else:
            var = 0
        for nl in nls:
            jn = nl.jono
            if isinstance(nl, Number):
                nl.jono = _strfn(jn)
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
        shp = self._st_sns_abbr.get(sns)
        if not shp:
            if sns:
                sns, shp = sns[1:], sns[0]
        else:
            sns, shp = shp[0], shp[1]
        return sns, shp


class Writer(object):
    """ read/write data from excel """

    _ptn_pk = cmpl(r"([A-Z]{2})-([A-Z@])-([A-Z\d]{1,4})-([A-Z\d])")

    def __init__(self, hksvc, cnsvc):
        self._hksvc = hksvc
        self._cnsvc = cnsvc
        self._nl = None

        # vermail color detection map
        #VWhite and #VBlue
        self._st_vw, self._st_vb = set(("DD", "CZ", )), set("DF")
        self._vx, self._vx_white, self._vx_blue = "電", "電白", "電藍"
        self._vc_mp = {"咪": "_MICRON_"} # Micro has higher priority
        mp = {"白": "WHITE", "黑": "BLACK", "藍": "BLUE", "黃 王": "YELLOW", "玫 瑰": "ROSE"}
        self._vc_mp.update({x[0]: getattr(karatsvc, "COLOR_" + trimu(x[1])) for x in mp.items()})

        # the fixed stone
        # AZ is based on JO# 460049
        self._st_dd = None #DD
        mp = {"AZ":  "玉", "ON": "安力士"}
        self._st_fixed = {y: x[1] for x in mp.items() for y in x[0].split()}
        self._utilz = Utilz()
        self._init_stx()

    def _init_stx(self):
        """ load stone data from workflow """
        kw = "_INITED_"
        if kw not in self._st_fixed:
            mp = {"JADE": "玉", "PEARL": "珠"}
            q = Query((StoneMaster.name, StoneMaster.sttype, ))
            with self._cnsvc.sessionctx() as cur:
                lst = q.filter(StoneMaster.sttype.in_([x for x in mp])).with_session(cur).all()
                self._st_fixed.update({x.name: mp[x.sttype] for x in lst})
                lst = q.filter(StoneMaster.sttype == "DIAMOND").with_session(cur).all()
                self._st_dd = {x.name for x in lst}
        self._st_fixed[kw] = 1

    def _get_his(self, skuno):
        """ return the history data by SKU# """
        return []

    def _getf_other(self, jo):
        """ the _specified items, return _specified result """
        return None

    def _karat_cvt(self, karat):
        """ karat to the excel's term """
        return karatsvc.getfamily(karat).name

    def _write(self, lsts, sht):
        """ write the result back to the sheet, copy the formulas if required """
        rng0 = find(sht, "镶石费$")
        eidx = rng0.end("down").row
        # formula copy if required
        if  len(lsts) - (eidx - rng0.row) > 0:
            rng = sht.range(sht.cells(eidx, rng0.column), sht.cells(eidx, rng0.end("right").column))
            rng.api.copy
            rng = rng.last_cell.column
            rng = sht.range(sht.cells(eidx + 1, rng0.column), sht.cells(rng0.row + len(lsts), rng0.column))
            rng.row_height = rng0.row_height
            rng.select()
            sht.api.paste
        rng0 = find(sht, self._utilz.CN_JONO, lookat=LookAt.xlWhole)
        rng0.offset(1, 0).value = lsts

    def _from_his(self, nl, his, **kwds):
        """ return a list of list with filled data from history """
        return None

    def _from_jo(self, jo, **kwds):
        """
        return a list of list with filled data from JO, don't need to return
        the first list, because it's already in the master.
        Some BLogics will be performed during data filling
        """
        cat = self._hksvc.getjocatetory(jo)
        nl, nls, lsts, var = self._nl, kwds.get("dtls"), [], None
        if nls:
            # sort the details by name before operation, MIT always at the end
            mstr, pt = nl.data if cat == "EARRING" else None, False  #earring's Pin has _special offer
            nls, nl1, var = self._extract_jis(jo, cat, nls)
        if var:
            jo.remark = ";".join(var) + ";" + jo.remark
        var = self._hksvc.getjowgts(jo)
        kt = karatsvc.getkarat(var.main.karat)
        self._mstr_BL(jo, var, kt, cat)
        var = [x for x in (var.main, var.aux) if x]
        for x, wgt in enumerate(var):
            nl["f_karat%d" % x] = self._karat_cvt(wgt.karat)
            nl["f_wgt%d" % x] = wgt.wgt
        if nls:
            pt = False
            for var in nls:
                nl1.setdata(var)
                if mstr and nl1.setting.find("耳迫") >= 0:  # big5
                    pt = True
                for x in nl1.colnames:
                    nl[x] = nl1[x]
                lsts.append(nl.newdata(True))
            if pt:
                nl.setdata(mstr).f_parts = {
                    925: "银迫",
                    9: "9K迫"
                }.get(kt.karat, "9KRW迫")
            return lsts[:-1]
        return []

    def _mstr_BL(self, jo, pw, kt, cat):
        """ the business logic of master record """
        nl = self._nl
        if pw.part and pw.part.wgt:
            nl.f_chain = "是"
        if pw.main and pw.aux:
            nl.f_mtl2 = 1 #at least one
            nl.f_pen = "是"
        if cat == "BRACELET":
            nl.f_spec = "手链单20" # don't know if it's 20 or 21
        elif cat == "NECKLACE":
            nl.f_other = (nl.f_other or 0) + 3 # f_chain cutting cost
        elif cat == "PENDANT" and [
                1 for x in "相 盒".split() if jo.description.find(x) >= 0
        ]: #big5
            nl.f_spec = "相盒"
        vcs = self._extract_vcs(jo.remark)
        if vcs:
            for vc in vcs:
                if vc == "_MICRON_":
                    nl.f_micron = self._get_micron(jo)
                elif not nl.f_tt and kt.color != vc:
                    nl.f_pen = "是"
        if sum((1 for x in "打 噴 沙 砂".split() if jo.remark.find(x) >= 0)) > 1:
            nl.f_other = (nl.f_other or 0) + 5
        if jo.description.find("套") >= 0: # big5
            nl.f_other = (nl.f_other or 0) + 3 if cat == "PENDANT" else 5 if cat == "RING" else 0

    def _get_micron(self, jo):
        """ return the micron price of given JO# """
        q = Query((JOcn.name, Plating.uprice, Plating.filldate, )).join(Plating).join(Style)
        q = q.filter(Style.name == jo.style.name)
        with self._cnsvc.sessionctx() as cur:
            q = q.with_session(cur).all()
        if not q:
            return -1000
        q = sorted(q, key=lambda x: x.filldate, reverse=True)
        # TODO:: get same SKU# item or just return the newest one?
        return q[0].uprice

    def _extract_vcs(self, rmk):
        """ the vermail color, None or an list """
        idx, rc = rmk.find(self._vx), None
        if idx < 0:
            return rc
        lst = set()
        for var in rmk.split(self._vx)[1:]:
            var, rc = var[:6], None
            for x in self._vc_mp.items():
                if [1 for y in x[0].split() if var.find(y) >= 0]:
                    rc = x[1]
                    break
            if rc:
                lst.add(rc)
        return lst

    def _extract_pk(self, rmk):
        """ extract the PK# from the remark """
        mt = self._ptn_pk.search(rmk)
        if not mt:
            return None
        pts = [x for x in mt.groups()]
        if pts[2].isnumeric():
            pts[2] = "%04d" % int(pts[2])
        return "".join(pts)

    def _extract_ji(self, ji, nl_ji):
        """ extract the JOItem to the given NamedList """
        st = trimu(ji.stname)
        sto, shp, shpo = (None, ) * 3
        blk = lambda st, x: None if st == self._utilz.CN_MIT else x
        st = "".join([x for x in st if "A" <= x <= "Z"])
        sz, qty, wgt = [blk(st, x) for x in (ji.stsize, ji.qty, ji.wgt)]
        if st != self._utilz.CN_MIT:
            pk = self._extract_pk(ji.remark)
            if pk:
                st, shp = pk, None
                sto, shpo = st[:2], st[2]
            else:
                sto, shpo = self._utilz.extract_st_sns(st)
        else:
            sto = st
        if sz == ".":
            sz = None
        nl_ji.setdata([st, shp, sz, qty, wgt, ji.remark, stsizefmt(sz), None, sto, shpo])

    def _extract_jis(self, jo, cat, jis):
        """ calc the main/side stone, sort and calc the BL """
        # the last one is mainstone sign, can be one of M/S/None
        nl_ji = NamedList("stone shape stsize stqty stwgt setting szcal ms sto shpo".split())
        sts, mns, dscs, hasds = [], [], set(), (jo.remark.find("碟") >= 0 or jo.style.name.alpha.find("M") >= 0)
        for ji in jis:
            self._extract_ji(ji, nl_ji)
            if not nl_ji.sto:
                dscs.add(nl_ji.setting)
                continue
            pk = nl_ji.stqty
            # main stone detection
            if pk == (1 if cat != "EARRING" else 2) and nl_ji.szcal >= "0300":
                nl_ji.ms = "M"
                mns.append(nl_ji.data)
            else:
                nl_ji.ms = "S" if pk else "X"
            sts.append(nl_ji.data)
            # this stone need WHITE
            if nl_ji.sto in self._st_vw:
                dscs.add(self._vx_white)
            elif nl_ji.sto in self._st_vb:
                dscs.add(self._vx_blue)
                hasds = True
        if hasds:
            sts.append(nl_ji.newdata(True))
            pk = nl_ji.getcol("sto")
            pk, hasds = [x for x in sts if x[pk] in self._st_dd], "stone shape sto shpo stsize stqty".split()
            if pk:
                pk = [pk[0][nl_ji.getcol(x)] for x in hasds]
                pk[-1] = -pk[-1]
            else:
                pk = ["DF", "R", "DF", "R", None, -1]
            hasds.append("setting")
            pk.append("碟(无石)")
            for x in enumerate(hasds):
                nl_ji[x[1]] = pk[x[0]]
        if len(mns) > 1:
            # find out the actual MS:
            pk = nl_ji.getcol("szcal")
            mns = sorted(mns, key=lambda x: x[pk], reverse=True)
            for x in mns[1:]:
                nl_ji.setdata(x)["ms"] = "S"
        for ji in sts:
            nl_ji.setdata(ji)
            nl_ji.setting = self._calc_stset(cat, nl_ji) or nl_ji.setting
        mns = {"M": "0", "S": "1"}
        def srt_key(data):
            nl_ji.setdata(data)
            return "%s,%s,%s,%s" % (mns.get(nl_ji.ms, "Z"), nl_ji.sto, nl_ji.shpo, nl_ji.szcal)
        sts, dscs = sorted(sts, key=srt_key), tuple(dscs) if dscs else []
        pk, ji = [nl_ji.getcol(x) for x in "szcal setting".split()]
        return [x[:pk] for x in sts], NamedList(nl_ji.colnames[:ji + 1]), dscs

    def _calc_stset(self, cat, nl):
        """ calculate the setting by given arguments """
        if not (nl.stqty and nl.stone):
            return None
        if nl.stqty < 0:
            return None
        rc = None
        st, shp, qty, sz = [nl[x] for x in "sto shpo stqty szcal".split()]
        if st in self._st_fixed:
            rc = self._st_fixed[st]
        elif st in self._st_dd and shp == "R":
            if self._is_st_microset(cat, qty):
                rc = "手微(圆钻)"
            else:
                rc = ("腊" if qty >= 5 else "手") + "爪/钉(圆钻)"
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
                        rc = ("腊爪/钉" if qty >= 5 else "手爪")
                        rc += "(CZ 3mm%s)" % ("以上" if sz > "0300" else "或下")
                else:
                    rc = "手爪(副石)3mm或下" if sz <= "0300" else "手爪(副石)6x4mm或下"
        return rc

    @classmethod
    def _is_st_microset(cls, cat, qty):
        """ is microset ? """
        return qty >= 40 and cat in ("RING", "EARRING", "PENDANT") or (qty >= 50 and cat == "BANGLE")

    def _fetch_data(self, jns):
        """ fetch data from db to an list"""
        with self._hksvc.sessionctx() as cur:
            jos, jerrs = self._hksvc.getjos(jns)
            dtls = Query((
                JO.id,
                JI.stname,
                JI.stsize,
                JI.stsize,
                JI.qty,
                JI.unitwgt.label("wgt"),
                JI.remark,
            )).join(JI)
            dtls, lsts = dtls.filter(idsin(idset(jos),
                                           JO)).with_session(cur).all(), {}
            for jo in dtls:
                lsts.setdefault(jo.id, []).append(jo)
            dtls, lsts, nl = lsts, [], self._nl
            ftrs = {True: self._from_his, False: self._from_jo}
            jos = sorted(jos, key=lambda jo: "%s,%s" % (jo.style.name.value, jo.name.value))
            for jo in jos:
                lsts.append(nl.newdata(True))
                nl.jono, nl.styno = "'" + jo.name.value, jo.style.name.value
                var = self._get_his(jo.po.skuno)
                var = ftrs[bool(var)](var or jo, dtls=dtls.get(jo.id))
                if var:
                    lsts.extend(var)
            for jo in jerrs:
                lsts.extend(nl.newdata(True))
                nl.jono = jo
        return lsts

    def run(self, wb):
        """ read JOs and write report back """
        # there is at the most one data sheet in the book
        sht = self._utilz.find_sheet(wb)
        if not sht:
            return
        self._nl = sht[-1]
        # read the JO#s
        nls = {_strfn(nl.jono) for nl in NamedRanges(sht[1], alias=self._utilz.alias) if nl.jono}
        if nls:
            nls = self._fetch_data(nls)
            self._write(nls, sht[0])

class HisMgr(object):
    """
    read/save/query c1 cost history
    """
    _utilz = Utilz()

    def __init__(self, eng_crtr):
        self._eng_crtr = eng_crtr
        self._sessmgr = None

    def sessionctx(self):
        if not self._sessmgr:
            self._sessmgr = SessionMgr(self._eng_crtr())
        return ResourceCtx(self._sessmgr)

    def _norm_fn(self, fn):
        return trimu(path.splitext(path.basename(fn))[0])

    def persist(self, fldr):
        """ gather cc file from sub-folder of fldr and persist
        them to the db
        """
        _is_cc = lambda fn: fn.find("CC") >= 0 and fn.find("_F") > 0
        for sfldr in [fn for fn in listdir(fldr) if path.isdir(path.join(fldr, fn))]:
            fns = [fn for fn in listdir(path.join(fldr, sfldr)) if _is_cc(trimu(fn))]
            if not fns:
                continue
            logger.debug("handling file(%s)" % fns[0])
            self._persist_file(path.join(fldr, sfldr, fns[0]))

    def _persist_file(self, fn):
        """ persist the given file, if already persisted, do nothing """
        pps = self._is_persisted(fn)
        if pps == 1:
            logger.debug("file(%s) has already been persisted", path.basename(fn))
            return None
        elif pps == -1:
            self._clear_expired(fn)
        pps, jcs, feas, sts = self._read_file(fn), [], [], []
        mstr_map, st_map = ({}, ) * 2
        arr_mp = lambda arr: (arr[0], arr[-1], )
        mstr_map = dict((arr_mp(x.split(",")) for x in "labcost;setcost;basecost;name,jono;styno;docno".split(";")))
        st_map = dict((arr_mp(x.split(",")) for x in "stone;shape;size,stsize;qty,stqty;wgt,stwgt;setting".split(";")))
        cdate, mdate = datetime.today(), datetime.fromtimestamp(path.getmtime(fn))
        for x in pps:
            jc = C1JC()
            jcs.append(jc)
            for y in mstr_map.items():
                setattr(jc, y[0], x.get(y[1], 0))
            jc.createdate, jc.lastmodified, jc.skuno, jc.tag = cdate, mdate, na, 0
            for y in (y for y in x if y[:2] == "f_"):
                fea = C1JCFeature()
                feas.append(fea)
                fea.name, fea.jc = y[2:], jc
                fea.value = FeaSOrN(x[y], None)

            fea = x.get("stone")
            if not fea:
                continue
            for y in fea:
                st = C1JCStone()
                sts.append(st)
                st.jc = jc
                for z in st_map.items():
                    setattr(st, z[0], y.get(z[1]))
                self._nrm_jc_st(st)
        with self.sessionctx() as cur:
            cur.add_all(jcs)
            cur.flush()
            if feas:
                cur.add_all(feas)
            if sts:
                cur.add_all(sts)
            cur.commit()
        logger.debug("%d of records persisted from file(%s)" % (len(jcs), fn))
        return jcs

    def _nrm_jc_st(self, st):
        if not st.shape:
            st.stone, st.shape = self._utilz.extract_st_sns(st.stone)
        if st.wgt is None:
            st.wgt = 0

    def _read_file(self, fn):
        app, tk = appmgr.acq()
        wb, lsts = app.books.open(fn), []
        try:
            var = self._utilz.find_sheet(wb)
            nls = self._utilz.read_his(var, True)
            if not nls:
                return None
            var = nls[0].colnames
            jo = var.index("stone")
            var, var1 = var[:jo], var[jo:]
            n_fn, l_jn, jo, jnset = self._norm_fn(fn), None, {}, set()
            # in the early age, there might be duplicated JO# inside a file
            for nl in nls:
                jn = nl.jono or l_jn
                if jn != l_jn:
                    l_jn = jn
                    # master properties
                    jo = {x: nl[x] for x in var if nl[x]}
                    jo["stone"], jo["docno"] = [], n_fn
                    if jn not in jnset:
                        jnset.add(jn)
                        lsts.append(jo)
                if not nl.stone or not nl.setting or nl.stone == self._utilz.CN_MIT:
                    continue
                jo["stone"].append({x: nl[x] for x in var1 if nl[x]})
        finally:
            if wb:
                wb.close()
            appmgr.ret(tk)
        return lsts

    def _is_persisted(self, fn):
        """ check if given file has been persisted
        @return: 1 if persited
                 0 if not persisted
                 -1 if expired
        """
        flag, fnx = 0, self._norm_fn(fn)
        with self.sessionctx() as cur:
            q = Query(func.max(C1JC.lastmodified)).filter(C1JC.docno == fnx).with_session(cur).first()
            if q[0]:
                flag = 1 if q[0] >= datetime.fromtimestamp(path.getmtime(fn)) else -1
        return flag

    def _clear_expired(self, fn):
        fn = self._norm_fn(fn)
        with self.sessionctx() as cur:
            #qm = cur.query(C1JC.id).filter(C1JC.docno == fn).subquery()
            qm = cur.query(C1JC.id).filter(C1JC.docno == fn)
            for clz in (C1JCFeature, C1JCStone,):
                cur.query(clz).filter(clz.jcid.in_(qm)).delete(synchronize_session='fetch')
            qm.delete()
            cur.commit()
