'''
#! coding=utf-8
@Author:        zmFeng
@Created at:    2018-12-18
@Last Modified: 2018-12-18 4:42:37 pm
@Modified by:   zmFeng

c1 calculation excel data filler and so on...

'''
from re import compile as cmpl

from sqlalchemy.orm import Query
from xlwings.constants import LookAt

from hnjapp.dbsvcs import idset, idsin
from hnjcore.models.hk import JO
from hnjcore.models.hk import JOItem as JI
from utilz import NamedList, karatsvc, stsizefmt, trimu
from utilz.xwu import NamedRanges, find


class Writer(object):
    """ read/write data from excel """
    _CN_JONO = "工单"
    _CN_MIT = "MIT"
    _ptn_pk = cmpl(r"([A-Z]{2})-([A-Z@])-([A-Z\d]{1,4})-([A-Z\d])")

    def __init__(self, hksvc):
        self._hksvc = hksvc
        self._nl = None

    def _read(self, sht):
        """ read JO# from source sheet, return a set of JO#s """
        rng = find(sht, self._CN_JONO, lookat=LookAt.xlWhole)
        if not rng:
            return None
        ttls, idx = sht.range(rng, rng.end("right")).value, 0
        for idx, nl in enumerate(ttls):
            if nl.find("镶法") >= 0:
                break
        alias = {
            "jono": self._CN_JONO,
            "styno": "款号",
            "spec": "参数",
            "karat0": "成色1",
            "wgt0": "金重1",
            "karat1": "成色2",
            "parts": "配件",
            "wgt1": "金重2",
            "pen": "笔电",
            "chain": "链尾",
            "tt": "分色",
            "micron": "电咪",
            "other": "其它",
            "mtl2": "银夹金",
            "stone": "石料",
            "shape": "形状",
            "stsize": "尺寸",
            "stqty": "粒数",
            "stwgt": "重量",
            "remark": "镶法"
        }
        self._nl = NamedList(ttls[:idx + 1], alias=alias)
        return {nl.jono for nl in NamedRanges(rng, alias=alias) if nl.jono}

    def _get_his(self, skuno):
        """ return the history data by SKU# """
        return []

    def _get_other(self, jo):
        """ the specified items, return specified result """
        return None

    def _karat_cvt(self, karat):
        """ karat to the excel's term """
        return karatsvc.getfamily(karat).name

    def _write(self, lsts, sht):
        rng = find(sht, self._CN_JONO, lookat=LookAt.xlWhole)
        rng.offset(1, 0).value = lsts

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
            mstr, pt = nl.data if cat == "EARRING" else None, False  #earring's Pin has special offer
            nls, nl1, var = self._extract_jis(jo, cat, nls)
        if var:
            jo.remark = ";".join(var) + ";" + jo.remark
        var = self._hksvc.getjowgts(jo)
        kt = karatsvc.getkarat(var.main.karat)
        self._mstr_BL(jo, var, kt, cat)
        var = [x for x in (var.main, var.aux) if x]
        for x, wgt in enumerate(var):
            nl["karat%d" % x] = self._karat_cvt(wgt.karat)
            nl["wgt%d" % x] = wgt.wgt
        if nls:
            pt = False
            for var in nls:
                nl1.setdata(var)
                if mstr and nl1.remark.find("耳迫") >= 0:  # big5
                    pt = True
                for x in nl1.colnames:
                    nl[x] = nl1[x]
                lsts.append(nl.newdata(True))
            if pt:
                nl.setdata(mstr).parts = {
                    925: "银迫",
                    9: "9K迫"
                }.get(kt.karat, "9KRW迫")
            return lsts[:-1]
        return []

    def _mstr_BL(self, jo, pw, kt, cat):
        """ the business logic of master record """
        nl = self._nl
        if pw.part and pw.part.wgt:
            nl.chain = "是"
        if pw.main and pw.aux:
            nl.mtl2 = 1 #at least one
            nl.pen = "是"
        if cat == "BRACELET":
            nl.spec = "手链单21" # don't know if it's 20 or 21
        elif cat == "NECKLACE":
            nl.other = (nl.other or 0) + 3 # chain cutting cost
        elif cat == "PENDANT" and [
                1 for x in "相 盒".split() if jo.description.find(x) >= 0
        ]: #big5
            nl.spec = "相盒"
        vcs = self._extract_vcs(jo.remark)
        if vcs:
            for vc in vcs:
                if vc == "_MICRON_":
                    nl.micron = self._get_micron(jo)
                elif not nl.tt and kt.color != vc:
                    nl.pen = "是"
        if sum((1 for x in "打 噴 沙 砂".split() if jo.remark.find(x) >= 0)) > 1:
            nl.other = (nl.other or 0) + 5
        if jo.description.find("套") >= 0: # big5
            nl.other = (nl.other or 0) + 3 if cat == "PENDANT" else 5 if cat == "RING" else 0

    def _get_micron(self, jo):
        """ TODO return the micron price of given JO# """
        return -1000

    @classmethod
    def _extract_vcs(cls, rmk):
        """ the vermail color, None or an list """
        idx, rc = rmk.find("電"), None
        if idx < 0:
            return rc
        ss, lst = rmk.split("電"), []
        mp = {"白": "WHITE", "黑": "BLACK", "藍": "BLUE", "黃 王": "YELLOW", "玫 瑰": "ROSE"}
        mp = {x[0]: getattr(karatsvc, "COLOR_" + trimu(x[1])) for x in mp.items()}
        mp["咪"] = "_MICRON_"
        for var in ss[1:]:
            var, rc = var[:6], None
            for x in mp.items():
                if [1 for y in x[0].split() if var.find(y) >= 0]:
                    rc = x[1]
                    break
            if rc:
                lst.append(rc)
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
        blk = lambda st, x: None if st == self._CN_MIT else x
        st = "".join([x for x in st if "A" <= x <= "Z"])
        sz, qty, wgt = [blk(st, x) for x in (ji.stsize, ji.qty, ji.wgt)]
        if st != self._CN_MIT:
            pk = self._extract_pk(ji.remark)
            if pk:
                st, shp = pk, None
                sto, shpo = st[:2], st[2]
            else:
                shp = {
                    "RD": ("DD", "R"),
                    "TD": ("DD", "T"),
                    "RZ": ("CZ", "R")
                }.get(st)
                if not shp:
                    if st:
                        st, shp = st[1:], st[0]
                else:
                    st, shp = shp[0], shp[1]
                sto, shpo = st, shp
        else:
            sto = st
        if sz == ".":
            sz = None
        nl_ji.setdata([st, shp, sz, qty, wgt, ji.remark, stsizefmt(sz), None, sto, shpo])

    def _extract_jis(self, jo, cat, jis):
        """ calc the main/side stone, sort and calc the BL """
        # the last one is mainstone sign, can be one of M/S/None
        nl_ji = NamedList("stone shape stsize stqty stwgt remark szcal ms sto shpo".split())
        sts, mns, dscs, hasds = [], [], set(), (jo.remark.find("碟") >= 0 or jo.style.name.alpha.find("M") >= 0)
        for ji in jis:
            self._extract_ji(ji, nl_ji)
            if not nl_ji.sto:
                dscs.add(nl_ji.remark)
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
            if nl_ji.sto in ("DD", "CZ"):
                dscs.add("電白")
            elif nl_ji.sto == "DF":
                dscs.add("電藍")
                hasds = True
        if hasds:
            sts.append(nl_ji.newdata(True))
            pk = nl_ji.getcol("sto")
            pk, hasds = [x for x in sts if x[pk] in ("DD", "DF")], "stone shape sto shpo stsize stqty".split()
            if pk:
                pk = [pk[0][nl_ji.getcol(x)] for x in hasds]
                pk[-1] = -pk[-1]
            else:
                pk = ["DF", "R", "DF", "R", None, -1]
            hasds.append("remark")
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
            nl_ji.remark = self._calc_stset(cat, nl_ji) or nl_ji.remark
        mns = {"M": "0", "S": "1"}
        def srt_key(data):
            nl_ji.setdata(data)
            return "%s,%s,%s,%s" % (mns.get(nl_ji.ms, "Z"), nl_ji.sto, nl_ji.shpo, nl_ji.szcal)
        sts, dscs = sorted(sts, key=srt_key), tuple(dscs) if dscs else []
        pk, ji = [nl_ji.getcol(x) for x in "szcal remark".split()]
        return [x[:pk] for x in sts], NamedList(nl_ji.colnames[:ji + 1]), dscs

    def _calc_stset(self, cat, nl):
        """ calculate the setting by given arguments """
        if not (nl.stqty and nl.stone):
            return None
        if nl.stqty < 0:
            return None
        rc = None
        st, shp, qty, sz = [nl[x] for x in "sto shpo stqty szcal".split()]
        fixed = {"JD": "玉", "ON": "安力士", "PL": "珠"}
        if st in fixed:
            rc = fixed[st]
        elif st == "DD" and shp == "R":
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
        for sht in wb.sheets:
            rng = find(sht, self._CN_JONO, lookat=LookAt.xlWhole)
            if not rng:
                continue
            nls = self._read(sht)
            if nls:
                nls = self._fetch_data(nls)
                self._write(nls, sht)
