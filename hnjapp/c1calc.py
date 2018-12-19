'''
#! coding=utf-8
@Author:        zmFeng
@Created at:    2018-12-18
@Last Modified: 2018-12-18 4:42:37 pm
@Modified by:   zmFeng

c1 calculation excel data filler and so on...

'''
from xlwings.constants import LookAt
from sqlalchemy.orm import Query

from utilz import NamedList, karatsvc, trimu
from utilz.xwu import NamedRanges, appmgr, find
from hnjcore.models.hk import JOItem as JI, JO
from hnjapp.dbsvcs import jesin, idsin, idset

class Writer(object):
    """ read/write data from excel """
    _CN_JONO = "工单"
    _CN_MIT = "MIT"

    def __init__(self, hksvc):
        self._hksvc = hksvc
        self._nl = None

    def _read(self, sht):
        """ read JO# from source sheet, return a set of JO#s """
        rng = find(sht, self._CN_JONO, lookat=LookAt.xlWhole)
        if not rng:
            return None
        ttls = sht.range(rng, rng.end("right")).value
        for idx, nl in enumerate(ttls):
            if nl.find("镶法") >= 0:
                break
        alias = {
            "jono": self._CN_JONO,
            "styno": "款号",
            "karat0": "成色1",
            "wgt0": "金重1",
            "karat2": "成色2",
            "parts": "配件",
            "wgt1": "金重2",
            "pen": "笔电",
            "chain": "链尾",
            "tt": "分色",
            "mic": "电咪",
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
        return None

    def _write(self, lsts, sht):
        rng = find(sht, self._CN_JONO, lookat=LookAt.xlWhole)
        rng.offset(1, 0).value = lsts

    def _from_his(self, nl, his, **kwds):
        """ return a list of list with filled data from history """
        return None

    def _karat_cvt(self, karat):
        """ karat to the excel's term """
        return karatsvc.getkarat(karat).name
    
    def _fmt_st(self, pw):
        """ format a full row of stone data """
        st, shp = trimu(pw.stname), None
        st = "".join([x for x in st if "A" <= x <= "Z"])
        blk = lambda x: None if st == self._CN_MIT else x
        if st != self._CN_MIT:
            shp = {"RD": ("DD", "R"), "TD": ("DD", "T"), "RZ": ("CZ", "R")}.get(st)
            if not shp:
                st, shp = st[:2], st[2:]
            else:
                st, shp = shp[0], shp[1]
        sz, qty, wgt = [blk(x) for x in (pw.stsize, pw.qty, pw.wgt)]
        return (st, shp, sz, qty, wgt, pw.remark)

    def _from_jo(self, nl, jo, **kwds):
        """
        return a list of list with filled data from JO, don't need to return
        the first list, because it's already in the master.
        Some BLogics will be performed during data filling
        """
        pw = self._hksvc.getjowgts(jo)
        if pw.part and pw.part.wgt:
            nl.chain = "是"
        pw = [x for x in (pw.main, pw.aux) if x]
        for idx, wgt in enumerate(pw):
            nl["karat%d" % idx] = self._karat_cvt(wgt.karat)
            nl["wgt%d" % idx] = wgt.wgt
        if len(pw) > 1:
            nl.mtl2 = 1
        kt = karatsvc.getkarat(pw[0].karat)
        # plating white, not on white metal
        if kt.color != "WHITE" and [1 for x in "鑽 電白".split() if jo.remark.find(x) >= 0]:
            nl.pen = "是"
        # Micron, wait for result
        if [1 for x in "咪".split() if jo.remark.find(x) >= 0]:
            nl.micro = -1000
        lsts, dtls = [], kwds.get("dtls")
        if dtls:
            # sort the details by name before operation, MIT always at the end
            mstr, pt = nl.data if self._is_er(jo.style.name) else None, False #earring's Pin has special offer
            nl1 = NamedList("stone shape stsize stqty stwgt remark".split())
            dtls = sorted([self._fmt_st(x) for x in dtls], key=lambda x: "%s,%s,%s" % ("a" if x[0] == self._CN_MIT else x[0], x[1], x[2]))
            for pw in dtls:
                nl1.setdata(pw)
                if mstr and nl1.remark.find("耳拍") >= 0: # big5
                    pt = True
                for x in nl1.colnames:
                    nl[x] = nl1[x]
                lsts.append(nl.newdata(True))
            if pt:
                nl.setdata(mstr).parts = {925: "银迫", 9: "9K迫"}.get(kt.id, "9KRW迫")
            return lsts[:-1]
        return []

    @classmethod
    def _is_er(cls, se):
        """ is earring """
        return se.alpha.find("E") >= 0

    def _db_fetch(self, jns):
        """ fetch data from db to an list"""
        with self._hksvc.sessionctx() as cur:
            jos, jerrs = self._hksvc.getjos(jns)
            dtls = Query((JO.id, JI.stname, JI.stsize, JI.stsize, JI.qty, JI.unitwgt.label("wgt"), JI.remark, )).join(JI)
            dtls, lsts = dtls.filter(idsin(idset(jos), JO)).with_session(cur).all(), {}
            for jo in dtls:
                lsts.setdefault(jo.id, []).append(jo)
            dtls, lsts, nl = lsts, [], self._nl
            ftrs = {True: self._from_his, False: self._from_jo}
            for jo in jos:
                lsts.append(nl.newdata(True))
                nl.jono, nl.styno = "'" + jo.name.value, jo.style.name.value
                pw = self._get_his(jo.po.skuno)
                lsts.extend(ftrs[bool(pw)](nl, pw or jo, dtls=dtls.get(jo.id)))
            for jo in jerrs:
                lsts.extend(nl.newdata(True))
                nl.jono = jo
        return lsts

    def run(self, wb):
        for sht in wb.sheets:
            rng = find(sht, self._CN_JONO, lookat=LookAt.xlWhole)
            if not rng:
                continue
            nls = self._read(sht)
            if nls:
                nls = self._db_fetch(nls)
                self._write(nls, sht)
