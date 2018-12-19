# coding=utf-8
'''
Created on Apr 17, 2018

the replacement of the Paj Shipment Invoice Reader, which was implmented
in PAJQuickCost.xls#InvMatcher

@author: zmFeng
'''

import numbers
import random
import re
from collections import namedtuple
from datetime import date, datetime
from decimal import Decimal
from logging import DEBUG
from os import path

from sqlalchemy import and_, func
from sqlalchemy.orm import Query
from xlwings import Book
from xlwings.constants import LookAt

from hnjcore import JOElement, isvalidp17
from hnjcore.models.hk import JO as JOhk
from hnjcore.models.hk import Orderma, PajCnRev, PajInv, PajShp
from hnjcore.models.hk import Style as Styhk
from hnjcore.utils.consts import NA
from utilz import (NamedList, NamedLists, ResourceCtx, SessionMgr, splitarray,
                   triml, daterange, getfiles, isnumeric, appathsep, deepget,
                   karatsvc, xwu, getvalue)


from .common import _getdefkarat
from .common import _logger as logger
from .localstore import PajCnRev as PajCnRevSt
from .localstore import PajInv as PajInvSt
from .localstore import PajItem as PajItemSt
from .localstore import PajWgt as PrdWgtSt
from .pajcc import (MPS, PAJCHINAMPS, P17Decoder, PajCalc, PrdWgt, WgtInfo,
                    _tofloat, addwgt)
from .dbsvcs import HKSvc

_accdfmt = "%Y-%m-%d %H:%M:%S"
_appmgr = xwu.appmgr


def _accdstr(dt):
    """ make a date into an access date """
    return dt.strftime(_accdfmt) if dt and isinstance(dt, date) else dt


def _removenonascii(s0):
    """remove thos non ascii characters from given string"""
    if isinstance(s0, str):
        return "".join(
            [x for x in s0 if ord(x) > 31 and ord(x) < 127 and x != "?"])
    return s0


class PajBomHhdlr(object):
    """ class to read BOMs from PAJ
    @param part_chk_ver: the Part checker version,
        default is None or 0,
            That is, when there is (chain with length) and (lock exists),
            圈 will be treated as part of the chain
        1 stands for loose,
            That is, when there is chain with length,
            圈 will be treated as part of the chain
    """
    _ptn_oz = re.compile(r"\(\$\d*/OZ\)")
    _one_hit_mp = {
        925: (re.compile(r"(925)|(银)"), ),
        200: (re.compile(r"(BRONZE)|(铜)", re.I), ),
        9925: (re.compile(r"BONDED", re.I), re.compile(r"B&Gold", re.I))
    }
    _ptn_k_gold = re.compile(r"^(\d*)K")
    _ptn_digits = re.compile(r"[\(（](\d*)[\)）]")
    _ptn_chn_lck = re.compile(r"(弹簧扣)|(龙虾扣)|(狗仔头扣)")
    # the parts must have karat, if not, follow the parent
    _mtl_parts = u"金 银 耳勾 线圈 耳针 耳束 Chain".lower().split()
    # keywords for parts (that should belong to a chain)
    _pts_kws = "圈 牌".split()
    # belts or so, not metal
    _voids = "色 带 胶".split()
    _pcdec = P17Decoder()
    _part_chk_ver = None

    _nmps = {
        "mstr": {
            u"pcode": "十七位,",
            "mat": u"材质,",
            "mtlwgt": u"抛光,",
            "up": "单价",
            "fwgt": "成品重"
        },
        "parts": {
            "pcode": u"十七位,",
            "matid": "物料ID,",
            "name": u"物料名称",
            "spec": u"物料特征",
            "qty": u"数量",
            "wgt": u"重量",
            "unit": u"单位",
            "length": u"长度"
        }
    }

    def __init__(self, **kwargs):
        self._part_chk_ver = getvalue(kwargs, "part_chk_ver")

    def _parse_karat(self, mat, wis=None, ispol=True):
        """ return karat(int type) from material string """
        kt, karat = None, None
        if ispol:
            mt = self._ptn_oz.search(mat)
            if not mt:
                # no /oz sign
                for x in (j for y in (200, 9925) for j in self._one_hit_mp[y]):
                    karat = x.search(mat)
                    if karat:
                        break
                if not karat:
                    return None
        kt = max(x[0] if y.search(mat) else 0 for x in self._one_hit_mp.items() for y in x[1])
        if not kt:
            mt = self._ptn_k_gold.search(mat) or self._ptn_digits.search(mat)
            if mt:
                kt = int(mt.group(1))
        if not kt:
            # not found, has must have keyword? if yes, follow master
            voids = [1 for x in self._voids if mat.find(x) >= 0]
            if not voids and wis and any(wis):
                s0 = mat.lower()
                for x in self._mtl_parts:
                    if s0.find(x) < 0:
                        continue
                    if s0.find(u"金") >= 0:
                        for wi in (x for x in wis if x):
                            karat = karatsvc.getkarat(wi.karat)
                            if karat and karat.category == karatsvc.CATEGORY_GOLD:
                                return wi.karat
                    # finally no one is found, follow master
                    # kt = wis[0].karat

                    # but zhangyuting claimed in e-mail with title "配件的"物料名称"里没有金" on 2018/12/10 that the karat should be 925
                    # so let it to be 925
                    kt = 925
                    if kt:
                        break
            if not kt and wis:
                if logger.isEnabledFor(DEBUG) and not voids:
                    logger.error(
                        "No karat found for (%s) and no default provided" %
                        mat)
        if kt:
            karat = karatsvc.getkarat(kt) or karatsvc.getbyfineness(kt)
            kt = karat.karat if karat else None
        return kt

    def _ispendant(self, pcode):
        return self._pcdec.decode(pcode, "PRODTYPE").find("吊") >= 0

    def _isring(self, pcode):
        return self._pcdec.decode(pcode, "PRODTYPE").find("戒") >= 0

    def readbom(self, fldr):
        """
        read BOM from given folder
        @param fldr: the folder contains the BOM file(s)
        return a dict with "pcode" as key and dict as items
            the item dict has keys("pcode","mtlwgt")
        """
        pmap = {}
        if isinstance(fldr, Book):
            self._read_book(fldr, pmap)
        else:
            fns = getfiles(fldr, "xls") if path.isdir(fldr) else (fldr,)
            if not fns:
                return
            app, kxl = _appmgr.acq()
            try:
                for fn in fns:
                    wb = app.books.open(fn)
                    self._read_book(wb, pmap)
                    wb.close()
            finally:
                if kxl and app:
                    _appmgr.ret(kxl)

        self._adjust_wgts(pmap)

        return pmap

    def _part_chk(self, y, chns, lks, has_semi_chn):
        """
        determine if the given y is a part
        """
        nm = y["name"]
        return chns and y["_id"] in chns or lks and (lks.get(y["_id"]) or sum([1 for v in self._pts_kws if nm.find(v) >= 0]))

    def _part_chk_l(self, y, chns, lks, has_semi_chn):
        """
        loose rule for determining if the given y is a part
        """
        nm = y["name"]
        return chns and y["_id"] in chns or\
            lks and lks.get(y["_id"]) or\
            has_semi_chn and [1 for v in self._pts_kws if nm.find(v) >= 0]

    def _adjust_wgts(self, pmap):
        if not pmap:
            return
        part_ck = self._part_chk if not self._part_chk_ver else self._part_chk_l
        for pcode, prop in pmap.items():
            ch_lks, prdwgt = prop.get("mtlwgt"), None
            if ch_lks:
                for y in ch_lks:
                    prdwgt = addwgt(prdwgt, WgtInfo(y[0], y[1]))
            else:
                logger.debug("%s does not have master weight" % pcode)
                prdwgt = PrdWgt(WgtInfo(0, 0))
            if "parts" not in prop:
                prop["mtlwgt"] = prdwgt._replace(netwgt=prop.get("netwgt"))
                continue
            ch_lks = {}
            if self._ispendant(pcode):
                for y in prop["parts"][::]:
                    nm = y["name"]
                    if triml(nm).find("chain") >= 0:
                        ch_lks.setdefault("chain", {})[y["_id"]] = y
                    elif self._ptn_chn_lck.search(nm):
                        ch_lks.setdefault("lock", {})[y["_id"]] = y
            chns, lks = tuple(ch_lks.get(x) or {} for x in "chain lock".split())
            has_semi_chn, subs = self._has_semi_chn(chns), 0
            for y in prop["parts"]:
                nm = y["name"]
                kt = self._parse_karat(nm, prdwgt.wgts, False)
                if not kt:
                    subs += y["wgt"]
                    continue
                y["karat"], is_part = kt, False
                is_part = part_ck(y, chns, lks, has_semi_chn)
                if is_part:
                    #make sure part candidate has the same karat with old
                    wgt0 = prdwgt.part
                    is_part = not wgt0 or kt == wgt0.karat
                prdwgt = addwgt(prdwgt, WgtInfo(kt, y["wgt"]),\
                    is_part, autoswap=False)
            if has_semi_chn:
                prdwgt = prdwgt._replace(part=WgtInfo(prdwgt.part.karat, -prdwgt.part.wgt * 100))
            prop["mtlwgt"] = prdwgt._replace(netwgt=round(prop.get("netwgt", 0) - subs, 2))

    def _has_semi_chn(self, chns):
        """
        check if the chains contains semi-chain, that is, 成品链
        """
        lc = 0
        for y in chns.values():
            lc = y["length"]
            if lc is None:
                continue
            try:
                lc = float(lc)
                break
            except:
                pass
        return lc

    def _read_book(self, wb, pmap):
        """
        read bom in the given wb to pmap
        """
        shts, bg_sht = [[], []], None
        for sht in wb.sheets:
            rng = xwu.find(sht, u"十七位")
            if not rng:
                continue
            if xwu.find(sht, u"抛光后"):
                shts[0] = (sht, rng)
            elif xwu.find(sht, u"物料特征"):
                shts[1] = (sht, rng)
            else:
                if xwu.find(sht, u"录入日期"):
                    bg_sht = sht
            if all(shts) and bg_sht:
                break
        if not all(shts):
            return
        if bg_sht:
            self._append_bd(bg_sht, shts[0][0])
        for sht, rdr in {shts[0]: self._read_mstr, shts[1]: self._read_pts}.items():
            rdr(sht, pmap)

    def _get_data(self, sht_rng, nmp):
        vvs = sht_rng[1].end("left").expand("table").value
        return NamedLists(vvs, nmp)

    def _read_mstr(self, sht_rng, pmap):
        """ read the bom master to pmap(dict) """
        sht_rng[0].name = "BOM_mstr"
        nl, mstrs, netwgts = self._nmps["mstr"], set(), {}
        for nl in self._get_data(sht_rng, nl):
            pcode = nl.pcode
            if not isvalidp17(pcode):
                break
            #dup check
            key = tuple(nl[x] or 0 for x in "pcode mat up mtlwgt fwgt".split())
            key = ("%s" * len(key)) % key
            if key in mstrs:
                logger.debug("duplicated bom_mstr found(%s, %s)" %
                                (nl.pcode, nl.mat))
                continue
            kt = nl.fwgt
            if kt:
                netwgts[pcode] = netwgts.get(pcode, 0) + kt
            kt = self._parse_karat(nl.mat)
            if not kt:
                continue
            mstrs.add(key)
            it = pmap.setdefault(pcode, {"pcode": pcode})
            it.setdefault("mtlwgt", []).append((kt, nl.mtlwgt))
        for pcode, kt in netwgts.items():
            pmap[pcode]["netwgt"] = round(kt, 2)

    def _read_pts(self, sht_rng, pmap):
        """ read parts from the sheet to pmap(dict) """
        sht_rng[0].name, pts = "BOM_part", set()
        nmp = [x for x in self._nmps["parts"] if x.find("pcode") < 0]
        _mat_id = lambda x: "%s,%f" % (x.matid, x.wgt or 0)
        for nl in self._get_data(sht_rng, self._nmps["parts"]):
            pcode = nl.pcode
            if not isvalidp17(pcode):
                break
            key = tuple(nl[x] or 0 for x in "pcode matid name spec qty wgt unit length".split())
            key = ("%s" * len(key)) % key
            if key in pts:
                logger.debug("duplicated bom_part found(%s, %s)" %
                                (nl.pcode, nl.name))
                continue
            pts.add(key)
            it = pmap.setdefault(pcode, {"pcode": pcode})
            mats, it = it.setdefault("parts", []), {}
            mats.append(it)
            for cn in nmp:
                it[cn] = nl[cn]
            it["_id"] = _mat_id(nl)
            if not it["wgt"]:
                it["wgt"] = 0

    def _append_bd(self, bg_sht, mstr_sht):
        """ append the single_bonded_gold sheet to bom-mstr sheet """
        bgs = xwu.NamedRanges(
            xwu.usedrange(bg_sht),
            name_map={
                "pcode": "十七,",
                "mtlwgt": "金银重,",
                "stwgt": "石头,"
            })
        nls = [
            x for x in xwu.NamedRanges(
                xwu.usedrange(mstr_sht), name_map=self._nmps["mstr"])
        ]
        nl, ridx = nls[0], len(nls)
        if isvalidp17(nls[-1].pcode):
            ridx += 1
        for bg in (x for x in bgs if x.pcode):
            vals = (bg.pcode, "BondedGold($0/OZ)", bg.mtlwgt or 0,
                    (bg.mtlwgt or 0) + (bg.stwgt or 0))
            for x in zip("pcode,mat,mtlwgt,fwgt".split(","), vals):
                mstr_sht[ridx, nl.getcol(x[0])].value = x[1]
            ridx += 1
        #bg_sht.name = "BG.Wgt"
        bg_sht.delete()

    def readbom2jos(self, fldr, hksvc, fn=None, mindt=None):
        """ build a jo collection list based on the BOM file provided
            @param fldr: the folder contains the BOM file(s)
            @param hksvc: the HK db service
            @param fn: save the file to
            @param mindt: the minimum datetime the query fetch until
            if None is provided, it will be 2017/01/01
            return a workbook contains the result
        """

        def _fmtwgt(prdwgt):
            wgt = (prdwgt.main, prdwgt.aux, prdwgt.part)
            lst = []
            [lst.extend((x.karat, x.wgt) if x else (0, 0)) for x in wgt]
            return lst

        def _samewgt(wgt0, wgt1):
            wis = []
            for x in (wgt0, wgt1):
                wis.append((x.main, x.aux, x.part))
            for i in range(3):
                wts = (wis[0][i], wis[1][i])
                eq = all(wts) or not any(wts)
                if not eq:
                    break
                if not all(wts):
                    continue
                eq = wts[0].karat == wts[0].karat or \
                    karatsvc.getfamily(wts[0].karat) == karatsvc.getfamily(wts[1].karat)
                if not eq:
                    break
                eq = abs(round(wis[0][i].wgt - wis[1][i].wgt, 2)) <= 0.02
            return eq

        pmap = self.readbom(fldr)
        ffn = None
        if not pmap:
            return ffn
        vvs = ["pcode,m.karat,m.wgt,p.karat,p.wgt,c.karat,c.wgt".split(",")]
        jos = [
            "Ref.pcode,JO#,Sty#,Run#,m.karat,m.wgt,p.karat,p.wgt,c.karat,c.wgt,rm.wgt,rp.wgt,rc.wgt"
            .split(",")
        ]
        if not mindt:
            mindt = datetime(2017, 1, 1)
        qp = Query(Styhk.id).join(Orderma, Orderma.styid == Styhk.id) \
            .join(JOhk, Orderma.id == JOhk.orderid).join(PajShp, PajShp.joid == JOhk.id)
        qj = Query([JOhk.name.label("jono"), Styhk.name.label("styno"), JOhk.running]) \
            .join(Orderma, Orderma.id == JOhk.orderid).join(Styhk).filter(JOhk.createdate >= mindt) \
            .order_by(JOhk.createdate)
        with hksvc.sessionctx() as sess:
            cnt, ln = 0, len(pmap)
            for x in pmap.values():
                lst, wgt = [x["pcode"]], x["mtlwgt"]
                if isinstance(wgt, PrdWgt):
                    lst.extend(_fmtwgt((wgt)))
                else:
                    lst.extend((0, 0, 0, 0, 0, 0))
                vvs.append(lst)

                pcode = x["pcode"]
                q = qp.filter(PajShp.pcode == pcode).limit(1).with_session(sess)
                try:
                    sid = q.one().id
                    q = qj.filter(Orderma.styid == sid).with_session(sess)
                    lst1 = q.all()
                    for jn in lst1:
                        jowgt = hksvc.getjowgts(jn.jono)
                        if not _samewgt(jowgt, wgt):
                            lst = [
                                pcode, jn.jono.value, jn.styno.value, jn.running
                            ]
                            lst.extend(_fmtwgt(jowgt))
                            lst.extend(_fmtwgt(wgt)[1::2])
                            jos.append(lst)
                        else:
                            logger.debug("JO(%s) has same weight as pcode(%s)" %
                                         (jn.jono.value, pcode))
                except:
                    pass

                cnt += 1
                if cnt % 20 == 0:
                    print("%d of %d done" % (cnt, ln))

            app, kxl = _appmgr.acq()
            wb = app.books.add()
            sns, data = "BOMData,JOs".split(","), (vvs, jos)
            for cnt, pcode in enumerate(sns):
                sht = wb.sheets[cnt]
                sht.name = pcode
                sht.range(1, 1).value = data[cnt]
                sht.autofit("c")
            wb.save(fn)
            ffn = wb.fullname
            _appmgr.ret(kxl)
        return ffn


class PajShpHdlr(object):
    """
    Tasks:
        .integrating PajShpRdr/PajInvRdr into one to generate data from HKdb
        .genereate shipment data for py.mm/py.bc
    """

    def __init__(self, hksvc):
        self._hksvc = hksvc

    @classmethod
    def get_shp_date(cls, fn, isfile=True):
        """
        extract the shipdate from file name
        """
        import datetime as dt
        ptnfd = re.compile(r"\d+")
        parts = ptnfd.findall(path.basename(fn))
        if not parts:
            return

        try:
            d0 = None
            parts = [int(x) for x in parts]
            lg = len(parts)
            if lg >= 3:
                d0 = dt.date(parts[0], parts[1], parts[2])
                # mm-ddxxxx_01.xlsx case, the first not year
                if d0.year < 2017:
                    parts = (parts[0], parts[1])
                    lg = 2
                    d0 = None
            if lg == 2 and isfile:
                # only month/date,guess the year
                d1 = dt.date.fromtimestamp(path.getmtime(fn))
                d0 = dt.date(d1.year, parts[0], parts[1])
                df = d1 - d0
                if df.days < -3:
                    d0 = dt.date(d0.year - 1, d0.month, d0.day)
        except:
            d0 = None
        return d0

    @classmethod
    def _getfmd(cls, fn):
        return datetime.fromtimestamp(
            path.getmtime(fn)).replace(microsecond=0).replace(second=0)

    @classmethod
    def _readquodata(cls, sht, qmap):
        """extract gold/stone weight data from the QUOXX sheet
        @param sht:  the DL_QUOTATION sheet that need to read data from
        @param qmap: the dict with p17 as key and (goldwgt,stwgt) as value
        """
        rng = xwu.find(sht, "Item*No", lookat=LookAt.xlPart)
        if not rng:
            return
        # because there is merged cells rng.expand('table').value
        # or sht.range(rng.end('right'),rng.end('down')).value failed
        _ptngwt = re.compile(r"[\d.]+")
        vvs = sht.range(rng, rng.current_region.last_cell).value
        nls = NamedLists(vvs, {
            "pcode": "Item,",
            "stone": "stone,",
            "metal": "metal ,"
        }, False)
        for tr in nls:
            p17 = tr.pcode
            if not p17:
                continue
            if isvalidp17(p17) and p17 not in qmap:
                sw = 0 if not tr.stone else \
                    sum([float(x)
                         for x in _ptngwt.findall(tr.stone)])
                mtls = tr.metal
                if isinstance(mtls, numbers.Number):
                    mw = (WgtInfo(0, mtls),)
                else:
                    s0, mw = tr.metal.replace("：", ":"), []
                    if s0.find(":") > 0:
                        for x in s0.split("\n"):
                            ss = x.split(":")
                            mt = _ptngwt.search(ss[0])
                            karat = 925 if not mt else int(mt.group())
                            mt = _ptngwt.search(ss[1])
                            wgt = float(mt.group()) if mt else 0
                            mw.append(WgtInfo(karat, wgt))
                    else:
                        mt = _ptngwt.search(s0)
                        mw.append(WgtInfo(0, float(mt.group()) if mt else None))
                qmap[p17] = (mw, sw)

    def _hasread(self, fmd, fn, invno=None):
        """
            check if given file(in shpment case) or inv#(in invoice case) has been read
            @param fn: the full-path filename
            return:
                1 if hasread and up to date
                2 if expired
                0 if not read
        """
        rc = 0
        if not invno:
            fn = _removenonascii(path.basename(fn))
            with self._hksvc.sessionctx() as cur:
                shp = Query([PajShp.fn,
                             func.min(PajShp.lastmodified)]).group_by(
                                 PajShp.fn).filter(
                                     PajShp.fn == fn).with_session(cur).first()
                if shp:
                    rc = 2 if shp[1] < fmd else 1
        else:
            with self._hksvc.sessionctx() as cur:
                inv = Query(
                    [PajInv.invno, func.min(PajInv.lastmodified)]).group_by(
                        PajInv.invno).filter(
                            PajInv.invno == invno).with_session(cur).first()
                if inv:
                    rc = 2 if inv[1] < fmd else 1
        return rc

    @classmethod
    def read_inv_raw(cls, sht, invno=None, fmd=None):
        """
        read the invoice, return a map with inv#+jo# as key and PajInvItem as item
        """
        PajInvItem = namedtuple(
            "PajInvItem", "invno,pcode,jono,qty,uprice,mps,stspec,lastmodified")
        mp = {}
        rng = xwu.find(sht, "Item*No", lookat=LookAt.xlWhole)
        if not rng:
            return None
        if not invno:
            invno = cls.read_invno(sht)
        if sht.name != invno:
            sht.name = invno
        rng = rng.expand("table")
        nls = tuple(
            NamedLists(
                rng.value, {
                    "pcode": "item,",
                    "gold": "gold,",
                    "silver": "silver,",
                    "jono": u"job#,工单",
                    "uprice": "price,",
                    "qty": "unit,",
                    "stspec": "stone,"
                }))
        if not nls:
            return None
        th = nls[0]
        x = [x for x in "uprice,qty,stspec".split(",") if not th.getcol(x)]
        if x:
            logger.debug("key columns(%s) missing in invoice sheet('%s')" %
                         (x, sht.name))
            return None
        for tr in nls:
            if not tr.uprice:
                continue
            p17 = tr.pcode
            if not (isvalidp17(p17) and
                    not tuple(1 for y in "qty,uprice,silver,gold".split(",")
                              if not isnumeric(tr[y]))):
                logger.debug(
                    "invalid p17 code(%s) or wgt/qty/uprice data in invoice sheet(%s)"
                    % (p17, invno))
                continue
            jn = JOElement(tr.jono).value if th.getcol("jono") else None
            if not jn:
                logger.debug("No JO# found for p17(%s)" % p17)
                continue
            key = invno + "," + jn
            if key in mp.keys():
                it = mp[key]
                mp[key] = it._replace(qty=it.qty + tr.qty)
            else:
                mps = MPS("S=%3.2f;G=%3.2f" % (tr.silver, tr.gold)).value \
                    if th.getcol("gold") and th.getcol("silver") else "S=0;G=0"
                it = PajInvItem(invno, p17, jn, tr.qty, tr.uprice, mps,
                                tr.stspec, fmd)
                mp[key] = it
        return mp

    @classmethod
    def read_invno(cls, sht):
        """ get the inv# inside the sheet(if there is) """
        rng = xwu.find(sht, "Inv*No:")
        return rng.offset(0, 1).value if rng else None

    def _readinv(self, fn, sht, fmd):
        """
        read files back, instead of using os.walk(root), use os.listdir()
        @param invfldr: the folder contains the invoices
        """

        invno, dups = self.read_invno(sht), []
        idx = self._hasread(fmd, fn, invno)
        if idx == 1:
            return
        elif idx == 2:
            dups.append(invno)
        items = self.read_inv_raw(sht, invno, fmd)
        return items, dups

    @classmethod
    def read_shp(cls, fn, fshd, fmd, sht, bomwgts=None):
        """
        @param fshd: the shipdate extracted by the file name
        @param fmd: the last-modified date
        @param fn: the full-path filename
        """

        vvs = xwu.usedrange(sht).value
        if not vvs:
            return None
        PajShpItem = namedtuple(
            "PajShpItem",
            "fn,orderno,jono,qty,pcode,invno,invdate,mtlwgt,stwgt,shpdate,lastmodified,filldate"
        )

        def _extring(x):
            return x[:8] + x[10:]

        items, td0 = {}, datetime.today()
        shd = {
            "odx": u"订单号",
            "invdate": u"发票日期",
            "odseq": u"订单,序号",
            "stwgt": u"平均单件石头,XXX",
            "invno": u"发票号",
            "orderno": u"订单号序号",
            "pcode": u"十七位,十七,物料",
            "mtlwgt": u"平均单件金,XX",
            "jono": u"工单,job",
            "qty": u"数量",
            "cost": u"cost"
        }
        nls = tuple(NamedLists(vvs, shd))
        th, bfn = nls[0], "invno,pcode,jono,qty,invdate".split(",")
        x = [x for x in bfn if th.getcol(x) is None]
        if x:
            if len(x) / len(bfn) < 0.5:
                bfn = "工作表(%s)漏掉关键列:(%s)" % (sht.name, [shd[x] for x in x])
                logger.debug(bfn)
                return {"_ERROR_": bfn}
            return None
        for x in nls:
            x.jono = JOElement(x.jono).value

        _is_ring = lambda pcode: isvalidp17(pcode) and pcode[1] == "4"

        def _getbomwgt(bomap, bomapring, pcode):
            """ in the case of ring, there is only one code there
            """
            if not (bomap and pcode):
                return None
            prdwgt = bomap.get(pcode)
            if not prdwgt:  # and is ring
                if bomapring and _is_ring(pcode):
                    pcode0 = pcode
                    pcode = _extring(pcode)
                    prdwgt = bomapring.get(pcode)
                    pcode = pcode0
            if not prdwgt:
                logger.debug("failed to get bom wgt for pcode(%s)" % pcode)
            return prdwgt

        def _str2date(s_date):
            if isinstance(s_date, str):
                s_date = datetime.strptime(s_date, "%Y-%m-%d").date()
            return s_date

        bfn = path.basename(fn).replace("_", "")
        shd = PajShpHdlr.get_shp_date(sht.name, False) or fshd
        # when sheet's shpdate differs from file's shpdate, use the maximum one
        shd = max(shd, fshd)
        if bomwgts is None:
            bomwgts = PajBomHhdlr().readbom(sht.book)
        if bomwgts:
            bomwgtsrng = {
                _extring(x[0]): x[1]["mtlwgt"]
                for x in bomwgts.items()
                if _is_ring(x[0])
            }
            bomwgts = {x[0]: x[1]["mtlwgt"] for x in bomwgts.items()}
        else:
            bomwgtsrng, bomwgts = (None,) * 2
        if not th.getcol("cost"):
            for tr in nls:
                if not tr.pcode:
                    break
                if not tr.odseq or tr.odseq[:2] == "CR" or not isvalidp17(
                        tr.pcode):
                    logger.debug("repairing(%s) item found, skipped", tr.pcode)
                    continue
                jono = tr.jono
                mwgt = _getbomwgt(bomwgts, bomwgtsrng, tr.pcode)
                bomsrc = bool(mwgt)
                if not bomsrc:
                    mwgt, bomsrc = tr.get("mtlwgt", 0), False
                    mwgt = PrdWgt(WgtInfo(_getdefkarat(jono), mwgt, 4))
                invno = tr.invno or "N/A"
                if th.getcol('orderno'):
                    odno = tr.orderno
                elif len(
                    [x for x in "odx,odseq".split(",") if th.getcol(x)]) == 2:
                    odno = tr.odx + "-" + tr.odseq
                else:
                    odno = "N/A"
                stwgt = tr.get("stwgt")
                if stwgt is None or isinstance(stwgt, str):
                    stwgt = 0
                thekey = "%s,%s,%s" % (jono, tr.pcode, invno)
                if thekey in items:
                    # order item's weight does not have karat event, so append it to
                    # the main, but bom source case, no weight-recalculation is neeeded
                    si = items[thekey]
                    wi = si.mtlwgt
                    if not bomsrc:
                        wi = wi._replace(
                            main=wi.main._replace(
                                wgt=_tofloat((wi.main.wgt * si.qty +
                                              mwgt.main.wgt * tr.qty) /
                                             (si.qty + tr.qty), 4)))
                    items[thekey] = si._replace(qty=si.qty + tr.qty, mtlwgt=wi)
                else:
                    ivd = _str2date(tr.invdate)
                    si = PajShpItem(bfn, odno, jono, tr.qty, tr.pcode, invno,
                                    ivd, mwgt, stwgt, ivd, fmd, td0)
                    items[thekey] = si
        else:
            # new sample case, extract weight data from the quo sheet, but deprecated
            # get from bom instead
            """
            if not qmap:
                wb, qmap = sht.book, {}
                for x in [xx for xx in wb.sheets if xx.api.Visible == -1 and xx.name.lower().find('dl_quotation') >= 0]:
                    PajShpHdlr._readquodata(x, qmap)
            """
            for tr in nls:
                # no cost item means repairing
                if not tr.get("cost"):
                    continue
                p17 = tr.pcode
                if not p17:
                    break
                ivd, odno = _str2date(tr.invdate), tr.get("orderno", NA)
                prdwgt = _getbomwgt(bomwgts, bomwgtsrng, p17)
                if not prdwgt:
                    """
                    mtl_stone = qmap[p17] if p17 in qmap else ((None, ), 0)
                    wis = list(mtl_stone[0])
                    for idx in range(len(wis)):
                        wi = wis[idx]
                        if not wi:
                            continue
                        if not wi.karat:
                            wis[idx] = wi._replace(karat=_getdefkarat(tr.jono))
                    prdwgt = PrdWgt(*wis)
                    """
                    prdwgt = PrdWgt(WgtInfo(0, 0))
                mtl_stone = (0, 0)
                si = PajShpItem(bfn, odno,
                                JOElement(tr.jono).value, tr.qty, p17, tr.invno,
                                ivd, prdwgt, mtl_stone[1], ivd, fmd, td0)
                # new sample won't have duplicated items
                items[random.random()] = si
        return items

    def _persist(self, shps, invs):
        """save the data to db
        @param dups: a list contains file names that need to be removed
        @param items: all the ShipItems that need to be persisted
        """

        err = True
        with self._hksvc.sessionctx() as sess:
            if shps[0]:
                sess.query(PajShp).filter(
                    PajShp.fn.in_([
                        _removenonascii(path.basename(x)) for x in shps[0]
                    ])).delete(synchronize_session=False)
            if invs[0]:
                sess.query(PajInv).filter(PajInv.invno.in_(
                    invs[0])).delete(synchronize_session=False)
            jns = set()
            if shps[1]:
                jns.update([JOElement(x.jono) for x in shps[1].values()])
            if invs[1]:
                jns.update([JOElement(x.jono) for x in invs[1].values()])
            if jns:
                jns = self._hksvc.getjos(jns)[0]
                jns = dict([(x.name, x) for x in jns])
                if shps[1]:
                    for dct in [x._asdict() for x in shps[1].values()]:
                        je = JOElement(dct["jono"])
                        if je not in jns or not isvalidp17(dct["pcode"]):
                            logger.info(
                                "Item(%s) does not contains valid JO# or pcode"
                                % dct)
                            continue
                        dct["fn"] = _removenonascii(dct["fn"])
                        dct["joid"] = jns[je].id
                        dct["mtlwgt"] = dct["mtlwgt"].metal_jc
                        # the stone weight field might be str only, set it to zero
                        shp = PajShp()
                        for x in dct.items():
                            k = x[0]
                            lk = k.lower()
                            if hasattr(shp, lk):
                                setattr(shp, lk, dct[k])
                        sess.add(shp)
                if invs[1]:
                    for dct in [x._asdict() for x in invs[1].values()]:
                        if not dct["stspec"]:
                            dct["stspec"] = NA
                        else:
                            dct["stspec"] = _removenonascii(dct["stspec"])
                        dct["china"] = 0
                        dct["joid"] = jns[JOElement(dct["jono"])].id
                        iv = PajInv()
                        for it in dct.items():
                            k, lk = it[0], it[0].lower()
                            if hasattr(iv, lk):
                                iv.__setattr__(lk, dct[k])
                        iv = sess.add(iv)
            sess.commit()
            err = False
        return -1 if err else 1, err

    def process(self, fldr):
        """
        read the shipment file and send shipment/invoice to hkdb
        @param fldr: the folder contains the files. sub-folders will be ignored
        """
        ptn = re.compile(r"HNJ\s+\d*-", re.IGNORECASE)
        fns = getfiles(fldr, "xls", True)
        if fns:
            p = appathsep(fldr)
            fns = [p + x for x in fns if ptn.match(x)]
        if not fns:
            return
        errors = list()
        app, kxl = _appmgr.acq()
        try:
            # when excel open a file, the file's modified date will be changed, so, in
            # order to get the actual modified date, get it first
            fmds = {x: self._getfmd(x) for x in fns}
            fns = sorted([(x, self.get_shp_date(x)) for x in fns],
                         key=lambda x: x[1])
            fns = [x[0] for x in fns]
            for fn in fns:
                rflag = self._hasread(fmds[fn], fn)
                if rflag == 1:
                    logger.debug(
                        "data in file(%s) is up-to-date" % path.basename(fn))
                    continue
                shptorv, invtorv = [], []
                shps, invs = {}, {}
                shtshps, shtinvs = [], []
                if rflag == 2:
                    shptorv.append(fn)
                shd0, fmd, wb = self.get_shp_date(fn), fmds[fn], app.books.open(
                    fn)
                try:
                    bomwgts = PajBomHhdlr().readbom(wb)
                    for sht in wb.sheets:
                        if sht.name.find(u"返修") >= 0:
                            continue
                        rng = xwu.find(sht, u"十七*", lookat=LookAt.xlPart)
                        if not rng:
                            rng = xwu.find(sht, u"物料*", lookat=LookAt.xlPart)
                        if not rng:
                            if xwu.find(sht, "PAJ"):
                                shtinvs.append(sht)
                        else:
                            shtshps.append(sht)
                    if shtshps and shtinvs:
                        if rflag != 1:
                            for sht in shtshps:
                                its = PajShpHdlr.read_shp(
                                    fn, shd0, fmd, sht, bomwgts)
                                if its:
                                    if "_ERROR_" in its:
                                        errors.append(its["_ERROR_"])
                                        break
                                    else:
                                        shps.update(its)
                        for sht in shtinvs:
                            its = self._readinv(fn, sht, fmd)
                            if its:
                                if its[0]:
                                    invs.update(its[0])
                                if its[1]:
                                    invtorv.extend(its[1])
                    elif bool(shtshps) ^ bool(shtinvs):
                        logger.info(
                            "Error::Not both shipment and invoice in file(%s), No data updated"
                            % path.basename(fn))
                finally:
                    if wb:
                        wb.close()
                if sum((len(x) for x in (shptorv, shps, invtorv, invs))) == 0:
                    logger.debug("no valid data returned from file(%s)" %
                                 path.basename(fn))
                logger.debug(
                    "counts of file(%s) are: Shp2Rv=%d, Shps=%d, Inv2Rv=%d, Invs=%d"
                    % (path.basename(fn), len(shptorv), len(shps), len(invtorv),
                       len(invs)))
                # sometimes the shipmentdata does not have inv# data
                its = {x[0]: x[1] for x in shps.items() if not x[1].invno}
                if its:
                    xmp = {x.jono: x for x in invs.values()}
                    for it in its.items():
                        x = xmp.get(it[1].jono)
                        if not x:
                            logger.debug("failed to get invoice for JO#(%s)" %
                                         it[1].jono)
                            return -1
                        else:
                            shps[it[0]] = it[1]._replace(invno=x.invno)
                x = self._persist((shptorv, shps), (invtorv, invs))
                if x[0] != 1:
                    errors.append(x[1])
                    logger.info("file(%s) contains errors", path.basename(fn))
                    logger.info(x[1])
                else:
                    logger.debug("data in file(%s) were committed to db",
                                 (path.basename(fn)))
        finally:
            _appmgr.ret(kxl)
        return -1 if errors else 1, errors


class PajJCMkr(object):
    """
    the JOCost maker of Paj for HK accountant, the twin brother of C1JCMkr
    """

    def __init__(self, hksvc=None, cnsvc=None, bcsvc=None):
        self._hksvc, self._cnsvc, self._bcsvc = hksvc, cnsvc, bcsvc

    def run(self, year, month, day=1, tplfn=None, tarfn=None):
        """ create report file of given year/month"""

        def _makemap(sht=None):
            coldefs = (
                u"invoice date=invdate;invoice no.=invno;order no.=orderno;customer=cstname;"
                u"job no.=jono;style no.=styno;running no.=running;paj item no.=pcode;karat=karat;"
                u"描述=cdesc;in english=edesc;job quantity=joqty;quantity received=shpqty;"
                u"total cost=ttlcost;cost=uprice;平均单件金银重g=umtlwgt;平均单件石头重g=ustwgt;石头=stspec;"
                u"mm program in#=iono;jmp#=jmpno;date=shpdate;remark=rmk;has dia=hasdia"
            )
            vvs = sht.range("A1").expand("right").value
            vvs = [x.lower() if isinstance(x, str) else x for x in vvs]
            imap, nmap = {}, {}
            for s0 in coldefs.split(";"):
                ss0 = s0.split("=")
                x = [
                    x for x in range(len(vvs))
                    if x not in imap and vvs[x].find(ss0[0]) >= 0
                ]
                if x:
                    imap[x[0]] = ss0[1]
                    nmap[ss0[1]] = x[0]
                else:
                    print("failed to get colname %s" % s0)

            return nmap, imap

        dfmt = "%Y/%m/%d"
        df, dt = daterange(year, month, day)

        runns, jes = set(), set()
        bcsvc = self._bcsvc

        mms = self._cnsvc.getshpforjc(df, dt)
        for x in mms:
            jo = x.JO
            rn = str(jo.running)
            if rn not in runns:
                runns.add(rn)
            jn = jo.name
            if jn not in jes:
                jes.add(jn)
        runns = tuple(runns)
        bcs = dict([(x.runn, x.desc) for x in bcsvc.getbcsforjc(runns)])
        lst = self._hksvc.getpajinvbyjes(jes)
        pajs, pajsjn = {}, {}
        for x in lst:
            jn = x.JO.name
            pajs["%s,%s" % (jn, x.PajShp.invdate.strftime(dfmt))] = x
            jn = jn.value
            if jn not in pajsjn:
                pajsjn[jn] = []
            lst1 = pajsjn[jn]
            lst1.append(x)
        ios = dict([("%s,%s,%s" % (x.running, x.jmp, x.shpdate.strftime(dfmt)),
                     x) for x in self._hksvc.getmmioforjc(df, dt, runns)])
        app, kxl = _appmgr.acq()
        lst = []
        try:
            wb = xwu.fromtemplate(tplfn, app)
            sht = wb.sheets("Data")
            nmps = _makemap(sht)
            ss = (
                "cstname,Customer.name,karat,JO.karat,cdesc,JO.description,joqty"
                ",JO.qty,jmpno,MM.name,shpdate,MMMa.refdate,shpqty,MM.qty"
            ).split(",")
            dtmap0 = dict(zip(ss[0:len(ss) - 1:2], ss[1:len(ss):2]))
            ss = (
                "invdate,PajShp.invdate,invno,PajShp.invno,orderno,PajShp.orderno"
                ",pcode,PajShp.pcode,uprice,PajInv.uprice,umtlwgt,PajShp.mtlwgt"
                ",ustwgt,PajShp.stwgt,stspec,PajInv.stspec").split(",")
            dtmap1 = dict(zip(ss[0:len(ss) - 1:2], ss[1:len(ss):2]))

            for row in mms:
                mp = {}
                rn = str(row.JO.running)
                jn = row.JO.name.value
                for x in dtmap0.items():
                    mp[x[0]] = deepget(row, x[1])
                mp["running"], mp["jono"], mp["styno"], mp["edesc"] = rn, "'" + \
                    jn, row.Style.name.value, bcs[rn] if rn in bcs else "N/A"

                key1, key, fnd = jn, "%s,%s" % (
                    jn, mp["shpdate"].strftime(dfmt)), False
                if key in pajs:
                    x = pajs[key]
                    for y in dtmap1.items():
                        mp[y[0]] = deepget(x, y[1])
                    fnd = True
                elif key1 in pajsjn:
                    lst1 = pajsjn[key1]
                    if lst1:
                        hts = []
                        shpd = mp["shpdate"]
                        for x in lst1:
                            ddiff = x.PajShp.shpdate - shpd
                            if (abs(ddiff.days) <= 5 and
                                    abs(x.PajShp.qty - float(mp["shpqty"])) <
                                    0.1):
                                hts.append(x)
                                for y in dtmap1.items():
                                    mp[y[0]] = deepget(x, y[1])
                                fnd = True
                                break
                        if hts:
                            for x in hts:
                                lst1.remove(x)
                        if not lst1:
                            del pajsjn[key1]
                if fnd:
                    mp["ttlcost"] = mp["uprice"] * mp["shpqty"]
                if not fnd:
                    for x in dtmap1.keys():
                        mp[x] = None
                    mp["ttlcost"] = None
                key = "%s,%s,%s" % (rn, mp["jmpno"],
                                    mp["shpdate"].strftime(dfmt))
                mp["rmk"] = ("QtyError"
                             if not (mp["joqty"] and mp["shpqty"]) else
                             "" if mp["joqty"] == mp["shpqty"] else "Partial")
                mp["iono"] = ios[key].inoutno if key in ios else "N/A"
                hasdia = (mp["cdesc"].find(u"钻") >= 0 or
                          mp["cdesc"].find(u"占") >= 0 or
                          (mp["edesc"] and
                           mp["edesc"].lower().find("dia") >= 0))
                mp["hasdia"] = "D" if hasdia else "N"

                x = [mp[nmps[1][x]] for x in range(len(nmps[1]))]
                lst.append([
                    "" if not y else y.strip() if isinstance(y, str) else
                    y.strftime(dfmt) if isinstance(y, datetime) else
                    float(y) if isinstance(y, Decimal) else y for y in x
                ])
            sht.range("A2").value = lst
            for x in [x for x in wb.sheets if x != sht]:
                x.delete()
            if tarfn:
                wb.save(tarfn)
        finally:
            _appmgr.ret(kxl)
        return lst, tarfn

def _read_pcodes(fn):
    if not (fn and path.exists(fn)):
        return None
    pcodes = None
    with open(fn, "r+t") as fh:
        pcodes = list({x[:-1] for x in fh.readlines() if x[0] != "#"})
    return pcodes

class PajCache(object):
    """
    Paj unit-price tracer
    to use this method, put a dat file inside a folder which should contains sty#
    then I will read and show the price trends

    to speed up the process of fetching data from hk, the key data(wgt/poprices) were cached by a sqlitedb.

    the original purpose is to track the stamping products, but in fact, can be use for any Paj items.
    Constructor Arguments:
    @srcsm: sessionMgr to the source db
    @localsm: sessionMgr to the local db(to cache the source)
    @srcfn: the text file contains the pcodes
    """

    def __init__(self, srcsm, localsm, srcfn):
        self._hksvc = HKSvc(srcsm)
        self._src_fn = srcfn
        self._local_sm = localsm
        PajInvSt.metadata.create_all(localsm.engine)

    def cache(self):
        """
            cache revcn/pajinv/weights
        """
        pcodes0 = self._cache_revcns()
        pcodes = _read_pcodes(self._src_fn) or pcodes0
        if not pcodes:
            return
        ttl = len(pcodes)
        logger.debug("totally %d pcodes need weight caching" % ttl)
        cnt, sz  = 0, 50
        for arr in splitarray(pcodes, sz):
            self._cache_wgts(arr)
            cnt += 1
            logger.debug("%d of %d weight records cached" % (cnt * sz, ttl))
        logger.debug("all weight of given pcodes cached")

    def _cache_wgts(self, pcodes):
        """
        create weight/invoice data from given product codes
        """
        mp = {"cc": PajCalc(), "td": datetime.today()}
        for pcode in pcodes:
            with ResourceCtx((self._local_sm, self._hksvc.sessmgr())) as curs:
                try:
                    self._cache_wgt(pcode, curs, **mp)
                except:
                    curs[0].rollback()
                    logger.debug("Error occur while persisting cache result")

    def _cache_wgt(self, pcode, curs, **kwds):
        """
        persist one pcode
        @param: curs: curs[0] is localDB, curs[1] is source(HK) db
        """
        var = Query([PajItemSt, PrdWgtSt]).outerjoin(PrdWgtSt)
        var = var.filter(PajItemSt.pcode == pcode)
        var = var.with_session(curs[0]).first()
        if var:
            if var[1]:
                logger.debug("weight of pcode(%s) already cached", pcode)
                return
            pi = var[0]
        else:
            pi = PajItemSt()
            pi.pcode, pi.createdate, pi.tag = pcode, kwds["td"], 0
            curs[0].add(pi)
            curs[0].flush()
        var = Query([PajShp.pcode, JOhk.name.label("jono"),
            Styhk.name.label("styno"), JOhk.createdate, PajShp.invdate,
            PajInv.uprice, PajInv.mps])
        var = var.join(JOhk).join(Orderma).join(Styhk).join(
            PajInv, and_(PajShp.joid == PajInv.joid,
            PajShp.invno == PajInv.invno))
        var = var.filter(PajShp.pcode == pcode)
        var = var.with_session(curs[1]).all()
        if not var:
            return
        td, jeset, wgts = kwds["td"], set(), None
        for jnv in var:
            je = jnv.jono
            if je in jeset:
                continue
            jeset.add(je)
            if not wgts:
                wgts = self._hksvc.getjowgts(je)
                if not wgts:
                    continue
                for ic, wgt in enumerate(wgts.wgts):
                    if not wgt:
                        continue
                    cn = PrdWgtSt()
                    cn.pid, cn.karat, cn.wgt, cn.remark, cn.tag = pi.id, wgt.karat, wgt.wgt, je.value, 0
                    cn.createdate, cn.lastmodified = td, td
                    cn.wtype = 0 if ic == 0 else 100 if ic == 2 else 10
                    curs[0].add(cn)
            cn = kwds["cc"].calchina(wgts, jnv.uprice, jnv.mps, PAJCHINAMPS)
            if cn:
                ic = PajInvSt()
                ic.pid, ic.uprice, ic.mps = pi.id, jnv.uprice, jnv.mps
                ic.cn = cn.china
                ic.jono, ic.styno = je.value, jnv.styno.value
                ic.jodate, ic.createdate, ic.invdate, ic.lastmodified = jnv.createdate, td, jnv.invdate, td
                ic.mtlcost, ic.otcost = cn.metalcost, cn.china - cn.metalcost
                curs[0].add(ic)
        curs[0].commit()
        logger.debug("weight of pcode(%s) cached", pcode)

    @classmethod
    def _get_tar_pis(clas, pcodes, cur_src):
        var = set((x for x in pcodes)) if pcodes else set()
        if not var:
            return None
        logger.debug(
            "Totally %d revised records return from HK, now begin copying",
            len(var))
        pcs = {}
        for arr in splitarray(list(var)):
            q0 = Query(PajItemSt).filter(PajItemSt.pcode.in_(arr))
            try:
                q0 = q0.with_session(cur_src).all()
                if not q0:
                    continue
                pcs.update({x.pcode: x for x in q0})
            except:
                pass
        return pcs

    @classmethod
    def _get_tar_revs(cls, src_its, cur_tar):
        if not src_its:
            return None
        var = []
        for arr in splitarray([x.id for x in src_its.values()]):
            try:
                var.extend(cur_tar.query([PajItemSt, PajCnRevSt]).join(PajCnRevSt).
                    filter(PajCnRevSt.id.in_(arr)).all())
            except:
                pass

    def _cache_revcns(self):
        """
        cache the rev history, only create the blank new, whose tag changed
        from 0 to lt 0 won't be caught. So when the last cache is too far
        away, clear the revcn first.
        """
        affdate = datetime(2018, 4, 4)
        q0 = Query((func.max(PajCnRevSt.createdate),))
        with ResourceCtx(self._local_sm) as cur:
            var = q0.with_session(cur).first()[0]
            if not var:
                var = affdate
            else:
                if (datetime.today() - var).days > 2:
                    var = affdate
                    PajCnRevSt.query.delete()
                    cur.commit()
        q0 = Query(PajCnRev).filter(
            and_(PajCnRev.filldate > var, PajCnRev.tag == 0,
                 PajCnRev.revdate >= affdate))
        with ResourceCtx((self._local_sm, self._hksvc.sessmgr())) as curs:
            src_revs = q0.with_session(curs[1]).all()
            pcodes = {x.pcode for x in src_revs}
            tar_its = self._get_tar_pis(pcodes, curs[0])
            affdate, npis = datetime.today(), []
            for x in [y for y in src_revs if y.pcode not in tar_its]:
                var = PajItemSt()
                var.pcode, var.createdate, var.tag = x.pcode, affdate, 0
                curs[0].add(var)
                npis.append(var)
            if npis:
                curs[0].flush()
                tar_its.update({x.pcode: x for x in npis})
            for x in src_revs:
                var = PajCnRevSt()
                var.pid, var.uprice, var.revdate, var.createdate, var.tag = tar_its[
                    x.pcode].id, x.uprice, x.revdate, affdate, x.tag
                curs[0].add(var)
            curs[0].commit()
        return pcodes

class PajUPTracker(object):
    """
    Keep track of the PAJ UnitPrice changes history (based on a text file
    which contains the pcodes)
    @Arguments:
    @srcsm: the sessionMgr for the source db
    @localsm: the sessionMgr for the local db

    @optional Arguments:
    @file: the file contains the pcodes that need to analyse
    @cutdate: the cutting date to analyse. default is 2018/05/01
    """
    nl_mix = NamedList("oc cn invdate jono".split())

    def __init__(self, srcsm, localsm, **kwargs):
        self._source_sm, self._local_sm = srcsm, localsm
        self._file = kwargs.get("file")
        self._cut_date = kwargs.get("cutdate") or datetime(2018, 5, 1)

    @classmethod
    def minmax(cls, arr):
        """
        return a 3 element tuple, each element contains mixcols data
        first   -> min
        second  -> max
        third   -> last
        """
        if not arr:
            return None
        fill = lambda ar: [float(ar.otcost), float(ar.cn), ar.invdate, ar.jono]
        li, lx = 9999, -9999
        mi = mx = None
        for ar in arr:
            lb = float(ar.otcost)
            if lb > lx:
                mx, lx = fill(ar), lb
            if lb < li:
                mi, li = fill(ar), lb
        df = lx - li
        if df > 0 and (df < 0.1 or df / li < 0.01):
            df = (lx + li) / 2.0
            mi[0], mx[0] = (df,) * 2
        return mi, mx, fill(arr[-1])

    def getonly(self, cns, arr):
        if isinstance(cns, str):
            cns = cns.split(",")
        if not arr:
            return [
                None,
            ] * (3 * len(cns))
        lst = []
        for ar in arr:
            self.nl_mix.setdata(ar)
            lst.extend([self.nl_mix[cn] for cn in cns])
        return lst

    @classmethod
    def fetch(cls, cur, pcodes=None):
        """
        fetch the pajitem/pajinv data from cache
        """
        q0 = Query([
            PajItemSt.pcode, PajInvSt.jono, PajInvSt.styno, PajInvSt.invdate,
            PajInvSt.cn, PajInvSt.otcost])
        q0 = q0.join(PajInvSt).order_by(PajItemSt.pcode).order_by(PajInvSt.invdate)
        if pcodes:
            lst = []
            for mp in splitarray(pcodes):
                try:
                    lst.extend(q0.filter(PajItemSt.pcode.in_(mp)).with_session(cur).all()
                    )
                except:
                    pass
            q0 = lst
        else:
            q0 = q0.with_session(cur).all()
        if not q0:
            return (None,) * 2
        mp, revdates = {}, {}
        for arr in q0:
            mp.setdefault(arr.pcode, []).append(arr)
        q0 = Query([PajItemSt.pcode, PajCnRevSt.revdate,
                    PajCnRevSt.uprice]).join(PajCnRevSt)
        for arr in splitarray(list(mp)):
            try:
                revdates.update({
                    y.pcode: (y.revdate, float(y.uprice)) for y in [
                        x for x in q0.filter(PajItemSt.pcode.in_(arr)).
                        with_session(cur).all()
                    ]
                })
            except:
                logger.debug("exception occur while caching revcn")
        return mp, revdates

    def _group(self, mp, revdates):
        if not mp:
            return None
        data, nl_q = {}, NamedList("pcode,jono,styno,invdate,cn,otcost")
        m_x = self.__class__.minmax
        for arr in mp.items():
            lst = arr[1]
            nl_q.setdata(lst[0])
            flag, idx = lst and len(lst) > 1, 0
            acutdate, revcn = revdates.get(arr[0], self._cut_date), 0
            if isinstance(acutdate, tuple):
                revcn, acutdate = acutdate[1], acutdate[0]
            if flag:
                for idx, x in enumerate(lst):
                    flag = x.invdate > acutdate
                    if flag:
                        break
            if not flag:
                flag = data.setdefault("noaff", [])
                flag.append((nl_q.pcode, nl_q.styno, acutdate, revcn, m_x(lst)))
            else:
                mix0, mix1 = m_x(lst[:idx]), m_x(lst[idx:])
                val = (nl_q.pcode, nl_q.styno, acutdate, revcn, mix0, mix1)
                if mix0 is None:
                    iot = "nochg"
                else:
                    iot = self.nl_mix.getcol("oc")
                    if mix0[0][iot] * 2.0 / 3.0 + 0.05 >= mix1[1][iot]:
                        iot = "drp"
                    elif self.__class__.almosteq(mix0[0][iot], mix1[1][iot]):
                        iot = "nochg"
                    elif mix0[0][iot] > mix1[1][iot]:
                        iot = "noeng"
                    elif mix0[0][iot] < mix1[1][iot]:
                        # old's max under new's min
                        iot = "pum"
                    else:
                        iot = "mixture"
                data.setdefault(iot, []).append(val)
        return data

    @classmethod
    def almosteq(cls, x, y):
        """ check if x and w are ver close """
        return abs(x - y) <= 0.1 or abs(x - y) / x < 0.01

    def analyse(self):
        """ do the Paj UnitPrice trend analyst """
        PajCache(self._source_sm, self._local_sm, self._file).cache()
        pcodes = _read_pcodes(self._file)
        with ResourceCtx(self._local_sm) as cur:
            grouped_data = self._group(*self.__class__.fetch(cur, pcodes))
            grps = {
                0: (("",), "Before,After".split(",")),
                1: "Min.,Max.,Last".split(","),
                2: "pcode,styno,revdate,cn,karat".split(",")
            }
            mp = {
                "noaff": "NotAffected",
                "nochg": "NoChanges",
                "mixture": "Mixture",
                "noeng": "NoEnough",
                "drp": "PriceDrop1of3",
                "pum": "PriceUp"
            }
            ctss = ("cn invdate".split(), "oc cn jono invdate".split(), )
            shts, pd = [], P17Decoder()
            app = xwu.appmgr.acq()[0]
            wb = app.books.add()
            for x in grouped_data.items():
                shts.append(wb.sheets.add())
                sht = shts[-1]
                sht.name, vvs = mp.get(x[0], "_" + str(random.randint(0, 99999))), []
                gidx = len(grps[2])
                ttl0, ttl1 = [None, ] * gidx, [None, ] * gidx
                gidx = 0 if x[0] == "noaff" else 1
                for z in grps[0][gidx]:
                    ttl0.append(z)
                    ttl0 += [None,] * (len(ctss[gidx]) * len(grps[1]) - 1)
                    for xx in grps[1]:
                        ttl1.append(xx)
                        ttl1 += [None, ] * (len(ctss[gidx]) - 1)
                if len(grps[0][gidx]) > 1:
                    vvs.append(ttl0)
                vvs.append(ttl1)

                ttl = grps[2].copy()
                ttlen = len(ttl) - 1
                ttl.extend(ctss[gidx] * (len(grps[1]) * len(grps[0][gidx])))
                vvs.append(ttl)
                for arr in x[1]:
                    ttl = list(arr[:ttlen])
                    ttl.append(pd.decode(ttl[0], "karat"))
                    for kk in arr[ttlen:]:
                        ttl.extend(self.getonly(ctss[gidx], kk))
                    vvs.append(ttl)
                sht.range(1, 1).value = vvs
                sht.autofit("c")
                # let the karat column smaller
                sht[1, grps[2].index("karat")].column_width = 10
                xwu.freeze(
                    sht.range(3 + (1 if len(grps[0][gidx]) > 1 else 0),
                              ttlen + 2))

            if shts:
                for sht in wb.sheets:
                    if sht not in shts:
                        sht.delete()
            app.visible = True
