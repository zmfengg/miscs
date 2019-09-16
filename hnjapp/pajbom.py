#! coding=utf-8
'''
* @Author: zmFeng
* @Date: 2019-01-16 09:18:55
* @Last Modified by:   zmFeng
* @Last Modified time: 2019-01-16 09:18:55
 class for PAJBom handling, including BomParse and BOMCaching
'''
from datetime import datetime
from logging import DEBUG
from os import path
from re import I as icase
from re import compile as cpl
from numbers import Number

from sqlalchemy import and_, func
from sqlalchemy.orm import Query
from xlwings import Book, apps
from hnjcore import JOElement, isvalidp17
from hnjcore.models.hk import JO as JOhk
from hnjcore.models.hk import Orderma, PajShp
from hnjcore.models.hk import Style as Styhk
from utilz import (NamedList, NamedLists, ResourceCtx, getfiles, getvalue,
                   karatsvc, tofloat, triml, xwu, NA)
from utilz.xwu import appmgr as _appmgr, esctext

from .common import _logger as logger, P17Decoder
from .localstore import PajBom, PajItem
from .pajcc import PrdWgt, WgtInfo, addwgt


class PajBomHdlr(object):
    """ class to read BOMs from PAJ
    @param part_chk_ver: the Part checker version,
        default is None or 0,
            That is, when there is (chain with length) and (lock exists),
            圈 will be treated as part of the chain
        1 stands for loose,
            That is, when there is chain with length,
            圈 will be treated as part of the chain
    @param cache: A sessionMgr pointing to a cache db
    """
    _ptn_oz = cpl(r"\(\$\d*/OZ\)")
    _one_hit_mp = {
        925: (cpl(r"(925)|(银)"), ),
        200: (cpl(r"(BRONZE)|(铜)", icase), ),
        9925: (cpl(r"BONDED", icase), cpl(r"B&Gold", icase))
    }
    _ptn_k_gold = cpl(r"^(\d*)K")
    _ptn_digits = cpl(r"[\(（](\d*)[\)）]")
    _ptn_chn_lck = cpl(r"(弹簧扣)|(龙虾扣)|(狗仔头扣)")
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
        self._dao = getvalue(kwargs, "cache,cache_db,sessmgr")
        if self._dao:
            self._dao = _PajBomDAO(self._dao)

    def _parse_karat(self, mat, wis=None, is_mstr=True):
        """ return karat(int type) from material string """
        if is_mstr:
            mt = self._ptn_oz.search(mat)
            if not mt:
                # no /oz sign, only 200 and 9925 is allowed
                if not [1 for y in (200, 9925) for j in self._one_hit_mp[y] if j.search(mat)]:
                    return None
        kt = max(x[0] if y.search(mat) else 0 for x in self._one_hit_mp.items() for y in x[1])
        if not kt:
            mt = self._ptn_k_gold.search(mat) or self._ptn_digits.search(mat)
            kt = int(mt.group(1)) if mt else None
        if kt:
            kt = karatsvc.getkarat(kt) or karatsvc.getbyfineness(kt)
            return kt.karat if kt else None
        # not found, has must have keyword? if yes, follow master
        voids = [1 for x in self._voids if mat.find(x) >= 0]
        if not voids and wis and any(wis):
            s0 = mat.lower()
            if s0.find(u"金") < 0:
                # finally no one is found, follow master
                # kt = wis[0].karat
                # but zhangyuting claimed in e-mail with title "配件的"物料名称"里没有金" on 2018/12/10 that the karat should be 925
                # so let it to be 925
                return 925
            if any([x for x in self._mtl_parts if s0.find(x) >= 0]):
                for karat in (karatsvc.getkarat(x.karat) for x in wis if x):
                    if not karat or karat.category != karatsvc.CATEGORY_GOLD:
                        continue
                    return karat.karat
        if logger.isEnabledFor(DEBUG) and not kt and wis and not voids:
            logger.error("No karat found for (%s) and no default provided" %
                mat)
        return None

    def _ispendant(self, pcode):
        return self._pcdec.decode(pcode, "PRODTYPE").find("吊") >= 0

    def _isring(self, pcode):
        return self._pcdec.decode(pcode, "PRODTYPE").find("戒") >= 0

    @staticmethod
    def wgtsForPaj():
        ''' read active book, write weigts. Request by Zhangyuting
        '''
        app = apps.active
        if not app:
            return
        try:
            wb = app.books.active
            mp = PajBomHdlr().readbom(wb)
            lsts = []
            lsts.append('17码 主成色 主重 副成色 副重 配件成色 配件重'.split())
            for pcode in mp:
                pwgt = mp[pcode]['mtlwgt']
                wgts = [(x.karat, x.wgt) if x else (None, None) for x in pwgt.wgts]
                wgts = [y for x in wgts for y in x]
                wgts.insert(0, pcode)
                lsts.append(wgts)
            sht = wb.sheets.add(after=wb.sheets[-1])
            sht.cells(1, 1).value = lsts
            sht.autofit('c')
            rng = xwu.usedrange(sht)
            rng.row_height = 18
            xwu.maketable(rng)
            xwu.freeze(sht.cells(2, 2))
        except:
            pass

    def readbom(self, fldr, cvt2tbl=True):
        """
        read BOM from given folder
        @param fldr: the folder contains the BOM file(s)
        @param cvt2tbl: convert the sheet to excel table
        return a dict with "pcode" as key and dict as items
            the item dict has keys("pcode","mtlwgt")
        """
        pmap = {}
        if isinstance(fldr, Book):
            self._read_book(fldr, pmap, cvt2tbl)
        else:
            fns = getfiles(fldr, "xls") if path.isdir(fldr) else (fldr,)
            if not fns:
                return None
            app, kxl = _appmgr.acq()
            try:
                for fn in fns:
                    wb = app.books.open(fn)
                    self._read_book(wb, pmap, cvt2tbl)
                    wb.close()
            finally:
                if kxl and app:
                    _appmgr.ret(kxl)
        self._adjust_wgts(pmap)
        return pmap

    def readbom_manual(self, wb, pcodeset, **kwds):
        '''
        from a shipment file, read the boms and return
        a tuple ready for inserting into sheet
        a tuple of (row(in excel term), rowcount) for place marker for each segment
        a namedlist for column lookup of the first returned argument
        '''
        if not pcodeset:
            return None
        pmap = {}
        main_offset, min_rowcnt = kwds.get("main_offset", 3), kwds.get("min_rowcnt", 7)
        bcwgts = kwds.get("bc_wgts") or {}
        self._read_book(wb, pmap)
        self._adjust_wgts(pmap)
        if NA in pcodeset:
            # the caller does not provided any pcode, try to build based on returned bom's semi-chain item(if there is).
            lsts = {x["pcode"] for x in pmap.values() if x["mtlwgt"].part and x["mtlwgt"].part.wgt < 0}
            if not lsts:
                return None
            pcodeset = {x[0]: x[1] for x in pcodeset.items() if x[0] in lsts}
        lsts = 'jono rcat mkarat mwgt mid mname pkarat pqty pwgt mpflag image'.split()
        main_offset = min(max(2, main_offset), 5)
        min_rowcnt = max(main_offset + 4 * (2 if bcwgts else 1), min_rowcnt)
        nl = NamedList(lsts)
        lsts, mkrs = [lsts, ], []
        pmap = sorted([x for x in pmap.items() if x[0] in pcodeset], key=lambda x: pcodeset.get(x[0])[0])
        for pcode, x in pmap:
            if pcode not in pcodeset:
                continue
            fn, kts = [], []
            pts, mstrs = [x.get(y, []) for y in "parts mstr".split()]
            # put the netwgt to the end of mstrs
            mstrs.append((0, x["mtlwgt"].netwgt))
            pjns = [y[0] for y in pcodeset[pcode]]
            pts = [x for x in pts if x.get("karat")]
            lns = [len(x) for x in (mstrs, pts, pjns)]
            for idx in range(max(lns)):
                fn.append(nl.newdata(True))
                if idx == 0:
                    nl.image = pcode
                    nl.jono, nl.rcat = pcodeset[pcode][0]
                    nl.jono = "'" + nl.jono
                elif idx < lns[2]:
                    # merge JOs with same pcode
                    nl.jono = "'" + pjns[idx]
                if idx < lns[0]:
                    self._fill_mstr(mstrs[idx], nl)
                    if nl.mkarat > 0:
                        kts.append(nl.mkarat)
                if idx == lns[0] - 1:
                    nl.rcat = "NetWgt"
                if idx < lns[1]:
                    self._fill_dtl(pts[idx], nl)
            idx += 1
            jn = fn[0][nl.getcol("jono")][1:]
            amrc = max(min_rowcnt - 0 if jn in bcwgts else 4, main_offset + 2 + 1)
            if idx < amrc:
                for idx in range(amrc - idx):
                    fn.append(nl.newdata(True))
            if kts:
                self._set_sub(nl, main_offset, "Main", fn, kts)
                self._set_sub(nl, main_offset + 2, "Part", fn, kts, False)
            kts = bcwgts.get(jn)
            if kts:
                self._fill_bc(nl, main_offset + 2 * 2, kts, fn)
            nl.setdata(fn[0])
            mkrs.append((nl.jono, nl.rcat, len(lsts) + 1, len(fn) - 1,))
            lsts.extend(fn)
            del pcodeset[pcode]
        # if there is still items in pcodeset, throw it to the result
        if pcodeset:
            for x in pcodeset.items():
                lsts.append(nl.newdata(True))
                pcode = x[1][0]
                nl.image, nl.jono, nl.rcat, nl.mname = x[0], "'" + pcode[0], pcode[1], "_NO_BOM_"
        return lsts, mkrs, nl, (main_offset, min_rowcnt)

    @staticmethod
    def _fill_mstr(qnw, nl):
        nl.mkarat, nl.mwgt = qnw

    @staticmethod
    def _fill_dtl(dtl, nl):
        nl.mid, nl.mname, nl.pkarat, nl.pqty, nl.pwgt = [dtl.get(x) for x in "matid name karat qty wgt".split()]
        nl.mpflag = 'N' if dtl.get("part_hints", False) else 'Y'

    @staticmethod
    def _fill_bc(nl, ridx, bcwgt, fn):
        wgts = [(x.karat, x.wgt) if x else (None, None) for x in bcwgt.wgts]
        for x, wgt in enumerate(wgts):
            nl.setdata(fn[ridx + x])
            nl.mkarat, nl.mwgt = wgt
            if x == 0:
                nl.rcat = "Main.BC"
            elif x == 2:
                nl.rcat = "Part.BC"

    @staticmethod
    def _set_sub(nl, idx, ttl, fn, kts, is_main=True):
        nl.setdata(fn[idx])
        nl.rcat, nl.mkarat = ttl, kts[0]
        if is_main and len(kts) > 1:
            nl.setdata(fn[idx + 1])
            nl.mkarat = kts[1]

    def _part_chk_strict(self, bi, chns, lks, has_semi_chn):
        """
        determine if the given bom item is a part
        """
        return chns and bi["_id"] in chns or lks and (lks.get(bi["_id"]) or sum([1 for v in self._pts_kws if bi["name"].find(v) >= 0]))

    def _part_chk_loose(self, bi, chns, lks, has_semi_chn):
        """
        loose rule for determining if the given bom item is a part
        """
        return chns and bi["_id"] in chns or lks and lks.get(bi["_id"]) or has_semi_chn and [1 for v in self._pts_kws if bi["name"].find(v) >= 0]

    def _adjust_wgts(self, pmap):
        if not pmap:
            return
        is_part = self._part_chk_strict if not self._part_chk_ver else self._part_chk_loose
        for pcode, prop in pmap.items():
            ch_lks, prdwgt = prop.get("mtlwgt"), None
            if ch_lks:
                pmap[pcode]["mstr"] = ch_lks
                for y in ch_lks:
                    prdwgt = addwgt(prdwgt, WgtInfo(y[0], y[1]))
            else:
                prdwgt = PrdWgt(WgtInfo(0, 0))
                logger.debug("%s does not have master weight" % pcode)
            if "parts" not in prop:
                prop["mtlwgt"] = prdwgt._replace(netwgt=prop.get("netwgt"))
                continue
            ch_lks = {}
            if self._ispendant(pcode):
                for y in prop["parts"][::]:
                    var = y["name"]
                    if triml(var).find("chain") >= 0:
                        ch_lks.setdefault("chain", {})[y["_id"]] = y
                    elif self._ptn_chn_lck.search(var):
                        ch_lks.setdefault("lock", {})[y["_id"]] = y
            chns, lks = tuple(ch_lks.get(x) or {} for x in "chain lock".split())
            has_semi_chn, subs = self._has_semi_chn(chns), 0
            for y in prop["parts"]:
                var = y["name"]
                kt = self._parse_karat(var, prdwgt.wgts, False)
                if not kt:
                    subs += y["wgt"]
                    continue
                y["karat"], var = kt, False
                var = is_part(y, chns, lks, has_semi_chn)
                if var:
                    #make sure part candidate has the same karat with old
                    wgt0 = prdwgt.part
                    var = not wgt0 or kt == wgt0.karat
                prdwgt = addwgt(prdwgt, WgtInfo(kt, y["wgt"]),\
                    var, autoswap=False)
                y["part_hints"] = var
            if has_semi_chn and prdwgt.part:
                prdwgt = prdwgt._replace(part=WgtInfo(prdwgt.part.karat, -prdwgt.part.wgt * 100))
            if self._from_his(pcode, prop):
                continue
            if prdwgt.aux and prdwgt.main.karat == 9925 and prdwgt.aux.karat == 925 and prdwgt.aux.wgt < 0.1:
                prdwgt = PrdWgt(WgtInfo(prdwgt.main.karat, prdwgt.main.wgt + prdwgt.aux.wgt), None, prdwgt.part)
            prop["mtlwgt"] = prdwgt._replace(netwgt=round(prop.get("netwgt", 0) - subs, 2))

    def _from_his(self, pcode, prop):
        '''
        @return True and setup the prop if there is valid history, else return False
        '''
        if not self._dao:
            return False
        pi = self._dao.get(pcode)
        if not pi or not pi[1]:
            return False
        pw = PrdWgt(None)
        for x in prop.get("mstr"):
            pw = addwgt(pw, WgtInfo(x[0], x[1]))
        pts, nl = [], _PajBomDAO.nl_mat
        for x in [x for x in prop.get("parts") if x.get("karat")]:
            pts.append(nl.newdata(True))
            # the flag is unknown, but need to be used by _new_boms(), so fill it as 'N'
            nl.id, nl.karat, nl.wgt, nl.flag = int(x["matid"]), int(x["karat"]), x["wgt"], "N"
        # extract the bom from pi with the netwgt removed
        pts, boms = _PajBomDAO.new_boms(None, pw, pts, _PajBomDAO.nl_mat), [x for x in pi[1] if x.mid]
        if len(boms) != len(pts):
            logger.debug("找到pcode(%s)半成品链历史记录, 但与当前BOM不一致，以当前为准", pcode)
            return False
        x = lambda x: (x.karat, x.wgt)
        pts, boms = [sorted(y, key=x) for y in (pts, boms)]
        pts = [1 for bom, nl in zip(boms, pts) if bom.mid != nl.mid or int(bom.karat) != int(nl.karat) or abs(float(bom.wgt) - nl.wgt) >= 0.01]
        if pts:
            return False
        boms, pi = pi[1], PrdWgt(None)
        for bi in boms:
            if bi.karat:
                pi = addwgt(pi, WgtInfo(bi.karat, float(bi.wgt)), isparts=(bi.flag == 0))
            else:
                pi = pi._replace(netwgt=float(bi.wgt))
        prop["mtlwgt"] = pi
        logger.debug("找到pcode(%s)半成品链历史记录", pcode)
        return True


    @staticmethod
    def _has_semi_chn(chns):
        """
        check if the chains contains semi-chain, that is, 成品链
        """
        for y in chns.values():
            if tofloat(y["length"]):
                return True
        return False

    def _read_book(self, wb, pmap, cvt2tbl=True):
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
            rdr(sht, pmap, cvt2tbl)

    @classmethod
    def _get_data(cls, sht_rng, nmp):
        vvs = sht_rng[1].end("left").expand("table").value
        return NamedLists(vvs, nmp)

    def _read_mstr(self, sht_rng, pmap, cvt2tbl=True):
        """ read the bom master to pmap(dict) """
        if cvt2tbl:
            sht_rng[0].name = "BOM_mstr"
            xwu.maketable(xwu.usedrange(sht_rng[0]), sht_rng[0].name)
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
            mstrs.add(key)
            kt = nl.fwgt
            if kt:
                netwgts[pcode] = netwgts.get(pcode, 0) + kt
            kt = self._parse_karat(nl.mat)
            if not kt:
                continue
            it = pmap.setdefault(pcode, {"pcode": pcode})
            it.setdefault("mtlwgt", []).append((kt, nl.mtlwgt))
        for pcode, kt in netwgts.items():
            pmap[pcode]["netwgt"] = round(kt, 2)

    def _read_pts(self, sht_rng, pmap, cvt2tbl=True):
        """ read parts from the sheet to pmap(dict) """
        if cvt2tbl:
            sht_rng[0].name = "BOM_part"
            xwu.maketable(xwu.usedrange(sht_rng[0]), sht_rng[0].name)
        pts = set()
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
            {
                "pcode": "十七,",
                "mtlwgt": "金银重,",
                "stwgt": "石头,"
            })
        nls = [
            x for x in xwu.NamedRanges(
                xwu.usedrange(mstr_sht), self._nmps["mstr"])
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
            return [(x.karat, x.wgt) if x else (0, 0) for x in prdwgt.wgts]

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
                eq = wts[0].karat == wts[0].karat or karatsvc.getfamily(wts[0].karat) == karatsvc.getfamily(wts[1].karat)
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
        qj = Query([JOhk.name.label("jono"), Styhk.name.label("styno"), JOhk.running]).join(Orderma, Orderma.id == JOhk.orderid).join(Styhk).filter(JOhk.createdate >= mindt).order_by(JOhk.createdate)
        with hksvc.sessionctx() as sess:
            cnt, ln = 0, len(pmap)
            for x in pmap.values():
                lst, wgt = [x["pcode"]], x["mtlwgt"]
                if isinstance(wgt, PrdWgt):
                    lst.extend(_fmtwgt((wgt)))
                else:
                    lst.extend((0, ) * 6)
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


class _PajBomDAO(object):
    ''' class help to cache and retrieve data for PajBomHdler
    caching PajItem and PajBom only
    '''

    _mat_nmp = {"id": "mid", "name": "mname", "karat": "pkarat", "wgt": "pwgt", "flag": "mpflag"}
    nl_mat = NamedList(tuple(_mat_nmp.keys()))
    def __init__(self, sessmgr):
        self._sessmgr = sessmgr

    def readsource(self, wb):
        '''
        read the commited result from the workbook as a list of dict.
        Only reading is perform, no validation/regulation will be perform
        '''
        tk = None
        def _ret_app(wb, tk):
            if not tk:
                return
            wb.close()
            xwu.appmgr.ret(tk)

        if isinstance(wb, str):
            if not path.exists(wb):
                return None
            app, tk = xwu.appmgr.acq()
            wb = app.books.open(wb)
        nms, sht = "Mkarat Image".split(), None
        for sht in wb.sheets:
            rngs = [xwu.find(sht, x) for x in nms]
            if all(rngs):
                break
            rngs = None

        if not rngs:
            _ret_app(wb, tk)
            return None
        nms, lsts = tuple(self._mat_nmp.values()), []
        nl_mat, phase = self.nl_mat, 1
        to_kt = lambda kt: karatsvc.getfamily(karatsvc.getkarat(kt)).karat if kt and isinstance(kt, str) else int(kt or 0)
        for nl in xwu.NamedRanges(xwu.usedrange(sht), {"jono": "JO#", "pcode": "Image"}):
            jn = nl.mkarat
            if jn and isinstance(jn, str):
                nl.mkarat = to_kt(jn)
            jn = JOElement.tostr(nl.jono)
            if jn:
                if phase > 0:
                    jn = esctext(jn)
                    phase, mp = 0, {"jono": jn, "pcode": nl.pcode, "mtlwgt": PrdWgt(WgtInfo(nl.mkarat, nl.mwgt))}
                    lsts.append(mp)
                else:
                    jn = None
            if nl.mid:
                lst = [nl.get(x) for x in nms]
                nl_mat.setdata(lst)
                nl_mat["id"] = int(nl_mat.id)
                nl_mat["karat"] = to_kt(nl_mat.karat)
                mp.setdefault("parts", []).append(lst)
            if jn:
                continue
            jn = triml(nl.rcat)
            phase = 1 if jn == "main" else 2 if jn == "part" else phase
            if isinstance(nl.mkarat, Number) and phase == 0:
                if nl.mkarat > 0:
                    mp["mtlwgt"] = addwgt(mp["mtlwgt"], WgtInfo(int(nl.mkarat), nl.mwgt))
                else:
                    mp["mtlwgt"] = mp["mtlwgt"]._replace(netwgt=nl.mwgt)
        _ret_app(wb, tk)
        return lsts, nl_mat

    def cache(self, wb):
        '''
        cache the given workbook
        @wb : A workbook object or a string point to the file
        @return : a list of cached objects
        '''
        res = self.readsource(wb)
        if not res:
            return
        res, nl_part = res[0], res[1]
        npis, nboms, n_oboms = [], [], []
        with ResourceCtx(self._sessmgr) as cur:
            for bom in res:
                pcode, jono = [bom[x] for x in "pcode jono".split()]
                pi, nbs = self.get(pcode), None
                if not pi:
                    pi = self._new_pi(pcode, jono)
                    npis.append(pi)
                else:
                    nbs = self.new_boms(pi[0], bom["mtlwgt"], bom["parts"], nl_part)
                    if self._is_same_boms(pi[1], nbs):
                        logger.debug("pcode(%s)'s boms are the same as before" % pcode)
                        continue
                    # save for history reference
                    if pi[1]:
                        logger.debug("pcode(%s)'s boms will be updated because there are changes" % pcode)
                    n_oboms.extend(self._apply_next_tag(pcode, pi[1], cur))
                    pi = pi[0]
                    pi.docno, pi.tag = jono, 0
                nboms.extend(nbs or self.new_boms(pi, bom["mtlwgt"], bom["parts"], nl_part))
            pi = 0
            for x in n_oboms:
                cur.add(x)
                pi += 1
            for x in npis:
                cur.add(x)
                pi += 1
            if pi:
                cur.flush()
            jono = 0
            for x in nboms:
                cur.add(x)
                jono += 1
            if pi + jono:
                cur.commit()
            jono = {}
            for x in nboms:
                jono.setdefault(x.item, []).append(x)
            return [(x[0], x[1]) for x in jono.items()]
        return None

    @staticmethod
    def _apply_next_tag(pcode, boms, cur):
        ''' return the next available tag for a pcode's history BOM '''
        q = Query(func.max(PajBom.tag)).join(PajItem).filter(PajItem.pcode == pcode).with_session(cur).first()
        nt, td = (q[0] or 0) + 1, datetime.today()
        for x in boms:
            x.tag, x.lastmodified = nt, td
        return boms

    @staticmethod
    def _is_same_boms(olds, news):
        if bool(olds) ^ bool(news):
            return False
        if not olds:
            return False
        cns = 'mid name karat wgt flag'.split()
        _eq = lambda x: abs(float(x[0]) - float(x[1])) < 0.01 if isinstance(x[0], Number) else x[0] == x[1]
        for o in zip(olds, news):
            for cn in cns:
                if not _eq(tuple(getattr(x, cn) for x in o)):
                    return False
        return True

    @staticmethod
    def _new_pi(pcode, jono):
        ''' create a new PajItem instance '''
        pi = PajItem()
        pi.pcode, pi.docno, pi.createdate, pi.tag = pcode, jono, datetime.today(), 0
        return pi

    @staticmethod
    def _get_boms(cur, pi):
        ''' the existing boms of given pajitem '''
        lst = cur.query(PajBom).filter(PajBom.pid == pi.id).all()
        return lst if lst else []

    @staticmethod
    def new_boms(pi, mtlwgt, parts, nl):
        ''' create a tuple of PajBom items, the mtlwgt will be also put as mat id = -karat '''
        for wgt in mtlwgt.wgts:
            if not wgt:
                continue
            parts.append(nl.newdata(True))
            nl.id, nl.name, nl.karat, nl.wgt, nl.flag = -wgt.karat, '_MAIN_', wgt.karat, wgt.wgt, 'Y'
        if mtlwgt.netwgt:
            parts.append(nl.newdata(True))
            nl.id, nl.name, nl.karat, nl.wgt, nl.flag = 0, '_NETWGT_', 0, mtlwgt.netwgt, 'Y'
        lsts, td = [], datetime.today()
        for x in sorted(parts, key=lambda x: x[0]):
            nl.setdata(x)
            if nl.flag not in ('Y', 'N'):
                continue
            bom = PajBom()
            lsts.append(bom)
            bom.item, bom.mid, bom.karat, bom.name, bom.wgt, bom.flag, bom.createdate,  bom.lastmodified, bom.tag = pi, nl.id, nl.karat, nl.name, nl.wgt, 1 if nl.flag == 'Y' else 0, td, td, 0
        return lsts

    def get(self, pcode):
        ''' return a tuple of PajItem and a list of PajBom or None '''
        pi = boms = None
        with ResourceCtx(self._sessmgr) as cur:
            pi = cur.query(PajItem).filter(PajItem.pcode == pcode).all()
            if pi:
                pids = [x.id for x in pi]
                # only get the current, history is for reference only
                boms = cur.query(PajBom).filter(and_(PajBom.pid.in_(pids), PajBom.tag == 0)).all()
                if boms:
                    boms = sorted(boms, key=lambda x: x.id)
                    pi = [x for x in pi if x.id == boms[0].pid][0]
                else:
                    pi = pi[0]
        return (pi, boms) if pi else None
