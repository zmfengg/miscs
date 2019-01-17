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

from sqlalchemy import and_
from sqlalchemy.orm import Query
from xlwings import Book
from xlwings.constants import (BorderWeight, Constants,
                               FormatConditionOperator, FormatConditionType,
                               LineStyle)

from hnjcore import JOElement, isvalidp17
from hnjcore.models.hk import JO as JOhk
from hnjcore.models.hk import Orderma, PajShp
from hnjcore.models.hk import Style as Styhk
from utilz import (NamedList, NamedLists, ResourceCtx, getfiles, getvalue,
                   karatsvc, tofloat, triml, xwu)
from utilz.xwu import NamedRanges
from utilz.xwu import appmgr as _appmgr
from utilz.xwu import col, find, findsheet, insertphoto

from .common import _logger as logger
from .localstore import PajBom, PajItem
from .pajcc import P17Decoder, PrdWgt, WgtInfo, addwgt


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
        kt, karat = None, None
        if is_mstr:
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

    def build_for_manul(self, fldr, **kwargs):
        '''
        from a shipment file, build an sheet for manully part judgement
        '''
        pmap, kxl = {}, None

        if isinstance(fldr, str):
            app, kxl = _appmgr.acq()
            fldr = app.books.open(fldr)
        jns = self._read_rpt(fldr)
        if not jns:
            return
        return
        self._read_book(fldr, pmap)
        self._adjust_wgts(pmap)

        def _fill_mstr(qnw, nl):
            nl.mkarat, nl.mwgt = qnw
        def _fill_dtl(dtl, nl):
            nl.mid, nl.mname, nl.pkarat, nl.pwgt = [dtl.get(x) for x in "matid name karat wgt".split()]
            nl.mpflag = 'N' if dtl.get("part_hints", False) else 'Y'
        fns = 'jono rcat mkarat mwgt mid mname pkarat pwgt mpflag image'.split()
        mcnt, nl = getvalue(kwargs, "min,min_rowcnt,rowcnt", 7), NamedList(fns)
        fns, mkrs = [fns], []
        for pcode, x in pmap.items():
            fn = []
            pts, mstrs = [x.get(y, []) for y in "parts mstr".split()]
            pts = [x for x in pts if x.get("karat")]
            lns = [len(x) for x in (mstrs, pts)]
            for idx in range(max(lns)):
                fn.append(nl.newdata(True))
                if idx == 0:
                    nl.image, nl.jono = pcode, "_JONO_"
                if idx < lns[0]:
                    _fill_mstr(mstrs[idx], nl)
                if idx < lns[1]:
                    _fill_dtl(pts[idx], nl)
            idx += 1
            if idx < mcnt:
                for idx in range(mcnt - idx):
                    fn.append(nl.newdata(True))
            mkrs.append((nl.jono, len(fns) + 1, len(fn) - 1,))
            fns.extend(fn)
        # TODO:: demo only
        if not kxl:
            app, kxl = _appmgr.acq()
        wb = app.books.add()
        self._write_manual(fns, mkrs, wb, nl)
        app.visible = True
        # if kxl :
        #    _appmgr.ret(kxl)
        return pmap

    def _read_rpt(self, wb):
        '''
        read the high-lighted wgt in rpt sheet
        return a {jn:styno} and a {pcode:jo}
        '''
        sht = findsheet(wb, "rpt")
        if not sht:
            return None
        rng, mkrs, idx = find(sht, "Wgt").expand("down"), [], -1
        for x in rng:
            if x.api.Interior.ColorIndex == 6:
                mkrs.append([idx,])
            idx += 1
        rng = [x for x in NamedRanges(sht.cells(1, 1), alias={"jono": "工单", "styno": "款号"})]
        for x in mkrs:
            fn, fns = x[0], None
            while fn >= 0:
                fns = rng[fn].styno
                if fns:
                    break
                fn -= 1
            x.extend((rng[fn].jono, fns))
        jns = {x[1]: x[2] for x in mkrs}
        return jns

    def _write_manual(self, lsts, mkrs, wb, nl):
        sht = wb.sheets[0]
        sht.cells(1, 1).value = lsts

        _col = lambda cn: nl.getcol(cn) + 1
        _cell = lambda r, cn: sht.cells(r, (nl.getcol(cn) + 1) if isinstance(cn, str) else cn)
        def _cond(api, con, clr=37):
            api.formatconditions.add(FormatConditionType.xlExpression, FormatConditionOperator.xlBetween, con)
            api.formatconditions(1).interior.colorindex = clr
        xwu.freeze(_cell(2, "mid"))
        _cell(2, "mpflag").select()
        # the Y/N validation
        # optional args need to be filled https://stackoverflow.com/questions/40722609/xlwings-range-api-validation-add-throws-com-error-2146827284
        idx, ln = nl.getcol("mpflag") + 1, len(lsts)
        rng = sht.range(_cell(2, idx), _cell(ln, idx)).api
        rng.Validation.Add(3, 1, 1, "Y,N")
        # Conditional formatting
        _cond(sht.range(_cell(2, "jono"), _cell(ln, "mwgt")).api, '=$%s2<>""' % col(_col("jono")))
        _cond(sht.range(_cell(2, "mid"), _cell(ln, "mpflag")).api, '=$%s2="Y"' % col(_col("mpflag")))
        sht.autofit()

        # formulas
        
        # images
        idx = col(_col("image"))
        for jono, frm, cnt in mkrs:
            # TODO:: get style name
            ln = r"\\172.16.8.91\Jpegs\style\BJ\BJ0\BJ001_580564.jpg"
            insertphoto(ln, sht.range("%s%d:%s%d" % (idx, frm, idx, frm + cnt)), margins=(2, 2))

    def _part_chk(self, bi, chns, lks, has_semi_chn):
        """
        determine if the given bom item is a part
        """
        return chns and bi["_id"] in chns or lks and (lks.get(bi["_id"]) or sum([1 for v in self._pts_kws if bi["name"].find(v) >= 0]))

    def _part_chk_l(self, bi, chns, lks, has_semi_chn):
        """
        loose rule for determining if the given bom item is a part
        """
        return chns and bi["_id"] in chns or lks and lks.get(bi["_id"]) or has_semi_chn and [1 for v in self._pts_kws if bi["name"].find(v) >= 0]

    def _adjust_wgts(self, pmap):
        if not pmap:
            return
        part_ck = self._part_chk if not self._part_chk_ver else self._part_chk_l
        for pcode, prop in pmap.items():
            if self._from_his(pcode, prop):
                continue
            ch_lks, prdwgt = prop.get("mtlwgt"), None
            if ch_lks:
                pmap[pcode]["mstr"] = ch_lks
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
                var = part_ck(y, chns, lks, has_semi_chn)
                if var:
                    #make sure part candidate has the same karat with old
                    wgt0 = prdwgt.part
                    var = not wgt0 or kt == wgt0.karat
                prdwgt = addwgt(prdwgt, WgtInfo(kt, y["wgt"]),\
                    var, autoswap=False)
                y["part_hints"] = var
            if has_semi_chn:
                prdwgt = prdwgt._replace(part=WgtInfo(prdwgt.part.karat, -prdwgt.part.wgt * 100))
            prop["mtlwgt"] = prdwgt._replace(netwgt=round(prop.get("netwgt", 0) - subs, 2))

    def _from_his(self, pcode, prop):
        ''' return the result from history if there is
        return True and setup the prop if there is history, else return False
        '''
        if not self._dao:
            return False
        pi = self._dao.get(pcode)
        if not pi:
            return False
        boms, pi = pi[1], PrdWgt(None)
        for bi in boms:
            pi = addwgt(pi, WgtInfo(bi.karat, float(bi.wgt)), isparts=(bi.flag == 0))
        prop["mtlwgt"] = pi
        return True


    @classmethod
    def _has_semi_chn(cls, chns):
        """
        check if the chains contains semi-chain, that is, 成品链
        """
        lc = 0
        for y in chns.values():
            lc = tofloat(y["length"])
            if lc:
                break
        return bool(lc)

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

    @classmethod
    def _get_data(cls, sht_rng, nmp):
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


class _PajBomDAO(object):
    ''' class help to cache and retrieve data for PajBomHdler
    caching PajItem and PajBom only
    '''
    def __init__(self, sessmgr):
        self._sessmgr = sessmgr

    def readsource(self, wb):
        ''' read the commited result from the workbook as a list of dict.
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
        phase = 0
        nls = [x for x in xwu.NamedRanges(xwu.usedrange(sht), alias={"jono": "JO#", "pcode": "Image"})]
        nmp = {"id": "mid", "name": "mname", "karat": "pkarat", "wgt": "pwgt", "flag": "mpflag"}
        nms, lsts, mp = tuple(nmp.values()), [], None
        nl_mat = NamedList(tuple(nmp.keys()))
        for nl in nls:
            jn = JOElement.tostr(nl.jono)
            if jn:
                if jn[0] == "'":
                    jn = jn[1:]
                phase, mp = 0, {"jono": jn, "pcode": nl.pcode, "mtlwgt": PrdWgt(WgtInfo(int(nl.mkarat), nl.mwgt))}
                lsts.append(mp)
            if nl.mid:
                lst = [nl.get(x) for x in nms]
                nl_mat.setdata(lst)
                nl_mat["id"], nl_mat["karat"] = [int(nl_mat[x]) for x in "id karat".split()]
                mp.setdefault("parts", []).append(lst)
            if jn:
                continue
            jn = triml(nl.rcat)
            phase = 1 if jn == "main" else 2 if jn == "part" else phase
            if nl.mkarat:
                if phase == 0:
                    mp["mtlwgt"] = addwgt(mp["mtlwgt"], WgtInfo(int(nl.mkarat), nl.mwgt))
        _ret_app(wb, tk)
        return lsts, nl_mat

    def cache(self, wb):
        ''' cache the given workbook
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
                pi = self.get(pcode)
                if not pi:
                    pi = self._new_pi(pcode, jono)
                    npis.append(pi)
                else:
                    if pi[0].docno == jono:
                        continue
                    else:
                        n_oboms.extend(pi[1])
                        pi = pi[0]
                        pi.docno, pi.tag = jono, 0
                nboms.extend(self._new_boms(pi, bom["mtlwgt"], bom["parts"], nl_part))
            pi = 0
            for x in n_oboms:
                cur.delete(x)
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


    def _new_pi(self, pcode, jono):
        ''' create a new PajItem instance '''
        pi = PajItem()
        pi.pcode, pi.docno, pi.createdate, pi.tag = pcode, jono, datetime.today(), 0
        return pi

    def _get_boms(self, cur, pi):
        ''' the existing boms of given pajitem '''
        lst = cur.query(PajBom).filter(PajBom.pid == pi.id).all()
        return lst if lst else []

    def _new_boms(self, pi, mtlwgt, parts, nl):
        ''' create a tuple of PajBom items, the mtlwgt will be also put as mat id = -karat '''
        for wgt in mtlwgt.wgts:
            if not wgt:
                continue
            parts.append(nl.newdata(True))
            nl.id, nl.name, nl.karat, nl.wgt, nl.flag = -wgt.karat, '_MAIN_', wgt.karat, wgt.wgt, 'Y'
        lsts = []
        for x in sorted(parts, key=lambda x: x[0]):
            nl.setdata(x)
            if nl.flag not in ('Y', 'N'):
                continue
            bom = PajBom()
            lsts.append(bom)
            bom.item, bom.mid, bom.karat, bom.name, bom.wgt, bom.flag = pi, nl.id, nl.karat, nl.name, nl.wgt, 1 if nl.flag == 'Y' else 0
        return lsts

    def get(self, pcode):
        ''' return a tuple of PajItem and a list of PajBom or None '''
        pi = boms = None
        with ResourceCtx(self._sessmgr) as cur:
            pi = cur.query(PajItem).filter(and_(PajItem.pcode == pcode, PajItem.tag == 0)).all()
            pi = pi[0] if pi else None
            if pi:
                boms = cur.query(PajBom).filter(PajBom.pid == pi.id).all()
                if boms:
                    boms = sorted(boms, key=lambda x: x.id)
        return (pi, boms) if pi else None