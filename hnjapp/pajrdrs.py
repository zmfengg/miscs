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
from os import path

from sqlalchemy import and_, func
from sqlalchemy.orm import Query
from xlwings.constants import (
    LookAt,
    FormatConditionOperator,
    FormatConditionType,
)

from hnjcore import JOElement, isvalidp17
from hnjcore.models.hk import JO as JOhk
from hnjcore.models.hk import Orderma, PajCnRev, PajInv, PajShp
from hnjcore.models.hk import Style as Styhk
from hnjcore.utils.consts import NA
from utilz import (NamedList, NamedLists, ResourceCtx, splitarray,
                   daterange, getfiles, isnumeric, appathsep, deepget, karatsvc,
                   trimu, xwu, updateopts, stsizefmt)
from utilz.xwu import find, findsheet, NamedRanges, insertphoto, col

from .common import _getdefkarat
from .common import _logger as logger, P17Decoder
from .localstore import PajCnRev as PajCnRevSt
from .localstore import PajInv as PajInvSt
from .localstore import PajItem as PajItemSt
from .localstore import PajWgt as PrdWgtSt
from .pajbom import PajBomHdlr, _PajBomDAO as PajBomDAO
from .pajcc import (MPS, PAJCHINAMPS, PajCalc, PrdWgt, WgtInfo,
                    _tofloat, addwgt)
from .dbsvcs import HKSvc
from .miscsvcs import StylePhotoSvc

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


PajShpItem = namedtuple(
        "PajShpItem",
        "fn,orderno,jono,qty,pcode,invno,invdate,mtlwgt,stwgt,shpdate,lastmodified,filldate"
    )

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
    def _read_stone_data(cls, sht):
        """
        extract stone data from quo sheet
        @param sht:  the DL_QUOTATION sheet that need to read data from
        """
        rng = xwu.find(sht, "Item*No", lookat=LookAt.xlPart)
        if not rng:
            return
        # because there is merged cells rng.expand('table').value
        # or sht.range(rng.end('right'),rng.end('down')).value failed
        _ptngwt = re.compile(r"[\d.]+")
        vvs = sht.range(rng, rng.current_region.last_cell).value
        qmap = {}
        nls = NamedLists(vvs, {
            "pcode": "Item,",
            "stone": "stone",
            "stshape": 'Stone\nShape',
            'stsize': 'Stone Size',
            'stwgt': 'Stone Weight'
        }, False)
        sm = _StMaker()
        for tr in nls:
            p17 = tr.pcode
            if not p17:
                continue
            if not isvalidp17(p17):
                continue
            st = tr.stone
            if not st or trimu(st).find('NO STONE') >= 0:
                continue
            qmap[p17] = sm.make_stones(tr.stshape, tr.stone, tr.stsize, tr.stwgt)
        return qmap

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
            return None
        if idx == 2:
            dups.append(invno)
        items = self.read_inv_raw(sht, invno, fmd)
        return items, dups

    @staticmethod
    def _is_ring(pcode):
        return isvalidp17(pcode) and pcode[1] == "4"

    @staticmethod
    def _getbomwgt(bomap, bomapring, pcode):
        """ in the case of ring, there is only one code there
        """
        if not (bomap and pcode):
            return None
        prdwgt = bomap.get(pcode)
        if not prdwgt:  # and is ring
            if bomapring and PajShpHdlr._is_ring(pcode):
                pcode0 = pcode
                pcode = PajShpHdlr._extring(pcode)
                prdwgt = bomapring.get(pcode)
                pcode = pcode0
        if not prdwgt:
            logger.debug("failed to get bom wgt for pcode(%s)" % pcode)
        return prdwgt

    @staticmethod
    def _str2date(s_date):
        if isinstance(s_date, str):
            s_date = datetime.strptime(s_date, "%Y-%m-%d").date()
        return s_date

    @staticmethod
    def _extring(x):
        return x[:8] + x[10:]

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

        bfn = path.basename(fn).replace("_", "")
        shd = PajShpHdlr.get_shp_date(sht.name, False) or fshd
        # when sheet's shpdate differs from file's shpdate, use the maximum one
        shd = max(shd, fshd)
        if bomwgts is None:
            bomwgts = PajBomHdlr().readbom(sht.book)
        if bomwgts:
            bomwgtsrng = {
                cls._extring(x[0]): x[1]["mtlwgt"]
                for x in bomwgts.items()
                if cls._is_ring(x[0])
            }
            bomwgts = {x[0]: x[1]["mtlwgt"] for x in bomwgts.items()}
        else:
            bomwgtsrng, bomwgts = (None,) * 2
        if not th.getcol("cost"):
            cls._read_order(fmd, td0, bomwgtsrng, locals())
        else:
            cls._read_sample(locals())
        return items

    @staticmethod
    def _read_order(fmd, td0, bomwgtsrng, kwds):
        th, items, bfn, bomwgts = [kwds.get(x) for x in 'th items bfn bomwgts'.split()]
        for tr in kwds['nls']:
            if not tr.pcode:
                break
            if not tr.odseq or tr.odseq[:2] == "CR" or not isvalidp17(
                    tr.pcode):
                logger.debug("repairing(%s) item found, skipped", tr.pcode)
                continue
            jono = tr.jono
            mwgt = PajShpHdlr._getbomwgt(bomwgts, bomwgtsrng, tr.pcode)
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
                ivd = PajShpHdlr._str2date(tr.invdate)
                si = PajShpItem(bfn, odno, jono, tr.qty, tr.pcode, invno,
                                ivd, mwgt, stwgt, ivd, fmd, td0)
                items[thekey] = si

    @staticmethod
    def _read_sample(kwds):
        # new sample case, extract weight data from the quo sheet, but deprecated
        # get from bom instead
        items, fmd, td0, bfn, bomwgts, bomwgtsrng = [kwds.get(x) for x in 'items fmd td0 bfn bomwgts bomwgtsrng'.split()]
        for tr in kwds['nls']:
            # no cost item means repairing
            if not tr.get("cost"):
                continue
            p17 = tr.pcode
            if not p17:
                break
            ivd, odno = PajShpHdlr._str2date(tr.invdate), tr.get("orderno", NA)
            prdwgt = PajShpHdlr._getbomwgt(bomwgts, bomwgtsrng, p17)
            if not prdwgt:
                prdwgt = PrdWgt(WgtInfo(0, 0))
            mtl_stone = (0, 0)
            si = PajShpItem(bfn, odno,
                            JOElement(tr.jono).value, tr.qty, p17, tr.invno,
                            ivd, prdwgt, mtl_stone[1], ivd, fmd, td0)
            # new sample won't have duplicated items
            items[random.random()] = si
        sht = [x for x in kwds['sht'].book.sheets if trimu(x.name).find('QUOTATION') >= 0]
        if sht:
            sts = PajShpHdlr._read_stone_data(sht[0])
            #TODO:send the results to pajshpitem's stones

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
                jns = {x.name: x for x in jns}
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
                    bomwgts = PajBomHdlr().readbom(wb)
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

    @staticmethod
    def build_bom_sheet(fn, **kwds):
        ''' build the bom check sheet based on the shipment file(with rpt/bc sheet inside)
        @param min_rowcnt: the minimum rows per item occupied, default is 7
        @param min_offset: the starting point of the main metal, default is 3
        @param bom_check_level: 0 for strict, 1 for loose. default is 1
        '''
        updateopts({"min_rowcnt": 7, "main_offset": 3, "bom_check_level": 1}, kwds)
        return _BomSheetBldr(**kwds).build(fn)


class _BomSheetBldr(object):
    '''
    class help to build a bom sheet for manual part check
    '''
    def __init__(self, **kwds):
        self._min_rowcnt, self._main_offset, self._bom_check_level = kwds.get("min_rowcnt", 7), kwds.get("main_offset", 3), kwds.get("bom_check_level", 1)

    def build(self, wb, **kwds):
        app = kxl = fn = None
        if isinstance(wb, str):
            fn = wb
            app, kxl = _appmgr.acq()
            wb = app.books.open(wb)
        else:
            fn = wb.fullname
        jns = self._read_request(wb)
        if not jns:
            return None
        pmp, td = {}, date.today()
        logger.debug("reading shipment data")
        for sht in wb.sheets:
            var = PajShpHdlr.read_shp(fn, td, td, sht, None)
            if var:
                # new sample has different storage format
                if isinstance(iter(var.keys()).__next__(), numbers.Number):
                    var = {"%s,%s,%s" % (x.jono, x.pcode, x.invno): x for x in var.values()}
                pmp.update(var)
        if pmp:
            logger.debug("totally %d shipment records returned" % len(pmp))
        else:
            logger.debug("no shipment data returned")
            return None
        td = {}
        for x in [x.split(",") for x in pmp if isinstance(x, str)]:
            if x[0] in jns:
                td.setdefault(x[1], []).append((
                    x[0],
                    jns[x[0]],
                ))
        pmp = self._read_bc_wgts(wb, td)
        logger.debug("begin to read bom")
        fns, mkrs, nl, mf_rc = PajBomHdlr(part_chk_ver=self._bom_check_level).readbom_manual(
            wb, td, main_offset=self._main_offset, min_rowcnt=self._min_rowcnt,
            bc_wgts=pmp)
        if fns:
            logger.debug("toally %d bom records returned" % len(fns))
        if not kxl:
            app, kxl = _appmgr.acq()
        if kwds.get("new_book"):
            wb = app.books.add()
        self._write_manual(fns, mkrs, wb, nl, mf_rc, ps=None, bcwgts=pmp, new_sheet=not kwds.get("new_book"))
        return wb

    @staticmethod
    def _read_bc_wgts(wb, pcodes):
        #read the bcdata if there is
        bcwgts = findsheet(wb, 'BCData')
        bcmp, bcwgts = {x.jobn: x for x in NamedRanges(bcwgts.cells(1, 1))}, {}
        for x in pcodes.values():
            jn = x[0][0]
            bc = bcmp[jn]
            pw = PrdWgt(WgtInfo(karatsvc.getkarat(int(bc.gmas)).karat, float(bc.gwgt)))
            rms = [bc["rem%d" % x] for x in range(1, 4) if bc["rem%d" % x] and bc["rem%d" % x][0] == "*"]
            if rms:
                for kt, wgt in [x.split() for x in rms]:
                    idx = kt.find("PTS")
                    kt = kt[1:idx if idx >= 0 else None]
                    pw = addwgt(pw, WgtInfo(karatsvc.getkarat(kt).karat, float(wgt)), idx >= 0)
            bcwgts[jn] = pw
        return bcwgts

    def _read_request(self, wb):
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
                mkrs.append([
                    idx,
                ])
            idx += 1
        rng = [
            x for x in NamedRanges(
                sht.cells(1, 1), alias={
                    "jono": "工单",
                    "styno": "款号"
                })
        ]
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

    def _write_manual(self, lsts, mkrs, wb, nl, mf_rc, **kwds):
        sht = wb.sheets.add("BOM_Check", after=wb.sheets[-1]) if kwds.get("new_sheet") else wb.sheets[0]
        # this is necessary because conditional criterias need to select target range
        sht.activate()
        sht.cells(1, 1).value = lsts

        _col = lambda cn: nl.getcol(cn) + 1
        _cell = lambda r, cn: sht.cells(r, (nl.getcol(cn) + 1) if isinstance(cn, str) else cn)

        _cols = lambda cn, f, t: "%s%d:%s%d" % (col(_col(cn)), f, col(_col(cn)), t - 1)
        xwu.freeze(_cell(2, "mid"))
        sht.autofit()
        sht.cells(1, _col('mname')).column_width = 28

        # Y/N validation
        idx, ln = nl.getcol("mpflag") + 1, len(lsts)
        rng = sht.range(_cell(2, idx), _cell(ln, idx)).api
        rng.Validation.Add(3, 1, 1, "Y,N")
        # Conditional formatting
        self._cond(
            sht.range(_cell(2, "jono"), _cell(ln, "mwgt")).api,
            '=$%s2="_NO_BOM_"' % col(_col("mname")), 3)
        self._cond(
            sht.range(_cell(2, "jono"), _cell(ln, "mwgt")).api,
            '=$%s2<>""' % col(_col("jono")))
        self._cond(
            sht.range(_cell(2, "mid"), _cell(ln, "mpflag")).api,
            '=$%s2="_NO_BOM_"' % col(_col("mname")), 3)
        self._cond(
            sht.range(_cell(2, "mid"), _cell(ln, "mpflag")).api,
            '=$%s2="Y"' % col(_col("mpflag")))
        # images and formula
        idx = col(_col("image"))
        bfs = {
            True:
            '=IF(ISBLANK(%(ref)s),"",SUMIF(%(mkarat)s,%(ref)s,%(mwgt)s) + SUMIFS(%(pwgt)s,%(pkarat)s,%(ref)s,%(mpflag)s,"Y"))',
            False:
            '=SUMIFS(%(pwgt)s,%(pkarat)s,%(ref)s,%(mpflag)s,"N")'
        }
        def _formula(idx, bf, mp, is_main=True):
            mp["ref"] = "%s%d" % (col(_col("mkarat")), idx)
            sht.cells(idx, _col('mwgt')).formula = bf % mp
            if not is_main:
                return
            # excel index - 1 == array index
            nl.setdata(lsts[idx])
            if nl.mkarat:
                mp["ref"] = "%s%d" % (col(_col("mkarat")), idx + 1)
                sht.cells(idx + 1, _col('mwgt')).formula = bf % mp
        bcwgts, ps = [kwds.get(x) for x in "bcwgts ps".split()]
        if not ps:
            ps = StylePhotoSvc()
        for jono, styno, frm, cnt in mkrs:
            rng, ln = frm + cnt + 1, frm + mf_rc[0]
            rng = {
                "mkarat": _cols('mkarat', frm, ln),
                "mwgt": _cols('mwgt', frm, ln),
                "pkarat": _cols('pkarat', frm, rng),
                "pwgt": _cols('pwgt', frm, rng),
                "mpflag": _cols('mpflag', frm, rng)
            }
            _formula(frm + mf_rc[0], bfs[True], rng)
            _formula(frm + mf_rc[0] + 2, bfs[False], rng, False)
            # for the bc validation, high-light differences
            if bcwgts:
                ln = frm + self._main_offset
                self._cond(
                    sht.range(_cell(ln, "mkarat"), _cell(ln + 3, "mwgt")).api,
                    "=" + _cols("mkarat", ln, ln + 5).replace(":", "<>"), 3)
            ln = ps.getPhotos(styno, hints=jono)
            if ln:
                insertphoto(
                    ln[0],
                    sht.range("%s%d:%s%d" % (idx, frm, idx, frm + cnt)),
                    margins=(2, 2), alignment="L,T")
        _cell(2, "mpflag").select()

    @staticmethod
    def _cond(api, con, clr=37):
        # https://support.microsoft.com/en-us/help/895562/the-conditional-formatting-may-be-set-incorrectly-when-you-use-vba-in
        try:
            xwu.apirange(api).select()
            api.formatconditions.add(FormatConditionType.xlExpression,
                                        FormatConditionOperator.xlEqual, con)
            api.formatconditions(api.formatconditions.count).interior.colorindex = clr
        except:
            pass


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
        cnt, sz = 0, 50
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
            pi.pcode, pi.docno, pi.createdate, pi.tag = pcode, kwds.get(
                "docno") or NA, kwds["td"], 0
            curs[0].add(pi)
            curs[0].flush()
        var = Query([
            PajShp.pcode,
            JOhk.name.label("jono"),
            Styhk.name.label("styno"), JOhk.createdate, PajShp.invdate,
            PajInv.uprice, PajInv.mps
        ])
        var = var.join(JOhk).join(Orderma).join(Styhk).join(
            PajInv,
            and_(PajShp.joid == PajInv.joid, PajShp.invno == PajInv.invno))
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
    def _get_tar_pis(cls, pcodes, cur_src):
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
                var.extend(
                    cur_tar.query([PajItemSt,
                                   PajCnRevSt]).join(PajCnRevSt).filter(
                                       PajCnRevSt.id.in_(arr)).all())
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
                    cur.query(PajCnRevSt).delete()
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
            PajInvSt.cn, PajInvSt.otcost
        ])
        q0 = q0.join(PajInvSt).order_by(PajItemSt.pcode).order_by(
            PajInvSt.invdate)
        if pcodes:
            lst = []
            for mp in splitarray(pcodes):
                try:
                    lst.extend(
                        q0.filter(
                            PajItemSt.pcode.in_(mp)).with_session(cur).all())
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
            ctss = (
                "cn invdate".split(),
                "oc cn jono invdate".split(),
            )
            shts, pd = [], P17Decoder()
            app = xwu.appmgr.acq()[0]
            wb = app.books.add()
            for x in grouped_data.items():
                shts.append(wb.sheets.add())
                sht = shts[-1]
                sht.name, vvs = mp.get(x[0],
                                       "_" + str(random.randint(0, 99999))), []
                gidx = len(grps[2])
                ttl0, ttl1 = [
                    None,
                ] * gidx, [
                    None,
                ] * gidx
                gidx = 0 if x[0] == "noaff" else 1
                for z in grps[0][gidx]:
                    ttl0.append(z)
                    ttl0 += [
                        None,
                    ] * (len(ctss[gidx]) * len(grps[1]) - 1)
                    for xx in grps[1]:
                        ttl1.append(xx)
                        ttl1 += [
                            None,
                        ] * (len(ctss[gidx]) - 1)
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

class _StMaker(object):
    _ptn_stwgt = re.compile(r'\d[\.]?\d*')
    _stn_div = ('*', '#', '(', '-')
    
    @classmethod
    def make_stones(cls, shps, sts, szs, wgts):
        '''
        make a tuple of stone data
        '''
        if not sts:
            return None
        sts = [cls._split(x) for x in (sts, wgts, szs)]
        if not all(sts):
            return None
        sts, wgts, szs = sts
        lst = []
        if not shps:
            shps = ('R', ) * len(sts) #default is round
        else:
            shps = shps.split('\n')
            while len(shps) < len(sts):
                shps.append(shps[-1])
        for idx, st in enumerate(sts):
            if not st:
                break
            qnz = cls._extract_qns(szs[idx])
            lst.append((qnz[0], cls._extract_shape(shps[idx]), cls._extract_name(st), qnz[1], cls._extract_wgt(wgts[idx])))
        return lst

    @staticmethod
    def _split(ss):
        for x in ('\n', '<br />', ):
            rc = ss.split(x)
            if len(rc) > 1:
                return rc
        return (ss, )

    @staticmethod
    def _extract_shape(shp):
        if not shp:
            return 'R'
        # TODO:: some normalization
        return shp

    @classmethod
    def _extract_wgt(cls, s0):
        '''
        one format is numeric only, the other is 'ttl \d[\.]?\d* cts'
        '''
        mt = cls._ptn_stwgt.search(s0)
        return float(mt.group()) if mt else -1

    @classmethod
    def _extract_name(cls, st):
        '''
        extract the stone name out from name by PAJ
        '''
        #dia is special, it use '-' to tell specify the grade while _stn_div contains '-', which make dia testing fail. So hardcode it here
        if st[:3] == 'DIA':
            return 'DIA'
        rc = cls._extract_name_all(st, cls._stn_div)
        if not rc:
            lst = []
            rc = trimu(st).split()
            for idx, x in enumerate(rc):
                pos = [ii for ii, y in enumerate(x) if y < 'A' or y > 'Z']
                if pos and not (idx == 0 and pos[0] == 0):
                    break
                lst.append(x)
            rc = ' '.join(lst)
        return rc

    @classmethod
    def _extract_name_all(cls, st, divs):
        if not divs:
            return st
        rc = None
        for ii, s in enumerate(divs):
            idx = st.find(s)
            if idx > 0: # if div appear in at the first place, ignore it
                rc = trimu(st[:idx])
                rc = cls._extract_name_all(rc, divs[ii+1:]) or rc
                break
        return rc

    @staticmethod
    def _extract_qns(sz):
        '''
        extract the qty and size out from the size string
        '''
        for idx in range(len(sz) - 1, -1, -1):
            ch = sz[idx]
            if ch < '0' or ch > '9':
                break
        if idx > 0:
            ch = sz[idx]
            if ch in ('m', 'M'):
                sz, qty = stsizefmt(sz[:idx + 1]), int(sz[idx + 1:])
            elif ch == '-':
                idx = idx * 2 + 1
                sz, qty = sz[:idx], int(sz[idx])
            elif ch == '"':
                sz, qty = sz[:idx], int(sz[idx + 1:])
            elif ch in ('p', 'P'):
                sz, qty = sz[:4], int(sz[4:])
        else:
            return (None, ) * 2
        return qty, "'" + sz
