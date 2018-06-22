# coding=utf-8
'''
Created on 2018-04-28
classes to read data from C1's monthly invoices
need to be able to read the 2 kinds of files: C1's original and calculator file
@author: zmFeng
'''

import numbers
import os
from os import path
import re
import sys
from collections import namedtuple

from sqlalchemy import and_, func
from sqlalchemy.orm import Query
from xlwings import constants

from hnjcore import JOElement, karatsvc
from hnjcore.models.cn import JO, MM, Customer, MMgd, MMMa, Style
from hnjcore.utils import appathsep, daterange, getfiles, isnumeric, xwu

from .common import _date_short
from .common import _logger as logger


class InvRdr():
    """
        read the monthly invoices from both C1 version and CC version
    """

    def __init__(self, c1log=None, cclog=None):
        self._c1log = c1log
        self._cclog = cclog
        self._cnstqnw = "stqty,stwgt".split(",")
        self._cnsnl = "setting,labor".split(",")

    def read(self, fldr):
        """
        perform the read action 
        @param fldr: the folder contains the invoice files
        @return: a list of C1InvItem
        """

        if not os.path.exists(fldr):
            return
        if os.path.isfile(fldr):
            fns = [fldr]
        else:
            root = appathsep(fldr)
            fns = getfiles(root)
        if not fns:
            return
        killxw, app = xwu.app(False)
        wb = None
        try:
            cnsc1 = u"工单号,镶工,胚底,备注".split(",")
            cnscc = u"镶石费$,胚底费$,工单,参数,备注".split(",")
            for fn in fns:
                wb = app.books.open(fn)
                items = list()
                for sht in wb.sheets:
                    rngs = list()
                    for s0 in cnsc1:
                        rng = xwu.find(sht, s0, lookat=constants.LookAt.xlPart)
                        if rng:
                            rngs.append(rng)
                    if len(cnsc1) == len(rngs):
                        items.extend(self._readc1(sht, rngs[0].row))
                    else:
                        for s0 in cnscc:
                            rng = xwu.find(
                                sht, s0, lookat=constants.LookAt.xlWhole)
                            if rng:
                                rngs.append(rng)
                        if len(cnsc1) == len(rngs):
                            items.extend(self._readcalc(sht))
        finally:
            if killxw:
                app.quit()
        return items

    def _readc1(self, sht, hdrow):
        """
        read c1 invoice file
        @param   sht: the sheet that is verified to be the C1 format
        @param hdrow: the row of the header 
        @return: a list of C1InvItem with source = "C1"
        """
        rng = sht.range("A%d" % hdrow).end("left")
        rng = sht.range(sht.range(rng.row, rng.column),
                        xwu.usedrange(sht).last_cell)
        vvs = rng.value
        tr = vvs[0]
        km = {u"工单号": "jono", u"镶工": "setting", u"胚底,": "labor", u"备注,": "remark",
              u"数量": "joqty", u"石名称": "stname", u"粒数": "stqty", u"石重,": "stwgt"}
        tm = xwu.list2dict(tr, km)
        if len(tm) < len(km):
            logger.debug("key columns(%s) not found in sheet(%s)" %
                         (tm, sht.name))
            return None

        C1InvItem = namedtuple(
            "C1InvItem", "source,jono,qty,labor,setting,remarks,stones,parts")
        C1InvStone = namedtuple("C1InvStone", "stone,qty,wgt,remark")

        items = list()
        for ridx in range(1, len(vvs)):
            tr = vvs[ridx]
            s0 = tr[tm["jono"]]
            if isinstance(s0, numbers.Number):
                s0 = str(int(s0))
            je = JOElement(s0)
            if je.isvalid():
                snl = []
                for x in self._cnsnl:
                    a0 = tr[tm[x]]
                    snl.append(float(a0) if isnumeric(a0) else 0)
                if any(snl):
                    c1 = C1InvItem(
                        "C1", je.value, tr[tm["joqty"]], snl[1], snl[0], tr[tm["remark"]], [], "N/A")
                    items.append(c1)
            qnw = []
            for x in self._cnstqnw:
                if not isnumeric(tr[tm[x]]):
                    break
                qnw.append(float(tr[tm[x]]))
            if len(qnw) == 2:
                s0 = tr[tm["stname"]]
                if s0 and isinstance(s0, str):
                    joqty = c1.qty
                    c1.stones.append(C1InvStone(
                        tr[tm["stname"]], qnw[0] / joqty, qnw[1] / joqty, "N/A"))
        return items

    def _readcalc(self, sht):
        """
        read cc file
        @param   sht: the sheet that is verified to be the CC format
        @return: a list of C1InvItem with source = "CC"
        """
        # todo::missing
        cns = u"镶石费$,胚底费$,工单,参数,配件,笔电,链尾,分色,电咪,其它,银夹金,石料,形状,尺寸,粒数,重量,镶法,备注".split(
            ",")
        rng = xwu.find(sht, cns[0], lookat=constants.LookAt.xlWhole)
        x = xwu.usedrange(sht)
        rng = sht.range((rng.row, x.columns.count),
                        (x.last_cell().row, x.last_cell().column))
        vvs = rng.value


class C1JCReader(object):

    #return refid by running, from existing list or db#
    def _getrefid(self, runn, mpss):
        refid = None
        if mpss:
            for x in mpss:
                if x[0][0] <= runn and x[0][1] >= runn:
                    refid = x[1]
                    break
        if not refid:
            x = self._cnsvc.getjcrefid(runn)
            if x:
                mpss.append((x[1], x[0]))
                refid = x[0]
        return refid

    #return mps of given refid #
    def _getmps(self, refid, mpsmp):
        if refid not in mpsmp:
            mp = self._cnsvc.getjcmps(refid)
            mpsmp[refid] = mp
        if refid in mpsmp:
            return mpsmp[refid]

    def __init__(self, cnsvc, bcsvc, invfldr):
        """
        @param cnsvc: the CNSvc instance
        @param dbsvc: the BCSvc instance
        @param invfldr: folder contains C1's invoices, one example is 
            \\172.16.8.46\pb\dptfile\quotation\2017外发工单工费明细\CostForPatrick\AIO_F.xlsx
        """
        self._cnsvc = cnsvc
        self._bcsvc = bcsvc
        self._invfldr = invfldr

    def read(self, year, month, day=1, tplfn=None, tarfldr=None):
        """class to create the C1 JOCost file for HK accountant"""
        df, dt = daterange(year, month, day)
        refs, mpsmp, runns = [], {}, set()
        invs = InvRdr().read(self._invfldr)
        if not invs:
            logger.debug(
                "failed to read C1's invoice data from folder(%s)" % self._invfldr)
            return
        invs = dict([(x.jono, x) for x in invs])
        ptncx = re.compile(r"C(\d)$")
        with self._cnsvc.sessionctx() as cur:
            mmids, vvs, refs = set(), {}, []
            gccols = [
                x.split(",") for x in "goldwgt,goldcost;extgoldwgt,extgoldcost".split(";")]
            ttls = ("mmid,lastmmdate,jobno,cstname,styno,running,mstone,description,joqty"
                    ",karat,goldwgt,goldcost,extgoldcost,stonecost,laborcost,extlaborcost,extcost,"
                    "totalcost,unitcost,extgoldwgt,cflag").split(",")
            cnmap = xwu.list2dict(ttls)
            q = Query([JO.name.label("jono"), Customer.name.label("cstname"),
                       Style.name.label("styno"), JO.running, JO.karat.label(
                           "jokarat"), MMgd.karat,
                       MM.id, MM.name.label("docno"), MM.qty, func.sum(MMgd.wgt).label("wgt"), func.max(MMMa.refdate).label("refdate")]).\
                join(Customer).join(MM).join(MMMa).join(MMgd).join(Style).\
                group_by(JO.name, Customer.name, Style.name, JO.running, JO.karat, MMgd.karat, MM.id, MM.name, MM.qty).\
                filter(and_(and_(MMMa.refdate >= df, MMMa.refdate < dt),
                            MM.name.like("%C[0-9]")))
            lst = q.with_session(cur).all()
            vvs["_TITLE_"] = ttls
            for x in lst:
                jn = x.jono.value
                # if jn != "580356": continue
                if x.id not in mmids:
                    mmids.add(x.id)
                    if jn not in vvs:
                        ll = [x.id, "'" + x.refdate.strftime(_date_short), "'" + x.jono.value, x.cstname.strip(),
                              x.styno.value, x.running, "_ST", "_EDESC", float(
                                  x.qty), karatsvc.getfamily(x.jokarat).karat, [],
                              0, 0, 0, 0, 0, 0, 0, 0, 0, "NA"]
                        mt = ptncx.search(x.docno)
                        if mt:
                            ll[cnmap["cflag"]] = "'" + mt.group()
                        vvs[jn] = ll
                        runns.add(int(x.running))
                    else:
                        vvs[jn][cnmap["joqty"]] += float(x.qty)
                vvs[jn][cnmap["goldwgt"]].append(
                    (karatsvc.getfamily(x.karat).karat, x.wgt))
            bcs = self._bcsvc.getbcsforjc(runns)
            bcs = dict([(x.runn, (x.desc, x.ston)) for x in bcs])
            for x in vvs.values():
                # the title
                if x[0] == ttls[0]:
                    continue
                joqty = x[cnmap["joqty"]]
                runn = str(x[cnmap["running"]])
                if runn in bcs:
                    dns = bcs[runn]
                    x[cnmap["description"]], x[cnmap["mstone"]] = dns[0], dns[1]

                runn = x[cnmap["jobno"]][1:]
                if runn in invs:
                    inv = invs[runn]
                    x[cnmap["laborcost"]] = (inv.setting + inv.labor) * joqty
                else:
                    logger.info("no labor data from c1 invoice(%s) for JO(%s)" %
                                (os.path.basename(self._invfldr), runn))

                lst1 = x[cnmap["goldwgt"]]
                if len(lst1) > 1:
                    mmids = {}
                    for knw in lst1:
                        if knw[0] in mmids:
                            mmids[knw[0]] += knw[1]
                        else:
                            mmids[knw[0]] = knw[1]
                    kt = x[cnmap["karat"]]
                    x[cnmap["goldwgt"]] = [(kt, mmids[kt])]
                    del mmids[kt]
                    if len(mmids) > 0:
                        x[cnmap["extgoldwgt"]] = list(mmids.items())
                lst1 = x[cnmap["goldwgt"]]
                refid = self._getrefid(x[cnmap["running"]], refs)
                if not refid:
                    logger.critical(("No refid found for running(%d),"
                                     " Pls. create one in codetable with (jocostma/costrefid) ") % x[cnmap["running"]])
                    vvs = None
                    break
                mp = self._getmps(refid, mpsmp)
                for vv in gccols:
                    lst1 = x[cnmap[vv[0]]]
                    if lst1:
                        ttlwgt = 0.0
                        for knw in lst1:
                            kt = knw[0]
                            # in extgold case, if different karat exists, the weight is merged
                            ttlwgt += float(knw[1])
                            if kt not in mp:
                                logger.critical("No MPS found for running(%d)'s karat(%d)" %
                                                (x[cnmap["running"]], kt))
                                x[cnmap[vv[1]]] = -1000
                            else:
                                x[cnmap[vv[1]]
                                  ] += round(float(mp[kt]) * float(knw[1]), 2)
                        x[cnmap[vv[0]]] = ttlwgt
                if vvs == None:
                    break
        return list(vvs.values()) if vvs else None

    def readst(self, fldr):
        if not path.exists(fldr): return
        if path.isfile(fldr):
            fns = [fldr]
        else:
            fns = getfiles(fldr,"xls")
        kxl, app = xwu.app(False)
        try:
            for fn in fns:
                wb = app.books.open(fn)
                for sht in wb.sheets:
                    if sht.name == u"进":
                        vvs = sht.range("A1").expand("table").value
                        #TODO::                        
                    elif sht.name == u"用":
                        pass
                wb.close()
        finally:
            if kxl: app.quit()
