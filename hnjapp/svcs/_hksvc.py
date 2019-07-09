# coding=utf-8
"""
 * @Author: zmFeng
 * @Date: 2018-05-25 14:21:01
 * @Last Modified by:   zmFeng
 * @Last Modified time: 2018-05-25 14:21:01
 * the database services, including HK's and py's, and the out-dated bc's
 """

import datetime
import re
from collections import Iterable
from collections.abc import Sequence
from operator import attrgetter

from sqlalchemy import and_, desc, or_
from sqlalchemy.orm import Query

from hnjcore import JOElement, KaratSvc, StyElement
from hnjcore.models.hk import JO
from hnjcore.models.hk import Invoice as IV
from hnjcore.models.hk import InvoiceItem as IVI
from hnjcore.models.hk import JOItem as JI
from hnjcore.models.hk import Orderma, PajCnRev, PajInv, PajShp, POItem
from hnjcore.models.hk import StockObjectMa as SO
from hnjcore.models.hk import Style
from utilz import splitarray

from .. import pajcc
from ..common import Utilz
from ..common import _logger as logger
from ..common import config
from ..pajcc import MPS, PrdWgt, WgtInfo, addwgt
from ._common import SvcBase, _getjos, idset, idsin, fmtsku

_ptn_mit = re.compile(config.get("pattern.mit"))


class HKSvc(SvcBase):
    _qcaches = {}
    _ktsvc = KaratSvc()

    def _samekarat(self, jo0, jo1):
        """ check if the 2 given jo's karat/auxkarat are the same
            this method compare the 2 karats
        """
        lst = ([(self._ktsvc.getfamily(x.orderma.karat),
                 self._ktsvc.getfamily(x.auxkarat) if x.auxwgt else 0)
                for x in (jo0, jo1)])
        return lst[0] == lst[1]

    def __init__(self, sqleng):
        """ init me with a sqlalchemy's engine """
        super(HKSvc, self).__init__(sqleng)
        # self._ptnmit = re.compile("M[iI]T")

    def _pjq(self):
        """ return a JO -> PajShp -> PajInv query, append your query
        before usage, then remember to execute using q.with_session(yousess)
        """
        return self._qcaches.setdefault(
            "jopajshp&inv",
            Query([PajShp, JO, PajInv]).join(JO, JO.id == PajShp.joid).join(
                PajInv,
                and_(PajShp.joid == PajInv.joid, PajShp.invno == PajInv.invno)))

    def _pjsq(self):
        """ return a cached Style -> JO -> PajShp -> PajInv query """
        return self._qcaches.setdefault(
            "jopajshp&inv",
            Query([PajShp, JO, Style, Orderma, PajInv]).join(
                JO, JO.id == PajShp.joid).join(
                    PajInv,
                    and_(PajShp.joid == PajInv.joid,
                         PajShp.invno == PajInv.invno)))

    def getjos(self, jesorrunns, extfltr=None):
        """get jos by a collection of JOElements/Strings or Integers
        when the first item is string or JOElement, it will be treated as getbyname, else by runn
        return a tuple, the first item is list containing hnjcore.models.hk.JO
                        the second item is a set of ids/jes/runns not found
        @param groupby: can be one of id/running/name, running should be a string
            starts with 'r' for example, 'r410100', id should be integer,
            name should be JOElement or string without 'r' as prefix
        """
        return _getjos(self, JO, Query(JO), jesorrunns, extfltr)

    def getjis(self, jos):
        with self.sessionctx() as cur:
            q = Query((
                JO.id,
                JI.stname,
                JI.stsize,
                JI.qty,
                JI.unitwgt.label("wgt"),
                JI.remark,
            )).join(JI)
            lst = q.filter(idsin(idset(jos), JO)).with_session(cur).all()
            if not lst:
                return None
            jis = {}
            for x in lst:
                jis.setdefault(x[0], []).append(x)
        return jis

    def getjo(self, jeorrunn):
        """ a convenient way for getjos """
        with self.sessionctx():
            jos = self.getjos([jeorrunn])
        return jos[0][0] if jos else None

    def getjocatetory(self, jo):
        return Utilz.getStyleCategory(jo.style.name.value, jo.description)

    def getrevcns(self, pcode, limit=0):
        """ return a list of revcns order by the affected date desc
        The current revcn is located at [0]
        @param limit: only return the given count of records. limit == 0 means no limits
        """
        with self.sessionctx() as cur:
            rows = None
            q = Query(PajCnRev).filter(PajCnRev.pcode == pcode).order_by(
                desc(PajCnRev.tag))
            if limit:
                q = q.limit(limit)
            rows = q.with_session(cur).all()
            if rows:
                rows = sorted(rows, key=attrgetter("tag"))
        return rows

    def getrevcn(self, pcode, calcdate=None):
        """ return the revcn of given calcdate, if no calcdate provided, return the current
        revcn that affect current items
        @param calcdate: the date you want the revcn to affect
        """
        rc = None
        with self.sessionctx() as cur:
            if not calcdate:
                rows = self.getrevcns(pcode, limit=1)
                if rows and rows[0] == 0:
                    rc = rows[0]
            else:
                q = Query(PajCnRev).filter(PajCnRev.pcode == pcode).filter(
                    PajCnRev.revdate <= calcdate)
                q = q.order_by(desc(PajCnRev.revdate)).limit(1)
                rc = q.with_session(cur).one()
        return rc.uprice if rc else 0

    def findsimilarjo(self, jo, level=1, mindate=datetime.datetime(2015, 1, 1)):
        """ return an list of JO based on below criteria
        @param level:   0 for extract SKU match
                        1 for extract karat match
                        1+ for extract style match
        @param mindate: the minimum date to fetch data
        """
        rc = None
        level = 0 if level < 0 else level
        je = jo.name
        jns, jn = None, None
        with self.sessionctx() as cur:
            #don't lookup too much, only return data since 2015
            q = Query([JO,POItem.skuno]).join(Orderma).join(POItem,POItem.id == JO.poid).filter(Orderma.styid == jo.orderma.style.id)
            if mindate:
                q = q.filter(JO.createdate >= mindate)
            try:
                rows = q.with_session(cur).all()
                if (rows):
                    jns = {}
                    for x in rows:
                        jn = x.JO.name
                        if jn == je:
                            continue
                        key = fmtsku(x.JO.po.skuno)
                        jns.setdefault(key, []).append(x.JO)
                if jns:
                    skuno = fmtsku(jo.po.skuno)
                    sks = jns[skuno] if skuno and skuno in jns else None
                    if not sks and level > 0:
                        sks = [
                            x.JO
                            for x in rows
                            if je != x.JO.name and self._samekarat(jo, x.JO)
                        ]
                        if not sks and level > 1:
                            sks = [x.JO for x in rows]
                    rc = sks
            except UnicodeDecodeError as e:
                logger.debug(
                    "description/edescription/po.description of JO#(%s) contains invalid Big5 character "
                    % jn.value)
        if rc and len(rc) > 1:
            rc = sorted(rc, key=attrgetter("createdate"), reverse=True)
        return rc

    def getjowgts(self, jo):
        """ return a PrdWgt object of given JO """
        if not jo:
            return None

        with self.sessionctx() as cur:
            if isinstance(jo, str) or isinstance(jo, JOElement):
                jo = self.getjos([jo])[0]
                if not jo:
                    return None
                jo = jo[0]
            knws = [WgtInfo(jo.orderma.karat, jo.wgt), None, None]
            rk, oth_chn = knws[0], []
            if jo.auxwgt:
                knws[1] = WgtInfo(jo.auxkarat, float(jo.auxwgt))
                if (knws[1].karat == 925):  # most of 925's parts is 925
                    rk = knws[1]
            #only pendant's parts weight should be returned
            if jo.style.name.alpha.find("P") >= 0:
                # attention, bug of the driver/pyodbc or encoding issue, Query(JI)
                # returns only one record even if there are several, so get one by one
                #lst = Query(JI).filter(JI.joid == jo.id, JI.stname.like("M%T%"))
                lst = Query((
                    JI.remark,
                    JI.stname,
                    JI.unitwgt,
                )).filter(JI.joid == jo.id, JI.stname.like("%M%T%"))
                lst = lst.with_session(cur).all()
                if lst:
                    for row in lst:
                        if row.unitwgt and _ptn_mit.search(row.stname) and [
                                1 for x in "\" 寸 吋".split()
                                if row.remark.find(x) >= 0
                        ]:
                            kt = rk.karat
                            if row.remark.find("銀") >= 0 or row.remark.find(
                                    "925") >= 0:
                                kt = 925
                            kt = pajcc.WgtInfo(kt, float(row.unitwgt))
                            if not knws[2]:
                                knws[2] = kt
                            else:
                                oth_chn.append(kt)
        if any(knws):
            kt = PrdWgt(knws[0], knws[1], knws[2])
            if oth_chn:
                for row in oth_chn:
                    kt = addwgt(kt, row, True, True)
            kt = kt._replace(netwgt=sum([x.wgt for x in knws if x]))
        else:
            kt = None
        return kt

    def calchina(self, je):
        """ get the weight of given JO# and calc the china
            return a map with keys (JO,PajShp,PajInv,china,wgts)
         """
        if isinstance(je, str):
            je = JOElement(je)
        elif isinstance(je, JO):
            je = je.name
        rmap = {
            "JO": None,
            "china": None,
            "PajShp": None,
            "PajInv": None,
            "wgts": None
        }
        if (not je.isvalid):
            return rmap
        with self.sessionctx():
            ups = self.getpajinvbyjes([je])
            if not ups:
                return rmap
            hnp = ups[0]
            jo = hnp.JO
            prdwgt = self.getjowgts(jo)
            if prdwgt:
                rmap["wgts"] = prdwgt
            rmap["JO"] = jo
            ups = hnp.PajInv.uprice
            rmap["PajInv"] = hnp.PajInv
            rmap["PajShp"] = hnp.PajShp
            revcn = self.getrevcn(hnp.PajShp.pcode)
        rmap["china"] = pajcc.newchina(revcn, prdwgt) if revcn else \
            pajcc.PajCalc.calchina(prdwgt, float(
                hnp.PajInv.uprice), MPS(hnp.PajInv.mps))
        return rmap

    def getmmioforjc(self, df, dt, runns):
        """return the mmstock's I/O# for PajJCMkr"""
        lst = list()
        if not isinstance(runns[0], str):
            runns = [str(x) for x in runns]
        with self.sessionctx() as cur:
            for x in splitarray(runns, self._querysize):
                q = cur.query(SO.running, IV.remark1.label("jmp"), IV.docdate.label("shpdate"), IV.inoutno).join(IVI).join(IV).filter(IV.inoutno.like("N%"))\
                    .filter(IV.remark1 != "").filter(and_(IV.docdate >= df, IV.docdate < dt))\
                    .filter(SO.running.in_(x))
                rows = q.all()
                if rows:
                    lst.extend(rows)
        return lst

    def getpajinvbyjes(self, jes):
        """ get the paj data for jocost
        @param jes: a list of JOElement/string or just one JOElement/string of JO#
        return a list of object contains JO/PajShp/PajInv objects
        """

        if not jes:
            return
        lst = []
        if not isinstance(jes, Sequence):
            if isinstance(jes, Iterable):
                jes = tuple(jes)
            elif isinstance(jes, JOElement):
                jes = (jes)
            elif isinstance(jes, str):
                jes = (JOElement(jes))
        jes = [x if isinstance(x, JOElement) else JOElement(x) for x in jes]

        with self.sessionctx() as cur:
            q0 = self._pjq()
            for ii in splitarray(jes, self._querysize):
                q = JO.name == ii[0]
                for yy in ii[1:]:
                    q = or_(JO.name == yy, q)
                q = q0.filter(q).with_session(cur)
                rows = q.all()
                if rows:
                    lst.extend(rows)
        return lst

    def getpajinvbypcode(self, pcode, maxinvdate=None, limit=0):
        """ return a list of pajinv history order by pajinvdate descendantly
        @param maxinvdate: the maximum invdate, those greater than that won't be return
        @param limit: the maximum count of results returned
        """
        rows = None
        with self.sessionctx() as cur:
            q0 = self._pjq()
            q = q0.filter(PajShp.pcode == pcode)
            if maxinvdate:
                q = q.filter(PajShp.invdate <= maxinvdate)
            q = q.order_by(desc(PajShp.invdate))
            if limit > 0:
                q.limit(limit)
            rows = q.with_session(cur).all()
        return rows

    def getpajinvbyse(self, styno):
        """return a list by PajRevcn as first element, then the cost sorted by joData
        @param styno: string type or StyElement type sty#
        """

        if isinstance(styno, str):
            styno = StyElement(styno)
        elif isinstance(styno, Style):
            styno = styno.name
        with self.sessionctx() as cur:
            lst = None
            q0 = self._pjq().join(Orderma).join(Style)
            q0 = q0.filter(Style.name == styno).order_by(desc(JO.deadline))
            lst = q0.with_session(cur).all()
        return lst
