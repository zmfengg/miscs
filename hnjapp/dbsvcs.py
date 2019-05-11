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

from sqlalchemy import and_, desc, func, or_
from sqlalchemy.orm import Query

from hnjcore import JOElement, KaratSvc, StyElement
from hnjcore.models.cn import JO as JOcn
from hnjcore.models.cn import MM, Codetable
from hnjcore.models.cn import Customer as Customercn
from hnjcore.models.cn import MMMa, StoneIn, StoneOut, StoneOutMaster, StonePk, StoneMaster
from hnjcore.models.cn import Style as Stylecn
from hnjcore.models.hk import JO, JOItem as JI
from hnjcore.models.hk import Invoice as IV
from hnjcore.models.hk import InvoiceItem as IVI
from hnjcore.models.hk import Orderma, PajCnRev, PajInv, PajShp, POItem
from hnjcore.models.hk import StockObjectMa as SO
from hnjcore.models.hk import Style
from hnjcore.utils.consts import NA
from utilz import ResourceCtx, splitarray, karatsvc, trimu, NamedList, stsizefmt
from functools import partial, cmp_to_key

from . import pajcc
from .common import _logger as logger
from .common import splitjns, config, Utilz
from .pajcc import MPS, PrdWgt, WgtInfo, addwgt

__all__ = [
    "CNSvc", "formatsn", "HKSvc", "idsin", "idset", "jesin", "namesin",
    "nameset"
]

_ptn_mit = re.compile(config.get("pattern.mit"))

def fmtsku(skuno):
    if not skuno:
        return None
    skuno = skuno.strip()
    if skuno.upper() == NA:
        return None
    return skuno


class SNFmtr(object):
    _ptn_rmk = re.compile(r"\(.*\)")
    _voidset = set("SN;HB".split(";"))
    _crmp = {
        "）": ")",
        "（": "(",
        ";": ",",
        "六號瓜子耳": "6號瓜子耳",
        "小心形": "小心型",
        "細心": "小心型",
        "中號光銀光子耳": "中號光金瓜子耳",
        "中號光銀瓜子耳": "中號光金瓜子耳"
    }
    _rvlst = sorted(
        "不用封底;執模加圈;圈仔執模做加啤件;不用封底;加啤件;耳針位;請提供;飛邊;占位;光底;夾片;#; ;底;針夾;捲底;相盒;吊墜;吊咀;面做相同花紋;面同一花紋;較用;請提供模號;樣品號;模號請提供;夾底片;片夾底;小占位;有字;用;啤沒有耳仔的"
        .split(";"),
        key=lambda x: len(x),
        reverse=True)
    _pfx = sorted("大2號;中號".split(";"), key=lambda x: len(x), reverse=True)
    _sfx = sorted(
        "夾層;針;小心型;瓜子耳;花".split(";"), key=lambda x: len(x), reverse=True)

    def _splitsn(self, sn):
        if not sn:
            return
        sfx, ots = "", ""
        for x in sn:
            if ord(x) > 128:
                sfx += x
            else:
                ots += x
        je = JOElement(ots)
        if je.digit > 0:
            pfx = "%s%d" % (je.alpha, je.digit)
            sfxx = [x for x in je.suffix]
            if sfx: sfxx.append(sfx)
            sn = tuple(pfx + x for x in sfxx) if sfxx else (sn,)
        else:
            sn = (sn,)
        return sn

    def formatsn(self, sn, parsemode=2, retuple=False):
        """
        parse/formatted/sort a sn string to tuple or a string

        @parsemode: #0 for keep SN like "BT1234ABC" as it was
                    #1 for set SN like "BT1234ABC" to BT1234
                    #2 for split SN like "BT1234ABC" to BT1234A,BT1234B,BT1234C
        @retuple:   return the result as a tuple instead of string
        """
        for x in self._crmp.items():
            sn = sn.replace(x[0], x[1])
        for x in self._rvlst:
            sn = sn.replace(x, ",")
        for x in self._pfx:
            sn = sn.replace(x, "," + x)
        for x in self._sfx:
            sn = sn.replace(x, x + ",")
        sn = re.sub(self._ptn_rmk, ",", sn)
        if not sn or sn in self._voidset: return
        lst = sorted([x for x in sn.split(",") if x])
        buf, dup = [], set()
        for x in lst:
            if x in self._voidset: continue
            if x in dup: continue
            if parsemode != 0:
                je = JOElement(x)
                if je.alpha >= 'A' and je.alpha <= 'Z' and je.suffix:
                    if parsemode == 1:
                        x = (je.alpha + str(je.digit),)
                    else:
                        x = self._splitsn(x)
                else:
                    x = (x,)
            else:
                x = (x,)
            for y in x:
                if not y or y in dup: continue
                dup.add(y)
                buf.append(y)
        return buf if retuple else ",".join(buf)


formatsn = SNFmtr().formatsn


def jesin(jes, objclz):
    """ simulate a in operation for jo.name """
    if not isinstance(jes, (
            tuple,
            list,
    )):
        jes = list(jes)
    q = objclz.name == jes[0]
    for y in jes[1:]:
        q = or_(objclz.name == y, q)
    return q


#these 4 object for sqlalcehmy's query maker for ids/names
idsin = lambda ids, objclz: objclz.id.in_(ids)
idset = lambda ids: {y.id for y in ids}
namesin = lambda names, objclz: objclz.name.in_(names)
nameset = lambda names: {y.name for y in names}


def _getjos(self, objclz, q0, jns, extfltr=None):
    ss = splitjns(jns)
    if not (ss and any(ss)):
        return None
    jes, rns, ids = ss[0], ss[1], ss[2]
    rsts = [None, None, None]
    if ids:
        rsts[0] = self._getbyids(q0, ids, idsin, objclz, idset, extfltr)
    if rns:
        rsts[1] = self._getbyids(q0, rns, lambda x, y: y.running.in_(x), objclz,
                                 lambda x: set([y.running for y in x]), extfltr)
    if jes:
        rsts[2] = self._getbyids(q0, jes, jesin, objclz, nameset, extfltr)
    its, failed = dict(), []
    for x in rsts:
        if not x: continue
        if x[0]: its.update(dict([(y.id, y) for y in x[0]]))
        if x[1]: failed.extend(x[1])
    return list(its.values()), failed


class SvcBase(object):
    _querysize = 20

    def __init__(self, trmgr):
        self._trmgr = trmgr

    def sessmgr(self):
        return self._trmgr

    def sessionctx(self):
        return ResourceCtx(self._trmgr)

    def _getbyids(self, q0, objs, qmkr, objclz, smkr, extfltr=None):
        """
        get object by providing a list of vars, return tuple with valid object tuple and not found set
        """
        if not objs: return
        if not isinstance(objs, Sequence):
            objs = tuple(objs)
        objss = splitarray(objs, self._querysize)
        al = []
        with self.sessionctx() as cur:
            for x in objss:
                q = q0.filter(qmkr(x, objclz))
                if extfltr is not None: q = q.filter(extfltr)
                lst = q.with_session(cur).all()
                if lst: al.extend(lst)
        if al:
            if len(al) < len(objs):
                a0 = set(objs)
                x = smkr(al)
                na = a0.difference(x)
            else:
                na = None
        else:
            na = set(objs)
        return al, na


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
        from operator import attrgetter
        with self.sessionctx() as cur:
            rows = None
            q = Query(PajCnRev).filter(PajCnRev.pcode == pcode).order_by(
                desc(PajCnRev.tag))
            if limit:
                q = q.limit(limit)
            rows = q.with_session(cur).all()
            if rows and len(rows) > 0:
                rows = sorted(rows, key=attrgetter("tag"))
                #rows = sorted(rows, key=attrgetter("revdate"))
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
                if rows:
                    rc = rows[0]
                    if rc.tag != 0: rc = None
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
            q = Query([JO,POItem.skuno]).join(Orderma).join(POItem,POItem.id == JO.poid)\
            .filter(Orderma.styid == jo.orderma.style.id)
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
            except Exception as e:
                if isinstance(e, UnicodeDecodeError):
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
        if (isinstance(je, str)):
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
            if maxinvdate: q = q.filter(PajShp.invdate <= maxinvdate)
            q = q.order_by(desc(PajShp.invdate))
            if limit > 0: q.limit(limit)
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


class CNSvc(SvcBase):

    def getshpforjc(self, df, dt):
        """return py shipment data for PajJCMkr
        @param df: start date(include) a date ot datetime object
        @param dt: end date(exclude) a date or datetime object
        """
        lst = None
        with self.sessionctx() as cur:
            if True:
                q0 = Query([JOcn, MMMa, Customercn, Stylecn, JOcn, MM])
            else:
                q0 = Query([
                    JOcn.id, MMMa.refdate,
                    Customercn.name.label("cstname"),
                    JOcn.name.label("jono"),
                    Stylecn.name.label("styno"), JOcn.running, JOcn.karat,
                    JOcn.description,
                    JOcn.qty.label("joqty"),
                    MM.qty.label("shpqty"),
                    MM.name.label("mmno")
                ])
            q = q0.join(Customercn).join(Stylecn).join(MM).join(MMMa)\
                .filter(and_(MMMa.refdate >= df, MMMa.refdate < dt)).with_session(cur)
            lst = q.all()
        return lst

    def getjcrefid(self, runn):
        """ return the referenceId of given runn#, return tuple
        tuple[0] = the refid, tuple[1] = (runnf,runnt)
        """
        x = None
        with self.sessionctx() as cur:
            q = Query([Codetable.coden0,Codetable.coden1,Codetable.coden2]).filter(and_(Codetable.tblname == "jocostma",Codetable.colname == "costrefid")).\
                filter(and_(Codetable.coden1 <= runn,Codetable.coden2 >= runn))
            x = q.with_session(cur).one_or_none()
        if x:
            return int(x.coden0), (int(x.coden1), int(x.coden2))

    def getjcmps(self, refid):
        """ return the metal ups of given refid as dict """
        x = None
        with self.sessionctx() as cur:
            q = Query(Codetable).filter(and_(Codetable.tblname == "metalma",Codetable.colname == "goldprice")).\
                filter(Codetable.tag == refid)
            lst = q.with_session(cur).all()
            return dict([(int(x.coden0), float(x.coden1)) for x in lst])

    def getstins(self, btnos):
        """
        return the stonein's by provding a list of btchnos or btchid
        btchno or id is determined by the fist item in btchnos
        @param btnos: should be a collection of btchno(str) or btchid(int)
        """
        if not isinstance(btnos, Sequence):
            btnos = tuple(btnos)
        isstr = isinstance(btnos[0], str)
        if isstr:
            qmkr, smkr = namesin, nameset
        else:
            qmkr, smkr = idsin, idset
        return self._getbyids(Query(StoneIn), btnos, qmkr, StoneIn, smkr)

    def getstpks(self, pknos):
        if not isinstance(pknos, Sequence):
            pknos = tuple(pknos)
        isstr = isinstance(pknos[0], str)
        if isstr:
            qmkr, smkr = namesin, nameset
        else:
            qmkr, smkr = idsin, idset
        return self._getbyids(Query(StonePk), pknos, qmkr, StonePk, smkr)

    def getjos(self, jns):
        return _getjos(self, JOcn, Query(JOcn), jns)

    def getjostcosts(self, runns):
        """
        return the stone costs by map, running or je as key and cost as value
        """
        if not runns: return None
        if not isinstance(runns, Sequence): runns = tuple(runns)
        isjn, jnlv = isinstance(runns[0], str) or isinstance(
            runns[0], JOElement), 0
        if isjn and isinstance(runns[0], str):
            runnsx = tuple(runns)
            runns = [JOElement(x) for x in runns]
            jnlv = 1
        else:
            runnsx = runns
        lst, cdmap = [], None
        sign = lambda x: 0 if x == 0 else 1 if x > 0 else -1
        cols = [
            JOcn.name, StoneOutMaster.isout, StonePk.pricen, StonePk.unit,
            func.sum(StoneOut.qty).label("qty"),
            func.sum(StoneOut.wgt).label("wgt")
        ]
        gcols = [JOcn.name, StoneOutMaster.isout, StonePk.pricen, StonePk.unit]
        if not isjn:
            cols[0], gcols[0] = JOcn.running, JOcn.running
        q0 = Query(cols).join(StoneOutMaster).join(StoneOut).join(StoneIn).join(
            StonePk).group_by(*gcols)
        with self.sessionctx() as cur:
            for arr in splitarray(runns, self._querysize):
                try:
                    if isjn:
                        q = q0.filter(jesin(arr, JOcn))
                    else:
                        q = q0.filter(JOcn.running.in_(arr))
                    lst1 = q.with_session(cur).all()
                    if lst1: lst.extend(lst1)
                except:
                    pass
            if lst:
                lst1 = Query([Codetable.coden0, Codetable.tag]).filter(
                    and_(Codetable.tblname == "stone_pkma",
                         Codetable.colname == "unit")).with_session(cur).all()
                cdmap = dict([(int(x.coden0), x.tag) for x in lst1])
        if lst and cdmap:
            costs = dict(zip(runnsx, [0] * len(runns)))
            for x in lst:
                costs[int(x.running) if not isjn else x.name.value if jnlv ==
                      1 else x.name] += round(
                          sign(float(x.isout)) * float(x.pricen) * (float(
                              x.qty) if cdmap[x.unit] == 0 else float(x.wgt)),
                          2)
            return costs


class BCSvc(object):
    """a handy Hnjhk dao for data access in this tests
    now it's only bc services
    """
    _querysize = 20  # batch query's batch, don't be too large

    def __init__(self, bcdb=None):
        self._bcdb = bcdb

    def getbcsforjc(self, runns):
        """return running and description from bc with given runnings """
        if not (self._bcdb and runns):
            return None
        runns = [str(x) for x in runns]
        s0 = "select runn,desc,ston from stocks where runn in (%s)"
        lst = []
        cur = self._bcdb.cursor()
        try:
            for x in splitarray(runns, self._querysize):
                cur.execute(s0 % ("'" + "','".join(x) + "'"))
                rows = cur.fetchall()
                if rows:
                    lst.extend(rows)
        except:
            pass
        finally:
            if cur:
                cur.close()
        return self._trim(lst)

    @classmethod
    def _trim(cls, lst):
        if not lst:
            return None
        for x in lst:
            for idx, s0 in enumerate(x):
                if s0 and isinstance(s0, str) and s0[-1] == " ":
                    x[idx] = s0.strip()
        return lst

    def getbcs(self, jnorunn, isstyno=False):
        """ should be jns or runnings
        runnings should be of numeric type
        """
        if not (self._bcdb and jnorunn):
            return None
        if not isinstance(jnorunn, Sequence):
            jnorunn = tuple(jnorunn)
        if isstyno:
            cn = "styn"
        else:
            cn = "jobn" if isinstance(jnorunn[0], str) else "runn"
            if cn == "runn":
                jnorunn = [str(x) for x in jnorunn]

        s0, lst = "select * from stocks where %s in (%%s)" % cn, []
        cur = self._bcdb.cursor()
        try:
            for x in splitarray(jnorunn, self._querysize):
                cur.execute(s0 % ("'" + "','".join(x) + "'"))
                rows = cur.fetchall()
                if rows:
                    lst.extend(rows)
        finally:
            if cur:
                cur.close()
        return self._trim(lst)


    def build_from_jo(self, jn, hksvc, cnsvc, hints=None):
        '''
        create a bc item based on provided jn or provided hints(dict)

        @param jn: A JO#(str)
        @param hksvc: An HKsvc instance(to fetch JO data from)
        @param hints: a dict contains necessary element, where
            '_raw_data' contains a bc item, when this exists, I use it instead of creating a new one.
            '_stone_data' contains a tuple of list, where stone is defined. First item in the tuple is a namedlist to operate on the stone data. The available columns can be found in PajShpHdlr.read_stone_data(). Up to 2019/04/04, the columns are:
                pcode,stone,stshape,stsize,stwgt

        description split into 4 parts:
        1.karat + plating where plating can be from JO.remark or stones
        2.stones
        3.category
        4.suffix
        then data that should also be send to rem1 to rem8 is:
        1.1 aux wgt(if there is)
        1.2 parts wgt(if there is)
        2.1 Any stone data
        99.1 SN#
        '''

        with hksvc.sessionctx():
            jo = hksvc.getjos((jn, ))[0][0]
            if not jo:
                return None
            mp = {"jo": jo}
            # karat in hints is not reliable, so at least get karat from JO system
            mp["wgts"] = self._merge_wgts(hksvc.getjowgts(jo), hints.get("mtlwgt") if hints else None)
            if hints and '_stone_data' in hints:
                var = self._adopt_stone(hints['_stone_data'], Utilz.nl_ji())
                if var:
                    mp["stones"], mp["nl_stone"] = var, Utilz.nl_ji()
            else:
                var = Utilz.extract_jis((jo, ), hksvc)
                if var:
                    var = var[jo.id]
                    for idx, x in enumerate("stones nl_stone miscs".split()):
                        mp[x] = var[idx]
            if var:
                var = mp["nl_stone"]
                var = [var.setdata(x).sto for x in mp["stones"]]
            for fn, x in zip((partial(Utilz.extract_vcs, jo, var, mp["wgts"]), partial(Utilz.extract_micron, jo.remark)), ("plating", "micron")):
                mp[x] = fn()
            if hints:
                for key in ('_raw_data', '_snno',):
                    if key in hints:
                        mp[key] = hints[key]
            return _JO2BC(cnsvc).build(mp)

    @classmethod
    def _merge_wgts(cls, jowgts, hintwgts):
        if not (jowgts and hintwgts):
            return None if not (jowgts or hintwgts) else jowgts or hintwgts
        rst = []
        for j, h in zip(jowgts.wgts, hintwgts.wgts):
            if j or h:
                rst.append(WgtInfo(j.karat if j else h.karat, h.wgt if h else j.wgt))
            else:
                rst.append(None)
        return PrdWgt(*rst)

    @classmethod
    def _adopt_stone(cls, sts, nl_ji):
        nl_src, rsts = sts[0], []
        for x in sts[1:]:
            nl_src.setdata(x)
            rsts.append(nl_ji.newdata())
            nl_ji.stone = nl_ji.sto = nl_src.stone
            nl_ji.shape = nl_ji.shpo = nl_src.shape
            nl_ji.stsize, nl_ji.stqty, nl_ji.stwgt = nl_src.size, nl_src.qty, nl_src.wgt
            nl_ji.szcal = stsizefmt(nl_src.size)
        return rsts


class _JO2BC(object):
    def __init__(self, cnsvc=None):
        # pp holds properties except (locket, locket_pic)
        self._pp = self._pp_x = None
        self._v_c2n = {x["color"]: x for x in config.get("vermail.defs")}
        x = config.get("bc.description.stone.categories")
        self._snno_mp = config.get("snno.translation")
        self._stcats = {z: x for x, y in x.items() for z in y}
        self._cnsvc = cnsvc
        x = config.get("bc.sns")
        if x:
            self._sns_mp = {(x[0], x[1]): x[2] for x in x}
        else:
            self._sns_mp = None
        self._st_bywgt = {x for x in config.get("bc.stone.bywgt")}


    def build(self, pp):
        '''
        Build a bc instance based on the dict provided
        '''
        self._pp = pp
        self._pp_x = {}
        nl_bc = NamedList(config.get("bc.colnames.alias"))
        vx = self._pp.get("_raw_data")
        if vx:
            nl_bc.setdata(vx)
        else:
            nl_bc.newdata()
            self._fill_basic(nl_bc)
        self._sort_st_ms(nl_bc)
        self._make_desc(nl_bc)
        self._make_rmks(nl_bc)
        return nl_bc

    def _fill_basic(self, nl):
        jo, wgts = (self._pp[x] for x in "jo wgts".split())
        nl.styno = nl.location = jo.style.name.value
        nl.qty, nl.wgt, nl.karat = jo.qty, wgts.main.wgt, wgts.main.karat
        td = datetime.date.today()
        nl.date, nl.mp_year, nl.mp_month = td.strftime("%Y%m%d 00:0000"), td.strftime("%y"), td.strftime("%m")
        nl.cstname, nl.jono, nl.descn = trimu(jo.orderma.customer.name), jo.name.value, jo.description

    def _sort_st_ms(self, nl_bc):
        '''
        sort the stone, fill the main stone if there is
        '''
        ms = "--"
        vx, nl = [self._pp.get(x) for x in ("stones", "nl_stone")]
        if vx:
            cat = self._pp["jo"]
            def srt_cmp(st0, st1):
                rc = [nl.setdata(x).stsize or '' for x in (st0, st1)]
                rc = 0 if rc[0] == rc[1] else 1 if rc[0] > rc[1] else -1
                if rc == 0:
                    sts = [nl.setdata(x).stone for x in (st1, st0)]
                    rc = 0 if sts[0] == sts[1] else 1 if sts[0] > sts[1] else -1
                return rc
            cat = Utilz.getStyleCategory(cat.style.name.value, cat.description)
            for x in vx:
                nl.setdata(x)
                if Utilz.is_main_stone(cat, nl.stqty, nl.szcal):
                    ms = nl.sto
                    break
            # sort the stone by MS + "SZ(A)+ST(D)"
            sts = sorted(vx, key=cmp_to_key(srt_cmp))
            st_p = {nl.setdata(st).stone: idx for idx, st in enumerate(sts)}
            sts = sorted(vx, key=lambda x: (st_p[nl.setdata(x).stone], nl.stsize or ''), reverse=True)
            self._pp["stones"] = sts
            if ms == '--':
                ms = nl.setdata(sts[0]).sto
        nl_bc.stone = ms

    def _make_desc(self, nl_bc):
        rc = self._make_karat_plating()
        rc += self._make_desc_stones(nl_bc)
        tc = self._make_style_cat()
        rc, tc = rc + " " + tc[0], tc[1]
        #<pendant's chain>
        # ^E customer don't need wo chain except ESO/ET/EJE/ELH
        if tc in ("PENDANT", 'LOCKET'):
            nl, wgts, jo = [self._pp.get(x) for x in ('nl_stone wgts jo'.split())]
            chns = config.get("jo.description.chain")
            if wgts.part or \
                [1 for x in chns if jo.description.find(x) >= 0] or \
                [1 for x in self._pp.get("miscs", []) if nl.setdata(x).sto == Utilz.MISC_MIT and [1 for x in chns if nl.setting.find(x) >= 0]]:
                rc += ' ROPE CHAIN'
                # check if wgts contains part, if not, insert one
                if not wgts.part:
                    self._pp["wgts"] = PrdWgt(wgts.main, wgts.aux, WgtInfo(wgts.main.karat, 0))
            else:
                vx = trimu(jo.orderma.customer.name)
                # rule according to e-mail title('EJE 新樣品 P41001_P41002') sent in Jul 12, 2018
                if vx[0] != 'E' or vx in config.get('bc.description.wochain.customers'):
                    rc += ' W/O CHAIN'
        #</pendant's chain>
        nl_bc.description = rc

    def _make_karat_plating(self):
        '''
        mixed the gold/color based on rule
        '''
        return self.knc_mix(self._pp.get('wgts'), self._pp.get("plating"))

    def knc_mix(self, wgts, vcs):
        '''
        9K&VW means partial VW
        9K VW measn all VW
        '''
        kts = [x.karat for x in (wgts.main, wgts.aux) if x and x.karat]
        if vcs is None:
            vcs = set() #don't return directly, below has 200 case
        else:
            vcs = {x[0] for x in vcs}
        tt = len(kts) > 1
        vx = lambda x: self._v_c2n[x]["name"]
        vxx = lambda y: [vx(x) for x in sorted(tuple(y), key=lambda x: self._v_c2n[x]["priority"])]
        kn = lambda x: "&".join((karatsvc.getkarat(y).name for y in x))
        if not tt:
            rc = kn(kts)
            if vcs:
                rc += "%s%s" % (" " if kts[0] in (925, 200) else "&", "&".join(vxx(vcs)))
        else:
            # when there is bronze, follow master's color
            if kts[1] == 200:
                vcs.add(karatsvc.getkarat(kts[0]).color)
            if kts[0] == 925:
                rc = karatsvc.getkarat(kts[0]).name + " " + "&".join(vxx(vcs)) + "&" + karatsvc.getkarat(kts[1]).name
            else:
                # gold + X case
                rc = kn(kts) + " " + "&".join(vxx(vcs))
        return rc

    def _make_style_cat(self):
        jo, vx = self._pp["jo"], None
        tc = Utilz.getStyleCategory(jo.style.name.value, jo.description)
        if tc in ("PENDANT", "LOCKET"):
            pc = re.compile(config.get("pattern.locket.pics"))
            pc = pc.search(jo.description)
            if pc:
                pc = pc.group(1)
            nl = self._pp.get("nl_stone")
            vx = [x for x in self._pp.get("miscs", []) if nl.setdata(x).sto == Utilz.MISC_LKSZ]
            if not vx:
                vx = jo.description.find("相盒") >= 0 or jo.remark.find("相盒") >= 0
            if not (vx or pc):
                vx = ""
            elif vx and isinstance(vx, (tuple, list)):
                nl.setdata(vx[0])
                self._pp_x["locket"], vx = "%s:%s" % (nl.stone, nl.stsize), 'LOCKET '
            else:
                self._pp_x["locket"], vx = "LKSZ:TODO", 'LOCKET '
            if pc:
                self._pp_x['locket_pic'] = "%s PIC" % pc
        elif tc == "EARRING":
            if jo.description.find("啤") > 0 or jo.remark.find("啤") >= 0:
                vx = 'CREOLE '
        return (vx or "") + tc, tc

    def _make_rmks(self, nl_bc):
        lst = []
        wgts = self._pp["wgts"]
        wgts = (wgts.aux, wgts.part)
        if wgts[0]:
            lst.append("%s %4.2f" % (karatsvc.getkarat(wgts[0].karat).name, wgts[0].wgt))
        if wgts[1]:
            lst.append("*%sPTS %4.2f" % (karatsvc.getkarat(wgts[1].karat).name, wgts[1].wgt))
        vx = self._pp.get("plating")
        if vx:
            vx = [x for x in vx if len(x) == 1], [x for x in vx if len(x) > 1]
            lst.append(self.knc_mix(self._pp.get('wgts'), vx[0]))
            if vx[1]:
                # vk with height case
                for x in vx[1]:
                    lst.append(x[0] + " " + x[1])
        for vx in (self._pp_x.get(x) for x in ("locket", "locket_pic")):
            if vx:
                lst.append(vx)
        vx, nl = [self._pp.get(x) for x in ("stones", "nl_stone")]
        if vx:
            for x in vx:
                lst.append(self._make_rem_stone(nl.setdata(x)))
        vx = self._pp.get("_snno")
        vx = self._pp["jo"].snno + ("," + vx if vx else "")
        if vx:
            vx = self._extract_snno(vx)
        if vx:
            vx[0] = "SN#" + vx[0]
            lst.extend(vx)
        else:
            lst.append('SN#%s' % NA)
        mc = 8
        # merge them to 8 items if there are more
        for idx, vx in enumerate(lst[:mc]):
            setattr(nl_bc, "rem%d" % (idx + 1), vx)
        if len(lst) > mc:
            setattr(nl_bc, "rem%d" % mc, ",".join(lst[mc - 1:]))
        return nl_bc

    def _make_rem_stone(self, nl):
        rc = str(int(nl.stqty)) if nl.stqty > 1 else ""
        rc += self._nrm_sns(nl.shpo, nl.sto)
        if nl.stwgt and nl.stwgt != nl.stqty and self._is_st_bywgt(nl.sto):
            rc += "-%r" % round(nl.stwgt, 3)
        if not self._is_st_bywgt(nl.sto):
            rc += "-" + (nl.stsize or "") + "MM"
        return rc

    def _nrm_sns(self, shp, stname):
        '''
        return RDD if shp == 'R' and stname == 'DD'
        '''
        return self._sns_mp.get((shp, stname), shp + stname)


    def _is_st_bywgt(self, stname):
        return stname in self._st_bywgt
        # following PK's unit is not a good point, use fixed answer from conf.json
        # with self._cnsvc.sessionctx() as cur:
        #    pk = cur.query(StonePk).join(StoneMaster).filter(StoneMaster.name == stname).first()
        #    if not pk:
        #        return False
        #    return pk.unit in (1, 3, 5)


    def _make_desc_stones(self, nl_bc):
        vx, nl = [self._pp.get(x) for x in ("stones", "nl_stone")]
        if not vx:
            nl_bc.stone = "--"
            return ""
        mc = config.get('bc.description.stone.maxcnt', 3)
        # whenever there is DIA, make it out
        cats, rc, sms, ms_cat = {}, [], {}, None
        lmb_sn = lambda pk: pk[:2]
        for x in vx:
            st = lmb_sn(nl.setdata(x).stone)
            cat = self._stcats.get(st, '*MULTI_STONE')
            if st == nl_bc.stone:
                ms_cat = cat = (" MS" if cat[0] == "*" else cat)
            if st not in sms:
                sms[st] = None
            cat = cats.setdefault(cat, [])
            if st not in cat:
                cat.append(st)
        if self._cnsvc:
            with self._cnsvc.sessionctx() as cur:
                lst = cur.query(StoneMaster).filter(StoneMaster.name.in_(sms.keys())).all()
                sms = {x.name: x.edesc for x in lst}
        else:
            sms = {}
        # sort the cats, ms as first, compositive as second, later by stone name
        #keys = sorted(cats, key=lambda x: ' ' if x == nl_bc.stone else '-' if x in self._stcats else x)
        keys = sorted(cats, key=lambda x: -10 if ms_cat == x else -1 if len(cats[x]) > mc else 1)
        for cn in keys:
            sts = cats.get(cn)
            if len(sts) > mc:
                rc.append(cn)
            else:
                for x in sts:
                    rc.append(sms.get(x, x))
        return ' ' + " ".join(rc)

    def _extract_snno(self, snno):
        # HB1729夾片底HB923,中號光金瓜子耳
        # HB2056底HB943夾片HB114夾層HB982,中號光金瓜子耳
        # #
        # HB1762底P27473,中號光金瓜子耳
        # #HB2600底夾片HB1485,PT5012BL
        if not snno:
            return None
        lst, part, ods, push = [], "", 0, False
        for ch in snno:
            if ch in ('#', ' '):
                continue
            od = min(max(ord(ch), 250), 251)
            if ods:
                push = ods != od or ch == ','
                if push and ch == ',':
                    ch = ""
            if push:
                lst.append((part, ods))
                part, push = "", False
            ods = od
            part += ch
        if part:
            lst.append((part, ods))
        ods = []
        for part, od in lst:
            if od == 251:
                #chinese, big5, need translation, if not translated, ignore it
                part = self._snno_mp.get(part)
            if part and part not in ods:
                ods.append(part)
        return ods
