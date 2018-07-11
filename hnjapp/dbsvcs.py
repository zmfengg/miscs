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
import threading
from collections import Iterable
from contextlib import contextmanager
from operator import attrgetter

from sqlalchemy import and_, desc, or_, true
from sqlalchemy.orm import Query, Session

from hnjcore import JOElement, KaratSvc, StyElement
from hnjcore.models.cn import JO as JOcn,StoneIn,StonePk
from hnjcore.models.cn import Customer as Customercn
from hnjcore.models.cn import Style as Stylecn
from hnjcore.models.cn import MM, MMMa, Codetable
from hnjcore.models.hk import Invoice as IV
from hnjcore.models.hk import InvoiceItem as IVI
from hnjcore.models.hk import JOItem as JI
from hnjcore.models.hk import StockObjectMa as SO
from hnjcore.models.hk import (JO, Customer, Orderma, PajCnRev, PajInv, PajShp,
                               POItem, Style)
from hnjcore.utils import samekarat, splitarray, ResourceCtx
from hnjcore.utils.consts import NA

from . import pajcc
from .pajcc import MPS, PrdWgt, WgtInfo
from .common import _logger as logger, splitjns

__all__ = ["HKSvc", "CNSvc","jesin","idsin","idset","namesin","nameset"]

def fmtsku(skuno):
    if not skuno:
        return None
    skuno = skuno.strip()
    if skuno.upper() == NA:
        return None
    return skuno

def jesin(jes,objclz):
    """ simulate a in operation for jo.name """
    if not (isinstance(jes,tuple) or isinstance(jes,list)):
        jes = list(jes)
    q = objclz.name == jes[0]
    for y in jes[1:]:
        q = or_(objclz.name == y,q)
    return q

#these 4 object for sqlalcehmy's query maker for ids/names
idsin = lambda ids,objclz: objclz.id.in_(ids)
idset = lambda ids: set([y.id for y in ids])
namesin = lambda names,objclz: objclz.name.in_(names)
nameset =lambda names: set([y.name for y in names])

def _getjos(self, objclz, q0, jns):
    ss = splitjns(jns)
    if not(ss and any(ss)): return
    jes, rns,ids= ss[0],ss[1],ss[2] 
    rsts = [None,None,None]
    if ids:
        rsts[0] = self._getbyids(q0,ids,idsin,objclz,idset)
    if rns:
        rsts[1] = self._getbyids(q0,rns,lambda x,y: y.running.in_(x),objclz,lambda x: set([y.running for y in x]))
    if jes:       
        rsts[2] = self._getbyids(q0,jes,jesin,objclz,nameset)
    its, failed = dict(), []
    for x in rsts:
        if not x: continue
        if x[0]: its.update(dict([(y.id,y) for y in x[0]]))
        if x[1]: failed.extend(x[1])
    return list(its.values()),failed

class SvcBase(object):
    _querysize = 20

    def __init__(self, trmgr):
        self._trmgr = trmgr
    
    def sessmgr(self):
        return self._trmgr

    def sessionctx(self):
        return ResourceCtx(self._trmgr)
    
    def _getbyids(self,q0, objs, qmkr, objclz, smkr):
        """
        get object by providing a list of vars, return tuple with valid object tuple and not found set
        """
        if not objs: return
        if not(isinstance(objs,tuple) or isinstance(objs,list)):
            objs = tuple(objs)
        objss = splitarray(objs,self._querysize)
        al = []
        with self.sessionctx() as cur:
            for x in objss:
                lst = q0.filter(qmkr(x,objclz)).with_session(cur).all()
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
        return al,na

class HKSvc(SvcBase):
    _qcaches = {}
    _ktsvc = KaratSvc()

    def _samekarat(self, jo0, jo1):
        """ check if the 2 given jo's karat/auxkarat are the same
            this method compare the 2 karats            
        """
        lst = ([(self._ktsvc.getfamily(x.orderma.karat), self._ktsvc.getfamily(x.auxkarat)
                 if x.auxwgt else 0) for x in (jo0, jo1)])
        return lst[0] == lst[1]

    def __init__(self, sqleng):
        """ init me with a sqlalchemy's engine """
        super(HKSvc, self).__init__(sqleng)
        self._ptnmit = re.compile("^M[A-Z]T")

    def _pjq(self):
        """ return a JO -> PajShp -> PajInv query, append your query
        before usage, then remember to execute using q.with_session(yousess)
        """
        return self._qcaches.setdefault("jopajshp&inv",
                                        Query([PajShp, JO, PajInv]).join(
                                            JO, JO.id == PajShp.joid)
                                        .join(PajInv, and_(PajShp.joid == PajInv.joid, PajShp.invno == PajInv.invno)))

    def _pjsq(self):
        """ return a cached Style -> JO -> PajShp -> PajInv query """
        return self._qcaches.setdefault("jopajshp&inv",
                                        Query([PajShp, JO, Style, Orderma, PajInv]).join(
                                            JO, JO.id == PajShp.joid)
                                        .join(PajInv, and_(PajShp.joid == PajInv.joid, PajShp.invno == PajInv.invno)))

    def getjos(self, jesorrunns):
        """get jos by a collection of JOElements/Strings or Integers
        when the first item is string or JOElement, it will be treated as getbyname, else by runn
        return a tuple, the first item is list containing hnjcore.models.hk.JO
                        the second item is a set of ids/jes/runns not found
        @param groupby: can be one of id/running/name, running should be a string
            starts with 'r' for example, 'r410100', id should be integer,
            name should be JOElement or string without 'r' as prefix
        """
        return _getjos(self,JO, Query(JO),jesorrunns)

    def getjo(self, jeorrunn):
        """ a convenient way for getjos """
        with self.sessionctx():
            jos = self.getjos([jeorrunn])
        return jos[0][0] if jos else None

    def getrevcns(self, pcode, limit=0):
        """ return a list of revcns order by the affected date desc        
        The current revcn is located at [0]
        @param limit: only return the given count of records. limit == 0 means no limits        
        """
        from operator import attrgetter
        with self.sessionctx() as cur:
            rows = None
            q = Query(PajCnRev).filter(PajCnRev.pcode == 
                                    pcode).order_by(desc(PajCnRev.tag))
            if limit:
                q = q.limit(limit)
            rows = q.with_session(cur).all()
            if rows and limit != 1:
                rows = sorted(rows, key=attrgetter("tag"))
                rows = sorted(rows, key=attrgetter("revdate"))
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

    def findsimilarjo(self, jo, level=1, mindate = datetime.datetime(2015,1,1)):
        """ return an list of JO based on below criteria
        @param level:   0 for extract SKU match
                        1 for extract karat match
                        1+ for extract style match
        @param mindate: the minimum date to fetch data
        """
        rc = None
        level = 0 if level < 0 else level
        je = jo.name
        jns = None
        with self.sessionctx() as cur:        
            #don't lookup too much, only return data since 2015
            q = Query([JO,POItem.skuno]).join(Orderma).join(POItem,POItem.id == JO.poid)\
                .filter(Orderma.styid == jo.orderma.style.id)
            if mindate:
                q = q.filter(JO.createdate >= mindate)
            try:
                rows = q.with_session(cur).all()
            except Exception as e:
                if isinstance(e,UnicodeDecodeError):
                    logger.debug("some jos of sty#(%s) contains invalid Big5 character " % jo.orderma.style.name.value)
                rows = None
            if(rows):
                jns = {}
                for x in rows:
                    if x.JO.name == je:
                        continue
                    key = fmtsku(x.JO.po.skuno)
                    jns.setdefault(key, []).append(x.JO)
            if jns:
                skuno = fmtsku(jo.po.skuno)
                sks = jns[skuno] if skuno and skuno in jns else None
                if not sks and level > 0:
                    sks = [x.JO for x in rows if je !=
                           x.JO.name and self._samekarat(jo, x.JO)]
                    if not sks and level > 1:
                        sks = [x.JO for x in rows]
                rc = sks
        if rc and len(rc) > 1:
            rc = sorted(rc,key = attrgetter("createdate"), reverse = True)
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
            rk = knws[0]
            if jo.auxwgt:
                knws[1] = WgtInfo(jo.auxkarat, float(jo.auxwgt))
                if(knws[1].karat == 925):  # most of 925's parts is 925
                    rk = knws[1]
            #only pendant's parts weight should be returned
            if jo.style.name.alpha.find("P") >= 0:
                lst = cur.query(JI).filter(
                    and_(JI.joid == jo.id, JI.stname.like("M%T"))).all()
                if lst:
                    row = lst[0]
                    if(row.unitwgt and self._ptnmit.search(row.stname)):
                        knws[2] = pajcc.WgtInfo(rk.karat, float(row.unitwgt))
        if any(knws):
            return PrdWgt(knws[0], knws[1], knws[2])

    def calchina(self, je):
        """ get the weight of given JO# and calc the china
            return a map with keys (JO,PajShp,PajInv,china,wgts)
         """
        if(isinstance(je, str)):
            je = JOElement(je)
        elif isinstance(je, JO):
            je = je.name
        rmap = {"JO": None, "china": None,
                "PajShp": None, "PajInv": None, "wgts": None}
        if(not je.isvalid):
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
        """return the mmstock's I/O# for PAJCReader"""
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
        if not(isinstance(jes, list) or isinstance(jes, tuple)):
            if isinstance(jes, Iterable):
                jes = tuple(jes)
            elif isinstance(jes, JOElement):
                jes = (jes)
            elif isinstance(jes, str):
                jes = (JOElement(jes))
        jes = [x if isinstance(x,JOElement) else JOElement(x) for x in jes]

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

    def getpajinvbypcode(self,pcode, maxinvdate = None, limit = 0):
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
        """return py shipment data for PAJCReader
        @param df: start date(include) a date ot datetime object
        @param dt: end date(exclude) a date or datetime object 
        """
        lst = None
        with self.sessionctx() as cur:
            if True:
                q0 = Query([JOcn, MMMa, Customercn, Stylecn, JOcn, MM])
            else:
                q0 = Query([JOcn.id, MMMa.refdate, Customercn.name.label("cstname"), JOcn.name.label("jono"), Stylecn.name.label(
                    "styno"), JOcn.running, JOcn.karat, JOcn.description, JOcn.qty.label("joqty"), MM.qty.label("shpqty"), MM.name.label("mmno")])
            q = q0.join(Customercn).join(Stylecn).join(MM).join(MMMa)\
                .filter(and_(MMMa.refdate >= df, MMMa.refdate < dt)).with_session(cur)
            lst = q.all()
        return lst
    
    def getjcrefid(self,runn):
        """ return the referenceId of given runn#, return tuple
        tuple[0] = the refid, tuple[1] = (runnf,runnt)
        """
        x = None
        with self.sessionctx() as cur:
            q = Query([Codetable.coden0,Codetable.coden1,Codetable.coden2]).filter(and_(Codetable.tblname == "jocostma",Codetable.colname == "costrefid")).\
                filter(and_(Codetable.coden1 <= runn,Codetable.coden2 >= runn))
            x = q.with_session(cur).one_or_none()
        if x:
            return int(x.coden0),(int(x.coden1),int(x.coden2))

    def getjcmps(self,refid):
        """ return the metal ups of given refid as dict """
        x = None
        with self.sessionctx() as cur:
            q = Query(Codetable).filter(and_(Codetable.tblname == "metalma",Codetable.colname == "goldprice")).\
                filter(Codetable.tag == refid)
            lst = q.with_session(cur).all()
            return dict([(int(x.coden0),float(x.coden1)) for x in lst])
    
    def getstins(self,btnos):
        """
        return the stonein's by provding a list of btchnos or btchid
        btchno or id is determined by the fist item in btchnos
        @param btnos: should be a collection of btchno(str) or btchid(int)
        """
        if not(isinstance(btnos,tuple) or isinstance(btnos,list)):
            btnos = tuple(btnos)
        isstr = isinstance(btnos[0],str)
        if isstr:
            qmkr, smkr = namesin, nameset
        else:
            qmkr, smkr = idsin, idset
        return self._getbyids(Query(StoneIn),btnos,qmkr,StoneIn,smkr)

    def getstpks(self,pknos):
        if not(isinstance(pknos,tuple) or isinstance(pknos,list)):
            pknos = tuple(pknos)
        isstr = isinstance(pknos[0],str)
        if isstr:
            qmkr,smkr =  namesin,nameset
        else:
            qmkr,smkr = idsin,idset
        return self._getbyids(Query(StonePk),pknos,qmkr,StonePk,smkr)
    
    def getjos(self, jns):
        return _getjos(self,JOcn, Query(JOcn),jns)

class BCSvc(object):
    """a handy Hnjhk dao for data access in this tests
    now it's only bc services
    """
    _querysize = 20  # batch query's batch, don't be too large

    def __init__(self, bcdb=None):
        if bcdb: self._bcdb = bcdb
    
    def getbcsforjc(self, runns):
        """return running and description from bc with given runnings """
        if not (self._bcdb and runns): return
        runns = [str(x) for x in runns]
        s0 = "select runn,desc,ston from stocks where runn in (%s)";lst = []
        cur = self._bcdb.cursor()
        try:
            for x in splitarray(runns, self._querysize):
                cur.execute(s0 % ("'" + "','".join(x) + "'"))
                rows = cur.fetchall()
                if rows: lst.extend(rows)
        except:
            pass
        finally:
            if cur: cur.close()
        #trim the spaces
        if lst:
            for x in lst:
                for idx in range(len(x)):
                    s0 = x[idx]
                    if s0 and isinstance(s0,str) and s0.find(" ") >= 0:
                        x[idx] = s0.strip()
        return lst