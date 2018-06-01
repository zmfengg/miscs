# coding=utf-8
"""
 * @Author: zmFeng 
 * @Date: 2018-05-25 14:21:01 
 * @Last Modified by:   zmFeng 
 * @Last Modified time: 2018-05-25 14:21:01 
 * the database services, including HK's and py's, and the out-dated bc's
 """

import re
from logging import Logger

from sqlalchemy.orm import Session,Query
from sqlalchemy import and_, desc, or_, true
import pajcc as pc
from pajcc import MPS,WgtInfo,PrdWgt,fmtkarat
from hnjcore import JOElement
from hnjcore.models.hk import JO, Customer, Style, Orderma, JOItem as JI
from hnjcore.models.hk import POItem, PajShp, PajInv, PajCnRev
from hnjcore.models.cn import JO as JOcn, Customer as Customercn,Style as Stylecn
from hnjcore.models.hk import StockObjectMa as SO,Invoice as IV ,InvoiceItem as IVI
from hnjcore.models.cn import MMMa,MM
from hnjcore.utils import samekarat, splitarray
from contextlib import contextmanager
from collections import Iterable


__all__ = ["HKSvc", "CNSvc"]

NA = "N/A"

def fmtsku(skuno):
    if not skuno: return None
    skuno = skuno.strip()
    if skuno.upper() == NA: return None
    return skuno

class SvcBase(object):
    _querysize = 20
    
    def __init__(self,sqleng):
        self._engine = sqleng
        self._logger = Logger(self.__class__.__name__)

    @contextmanager
    def sessionctx(self, autocommit = True):
        sess =  self.session()
        err = False
        try:
            yield sess
        except:
            err = True
        finally:
            if autocommit and not err:
                sess.commit()
            else:
                sess.rollback()
            sess.close()
    
    def session(self):
        """ a raw session """
        return Session(self._engine)

class HKSvc(SvcBase):
    _qcaches = {}
    
    def _samekarat(self,jo0,jo1):
        """ check if the 2 given jo's karat/auxkarat are the same
            this method compare the 2 karats            
        """
        lst = ([(fmtkarat(x.orderma.karat),fmtkarat(x.auxkarat) \
            if x.auxwgt else 0) for x in (jo0,jo1)])
        return lst[0] == lst[1]

    def __init__(self, sqleng):
        """ init me with a sqlalchemy's engine """
        super(HKSvc,self).__init__(sqleng)
        self._ptnmit = re.compile("^M[A-Z]T")
    
    def _pjq(self):
        """ return a cached JO -> PajShp -> PajInv query """
        return self._qcaches.setdefault("jopajshp&inv", \
            Query([PajShp, JO, PajInv]).join(JO, JO.id == PajShp.joid) \
                .join(PajInv, and_(PajShp.joid == PajInv.joid, PajShp.invno == PajInv.invno)))
    
    def _pjsq(self):
        """ return a cached Style -> JO -> PajShp -> PajInv query """
        return self._qcaches.setdefault("jopajshp&inv", \
            Query([PajShp, JO, Style, Orderma, PajInv]).join(JO, JO.id == PajShp.joid) \
                .join(PajInv, and_(PajShp.joid == PajInv.joid, PajShp.invno == PajInv.invno)))

    def getjos(self,jesorrunns, psess = None):
        """get jos by a collection of JOElements/Strings or Integers
        when the first item is string or JOElement, it will be treated as getbyname, else by runn
        return a tuple, the first item is list containing hnjcore.models.hk.JO
                        the second item is a set of ids/jes/runns not found
        @param groupby: can be one of id/running/name, running should be a string
            starts with 'r' for example, 'r410100', id should be integer,
            name should be JOElement or string without 'r' as prefix
        """
        #todo:append a pattern to match a r\d{6} case to extract running
        #else treatd them as joid
        if not jesorrunns: return
        jes = set();rns = set();ids = set(); jos = {}
        for x in jesorrunns:
            if isinstance(x,JOElement):
                jes.add(x)
            elif isinstance(x,int):
                ids.add(x)
            elif isinstance(x,basestring):
                if x.find("r") >= 0:
                    rns.add(int(x[1:]))
                else:
                    je = JOElement(x)
                    if(je.isvalid): jes.add(je)
        if not any((jes,rns,ids)): return
        
        def _putjos(jos,mp,groupby):
            if not jos: return
            #groupby = "name" if not groupby else groupby.lower()
            if groupby.find("name") >= 0:
                mp1 = [(x.name,x) for x in jos]
            elif groupby.find("runn") >= 0:
                mp1 = [(x.running,x) for x in jos]
            else:
                mp1 = [(x.id,x) for x in jos]
            mp.update(dict(mp1))

        cur = psess if psess else self.session()
        q0 = Query(JO)
        try:
            if jes:
                for ii in splitarray(list(jes), self._querysize):
                    q = JO.name == ii[0]
                    for yy in ii[1:]:
                        q = or_(JO.name == yy,q)
                    q = q0.filter(q).with_session(cur)
                    _putjos(q.all(),jos,"id")
            if rns:
                for ii in splitarray(list(rns), self._querysize):
                    q = q0.filter(JO.running.in_(ii)).with_session(cur)
                    _putjos(q.all(),jos,"id")
            if ids:
                for ii in splitarray(list(ids), self._querysize):
                    q = q0.filter(JO.id.in_(ii)).with_session(cur)
                    _putjos(q.all(),jos,"id")
        except:
            pass
        finally:
            if cur and not psess: cur.close()
        failed = set()
        #check what's not got        
        its = jos.values()
        if jes:
            failed.update(jes.difference(set([x.name for x in its])))
        if rns:
            failed.update(rns.difference(set([x.running for x in its])))
        if ids:
            failed.update(ids.difference(set([x.id for x in its])))
        if not failed: failed = None        
        return (its,failed)

    def getjo(self,jeorrunn, psess = None):
        """ a convenient way for getjos """
        jos = self.getjos([jeorrunn],psess)
        return jos[0][0] if jos else None

    def getrevcn(self, pcode, psess = None):
        """return the revised for given pcode"""
        revcn = 0
        cur = psess if psess else self.session()
        try:
            q = cur.query(PajCnRev.uprice).order_by(desc(PajCnRev.tag))\
                .filter(PajCnRev.pcode == pcode).limit(2)
            rows = q.one()
            if rows:
                revcn = float(rows[0])
        except:
            pass
        finally:
            if cur and not psess: cur.close()
        return revcn

    def findsimilarjo(self, jo, level = 1, psess = None):
        """ return an list of JO based on below criteria
        @param level:   0 for extract SKU match
                        1 for extract karat match
                        1+ for extract style match
        """ 
        rc = None;level = 0 if level < 0 else level
        je = jo.name;jns = None
        cur = psess if psess else self.session()
        try:
            rows = cur.query(JO,POItem.skuno).join(Orderma).join(POItem,JO.poid == POItem.id)\
                .filter(Orderma.styid == jo.orderma.style.id).all()
            if(rows):                
                jns = {}
                for x in rows:
                    if x.JO.name == je: continue
                    key = fmtsku(x.skuno)
                    lst = jns.setdefault(key,[])
                    lst.append(x.JO)
            if jns:
                skuno = fmtsku(jo.po.skuno)
                sks = jns[skuno] if skuno and skuno in jns else None
                if not sks and level > 0:
                    sks = [x.JO for x in rows if je != x.JO.name and self._samekarat(jo,x.JO)]
                    if not sks and level > 1:
                        sks =[x.JO for x in rows]
                rc = sks
        except Exception as e:
            pass
        finally:
            if cur and not psess: cur.close()
        return rc

    def getjowgts(self,jo,psess = None):
        """ return a PrdWgt object of given JO """
        if not jo: return None
        if isinstance(jo,basestring) or isinstance(jo,JOElement):
            jo = self.getjos([jo],psess)[0]
            if not jo: return None
            jo = jo[0].values()[0]
        knws = [WgtInfo(jo.orderma.karat,jo.wgt), None, None]
        rk = knws[0]
        cur = psess if psess else self.session()
        try:
            if jo.auxwgt:
                knws[1] = WgtInfo(jo.auxkarat, float(jo.auxwgt))
                if(knws[1].karat == 925): #most of 925's parts is 925
                    rk = knws[1]
            lst = cur.query(JI).filter(and_(JI.joid == jo.id,JI.stname.like("M%T"))).all()
            if lst:
                row = lst[0]
                if(row.unitwgt and self._ptnmit.search(row.stname)):
                    knws[2] = pc.WgtInfo(rk.karat, float(row.unitwgt))
        except:
            pass
        finally:
            if cur and not psess: cur.close()
        if any(knws): return PrdWgt(knws[0],knws[1],knws[2])

    def calchina(self, je, psess = None):
        """ get the weight of given JO# and calc the china
            return a map with keys (JO,PajShp,PajInv,china,wgts)
         """        
        if(isinstance(je, basestring)):
            je = JOElement(je)
        elif isinstance(je,JO):
            je = je.name
        rmap = {"JO":None, "china":None, "PajShp":None, "PajInv": None, "wgts": None}
        if(not je.isvalid): return rmap
        cur = psess if psess else self.session()
        try:
            ups = self.getpajshpinv([je],psess = cur)
            if not ups: return rmap
            hnp = ups[0]; jo = hnp.JO
            prdwgt = self.getjowgts(jo)
            if prdwgt: rmap["wgts"] = prdwgt
            rmap["JO"] = jo; ups = hnp.PajInv.uprice
            rmap["PajInv"] = hnp.PajInv; rmap["PajShp"] = hnp.PajShp                        
            revcn = self.getrevcn(hnp.PajShp.pcode, psess = cur)
        except:
            pass
        finally:
            if cur and not psess: cur.close()
        rmap["china"] = pc.newchina(revcn, prdwgt) if revcn else \
            pc.PajCalc.calchina(prdwgt, float(hnp.PajInv.uprice), MPS(hnp.PajInv.mps))
        return rmap

    def getmmioforjc(self, df, dt, runns,psess = None):
        """return the mmstock's I/O# for PAJCReader"""
        lst = list()
        if not isinstance(runns[0], basestring): runns = [str(x) for x in runns]
        cur = psess if psess else self.session()
        try:
            for x in splitarray(runns, self._querysize):
                q = cur.query(SO.running, IV.remark1.label("jmp"), IV.docdate.label("shpdate") \
                    ,IV.inoutno).join(IVI).join(IV).filter(IV.inoutno.like("N%"))\
                    .filter(IV.remark1 != "").filter(and_(IV.docdate >= df,IV.docdate < dt))\
                    .filter(SO.running.in_(x))
                rows = q.all()
                if rows: lst.extend(rows)
        except:
            pass
        finally:
            if cur and not psess: cur.close()
        return lst
    
    def getpajshpinv(self, jes, psess = None):
        """ get the paj data for jocost
        @param jes: one or a list of JOElement()
        return a list of object contains JO/PajShp/PajInv objects
        """
        
        if not jes: return
        lst = []
        if not(isinstance(jes,list) or isinstance(jes,tuple)):
            if isinstance(jes,Iterable):
                jes = tuple(jes)
            elif isinstance(jes,JOElement):
                jes = [jes]

        cur = psess if psess else self.session()
        try:
            q0 = self._pjq()
            for ii in splitarray(jes, self._querysize):
                q = JO.name == ii[0]
                for yy in ii[1:]:
                    q = or_(JO.name == yy,q)
                q = q0.filter(q).with_session(cur)
                rows = q.all()
                if rows: lst.extend(rows)
        except:
            pass
        finally:
            if cur and not psess: cur.close()
        return lst
    
    def getpajprices(self, styno, psess = None):
        """return a list by PajRevcn as first element, then the cost sorted by joData"""
        cur = psess if psess else self.session()
        try:
            #todo::
            s0 = ("select sty.alpha,sty.digit,shp.pcode,jo.joalpha,jo.jodigit,shp.shpdate" 
                " from styma sty inner orderma od on sty.styid = od.styid"
                " inner join jo on od.orderid = jo.orderid inner join pajshp shp "
                " on jo.joid = shp.joid where (%s) order by jo.shipdate")
            cur.execute("select from")
            #lst = cur.fetchall()
        except:
            pass
        finally:
            if cur and not psess: cur.close()
 
    
class CNSvc(SvcBase):

    def getshpforjc(self, df, dt, psess = None):
        """return py shipment data for PAJCReader
        @param df: start date(include) a date ot datetime object
        @param dt: end date(exclude) a date or datetime object 
        """
        lst = None
        cur = psess if psess else self.session()
        try:
            if True:
                q0 = Query([JOcn,MMMa,Customercn,Stylecn,JOcn,MM])
            else:
                q0 = Query([JOcn.id,MMMa.refdate,Customercn.name.label("cstname"),JOcn.name.label("jono")\
                    ,Stylecn.name.label("styno"),JOcn.running,JOcn.karat,JOcn.description\
                    ,JOcn.qty.label("joqty"),MM.qty.label("shpqty"),MM.name.label("mmno")])
            q = q0.join(Customercn).join(Stylecn).join(MM).join(MMMa)\
                .filter(and_(MMMa.refdate >= df, MMMa.refdate < dt)).with_session(cur)
            lst = q.all()
        except:
            pass
        finally:
            if cur and not psess: cur.close()
        return lst

