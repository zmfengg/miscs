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

from sqlalchemy.orm import Session
from sqlalchemy import and_, desc
import pajcc as pc
from hnjcore import JOElement
from hnjcore.models.hk import JO, Customer, Style, Orderma, JOItem as JI
from hnjcore.models.hk import POItem, PajShp, PajInv, PajCnRev
from hnjcore.utils import samekarat, splitarray


__all__ = ["HKSvc", "CNSvc"]


class HKSvc(object):
    _querysize = 20
    _logger = Logger(__name__)

    def __init__(self, sqleng):
        """ init me with a sqlalchemy's engine """
        self._engine = sqleng
        self._ptnmit = re.compile("^M[A-Z]T")

    def getjo(self, je):
        """todo:: rename this function to sth. else, for example, prdwgt"""
        knws = [None, None, None]
        jo = None
        cur = Session(self._engine)
        try:
            # wgt info including mit
            if isinstance(je, basestring):
                je = JOElement(je)
            qry = cur.query(JO, POItem.skuno, JI.stname, JI.stsize, JI.unitwgt).join(POItem)\
                .outerjoin(JI, and_(JO.id == JI.joid, JI.stname.like("M%T"))).filter(JO.name == je)
            rows = qry.all()
            if rows:
                for row in rows:
                    jo = row.JO
                    if(not knws[0]):
                        knws[0] = pc.WgtInfo(jo.karat, float(jo.wgt))
                        rk = knws[0]
                        joid = jo.id
                        skuno = row.skuno
                        styid = jo.style.id
                        cstname = jo.customer.name.strip()
                        styno = jo.style.name
                        if(skuno):
                            skuno = skuno.strip()
                            if skuno in ("", "N/A"):
                                skuno = None
                            if skuno and [x for x in skuno if ord(x) <= 31 or ord(x) > 127]:
                                skuno = None
                        if(jo.auxwgt and jo.auxwgt > 0):
                            knws[1] = pc.WgtInfo(
                                jo.auxkarat, float(jo.auxwgt))
                            if(knws[1].karat == 925):
                                rk = knws[1]
                    if(not row.stname):
                        break
                    if(row.unitwgt > 0 and self._ptnmit.search(row.stname)):
                        knws[2] = pc.WgtInfo(rk.karat, float(row.unitwgt))
                        break
                jo = {"id": joid, "name": je, "styid": styid, "skuno": skuno,
                      "wgts": pc.PrdWgt(knws[0], knws[1], knws[2]), "cstname": cstname, "styno": styno} if any(knws) else None
        finally:
            cur.close()
        return jo

    def getpaj(self, jo):
        """ return the je,pcode,uprice,mps
        @param jo: a dict contains jo data, the dict can be returned by this.getjo(je)
        @return:  a map with keys(jono,pcode,uprice,mps)  
        todo:: change jo type to JE
        """
        def _mapx(x, je):
            return dict(zip("jono,pcode,uprice,mps".split(","), (je, x.pcode.strip(), x.uprice, x.mps)))

        ups = None
        cur = Session(self._engine)
        try:
            je = jo["name"]
            q = cur.query(PajShp.pcode, PajInv.uprice, PajInv.mps).join(JO, JO.id == PajShp.joid) \
                .join(PajInv, and_(PajShp.joid == PajInv.joid, PajShp.invno == PajInv.invno)) \
                .filter(JO.name == je)
            rows = q.all()
            ups = [_mapx(x, je)
                   for x in rows if x.uprice and x.mps] if(rows) else None
        finally:
            cur.close()
        return ups

    def getrevcn(self, pcode):
        """return the revised for given pcode"""
        revcn = 0
        cur = Session(self._engine)
        try:
            q = cur.query(PajCnRev.uprice).order_by(desc(PajCnRev.tag))\
                .filter(PajCnRev.pcode == pcode).limit(2)
            rows = q.all()
            if rows:
                revcn = float(rows[0].uprice)
        finally:
            cur.close()
        return revcn

    def extsearch(self, jo, level = 1):
        """find the JOs with the same sty# of given jo, which can be obtained by this.getjo(je). 
        return an list of JO based on below criteria
        @param level:   0 for extract SKU match
                        1 for extract karat match
                        1+ for extract style match
        """ 
        rc = None;level = 0 if level < 0 else level
        je = jo["name"];jns = None
        cur = Session(self._engine)        
        try:
            rows = cur.query(JO.name,POItem.skuno).join(Orderma).join(POItem,JO.poid == POItem.id)\
                .filter(Orderma.styid == jo["styid"]).all()
            if(rows): 
                jns = dict((x.name, x.skuno.strip()) for x in rows if x.name != je)
            if jns:
                sks = [x[0] for x in jns.iteritems() if x[1] == jo["skuno"]]
                if not sks and level > 0:
                    sks = [x[0] for x in jns.iteritems() if samekarat(je, x[0])]
                    if not sks and level > 1: sks = jns.keys
                if not sks and level > 1:
                    sks = sks.keys()
                rc = sks
        finally:
            cur.close()
        return [self.getjo(x) for x in rc] if rc else None

    def getjosbyrunns(self, runns):
        self._logger.debug("begin to fetch JO#s for running, count = %d" % len(runns))
        cur = Session(self._engine)
        mp = {}        
        try:            
            for x in splitarray(runns, self._querysize):
                q = cur.query(JO.running,JO.name).filter(JO.running.in_(x))
                rows = q.all()
                if(rows):
                    for pr in [(row.name, str(row.running)) for row in rows]:
                        if(not mp.has_key(pr[1])): mp[pr[1]] = pr[0]
                else:
                    break
        finally:
            cur.close()
        runns = set(runns)
        df = runns.difference(mp.keys()) if len(runns) > len(mp) else None
        self._logger.debug("Running -> JO done")
        return mp if mp else None, df

class CNSvc(object):
    pass
