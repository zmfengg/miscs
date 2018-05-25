# coding=utf-8
"""
 * @Author: zmFeng 
 * @Date: 2018-05-25 14:21:01 
 * @Last Modified by:   zmFeng 
 * @Last Modified time: 2018-05-25 14:21:01 
 * the database services, including HK's and py's, and the out-dated bc's
 """

import re

from sqlalchemy.orm import Session
from sqlalchemy import and_

import pajcc as pc
from hnjcore import JOElement
from hnjcore.models.hk import JO, Customer, Style, Orderma, JOItem as JI, POItem


__all__ = ["HKSvc", "CNSvc"]


class HKSvc(object):
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
                    if(row.wgt > 0 and self._ptnmit.search(row.stname)):
                        knws[2] = pc.WgtInfo(rk.karat, float(row.wgt))
                        break
                jo = {"id": joid, "name": je, "styid": styid, "skuno": skuno,
                      "wgts": pc.PrdWgt(knws[0], knws[1], knws[2]), "cstname": cstname, "styno": styno} if any(knws) else None
        finally:
            cur.close()
        return jo


class CNSvc(object):
    pass
