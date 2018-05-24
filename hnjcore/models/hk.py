# coding=utf-8 
"""
 * @Author: zmFeng 
 * @Date: 2018-05-24 14:36:46 
 * @Last Modified by:   zmFeng 
 * @Last Modified time: 2018-05-24 14:36:46 
 """
from sqlalchemy.ext.declarative.api import declarative_base
from sqlalchemy.sql.sqltypes import Integer,VARCHAR,Float,DateTime,DECIMAL,Numeric
from sqlalchemy.sql.schema import Column,UniqueConstraint,ForeignKey
from sqlalchemy.orm import composite,relationship
from hnjcore import JOElement

HKBase = declarative_base()

class JO(HKBase):
    """ jo table """
    __tablename__ = "jo"
    id = Column(Integer,name = "joid",primary_key = True)
    alpha = Column(VARCHAR(2), nullable = False,name = 'alpha')
    digit = Column(Integer,name = "digit", nullable = False)
    description = Column(VARCHAR(50),name = "description", nullable = False)
    qty = Column(Float,name = 'qty')

    name = composite(JOElement,alpha,digit)
   
    UniqueConstraint(alpha,digit,name = 'idx_jono')

class PajShp(HKBase):
    """ pajshp table """
    __tablename__ = "pajshp"
    id = Column(Numeric(9,0), name = "id", primary_key = True)
    fn = Column(VARCHAR(100), name = "fn")
    pcode = Column(VARCHAR(30), name = "pcode")
    invno = Column(VARCHAR(10), name = "invno")
    qty = Column(Float, name = "qty")
    orderno = Column(VARCHAR(20), name = "orderno")
    mtlwgt = Column(Float, name = "mtlwgt")
    stwgt = Column(Float, name = "stwgt")
    invdate = Column(DateTime, name = "invdate")
    shpdate = Column(DateTime, name = "shpdate")
    filldate = Column(DateTime, name = "filldate")
    lastmodified = Column(DateTime, name = "lastmodified")

    joid = Column(Integer,ForeignKey("jo.joid"),name = "joid")    
    jo = relationship("JO")

class PajInv(HKBase):
    __tablename__ = "pajinv"
    id = Column(Numeric(9,0), name = "id", primary_key = True)
    invno = Column(VARCHAR(10), name = "invno")
    qty = Column(Float, name = "qty")
    uprice = Column(DECIMAL, name = "uprice")
    mps = Column(VARCHAR(50), name = "mps")
    stspec = Column(VARCHAR(200), name = "stspec")
    lastmodified = Column(DateTime, name = "lastmodified")
    china = Column(DECIMAL, name = "china")

    joid = Column(Integer,ForeignKey("jo.joid"),name = "joid")
    jo = relationship("JO")