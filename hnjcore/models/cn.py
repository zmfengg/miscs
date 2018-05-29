# coding=utf-8
'''
Created on Mar 5, 2018
models for hnjcn
@author: zmFeng
'''

#from sqlalchemy.orm import relationship, relation
from sqlalchemy.ext.declarative.api import declarative_base
from sqlalchemy.orm import composite,relationship
from sqlalchemy.sql.schema import Column, ForeignKey, UniqueConstraint
from sqlalchemy.sql.sqltypes import VARCHAR, Float, Integer, DateTime, DECIMAL

#from main import hnjcnCtx
from .utils import JOElement, StyElement

#Base = hnjcnCtx.base
CNBase = declarative_base()

class Style(CNBase):
    """ the styma table """    
    __tablename__ = "styma"
    id = Column(Integer,name = "styid",primary_key = True)
    alpha = Column(VARCHAR(2), nullable = False,name = 'alpha')
    digit = Column(Integer,name = "digit", nullable = False)
    description = Column(VARCHAR(50),name = "description", nullable = False)
    
    name = composite(StyElement,alpha,digit)
        
    UniqueConstraint(alpha,digit,name = 'idx_styno')
    
class Customer(CNBase):
    __tablename__ = "cstinfo"
    id = Column(Integer,name = "cstid",primary_key = True)
    name = Column(VARCHAR(15), unique = True, nullable = False,name = 'cstname')

class JO(CNBase):
    """ b_cust_bill table """
    __tablename__ = "b_cust_bill"
    id = Column(Integer,name = "jsid",primary_key = True)
    alpha = Column(VARCHAR(2), nullable = False,name = 'cstbldid_alpha')
    digit = Column(Integer,name = "cstbldid_digit", nullable = False)
    name = composite(JOElement,alpha,digit)
    running = Column(Integer,name = "running")
    description = Column(VARCHAR(50),name = "description", nullable = False)
    qty = Column(Float,name = 'quantity')
    karat = Column(Integer,name = "karat")
    
    styid = Column(Integer,ForeignKey('styma.styid'))
    style = relationship("Style")
    cstid = Column(Integer,ForeignKey('cstinfo.cstid'))
    customer = relationship("Customer")    
    
   
    UniqueConstraint(alpha,digit,name = 'idx_jono')

class MMMa(CNBase):
    __tablename__ = "mmma"
    id = Column(Integer, name = "refid", primary_key = True, autoincrement = False)
    name = Column(VARCHAR(11), name = "refno")
    karat = Column(Integer, name = "karat")
    refdate = Column(DateTime, name = "refdate")
    tag = Column(Integer, name = "tag")

class MM(CNBase):
    __tablename__ = "mm"
    id = Column(Integer, name = "mmid", primary_key = True, autoincrement = False)
    refid = Column(Integer, ForeignKey("mmma.refid"), name = "refid")
    
    name = Column(VARCHAR(8), name = "docno")
    jsid = Column(Integer, ForeignKey("b_cust_bill.jsid"), name = "jsid")

    qty = Column(DECIMAL, name = "qty")
    tag = Column(Integer, name = "tag")

