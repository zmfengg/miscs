# coding=utf-8
'''
Created on Mar 5, 2018
models for hnjcn
@author: zmFeng
'''

#from main import hnjcnCtx
from utils import JOElement,StyElement
from sqlalchemy.sql.schema import Column, UniqueConstraint, ForeignKey
from sqlalchemy.sql.sqltypes import Integer, VARCHAR, Float
from sqlalchemy.orm import composite
#from sqlalchemy.orm import relationship, relation
from sqlalchemy.ext.declarative.api import declarative_base

#Base = hnjcnCtx.base
Base = declarative_base()

def f2in1(alpha,digit):
    return '%s%4d' % (alpha,digit) 

class JO(Base):
    """ b_cust_bill table """
    __tablename__ = "b_cust_bill"
    id = Column(Integer,name = "jsid",primary_key = True)
    alpha = Column(VARCHAR(2), nullable = False,name = 'cstbldid_alpha')
    digit = Column(Integer,name = "cstbldid_digit", nullable = False)
    description = Column(VARCHAR(50),name = "description", nullable = False)
    qty = Column(Float,name = 'quantity')
    '''
    _styid = Column(Integer, ForeignKey("styma.id"), name="styid")
    _cstid = Column(Integer, ForeignKey("customer.id"), name="cstid")
    
    customer = relationship('Customer', lazy = 'joined')
    customer = relation("Customer",backref="cstinfo")
    style = relation('Style', backref='style')    
    '''
    styid = Column(Integer,ForeignKey('styma.id'))
    cstid = Column(Integer,ForeignKey('customer.id'))
    
    name = composite(JOElement,alpha,digit)
   
    UniqueConstraint(alpha,digit,name = 'idx_jono')

class Style(Base):
    """ the styma table """    
    __tablename__ = "styma"
    id = Column(Integer,name = "styid",primary_key = True)
    alpha = Column(VARCHAR(2), nullable = False,name = 'alpha')
    digit = Column(Integer,name = "digit", nullable = False)
    description = Column(VARCHAR(50),name = "description", nullable = False)
    
    name = composite(StyElement,alpha,digit)
        
    UniqueConstraint(alpha,digit,name = 'idx_styno')
    
class Customer(Base):
    __tablename__ = "cstinfo"
    id = Column(Integer,name = "cstid",primary_key = True)
    name = Column(VARCHAR(15), unique = True, nullable = False,name = 'cstname')
    