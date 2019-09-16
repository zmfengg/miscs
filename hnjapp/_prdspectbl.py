'''
#! coding=utf-8
@Author:        zmFeng
@Created at:    2019-08-02
@Last Modified: 2019-08-02 3:18:52 pm
@Modified by:   zmFeng
temp place for holding product specification tables
'''

from numbers import Number

# from sqlalchemy.dialects.sqlite import DateTime, TIMESTAMP
from sqlalchemy.ext.declarative.api import declarative_base
from sqlalchemy.orm import composite, relationship
from sqlalchemy.sql.schema import Column, ForeignKey, Index
from sqlalchemy import text
from sqlalchemy.sql.sqltypes import (DECIMAL, VARCHAR, Integer,
                                     SmallInteger, Float, CHAR, DateTime, TIMESTAMP)
String = VARCHAR

_base = object # set this to declarative_base
# below several class for prdspec
class SNDice(_base):
    __tablename__ = 'sndice'
    __table_args__ = (
        Index('idx_sndice_name', 'prefix', 'name', unique=True),
    )
    id = Column(Integer, primary_key=True, autoincrement=True)
    prefix = Column(String(2), primary_key=False, nullable=False)
    name = Column(String(15))
    tag = Column(SmallInteger)


class Style(_base):
    __tablename__ = 'style'

    id = Column(Integer, primary_key=True, autoincrement=True)
    name = Column(String(20), unique=True)
    description = Column(String(255))
    netwgt = Column(DECIMAL(6, 2), server_default=text("'0.00'"))
    increment = Column(String(50))
    keywords = Column(String(255), index=True)
    dim = Column(String(255))
    creatorid = Column(Integer)
    createddate = Column(DateTime)
    lastuserid = Column(Integer, server_default=text("'0'"))
    modifieddate = Column(DateTime)
    remarks = Column(String(255))
    size = Column(String(255))
    tag = Column(Integer, server_default=text("'0'"))

class Mat(_base):
    __tablename__ = 'mat'

    id = Column(Integer, primary_key=True, autoincrement=True)
    type = Column(String(255))
    name = Column(String(255), unique=True)
    code = Column(String(10), unique=True)
    description = Column(String(255))
    creatorid = Column(Integer)
    createddate = Column(DateTime)
    lastuserid = Column(Integer, index=True)
    modifieddate = Column(DateTime)
    tag = Column(Integer)

class Stymat(_base):
    __tablename__ = 'stymat'

    id = Column(Integer, primary_key=True, autoincrement=True)
    styid = Column(Integer, ForeignKey("style.id"), index=True)
    style = relationship("Style")
    idx = Column(Integer, index=True)
    matid = Column(Integer, ForeignKey("mat.id"), index=True)
    mat = relationship("Mat")
    qty = Column(Integer)
    wgt = Column(DECIMAL(8, 4))
    remarks = Column(String(255))
    tag = Column(Integer)


class Stystone(_base):
    __tablename__ = 'stystone'

    id = Column(Integer, ForeignKey("stymat.id"), primary_key=True)
    stymat = relationship("Stymat")
    setting = Column(String(255))
    wgtunit = Column(String(255))
    ismain = Column(SmallInteger, server_default=text("'0'"))
    remarks = Column(String(255))
