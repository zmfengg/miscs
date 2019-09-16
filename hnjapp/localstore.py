#! coding=utf-8
'''
* @Author: zmFeng
* @Date: 2018-07-09 15:47:54
* @Last Modified by:   zmFeng
* @Last Modified time: 2018-07-09 15:47:54
classes to hold jodata into sqlite for faster calculation
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

_base = declarative_base()
''' a docno column was appended to PajItem on 2019/01/16,
    below is the script for existing table:
    alter table pajitem add column docno varchar(20)
'''


class PajItem(_base):
    ''' A Paj's product, use for pajrdrs.PajUPTracker and pajrdrs.PajBomHdlr '''
    __tablename__ = "pajitem"
    id = Column(Integer, primary_key=True, autoincrement=True)
    __table_args__ = (Index('idx_pajinv_pcode', 'pcode', unique=True),)
    pcode = Column(VARCHAR(20))
    docno = Column(VARCHAR(20))
    createdate = Column(TIMESTAMP)
    tag = tag = Column(SmallInteger)


class PajInv(_base):
    ''' use for pajrdrs.PajUPTracker '''
    __tablename__ = "pajinv"
    __table_args__ = (Index('idx_pajinv_pid', 'pid'),
                      Index('idx_pajinv_jono', 'jono'),
                      Index('idx_pajinv_pid_jono', 'pid', 'jono', unique=True))
    id = Column(Integer, primary_key=True, autoincrement=True)
    pid = Column(ForeignKey("pajitem.id"))
    jono = Column(VARCHAR(10))
    styno = Column(VARCHAR(10))
    mps = Column(VARCHAR(50))
    uprice = Column(DECIMAL(8, 3))
    cn = Column(DECIMAL(8, 2))
    mtlcost = Column(DECIMAL(8, 2))  #CN's metal cost
    otcost = Column(DECIMAL(8, 2))  #CN's other cost(Labour and stone(if))
    jodate = Column(DateTime)
    invdate = Column(DateTime)
    createdate = Column(TIMESTAMP)
    lastmodified = Column(TIMESTAMP)


class PajWgt(_base):
    ''' use for pajrdrs.PajUPTracker '''
    __tablename__ = "pajwgt"
    id = Column(Integer, primary_key=True, autoincrement=True)
    pid = Column(ForeignKey("pajitem.id"))
    wtype = Column(
        SmallInteger)  #wgt type, 0 -> main, 10 for sub, 100 for parts
    karat = Column(SmallInteger)
    wgt = Column(DECIMAL(6, 2))
    createdate = Column(DateTime)
    lastmodified = Column(DateTime)
    remark = VARCHAR(100)
    tag = Column(Integer)


class PajCnRev(_base):
    ''' use for pajrdrs.PajUPTracker '''
    __tablename__ = "pajcnrev"
    id = Column(Integer, primary_key=True, autoincrement=True)
    pid = Column(ForeignKey("pajitem.id"))
    uprice = Column(DECIMAL(6, 2))
    revdate = Column(DateTime)
    createdate = Column(DateTime)
    tag = Column(Integer)


class C1JC(_base):
    """ jo item for c1calc.C1JC """
    __tablename__ = "c1jc"
    __table_args__ = (
        Index('idx_c1jo_skuno', 'skuno'),
        Index('idx_c1jo_jono', 'jono', unique=True),
    )
    id = Column(Integer, primary_key=True, autoincrement=True)
    name = Column(VARCHAR(10), name="jono")
    styno = Column(VARCHAR(10))
    skuno = Column(VARCHAR(10))
    docno = Column(VARCHAR(50))
    setcost = Column(DECIMAL(6, 2))
    basecost = Column(DECIMAL(6, 2))
    labcost = Column(DECIMAL(6, 2))
    createdate = Column(DateTime)
    lastmodified = Column(DateTime)
    tag = Column(Integer)

    @property
    def jono(self):
        ''' alias of name '''
        return self.name


class FeaSOrN(object):
    """ the class for storing feature, can be numeric or string """
    _value = None

    def __init__(self, s, n):
        self._value = s or n

    @property
    def s(self):
        ''' the string value '''
        if not self._value:
            return None
        return self._value if isinstance(self._value, str) else None

    @property
    def n(self):
        ''' the numeric value '''
        if not self._value:
            return None
        return self._value if isinstance(self._value, Number) else None

    @property
    def v(self):
        ''' the value itself, no matter string or numeric '''
        return self._value

    def __composite_values__(self):
        return self.s, self.n


class C1JCFeature(_base):
    """ the features except stone of a c1jo, for c1calc.C1JC """
    __tablename__ = "c1jcfeature"
    id = Column(Integer, primary_key=True, autoincrement=True)
    jcid = Column(ForeignKey("c1jc.id"))
    jc = relationship("C1JC")
    name = Column(VARCHAR(30))
    values = Column(VARCHAR(30))
    valuen = Column(DECIMAL(6, 2))
    value = composite(FeaSOrN, values, valuen)


class C1JCStone(_base):
    """ stone of a c1jo, for c1calc.C1JC """
    __tablename__ = "c1jcstone"
    id = Column(Integer, primary_key=True, autoincrement=True)
    jcid = Column(ForeignKey("c1jc.id"))
    jc = relationship("C1JC")
    stone = Column(VARCHAR(15))
    shape = Column(VARCHAR(5))
    size = Column(VARCHAR(5))
    qty = Column(SmallInteger)
    wgt = Column(DECIMAL(6, 3))
    setting = Column(VARCHAR(50))


class PajBom(_base):
    ''' paj's bom item, used shipment main/part judgement in pajrdrs.PajBomHdlr '''
    __tablename__ = "pajbom"
    __table_args__ = (
        Index('idx_pajbom_itemid', 'pid'),
        Index('idx_pajbom_pid_mid', 'pid', 'mid', 'tag', unique=True),
    )
    id = Column(Integer, primary_key=True, autoincrement=True)
    pid = Column(ForeignKey("pajitem.id"))
    item = relationship("PajItem")
    mid = Column(Integer)
    name = Column(VARCHAR(100)) # _MAIN_ for wgt from bom_mstr, _NETWGT_ for netwgt(mtl+stone)
    karat = Column(Integer)
    wgt = Column(DECIMAL(6, 3))
    flag = Column(SmallInteger)  # 1 for main-part, 0 for chain
    createdate = Column(DateTime)
    lastmodified = Column(DateTime)
    tag = Column(SmallInteger) # 0 for current, gt 0 for revision

class Codetable(_base):
    '''
    codetable for storing misc data
    '''
    __tablename__ = "codetable"
    __table_args__ = (
        Index('idx_cd_name', 'name'),
    )
    id = Column(Integer, primary_key=True, autoincrement=True)
    name = Column(VARCHAR(100))
    coden0 = Column(Float)
    coden1 = Column(Float)
    coden2 = Column(Float)
    codec0 = Column(VARCHAR(250))
    codec1 = Column(VARCHAR(250))
    codec2 = Column(VARCHAR(250))
    coded0 = Column(DateTime)
    coded1 = Column(DateTime)
    coded2 = Column(DateTime)
    description = Column(VARCHAR(250))
    createdate = Column(DateTime)
    lastmodified = Column(DateTime)
    tag = Column(SmallInteger)

class Stysn(_base):
    ''' class for sty <-> sn lookup
    '''
    __tablename__ = "stysn"
    __table_args__ = (
        Index('idx_stysn_sty', 'styno'),
        Index('idx_stysn_sn', 'snno'),
        Index('idx_stysn_stysn', 'styno', 'snno', unique=True),
    )
    id = Column(Integer, primary_key=True, autoincrement=True)
    styno = Column(VARCHAR(10))
    snno = Column(VARCHAR(20))
    # S for SN#, snno is the snno of styno; P for hierachi, snno is the Parent of styno
    # K for keyword
    tag = Column(CHAR(1))

# product specification
