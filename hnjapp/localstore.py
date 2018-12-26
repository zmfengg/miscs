#! coding=utf-8
'''
* @Author: zmFeng
* @Date: 2018-07-09 15:47:54
* @Last Modified by:   zmFeng
* @Last Modified time: 2018-07-09 15:47:54
classes to hold jodata into sqlite for faster calculation
'''

from numbers import Number

from sqlalchemy.dialects.sqlite import DATETIME, TIMESTAMP
from sqlalchemy.ext.declarative.api import declarative_base
from sqlalchemy.orm import composite, relationship
from sqlalchemy.sql.schema import Column, ForeignKey, Index, UniqueConstraint
from sqlalchemy.sql.sqltypes import (DECIMAL, VARCHAR, Float, Integer,
                                     SmallInteger)

Base = declarative_base()


class PajItem(Base):
    __tablename__ = "pajitem"
    id = Column(Integer, primary_key=True, autoincrement=True)
    __table_args__ = (Index('idx_pajinv_pcode', 'pcode', unique=True),)
    pcode = Column(VARCHAR(20))
    createdate = Column(TIMESTAMP)
    tag = tag = Column(SmallInteger)


class PajInv(Base):
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
    jodate = Column(DATETIME)
    invdate = Column(DATETIME)
    createdate = Column(TIMESTAMP)
    lastmodified = Column(TIMESTAMP)


class PajWgt(Base):
    __tablename__ = "pajwgt"
    id = Column(Integer, primary_key=True, autoincrement=True)
    pid = Column(ForeignKey("pajitem.id"))
    wtype = Column(
        SmallInteger)  #wgt type, 0 -> main, 10 for sub, 100 for parts
    karat = Column(SmallInteger)
    wgt = Column(DECIMAL(6, 2))
    createdate = Column(DATETIME)
    lastmodified = Column(DATETIME)
    remark = VARCHAR(100)
    tag = Column(Integer)


class PajCnRev(Base):
    __tablename__ = "pajcnrev"
    id = Column(Integer, primary_key=True, autoincrement=True)
    pid = Column(ForeignKey("pajitem.id"))
    uprice = Column(DECIMAL(6, 2))
    revdate = Column(DATETIME)
    createdate = Column(DATETIME)
    tag = Column(Integer)


class C1JC(Base):
    """ jo item for C1 """
    __tablename__ = "c1jc"
    __table_args__ = (Index('idx_c1jo_skuno', 'skuno'),
                    Index('idx_c1jo_jono', 'jono', unique=True), )
    id = Column(Integer, primary_key=True, autoincrement=True)
    name = Column(VARCHAR(10), name="jono")
    styno = Column(VARCHAR(10))
    skuno = Column(VARCHAR(10))
    docno = Column(VARCHAR(50))
    setcost = Column(DECIMAL(6, 2))
    basecost = Column(DECIMAL(6, 2))
    labcost = Column(DECIMAL(6, 2))
    createdate = Column(DATETIME)
    lastmodified = Column(DATETIME)
    tag = Column(Integer)

class FeaSOrN(object):
    """ the class for storing feature, can be numeric or string """
    _value = None

    def __init__(self, s, n):
        self._value = s or n

    @property
    def s(self):
        if not self._value:
            return None
        return self._value if isinstance(self._value, str) else None

    @property
    def n(self):
        if not self._value:
            return None
        return self._value if isinstance(self._value, Number) else None

    @property
    def v(self):
        return self._value

    def __composite_values__(self):
        return self.s, self.n

class C1JCFeature(Base):
    """ the features except stone of a c1jo """
    __tablename__ = "c1jcfeature"
    id = Column(Integer, primary_key=True, autoincrement=True)
    jcid = Column(ForeignKey("c1jc.id"))
    jc = relationship("C1JC")
    name = Column(VARCHAR(30))
    values = Column(VARCHAR(30))
    valuen = Column(DECIMAL(6, 2))
    value = composite(FeaSOrN, values, valuen)

class C1JCStone(Base):
    """ stone of a c1 JO """
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

"""
statement for sqlite table creation

create table if not exists pajinv(
    id INTEGER primary key asc,
    name varchar(30),
    jono VARCHAR(10),
    mps VARCHAR(50),
    uprice DECIMAL(8,2),
    invdate Integer,
    createdate Integer,
    lastmodified Integer
)

create index if not exists idx_pajinv_name on pajinv(name)
create index if not exists idx_pajinv_jono on pajinv(jono)
create unique index if not exists idx_pajinv_name_jono on pajinv(name,jono)

create table if not exists prdwgt(
    id Integer primary key asc,
    pid Integer,
    wtype Tinyint,
    karat SmallInt,
    wgt DECIMAL(4,2),
    createdate Integer,
    lastmodified Integer,
    constraint fk_prdwgt_ref_pajinv foreign key (pid) references pajinv(id)
)
"""
