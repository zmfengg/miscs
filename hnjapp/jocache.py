#! coding=utf-8 
'''
* @Author: zmFeng 
* @Date: 2018-07-09 15:47:54 
* @Last Modified by:   zmFeng 
* @Last Modified time: 2018-07-09 15:47:54 
classes to hold jodata into sqlite for faster calculation
'''

from sqlalchemy.ext.declarative.api import declarative_base
from sqlalchemy.sql.schema import Column, ForeignKey, UniqueConstraint, Index
from sqlalchemy.sql.sqltypes import (DECIMAL, VARCHAR, Float,
                                     Integer, SmallInteger)
from sqlalchemy.dialects.sqlite import DATETIME,TIMESTAMP
Base = declarative_base()

class PajInv(Base):
    __tablename__ = "pajinv"
    __table_args__ = (
        Index('idx_pajinv_name', 'name'),
        Index('idx_pajinv_jono', 'jono'),
        Index('idx_pajinv_name_jono', 'name','jono', unique = True)
    )
    id = Column(Integer,primary_key = True,autoincrement = True)
    name = Column(VARCHAR(30))
    jono = Column(VARCHAR(10))
    mps = Column(VARCHAR(50))
    uprice = Column(DECIMAL)
    invdate = Column(DATETIME)
    createdate = Column(TIMESTAMP)
    lastmodified = Column(TIMESTAMP)

class PrdWgt(Base):
    __tablename__ = "prodwgt"
    id = Column(Integer,primary_key = True,autoincrement = True)
    pid = Column(ForeignKey("pajinv.id"))
    wtype = Column(SmallInteger) #wgt type, 0 -> main, 1 for sub, 100 for parts
    karat = Column(SmallInteger)
    wgt = Column(DECIMAL)
    createdate = Column(DATETIME)
    lastmodified = Column(DATETIME)

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