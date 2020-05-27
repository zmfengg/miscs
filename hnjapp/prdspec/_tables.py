'''
#! coding=utf-8
@Author:        zmFeng
@Created at:    2019-08-02
@Last Modified: 2019-08-02 3:18:52 pm
@Modified by:   zmFeng
tables for product specification
'''

from numbers import Number

from sqlalchemy import text
# from sqlalchemy.dialects.sqlite import DateTime, TIMESTAMP
from sqlalchemy.ext.declarative.api import declarative_base
from sqlalchemy.orm import relationship
from sqlalchemy.sql.schema import Column, ForeignKey, Index
from sqlalchemy.sql.sqltypes import (DECIMAL, VARCHAR, DateTime, Float,
                                     Integer, SmallInteger)

String = VARCHAR

_base = declarative_base()
_base.creatorid = Column(Integer)
_base.createddate = Column(DateTime)
_base.lastuserid = Column(Integer, server_default=text("'0'"))
_base.modifieddate = Column(DateTime)
_base.tag = Column(Integer, server_default=text("'0'"))


class Style(_base):
    '''
    the style table, holds the name and key data of a style
    '''
    __tablename__ = 'style'

    id = Column(Integer, primary_key=True, autoincrement=True)
    name = Column(String(20), unique=True)
    docno = Column(String(20), unique=False) 
    description = Column(String(255))
    netwgt = Column(DECIMAL(6, 2), server_default=text("'0.00'"))
    increment = Column(DECIMAL(8, 6), server_default=text("'0.00'"))
    dim = Column(String(255))
    size = Column(String(255))
    qclevel = Column(SmallInteger)


class Mat(_base):
    '''
    basic material info contains name/code/description only, other extended material should have its own table
    '''
    __tablename__ = 'mat'
    __table_args__ = (
        Index('idx_mat_name', 'name'),
        Index('idx_mat_code', 'code'),
    )

    id = Column(Integer, primary_key=True, autoincrement=True)
    type = Column(String(25)) # maybe METAL/STONE/PARTS/FINISHING or so on
    name = Column(String(255), unique=True)
    code = Column(String(10), unique=True)
    spec = Column(String(255))
    unit = Column(String(20))
    description = Column(String(255))


class Stymat(_base):
    '''
    materials inside a style
    '''
    __tablename__ = 'stymat'
    __table_args__ = (
        Index('idx_stymat_sty_mat', 'styid', 'matid', unique=True),
        Index('idx_stymat_sty_idx', 'styid', 'idx', unique=True),
    )

    id = Column(Integer, primary_key=True, autoincrement=True)
    styid = Column(Integer, ForeignKey("style.id"), index=True)
    style = relationship("Style")
    idx = Column(Integer, index=True)
    matid = Column(Integer, ForeignKey("mat.id"), index=True)
    mat = relationship("Mat")
    qty = Column(Integer)
    wgt = Column(DECIMAL(8, 4))
    remarks = Column(String(255))


class Stystset(_base):
    '''
    stone settings data, also with qty and weight
    '''
    __tablename__ = 'stystset'

    id = Column(Integer, ForeignKey("stymat.id"), primary_key=True)
    stymat = relationship("Stymat")
    setting = Column(SmallInteger)  # codetable('stystset.setting').coden0
    ismain = Column(SmallInteger, server_default=text("'0'"))
    remarks = Column(String(50))


class Stypidef(_base):
    '''
    defs for Stypi.defid
    '''
    __tablename__ = 'stypidef'
    __table_args__ = (
        Index('idx_stypidef_name', 'name', unique=True),
        Index('idx_stypidef_tag', 'tag', unique=False)
    )

    id = Column(Integer, primary_key=True, autoincrement=True)
    name = Column(String(50))
    # group = Column(String(50))    # maybe a def belongs to more than one group
    type = Column(String(50))  # STRING/DATE/NUMERIC
    format = Column(String(50))  # STRING/DATE_FORMAT/...
    remarks = Column(String(255))

class Stypi(_base):
    '''
    Style property item
    '''
    __tablename__ = 'stypi'
    id = Column(Integer, primary_key=True, autoincrement=True)
    defid = Column(Integer, ForeignKey('stypidef.id'), index=True)
    pdef = relationship('Stypidef')
    # a formatted value, string object as it self, other formatted by stypropdef.format
    valuec = Column(String(255))
    # numeric value should be placed here, others to valuec
    valuen = Column(Float)


class Styp(_base):
    '''
    style properties
    '''
    __tablename__ = 'styp'
    __table_args__ = (
        Index('idx_styp_sty_prop', 'styid', 'propid', unique=True),
        Index('idx_stypp_sty_idx', 'styid', 'idx', unique=True),
    )

    id = Column(Integer, primary_key=True, autoincrement=True)
    styid = Column(Integer, ForeignKey("style.id"), index=True)
    style = relationship("Style")
    idx = Column(Integer, index=True)
    propid = Column(Integer, ForeignKey("stypi.id"), index=True)
    prop = relationship('Stypi')
