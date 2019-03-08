# coding=utf-8
'''
Created on Mar 5, 2018
models for hnjcn
@author: zmFeng
'''

from sqlalchemy import text
from sqlalchemy.dialects.sybase.base import TINYINT
#from sqlalchemy.orm import relationship, relation
from sqlalchemy.ext.declarative.api import declarative_base
from sqlalchemy.orm import composite, relationship
from sqlalchemy.sql.schema import Column, ForeignKey, UniqueConstraint
from sqlalchemy.sql.sqltypes import (DECIMAL, VARCHAR, DateTime, Float, Integer,
                                     SmallInteger)

#from main import hnjcnCtx
from .utils import JOElement, StyElement

#Base = hnjcnCtx.base
CNBase = declarative_base()


class Style(CNBase):
    """ the styma table"""
    __tablename__ = "styma"
    id = Column(Integer, name="styid", primary_key=True)
    alpha = Column(VARCHAR(2), nullable=False, name='alpha')
    digit = Column(Integer, name="digit", nullable=False)
    description = Column(VARCHAR(50), name="description", nullable=False)

    name = composite(StyElement, alpha, digit)

    UniqueConstraint(alpha, digit, name='idx_styno')


class Customer(CNBase):
    __tablename__ = "cstinfo"
    id = Column(Integer, name="cstid", primary_key=True)
    name = Column(VARCHAR(15), unique=True, nullable=False, name='cstname')


class JO(CNBase):
    """ b_cust_bill table"""
    __tablename__ = "b_cust_bill"
    id = Column(Integer, name="jsid", primary_key=True)
    alpha = Column(VARCHAR(2), nullable=False, name='cstbldid_alpha')
    digit = Column(Integer, name="cstbldid_digit", nullable=False)
    name = composite(JOElement, alpha, digit)
    running = Column(Integer, name="running")
    description = Column(VARCHAR(50), name="description", nullable=False)
    qty = Column(Float, name='quantity')
    qtyleft = Column(Float, name="qtyleft")
    karat = Column(Integer, name="karat")
    createdate = Column(DateTime, nullable=False, name="createdate")
    lastupdate = Column(DateTime, nullable=False, name="modi_date")
    deadline = Column(DateTime, nullable=False, name="dead_line")
    docno = Column(VARCHAR(7), nullable=False, name="dept_bill_id")
    ordertype = Column(VARCHAR(1), name="remark", nullable=False)
    tag = Column(Integer, name="tag")

    styid = Column(Integer, ForeignKey('styma.styid'))
    style = relationship("Style")
    cstid = Column(Integer, ForeignKey('cstinfo.cstid'))
    customer = relationship("Customer")

    UniqueConstraint(alpha, digit, name='idx_jono')


class MMMa(CNBase):
    __tablename__ = "mmma"
    id = Column(Integer, name="refid", primary_key=True, autoincrement=False)
    name = Column(VARCHAR(11), name="refno")
    karat = Column(Integer, name="karat")
    refdate = Column(DateTime, name="refdate")
    tag = Column(Integer, name="tag")


class MM(CNBase):
    __tablename__ = "mm"
    id = Column(Integer, name="mmid", primary_key=True, autoincrement=False)
    refid = Column(Integer, ForeignKey("mmma.refid"), name="refid")

    name = Column(VARCHAR(8), name="docno")
    jsid = Column(Integer, ForeignKey("b_cust_bill.jsid"), name="jsid")

    qty = Column(DECIMAL, name="qty")
    tag = Column(Integer, name="tag")


class MMgd(CNBase):
    __tablename__ = "mmgd"
    id = Column(
        Integer,
        ForeignKey("mm.mmid"),
        name="mmid",
        primary_key=True,
        autoincrement=False)
    karat = Column(Integer, name="karat", primary_key=True, autoincrement=False)
    wgt = Column(DECIMAL, name="wgt")


class Codetable(CNBase):
    __tablename__ = "codetable"
    id = Column(Integer, primary_key=True, autoincrement=False)
    tblname = Column(VARCHAR(20))
    colname = Column(VARCHAR(20))
    coden0 = Column(DECIMAL, name="coden")
    coden1 = Column(DECIMAL, name="coden1")
    coden2 = Column(DECIMAL, name="coden2")
    codec0 = Column(VARCHAR(255), name="codec")
    codec1 = Column(VARCHAR(255), name="codec1")
    codec2 = Column(VARCHAR(255), name="codec2")
    coded0 = Column(DateTime, name="coded")
    coded1 = Column(DateTime)
    coded2 = Column(DateTime)
    code = Column(VARCHAR(255))
    description = Column(VARCHAR(255))
    filldate = Column(DateTime, name="fill_date")
    tag = Column(Integer)
    pid = Column(Integer)


class StoneMaster(CNBase):
    __tablename__ = 'stone_master'

    id = Column(SmallInteger, primary_key=True, name="stid")
    name = Column(VARCHAR(4), nullable=False, unique=True, name="stname")
    edesc = Column(VARCHAR(50), nullable=False)
    cdesc = Column(VARCHAR(50), nullable=False)
    sttype = Column(VARCHAR(20), nullable=False)
    settype = Column(TINYINT, nullable=False)
    filldate = Column(DateTime, nullable=False, name="fill_date")
    tag = Column(TINYINT, nullable=False)
    lastuserid = Column(SmallInteger, server_default=text("0"))
    lastupdate = Column(DateTime, server_default=text("getdate()"))


class StoneIn(CNBase):
    __tablename__ = 'stone_in'

    id = Column(Integer, primary_key=True, name="btchid", autoincrement=False)
    pkid = Column(ForeignKey('stone_pkma.pkid'), nullable=False)
    name = Column(VARCHAR(15), nullable=False, unique=True, name="batch_id")
    docno = Column(VARCHAR(15), nullable=False, name="bill_id")
    qty = Column(Integer, nullable=False, name="quantity")
    wgt = Column(Float, nullable=False, name="weight")
    qtyused = Column(
        Integer, nullable=False, server_default=text("0"), name="qty_used")
    wgtused = Column(
        Float, nullable=False, server_default=text("0"), name="wgt_used")
    qtybck = Column(Integer, server_default=text("0"), name="qty_bck")
    wgtbck = Column(
        Float, nullable=False, server_default=text("0"), name="wgt_bck")
    size = Column(VARCHAR(60), nullable=False)
    filldate = Column(DateTime, nullable=False, name="fill_date")
    tag = Column(SmallInteger, nullable=False, index=True, name="is_used_up")
    wgtadj = Column(
        Float, nullable=False, server_default=text("0"), name="adjust_wgt")
    cstid = Column(ForeignKey('cstinfo.cstid'), nullable=False)
    cstref = Column(VARCHAR(60), nullable=False)
    qtytrans = Column(Integer, nullable=False, server_default=text("0"))
    wgttrans = Column(Float, nullable=False, server_default=text("0"))
    wgttmptrans = Column(
        Float, nullable=False, server_default=text("0"), name="wgt_tmptrans")
    wgtprep = Column(
        Float, nullable=False, server_default=text("0"), name="wgt_prepared")
    lastuserid = Column(SmallInteger, server_default=text("0"))
    lastupdate = Column(DateTime, server_default=text("getdate()"))

    #cstinfo = relationship('Cstinfo')
    package = relationship('StonePk')


class StonePk(CNBase):
    __tablename__ = 'stone_pkma'

    id = Column(Integer, primary_key=True, name="pkid")
    name = Column(VARCHAR(20), nullable=False, unique=True, name="package_id")
    unit = Column(SmallInteger, nullable=False)
    pricec = Column(VARCHAR(6), nullable=False, name="price")
    pricen = Column(Float, name="pricen")
    stid = Column(ForeignKey('stone_master.stid'))
    stshpid = Column(SmallInteger)
    tag = Column(TINYINT)
    filldate = Column(DateTime, name="fill_date")
    relpkid = Column(Integer)
    stsizeidf = Column(SmallInteger)
    stsizeidt = Column(SmallInteger)
    lastuserid = Column(SmallInteger, server_default=text("0"))
    lastupdate = Column(DateTime, server_default=text("getdate()"))
    wgtunit = Column(Float, server_default=text("0.2"))
    color = Column(VARCHAR(50), server_default=text("N/A"))

    stone_master = relationship('StoneMaster')


class StoneOutMaster(CNBase):
    __tablename__ = 'stone_out_master'

    id = Column(Integer, primary_key=True, autoincrement=False)
    name = Column(Integer, nullable=False, name="bill_id")
    isout = Column(SmallInteger, nullable=False, name="is_out")
    joid = Column(ForeignKey('b_cust_bill.jsid'), nullable=False, name="jsid")
    qty = Column(DECIMAL, nullable=False)
    filldate = Column(DateTime, nullable=False, name="fill_date")
    packed = Column(TINYINT, nullable=False, index=True)
    subcnt = Column(TINYINT, nullable=False)
    workerid = Column(SmallInteger, server_default=text("0"), name="worker_id")
    lastuserid = Column(SmallInteger, server_default=text("0"))
    lastupdate = Column(DateTime, server_default=text("getdate()"))

    jo = relationship('JO')


class StoneOut(CNBase):
    __tablename__ = 'stone_out'

    id = Column(
        ForeignKey('stone_out_master.id'),
        primary_key=True,
        nullable=False,
        autoincrement=False)
    idx = Column(TINYINT, primary_key=True, nullable=False)
    btchid = Column(ForeignKey('stone_in.btchid'), nullable=False, index=True)
    workerid = Column(SmallInteger, nullable=False, name="worker_id")
    qty = Column(Integer, nullable=False, name="quantity")
    wgt = Column(Float, nullable=False, name="weight")
    checkerid = Column(SmallInteger, nullable=False, name="checker_id")
    checkdate = Column(DateTime, nullable=False, name="check_date")
    joqty = Column(SmallInteger, name="qty")
    printid = Column(Integer)
    lastuserid = Column(SmallInteger, server_default=text("0"))
    lastupdate = Column(DateTime, server_default=text("getdate()"))

    stonein = relationship('StoneIn')


class StoneBck(CNBase):
    __tablename__ = 'stone_bck'

    btchid = Column(
        ForeignKey('stone_in.btchid'),
        primary_key=True,
        nullable=False,
        name="btchid")
    idx = Column(SmallInteger, primary_key=True, nullable=False)
    wgt = Column(Float, nullable=False, name="weight")
    docno = Column(VARCHAR(8), nullable=False, name="bill_id")
    filldate = Column(DateTime, nullable=False, name="fill_date")
    lastuserid = Column(SmallInteger, server_default=text("0"))
    lastupdate = Column(DateTime, server_default=text("getdate()"))
    qty = Column(Integer, server_default=text("0"), name="quantity")

    stonein = relationship('StoneIn')


class Plating(CNBase):
    __tablename__ = 'goldplating'
    id = Column(Integer, primary_key=True, nullable=False)
    joid = Column(ForeignKey('b_cust_bill.jsid'), nullable=False, name="jsid")
    vendorid = Column(SmallInteger)
    height = Column(VARCHAR(8))
    uprice = Column(Float)
    qty = Column(Float)
    remark = Column(VARCHAR(100))
    filldate = Column(DateTime, name="fill_date")
    tag = Column(SmallInteger)

    jo = relationship('JO')
