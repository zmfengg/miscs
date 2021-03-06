# coding=utf-8
"""
 * @Author: zmFeng
 * @Date: 2018-05-24 14:36:46
 * @Last Modified by:   zmFeng
 * @Last Modified time: 2018-05-24 14:36:46
 """
from sqlalchemy.ext.declarative.api import declarative_base
from sqlalchemy.orm import composite, relationship
from sqlalchemy.sql.schema import Column, ForeignKey, UniqueConstraint
from sqlalchemy.sql.sqltypes import (DECIMAL, VARCHAR, DateTime, Float, Integer)

from .utils import JOElement, StyElement

HKBase = declarative_base()


class Customer(HKBase):
    __tablename__ = "cstinfo"
    id = Column(Integer, name="cstid", primary_key=True, autoincrement=False)
    name = Column(VARCHAR(30), name="cstname")
    description = Column(VARCHAR(40), name="description")
    range = Column(VARCHAR(20), name="range")
    groupid = Column(Integer, name="groupid")
    tag = Column(Integer, name="tag")

    pcstid = Column(Integer, ForeignKey("cstinfo.cstid"), name="pcstid")
    pcstinfo = relationship("Customer")


class Style(HKBase):
    __tablename__ = "styma"
    id = Column(Integer, name="styid", primary_key=True, autoincrement=False)
    alpha = Column(VARCHAR(5), name="alpha")
    digit = Column(Integer, name="digit")
    name = composite(StyElement, alpha, digit)

    description = Column(VARCHAR(50), name="description")
    edescription = Column(VARCHAR(50), name="edescription")
    ordercnt = Column(Integer, name="ordercnt")
    filldate = Column(DateTime, name="fill_date")
    tag = Column(Integer, name="tag")
    suffix = Column(VARCHAR(50), name="suffix")
    name1 = Column(VARCHAR(20), name="name1")


class Orderma(HKBase):
    __tablename__ = "orderma"
    id = Column(Integer, name="orderid", primary_key=True, autoincrement=False)
    orderno = Column(VARCHAR(50), name="orderno")

    cstid = Column(Integer, ForeignKey("cstinfo.cstid"), name="cstid")
    customer = relationship("Customer")
    styid = Column(Integer, ForeignKey("styma.styid"), name="styid")
    style = relationship("Style")

    karat = Column(Integer, name="karat")
    soqty = Column(DECIMAL, name="soqty")
    soprice = Column(DECIMAL, name="soprice")
    poqty = Column(DECIMAL, name="poqty")
    poprice = Column(DECIMAL, name="poprice")
    joqty = Column(DECIMAL, name="joqty")
    jofqty = Column(DECIMAL, name="jofqty")
    joprice = Column(DECIMAL, name="joprice")
    filldate = Column(DateTime, name="fill_date")
    lastmodified = Column(DateTime, name="modi_date")
    tag = Column(Integer, name="tag")


class PO(HKBase):
    __tablename__ = "poma"
    id = Column(Integer, name="pomaid", primary_key=True)
    cstid = Column(Integer, ForeignKey("cstinfo.cstid"), name="cstid")
    customer = relationship("Customer")
    name = Column(VARCHAR(50), name="pono")
    ordertype = Column(Integer, name="ordertype")
    filldate = Column(DateTime, name="fill_date")
    orderdate = Column(DateTime, name="order_date")
    receiptdate = Column(DateTime, name="receipt_date")
    canceldate = Column(DateTime, name="cancel_date")
    tag = Column(Integer, name="tag")
    mps = Column(VARCHAR(50), name="mps")


class POItem(HKBase):
    __tablename__ = "po"
    id = Column(Integer, name="poid", primary_key=True)
    pomaid = Column(Integer, ForeignKey("poma.pomaid"), name="pomaid")
    po = relationship("PO")
    orderid = Column(Integer, ForeignKey("orderma.orderid"), name="orderid")
    orderma = relationship("Orderma")
    qty = Column(DECIMAL, name="qty")
    uprice = Column(DECIMAL, name="uprice")
    skuno = Column(VARCHAR(50), name="skuno")
    rmk = Column(VARCHAR(20), name="rmk")
    description = Column(VARCHAR(50), name="description")
    tag = Column(Integer, name="tag")
    joqty = Column(DECIMAL, name="joqty")


class JO(HKBase):
    """ jo table """
    __tablename__ = "jo"
    id = Column(Integer, name="joid", primary_key=True, autoincrement=False)
    alpha = Column(VARCHAR(2), nullable=False, name='alpha')
    digit = Column(Integer, name="digit", nullable=False)
    name = composite(JOElement, alpha, digit)

    running = Column(Integer, name="running")
    description = Column(VARCHAR(50), name="description", nullable=False)
    edescription = Column(VARCHAR(255), name="edescription")
    qty = Column(Float, name='qty')
    orderid = Column(Integer, ForeignKey("orderma.orderid"), name="orderid")
    orderma = relationship("Orderma")

    wgt = Column(DECIMAL, name="wgt")
    auxkarat = Column(Integer, name="auxkarat")
    auxwgt = Column(DECIMAL, name="auxwgt")

    createdate = Column(DateTime, name="createdate")
    filldate = Column(DateTime, name="fill_date")
    deadline = Column(DateTime, name="deadline")
    shipdate = Column(DateTime, name="shipdate")

    soid = Column(Integer, name="soid")
    poid = Column(Integer, ForeignKey("po.poid"), name="poid")
    po = relationship("POItem")
    ponohk = Column(VARCHAR(10), name="ponohk")
    remark = Column(VARCHAR(250), name="remark")
    snno = Column(VARCHAR(250), name="snno")
    tag = Column(Integer, name="tag")

    UniqueConstraint(alpha, digit, name='idx_jono')

    @property
    def style(self):
        return None if not self.orderma else self.orderma.style

    @property
    def customer(self):
        return None if not self.orderma else self.orderma.customer

    @property
    def karat(self):
        return self.orderma.karat if self.orderma else None

    @property
    def styid(self):
        return self.orderma.styid if self.orderma else None


class JOItem(HKBase):
    __tablename__ = "cstdtl"
    # TODO this table's primary key is malform
    joid = Column(
        Integer,
        ForeignKey("jo.joid"),
        name="jsid",
        primary_key=True,
        autoincrement=False)
    jo = relationship("JO")
    # stname = Column(VARCHAR(4), name="sttype", primary_key=True)
    stname = Column(VARCHAR(4), name="sttype")
    stsize = Column(VARCHAR(10), name="stsize")
    unitwgt = Column(DECIMAL, name="wgt")
    qty = Column(Integer, name="quantity")
    wgt = Column(DECIMAL, name="wgt_calc")
    remark = Column(VARCHAR(30), name="remark")


class PajShp(HKBase):
    __tablename__ = "pajshp"
    id = Column(Integer, name="id", primary_key=True, autoincrement=True)
    fn = Column(VARCHAR(100), name="fn")
    pcode = Column(VARCHAR(30), name="pcode")
    invno = Column(VARCHAR(10), name="invno")
    qty = Column(Float, name="qty")
    orderno = Column(VARCHAR(20), name="orderno")
    mtlwgt = Column(Float, name="mtlwgt")
    stwgt = Column(Float, name="stwgt")
    invdate = Column(DateTime, name="invdate")
    shpdate = Column(DateTime, name="shpdate")
    filldate = Column(DateTime, name="filldate")
    lastmodified = Column(DateTime, name="lastmodified")

    joid = Column(Integer, ForeignKey("jo.joid"), name="joid")
    jo = relationship("JO")


class PajInv(HKBase):
    __tablename__ = "pajinv"
    id = Column(Integer, name="id", primary_key=True, autoincrement=True)
    invno = Column(VARCHAR(10), name="invno")
    qty = Column(Float, name="qty")
    uprice = Column(DECIMAL, name="uprice")
    mps = Column(VARCHAR(50), name="mps")
    stspec = Column(VARCHAR(200), name="stspec")
    lastmodified = Column(DateTime, name="lastmodified")
    china = Column(DECIMAL, name="china")

    joid = Column(Integer, ForeignKey("jo.joid"), name="joid")
    jo = relationship("JO")


class PajCnRev(HKBase):
    __tablename__ = "pajcnrev"
    id = Column(Integer, name="id", primary_key=True)
    pcode = Column(VARCHAR(30), name="pcode")
    uprice = Column(DECIMAL, name="uprice")
    revdate = Column(DateTime, name="revdate")
    filldate = Column(DateTime, name="filldate")
    tag = Column(Integer, name="tag")


class PajAck(HKBase):
    __tablename__ = "pajack"
    id = Column(Integer, name="id", primary_key=True, autoincrement=True)
    joid = Column(ForeignKey("jo.joid"), nullable=False, index=True)
    pcode = Column(VARCHAR(30), name="pcode", index=True)
    docno = Column(VARCHAR(50))
    uprice = Column(DECIMAL, name="uprice")
    mps = Column(VARCHAR(50), name="mps")
    ackdate = Column(DateTime)
    filldate = Column(DateTime)
    lastmodified = Column(DateTime)
    tag = Column(Integer, name="tag")

    jo = relationship("JO")

"""
create table pajack
(
id numeric(8,0) identity not null,
joid integer not null,
pcode varchar(30) not null,
docno varchar(100) not null,
mps varchar(50) not null,
uprice smallmoney,
ackdate smalldatetime,
filldate smalldatetime,
lastmodified smalldatetime,
tag smallint,
constraint pk_pajack primary key(id)
)

create unique index idx_pajack_name on pajack(joid,pcode,ackdate)
"""


class StockObjectMa(HKBase):
    __tablename__ = "stockobjectma"
    id = Column(Integer, name="srid", primary_key=True, autoincrement=False)
    styno = Column(VARCHAR(30), name="styno")
    running = Column(VARCHAR(30), name="running")
    name = Column(VARCHAR(60), name="stockcode")
    description = Column(VARCHAR(250), name="description")
    tag = Column(Integer, name="tag")
    qtyleft = Column(DECIMAL, name="qtyleft")
    tag = Column(Integer, name="type")


class Invoice(HKBase):
    __tablename__ = "invoicema"
    id = Column(Integer, name="invid", primary_key=True, autoincrement=False)
    inoutno = Column(VARCHAR(50), name="inoutno")
    docno = Column(VARCHAR(50), name="docno")
    docdate = Column(DateTime, name="docdate")
    locationidfrm = Column(Integer, name="locationidfrm")
    locationidto = Column(Integer, name="locationidto")
    remark1 = Column(VARCHAR(100), name="remark1")
    remark2 = Column(VARCHAR(100), name="remark2")
    lastuserid = Column(Integer, name="lastuserid")
    lastupdate = Column(DateTime, name="lastupdate")
    tag = Column(Integer, name="tag")


class InvoiceItem(HKBase):
    __tablename__ = "invoicedtl"
    id = Column(Integer, name="invdid", primary_key=True, autoincrement=False)
    invid = Column(Integer, ForeignKey("invoicema.invid"), name="invid")
    jono = Column(VARCHAR(20), name="jono")
    srid = Column(Integer, ForeignKey("stockobjectma.srid"), name="srid")
    stockobject = relationship("StockObjectMa")
    qty = Column(DECIMAL, name="qty")
    lastuserid = Column(Integer, name="lastuserid")
    lastupdate = Column(DateTime, name="lastupdate")
    tag = Column(Integer, name="tag")
