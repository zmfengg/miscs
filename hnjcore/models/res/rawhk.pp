# coding: utf-8
from sqlalchemy import CHAR, Column, DECIMAL, DateTime, Float, ForeignKey, Index, Integer, Numeric, SmallInteger, String, Table, Unicode, text
from sqlalchemy.dialects.sybase.base import MONEY, SMALLMONEY, TINYINT
from sqlalchemy.orm import relationship
from sqlalchemy.ext.declarative import declarative_base

HKBase = declarative_base()
metadata = Base.metadata


class Codetable(HKBase):
    __tablename__ = 'codetable'

    id = Column(Integer, primary_key=True)
    tblname = Column(String(20), nullable=False, index=True)
    colname = Column(String(20), nullable=False)
    coden = Column(Float, nullable=False, index=True)
    coden1 = Column(Float, nullable=False)
    coden2 = Column(Float, nullable=False)
    codec = Column(String(255), nullable=False, index=True)
    codec1 = Column(String(255), nullable=False)
    codec2 = Column(String(255), nullable=False)
    coded = Column(DateTime, nullable=False)
    coded1 = Column(DateTime, nullable=False)
    coded2 = Column(DateTime, nullable=False)
    code = Column(String(30), nullable=False)
    description = Column(String(40), nullable=False)
    fill_date = Column(DateTime, nullable=False)
    tag = Column(SmallInteger, nullable=False, index=True)
    pid = Column(Integer, nullable=False, index=True)


class Cstbldrng(HKBase):
    __tablename__ = 'cstbldrng'

    cstbldid_alpha = Column(CHAR(1), primary_key=True, nullable=False)
    min_digit = Column(Integer, primary_key=True, nullable=False)
    max_digit = Column(Integer, nullable=False)
    tag = Column(TINYINT, nullable=False)


class Cstbldtpym(HKBase):
    __tablename__ = 'cstbldtpym'

    docid = Column(Integer, primary_key=True)
    docno = Column(String(10), nullable=False)
    fill_date = Column(DateTime, nullable=False)
    qty = Column(SmallInteger, nullable=False)
    tag = Column(Integer, nullable=False, index=True)


class Cstinfo(HKBase):
    __tablename__ = 'cstinfo'

    cstid = Column(SmallInteger, primary_key=True)
    cstname = Column(CHAR(30), nullable=False, unique=True)
    description = Column(CHAR(40), nullable=False)
    range = Column(CHAR(20), nullable=False)
    groupid = Column(SmallInteger, nullable=False)
    tag = Column(TINYINT, nullable=False)
    pcstid = Column(SmallInteger, server_default=text("-1"))


class Employee(HKBase):
    __tablename__ = 'employee'

    id = Column(Integer, primary_key=True)
    name = Column(String(20), nullable=False, unique=True)
    empname = Column(String(30), nullable=False)
    tag = Column(Integer, nullable=False)


class Fdmdm(HKBase):
    __tablename__ = 'fdmdm'

    sid = Column(Integer, primary_key=True)
    fill_date = Column(DateTime, nullable=False)
    modi_date = Column(DateTime, nullable=False)
    subcnt = Column(SmallInteger, nullable=False)
    tag = Column(Integer, nullable=False)


class InsertKey(HKBase):
    __tablename__ = 'insert_key'
    __table_args__ = (
        Index('idx_inskey_tblcolname', 'tblname', 'colname'),
    )

    id = Column(Integer, primary_key=True)
    tblname = Column(String(50), nullable=False)
    colname = Column(String(50))
    maxvalue = Column(String(255))
    remark = Column(String(255))
    tag = Column(Integer, index=True)


class InsertKey1(HKBase):
    __tablename__ = 'insert_key1'
    __table_args__ = (
        Index('idx_insertkey1_pk', 'tblname', 'colname'),
    )

    id = Column(Integer, primary_key=True)
    tblname = Column(String(50), nullable=False)
    colname = Column(String(50), nullable=False)
    nvalue = Column(Integer, nullable=False)
    nvalue1 = Column(Float, nullable=False)
    cvalue = Column(String(255), nullable=False)
    cvalue1 = Column(String(255), nullable=False)
    ttype = Column(Integer, nullable=False)
    processid = Column(Integer, nullable=False, index=True)
    tag = Column(Integer, nullable=False)


class Jogap(HKBase):
    __tablename__ = 'jogap'
    __table_args__ = (
        Index('idx_jogap_joidf_joidt', 'joidf', 'joidt'),
    )

    id = Column(SmallInteger, primary_key=True)
    joidf = Column(Integer, nullable=False)
    joidt = Column(Integer, nullable=False, index=True)


class Jopkprice(HKBase):
    __tablename__ = 'jopkprice'

    id = Column(SmallInteger, primary_key=True)
    name = Column(String(50), nullable=False)
    size = Column(String(50), nullable=False)
    clarity = Column(String(50), nullable=False, server_default=text(""))
    price = Column(String(50), nullable=False)
    filldate = Column(DateTime, nullable=False, server_default=text("getdate()"))


class Joprop(HKBase):
    __tablename__ = 'joprop'

    joid = Column(Integer, primary_key=True)
    smptype = Column(TINYINT, nullable=False)


t_key_buffer = Table(
    'key_buffer', metadata,
    Column('tblname', String(30), nullable=False),
    Column('colname', String(30), nullable=False),
    Column('nvalue', Integer, nullable=False),
    Column('processid', Integer, nullable=False),
    Column('createdate', DateTime, nullable=False),
    Column('tag', SmallInteger, nullable=False),
    Index('pk_key_buffer_tblname', 'tblname', 'colname', 'nvalue', unique=True)
)


class Locationma(HKBase):
    __tablename__ = 'locationma'

    locationid = Column(Integer, primary_key=True)
    location = Column(String(20), index=True)
    plocationid = Column(Integer)
    status = Column(SmallInteger)
    tag = Column(SmallInteger)


class Mitbck(HKBase):
    __tablename__ = 'mitbck'

    mitid = Column(Integer, primary_key=True, nullable=False)
    idx = Column(TINYINT, primary_key=True, nullable=False)
    qty = Column(SMALLMONEY, nullable=False)
    wgt = Column(Float, nullable=False)
    fill_date = Column(DateTime, nullable=False)


class Mitin(HKBase):
    __tablename__ = 'mitin'

    mitid = Column(Integer, primary_key=True)
    cstid = Column(SmallInteger, nullable=False, index=True)
    invid = Column(String(10), nullable=False)
    stcode = Column(String(20), nullable=False)
    mit = Column(String(15), nullable=False, unique=True)
    karat = Column(SmallInteger, nullable=False)
    qty = Column(SMALLMONEY, nullable=False)
    wgt = Column(Float, nullable=False)
    size = Column(String(7), nullable=False)
    description = Column(String(45), nullable=False)
    cost = Column(SMALLMONEY, nullable=False)
    unit = Column(TINYINT, nullable=False)
    in_date = Column(DateTime, nullable=False)
    qtyused = Column(SMALLMONEY, nullable=False)
    wgtused = Column(Float, nullable=False)
    qtybck = Column(SMALLMONEY, nullable=False)
    wgtbck = Column(Float, nullable=False)
    qtypy = Column(SMALLMONEY, nullable=False)
    wgtpy = Column(Float, nullable=False)
    tag = Column(Integer, nullable=False)
    qtyusedpy = Column(SMALLMONEY, nullable=False)
    wgtusedpy = Column(Float, nullable=False)
    remark = Column(String(255))
    aprice = Column(SMALLMONEY)
    laborcost = Column(SMALLMONEY, server_default=text("0"))
    ttype = Column(String(2))


class Mmma(HKBase):
    __tablename__ = 'mmma'

    refid = Column(Integer, primary_key=True)
    refno = Column(String(11), nullable=False)
    karat = Column(SmallInteger, nullable=False)
    refdate = Column(DateTime, nullable=False)
    tag = Column(Integer, nullable=False, index=True)


class Mmrep(HKBase):
    __tablename__ = 'mmrep'

    repid = Column(Numeric(18, 0), primary_key=True)
    docno = Column(String(10), nullable=False)
    cstbld = Column(String(10), nullable=False)
    styno = Column(String(10), nullable=False)
    cstname = Column(String(10), nullable=False)
    karat = Column(SmallInteger, nullable=False)
    qty = Column(SMALLMONEY, nullable=False)
    wgt = Column(SMALLMONEY, nullable=False)
    tag = Column(Integer, nullable=False, index=True)


class Mmsyn(HKBase):
    __tablename__ = 'mmsyn'

    mmid = Column(Integer, primary_key=True)
    userid = Column(SmallInteger, nullable=False)
    fill_date = Column(DateTime, nullable=False, server_default=text("getdate()"))
    cnt = Column(Integer, server_default=text("0"))


class P17ctbl(HKBase):
    __tablename__ = 'p17ctbl'
    __table_args__ = (
        Index('idx_p17ctbl_cid', 'catid', 'codec'),
    )

    id = Column(Numeric(5, 0), primary_key=True)
    catid = Column(SmallInteger, nullable=False)
    codec = Column(String(6), nullable=False)
    description = Column(String(250), nullable=False)


class Pajcnrev(HKBase):
    __tablename__ = 'pajcnrev'
    __table_args__ = (
        Index('idx_pajcnrev_pc', 'pcode', 'tag', unique=True),
    )

    id = Column(Numeric(18, 0), primary_key=True)
    pcode = Column(String(30), nullable=False)
    uprice = Column(SMALLMONEY, nullable=False)
    revdate = Column(DateTime, nullable=False)
    filldate = Column(DateTime, nullable=False)
    tag = Column(SmallInteger, nullable=False)


t_pbcatcol = Table(
    'pbcatcol', metadata,
    Column('pbc_tnam', CHAR(30)),
    Column('pbc_tid', Integer),
    Column('pbc_ownr', CHAR(30)),
    Column('pbc_cnam', CHAR(30)),
    Column('pbc_cid', SmallInteger),
    Column('pbc_labl', String(254)),
    Column('pbc_lpos', SmallInteger),
    Column('pbc_hdr', String(254)),
    Column('pbc_hpos', SmallInteger),
    Column('pbc_jtfy', SmallInteger),
    Column('pbc_mask', String(31)),
    Column('pbc_case', SmallInteger),
    Column('pbc_hght', SmallInteger),
    Column('pbc_wdth', SmallInteger),
    Column('pbc_ptrn', String(31)),
    Column('pbc_bmap', CHAR(1)),
    Column('pbc_init', String(254)),
    Column('pbc_cmnt', String(254)),
    Column('pbc_edit', String(31)),
    Column('pbc_tag', String(254)),
    Index('pbcatcol_idx', 'pbc_tid', 'pbc_cid', unique=True)
)


t_pbcatedt = Table(
    'pbcatedt', metadata,
    Column('pbe_name', String(30), nullable=False, unique=True),
    Column('pbe_edit', String(254)),
    Column('pbe_type', SmallInteger, nullable=False),
    Column('pbe_cntr', Integer),
    Column('pbe_seqn', SmallInteger, nullable=False),
    Column('pbe_flag', Integer),
    Column('pbe_work', CHAR(32))
)


t_pbcatfmt = Table(
    'pbcatfmt', metadata,
    Column('pbf_name', String(30), nullable=False, unique=True),
    Column('pbf_frmt', String(254), nullable=False),
    Column('pbf_type', SmallInteger, nullable=False),
    Column('pbf_cntr', Integer)
)


t_pbcattbl = Table(
    'pbcattbl', metadata,
    Column('pbt_tnam', CHAR(30)),
    Column('pbt_tid', Integer, unique=True),
    Column('pbt_ownr', CHAR(30)),
    Column('pbd_fhgt', SmallInteger),
    Column('pbd_fwgt', SmallInteger),
    Column('pbd_fitl', CHAR(1)),
    Column('pbd_funl', CHAR(1)),
    Column('pbd_fchr', SmallInteger),
    Column('pbd_fptc', SmallInteger),
    Column('pbd_ffce', CHAR(18)),
    Column('pbh_fhgt', SmallInteger),
    Column('pbh_fwgt', SmallInteger),
    Column('pbh_fitl', CHAR(1)),
    Column('pbh_funl', CHAR(1)),
    Column('pbh_fchr', SmallInteger),
    Column('pbh_fptc', SmallInteger),
    Column('pbh_ffce', CHAR(18)),
    Column('pbl_fhgt', SmallInteger),
    Column('pbl_fwgt', SmallInteger),
    Column('pbl_fitl', CHAR(1)),
    Column('pbl_funl', CHAR(1)),
    Column('pbl_fchr', SmallInteger),
    Column('pbl_fptc', SmallInteger),
    Column('pbl_ffce', CHAR(18)),
    Column('pbt_cmnt', String(254))
)


t_pbcatvld = Table(
    'pbcatvld', metadata,
    Column('pbv_name', String(30), nullable=False),
    Column('pbv_vald', String(254), nullable=False),
    Column('pbv_type', SmallInteger, nullable=False),
    Column('pbv_cntr', Integer),
    Column('pbv_msg', String(254))
)


class Poma(HKBase):
    __tablename__ = 'poma'

    pomaid = Column(Integer, primary_key=True)
    cstid = Column(SmallInteger)
    pono = Column(String(50), nullable=False, unique=True)
    ordertype = Column(SmallInteger, nullable=False)
    fill_date = Column(DateTime, nullable=False)
    order_date = Column(DateTime, nullable=False)
    receipt_date = Column(DateTime, nullable=False)
    cancel_date = Column(DateTime, nullable=False)
    tag = Column(SmallInteger, nullable=False)
    mps = Column(String(50))


t_quomemo = Table(
    'quomemo', metadata,
    Column('quoid', Integer, nullable=False, unique=True),
    Column('custid', Integer, nullable=False),
    Column('styid', Integer, nullable=False),
    Column('description', String(200), nullable=False),
    Column('stone_wt', String(100), nullable=False),
    Column('gold_karat', SmallInteger, nullable=False),
    Column('gold_wt', Float, nullable=False),
    Column('dia_quality', String(250), nullable=False),
    Column('karat_10', MONEY, nullable=False),
    Column('karat_14', MONEY, nullable=False),
    Column('karat_other', String(100), nullable=False),
    Column('gold_lock', MONEY, nullable=False),
    Column('remark', String(250), nullable=False),
    Column('createdate', DateTime),
    Column('userid', SmallInteger),
    Column('tag', SmallInteger, nullable=False),
    Column('docno', String(20), server_default=text("\"N/A\""))
)


class Quoq(HKBase):
    __tablename__ = 'quoq'

    id = Column(Integer, primary_key=True)
    name = Column(String(30), nullable=False, unique=True)
    mps = Column(String(50), nullable=False)
    quodate = Column(DateTime, nullable=False)
    cstid = Column(SmallInteger, nullable=False)
    fill_date = Column(DateTime, nullable=False)
    lastuserid = Column(SmallInteger, nullable=False)
    tag = Column(SmallInteger, nullable=False)


class Quoqitem(HKBase):
    __tablename__ = 'quoqitem'

    id = Column(Integer, primary_key=True)
    quoqid = Column(Integer, nullable=False)
    karat = Column(String(50), nullable=False)
    styid = Column(Integer, nullable=False)
    specs = Column(String(50), nullable=False)
    price = Column(SMALLMONEY, nullable=False)
    tag = Column(SmallInteger, nullable=False)


class Rpma(HKBase):
    __tablename__ = 'rpma'

    rpmaid = Column(Integer, primary_key=True)
    styid = Column(Integer, nullable=False)
    running = Column(Integer, nullable=False)
    qty = Column(SmallInteger, nullable=False)
    karat = Column(SmallInteger, nullable=False)
    rpno = Column(String(15), nullable=False)
    reljono = Column(String(20), nullable=False)
    description = Column(String(30), nullable=False)
    createdate = Column(DateTime, nullable=False)
    userid = Column(SmallInteger, nullable=False)
    tag = Column(SmallInteger, nullable=False)


class SecurityApp(HKBase):
    __tablename__ = 'security_apps'

    appid = Column(Integer, primary_key=True)
    application = Column(String(32), nullable=False, index=True)
    description = Column(String(64), nullable=False)


class SecurityUser(HKBase):
    __tablename__ = 'security_users'

    userid = Column(Integer, primary_key=True)
    name = Column(String(16), nullable=False, unique=True)
    description = Column(String(32), nullable=False)
    priority = Column(Integer, nullable=False)
    password = Column(String(20), nullable=False)
    status = Column(TINYINT, nullable=False)
    user_type = Column(Integer, nullable=False)


t_sopojo = Table(
    'sopojo', metadata,
    Column('id', Integer, nullable=False),
    Column('ttype', TINYINT, nullable=False),
    Column('joid', Integer, nullable=False),
    Column('sopoid', Integer, nullable=False),
    Column('qty', SMALLMONEY, nullable=False),
    Column('fill_date', DateTime, nullable=False),
    Column('tag', TINYINT, nullable=False),
    Index('idx_sopojo_for_jo_update_trig', 'ttype', 'joid', 'sopoid', 'tag')
)


class Sosum(HKBase):
    __tablename__ = 'sosum'
    __table_args__ = (
        Index('idx_sosum', 'cstname', 'yrmth', 'type'),
    )

    id = Column(Integer, primary_key=True)
    cstname = Column(String(5), nullable=False)
    yrmth = Column(String(8), nullable=False)
    type = Column(String(1), nullable=False)
    totqty = Column(SMALLMONEY)
    totamt = Column(Float)


class Stockobjectma(HKBase):
    __tablename__ = 'stockobjectma'

    srid = Column(Integer, primary_key=True, index=True)
    styno = Column(String(30), index=True)
    running = Column(String(30), index=True)
    stockcode = Column(String(60), unique=True)
    description = Column(String(250))
    tag = Column(SmallInteger)
    qtyleft = Column(Float)
    type = Column(SmallInteger, index=True)


class StoneMaster(HKBase):
    __tablename__ = 'stone_master'

    stid = Column(SmallInteger, primary_key=True)
    stname = Column(String(4), nullable=False, unique=True)
    edesc = Column(String(50), nullable=False)
    cdesc = Column(String(50), nullable=False)
    sttype = Column(String(20), nullable=False)
    settype = Column(TINYINT, nullable=False)
    fill_date = Column(DateTime, nullable=False)
    tag = Column(TINYINT, nullable=False)


t_stone_price = Table(
    'stone_price', metadata,
    Column('stpid', Integer, nullable=False),
    Column('package', String(30), nullable=False),
    Column('size', String(30), nullable=False),
    Column('price', SMALLMONEY, nullable=False),
    Column('unit', SmallInteger, nullable=False),
    Column('currid', SmallInteger, nullable=False),
    Column('tag', SmallInteger, nullable=False),
    Index('idx_stone_price_package_size', 'package', 'size', unique=True)
)


class StoneSizema(HKBase):
    __tablename__ = 'stone_sizema'

    stsizeid = Column(SmallInteger, primary_key=True)
    stsize = Column(String(20), nullable=False, index=True)
    wgt = Column(Float, nullable=False)
    grp = Column(String(10), nullable=False, index=True)
    pingrp = Column(SmallInteger, nullable=False)
    tag = Column(TINYINT, nullable=False)


class StoneSizemaster(HKBase):
    __tablename__ = 'stone_sizemaster'

    stsizeid = Column(SmallInteger, primary_key=True)
    stsize = Column(String(20), nullable=False, unique=True)
    wgt = Column(Float, nullable=False)
    grp = Column(CHAR(2))
    pingrp = Column(SmallInteger)
    tag = Column(TINYINT, nullable=False)


class Styma(HKBase):
    __tablename__ = 'styma'
    __table_args__ = (
        Index('idx_styma_alphadigit', 'alpha', 'digit'),
    )

    styid = Column(Integer, primary_key=True)
    alpha = Column(String(5), nullable=False)
    digit = Column(Integer, nullable=False)
    description = Column(String(50), nullable=False)
    edescription = Column(String(50), nullable=False)
    ordercnt = Column(SmallInteger, nullable=False)
    fill_date = Column(DateTime, nullable=False)
    tag = Column(SmallInteger, nullable=False)
    suffix = Column(String(50), server_default=text("\"\""))
    name1 = Column(String(20), index=True)


t_stymaster = Table(
    'stymaster', metadata,
    Column('id', Integer, nullable=False, unique=True),
    Column('styn', Unicode(10)),
    Column('runn', Unicode(10)),
    Column('goldwgt', DECIMAL(8, 3)),
    Column('jew_type', Unicode(20)),
    Column('mst_qty', Integer),
    Column('mst_type', Unicode(15)),
    Column('mst_size', Unicode(15)),
    Column('mst_wgt', DECIMAL(8, 3)),
    Column('fst_qty', Integer),
    Column('fst_type', Unicode(15)),
    Column('fst_size', Unicode(15)),
    Column('fst_wgt', DECIMAL(8, 3)),
    Column('est_qty', Integer),
    Column('est_type', Unicode(15)),
    Column('est_size', Unicode(15)),
    Column('est_wgt', DECIMAL(8, 3)),
    Column('tst_qty', Integer),
    Column('tst_type', Unicode(15)),
    Column('tst_size', Unicode(15)),
    Column('tst_wgt', DECIMAL(8, 3)),
    Column('tag', SmallInteger),
    Column('joid', Integer),
    Column('styid', Integer),
    Column('typeid', Integer)
)


class Stypropdef(HKBase):
    __tablename__ = 'stypropdef'

    propdefid = Column(SmallInteger, primary_key=True)
    propname = Column(String(50), nullable=False, unique=True)
    propgrp = Column(String(50), nullable=False)
    proptype = Column(SmallInteger, nullable=False, index=True)


t_sysquerymetrics = Table(
    'sysquerymetrics', metadata,
    Column('uid', Integer, nullable=False),
    Column('gid', Integer, nullable=False),
    Column('hashkey', Integer, nullable=False),
    Column('id', Integer, nullable=False),
    Column('sequence', SmallInteger, nullable=False),
    Column('exec_min', Integer),
    Column('exec_max', Integer),
    Column('exec_avg', Integer),
    Column('elap_min', Integer),
    Column('elap_max', Integer),
    Column('elap_avg', Integer),
    Column('lio_min', Integer),
    Column('lio_max', Integer),
    Column('lio_avg', Integer),
    Column('pio_min', Integer),
    Column('pio_max', Integer),
    Column('pio_avg', Integer),
    Column('cnt', Integer),
    Column('abort_cnt', Integer),
    Column('qtext', String(255))
)


t_t1 = Table(
    't1', metadata,
    Column('joid', Integer, nullable=False),
    Column('orderid', Integer, nullable=False),
    Column('poid', Integer, nullable=False),
    Column('soid', Integer, nullable=False),
    Column('alpha', CHAR(2), nullable=False),
    Column('digit', Integer, nullable=False),
    Column('description', String(255), nullable=False),
    Column('running', Integer, nullable=False),
    Column('qty', SMALLMONEY, nullable=False),
    Column('qtyleft', SMALLMONEY, nullable=False),
    Column('wgt', SMALLMONEY, nullable=False),
    Column('uprice', SMALLMONEY, nullable=False),
    Column('snno', String(250), nullable=False),
    Column('ponohk', String(10), nullable=False),
    Column('remark', String(160), nullable=False),
    Column('fill_date', DateTime, nullable=False),
    Column('deadline', DateTime, nullable=False),
    Column('shipdate', DateTime, nullable=False),
    Column('opid', SmallInteger, nullable=False),
    Column('tag', SmallInteger, nullable=False),
    Column('soqty', SMALLMONEY),
    Column('poqty', SMALLMONEY),
    Column('edescription', String(255)),
    Column('createdate', DateTime),
    Column('stid', SmallInteger),
    Column('auxkarat', SmallInteger),
    Column('ordertype', TINYINT),
    Column('auxwgt', SMALLMONEY)
)


class Tempjo(HKBase):
    __tablename__ = 'tempjo'

    jo = Column(String(10), primary_key=True)
    alpha = Column(String(1), nullable=False)
    digit = Column(Integer, nullable=False)


class Tmp(HKBase):
    __tablename__ = 'tmp'

    id = Column(Integer, primary_key=True)
    alpha = Column(String(10), nullable=False)
    digit = Column(Integer, nullable=False)
    remark = Column(String(250), nullable=False)


t_uv_jo = Table(
    'uv_jo', metadata,
    Column('joid', Integer, nullable=False),
    Column('orderid', Integer, nullable=False),
    Column('soid', Integer, nullable=False),
    Column('poid', Integer, nullable=False),
    Column('alpha', CHAR(2), nullable=False),
    Column('digit', Integer, nullable=False),
    Column('cstbldid', String(12), nullable=False),
    Column('description', String(255), nullable=False),
    Column('running', Integer, nullable=False),
    Column('qty', SMALLMONEY, nullable=False),
    Column('qtyleft', SMALLMONEY, nullable=False),
    Column('wgt', SMALLMONEY, nullable=False),
    Column('uprice', SMALLMONEY, nullable=False),
    Column('snno', String(250), nullable=False),
    Column('ponohk', String(10), nullable=False),
    Column('remark', String(250), nullable=False),
    Column('fill_date', DateTime, nullable=False),
    Column('deadline', DateTime, nullable=False),
    Column('shipdate', DateTime, nullable=False),
    Column('opid', SmallInteger, nullable=False),
    Column('tag', SmallInteger, nullable=False),
    Column('cstid', SmallInteger, nullable=False),
    Column('styid', Integer, nullable=False),
    Column('karat', SmallInteger, nullable=False),
    Column('cstname', CHAR(30), nullable=False),
    Column('styno', String(65)),
    Column('soqty', SMALLMONEY),
    Column('poqty', SMALLMONEY),
    Column('edescription', String(255)),
    Column('createdate', DateTime),
    Column('stid', SmallInteger),
    Column('stname', String(4), nullable=False),
    Column('auxkarat', SmallInteger),
    Column('auxwgt', SMALLMONEY)
)


t_uv_p17dc = Table(
    'uv_p17dc', metadata,
    Column('category', String(255), nullable=False),
    Column('digits', String(255), nullable=False),
    Column('id', Numeric(5, 0), nullable=False),
    Column('catid', SmallInteger, nullable=False),
    Column('codec', String(6), nullable=False),
    Column('description', String(250), nullable=False)
)


t_uv_paj = Table(
    'uv_paj', metadata,
    Column('cstname', CHAR(30), nullable=False),
    Column('jono', String(20), nullable=False),
    Column('deadline', DateTime, nullable=False),
    Column('styno', String(15), nullable=False),
    Column('joalpha', CHAR(2), nullable=False),
    Column('jodigit', Integer, nullable=False),
    Column('styalpha', String(5), nullable=False),
    Column('stydigit', Integer, nullable=False),
    Column('running', Integer, nullable=False),
    Column('joqty', SMALLMONEY, nullable=False),
    Column('karat', SmallInteger, nullable=False),
    Column('wgt', SMALLMONEY, nullable=False),
    Column('auxkarat', SmallInteger),
    Column('auxwgt', SMALLMONEY),
    Column('shpfn', String(100), nullable=False),
    Column('pcode', String(30), nullable=False),
    Column('invno', String(10), nullable=False),
    Column('shpqty', Float, nullable=False),
    Column('pajorderno', String(20), nullable=False),
    Column('shpdate', DateTime, nullable=False),
    Column('invqty', Float, nullable=False),
    Column('pajmps', String(50), nullable=False),
    Column('pajuprice', SMALLMONEY, nullable=False),
    Column('china', SMALLMONEY, nullable=False),
    Column('stspec', String(200), nullable=False),
    Column('pomps', String(50)),
    Column('pouprice', SMALLMONEY, nullable=False),
    Column('skuno', String(50), nullable=False),
    Column('poid', Integer, nullable=False),
    Column('joid', Integer, nullable=False)
)


t_uv_quoq = Table(
    'uv_quoq', metadata,
    Column('mid', Integer, nullable=False),
    Column('qid', Integer, nullable=False),
    Column('name', String(30), nullable=False),
    Column('quodate', DateTime, nullable=False),
    Column('cstname', CHAR(30), nullable=False),
    Column('karat', String(50), nullable=False),
    Column('mtag', SmallInteger, nullable=False),
    Column('alpha', String(5), nullable=False),
    Column('digit', Integer, nullable=False),
    Column('specs', String(50), nullable=False),
    Column('price', SMALLMONEY, nullable=False)
)


t_ven_quomemo = Table(
    'ven_quomemo', metadata,
    Column('stpid', Integer, nullable=False, unique=True),
    Column('vendor', String(50), nullable=False),
    Column('stone', String(100), nullable=False),
    Column('color', String(100), nullable=False),
    Column('shape', String(100), nullable=False),
    Column('size', String(100), nullable=False),
    Column('grade', String(100), nullable=False),
    Column('unit_wt', Numeric(10, 2), nullable=False),
    Column('price_pc', Numeric(10, 2), nullable=False),
    Column('price_ct', Numeric(10, 2), nullable=False),
    Column('originid', Integer, nullable=False),
    Column('sourcingid', Integer, nullable=False),
    Column('createdate', DateTime, nullable=False),
    Column('userid', SmallInteger, nullable=False),
    Column('tag', SmallInteger, nullable=False),
    Column('remark', String(255), server_default=text("\"\" null"))
)


class Cstbldsizema(HKBase):
    __tablename__ = 'cstbldsizema'

    sizeid = Column(SmallInteger, primary_key=True, unique=True)
    cstid = Column(ForeignKey('cstinfo.cstid'), nullable=False)
    rsize = Column(String(15), nullable=False)
    tag = Column(TINYINT, nullable=False)

    cstinfo = relationship('Cstinfo')


class Cstbldtpy(HKBase):
    __tablename__ = 'cstbldtpy'

    docid = Column(ForeignKey('cstbldtpym.docid'), nullable=False, index=True)
    alpha = Column(CHAR(1), primary_key=True, nullable=False)
    digitf = Column(Integer, primary_key=True, nullable=False)
    digitt = Column(Integer, primary_key=True, nullable=False)
    qty = Column(MONEY, nullable=False)

    cstbldtpym = relationship('Cstbldtpym')


class Fdmd(HKBase):
    __tablename__ = 'fdmd'

    id = Column(Integer, primary_key=True)
    sid = Column(ForeignKey('fdmdm.sid'), nullable=False)
    jsid = Column(Integer, nullable=False, index=True)
    qty = Column(SMALLMONEY, nullable=False)

    fdmdm = relationship('Fdmdm')


class Invoicema(HKBase):
    __tablename__ = 'invoicema'
    __table_args__ = (
        Index('idx_invoicema_inoutno_docno', 'inoutno', 'docno', unique=True),
    )

    invid = Column(Integer, primary_key=True)
    inoutno = Column(String(50), index=True)
    docno = Column(String(50), index=True)
    docdate = Column(DateTime, index=True)
    locationidfrm = Column(ForeignKey('locationma.locationid'), index=True)
    locationidto = Column(ForeignKey('locationma.locationid'), index=True)
    remark1 = Column(String(100))
    remark2 = Column(String(100))
    lastuserid = Column(SmallInteger)
    lastupdate = Column(DateTime)
    tag = Column(SmallInteger)

    locationma = relationship('Locationma', primaryjoin='Invoicema.locationidfrm == Locationma.locationid')
    locationma1 = relationship('Locationma', primaryjoin='Invoicema.locationidto == Locationma.locationid')


class Mitpy(HKBase):
    __tablename__ = 'mitpy'

    mitid = Column(ForeignKey('mitin.mitid'), primary_key=True, nullable=False)
    idx = Column(TINYINT, primary_key=True, nullable=False)
    qty = Column(SMALLMONEY, nullable=False)
    wgt = Column(Float, nullable=False)
    fill_date = Column(DateTime, nullable=False)

    mitin = relationship('Mitin')


class Mmstocktake(HKBase):
    __tablename__ = 'mmstocktake'
    __table_args__ = (
        Index('idx_mmstocktake_perno_docno', 'perno', 'docno'),
    )

    id = Column(Integer, primary_key=True, server_default=text("0"))
    perno = Column(String(50), nullable=False, index=True, server_default=text(""))
    docno = Column(String(20), nullable=False, server_default=text(""))
    locationid = Column(Integer, nullable=False, server_default=text("0"))
    srid = Column(ForeignKey('stockobjectma.srid'), ForeignKey('stockobjectma.srid'), nullable=False, index=True, server_default=text("0"))
    qty = Column(Float, nullable=False, server_default=text("0"))
    fill_date = Column(DateTime, nullable=False, server_default=text("getdate()"))
    lastuserid = Column(Integer, nullable=False, server_default=text("0"))
    lastupdate = Column(DateTime, nullable=False, server_default=text("getdate()"))
    customer = Column(String(30), nullable=False, server_default=text("N/A"))

    stockobjectma = relationship('Stockobjectma', primaryjoin='Mmstocktake.srid == Stockobjectma.srid')
    stockobjectma1 = relationship('Stockobjectma', primaryjoin='Mmstocktake.srid == Stockobjectma.srid')


class Orderma(HKBase):
    __tablename__ = 'orderma'
    __table_args__ = (
        Index('idx_orderma_cstid', 'cstid', 'styid', 'karat', unique=True),
    )

    orderid = Column(Integer, primary_key=True)
    orderno = Column(String(50), nullable=False)
    cstid = Column(ForeignKey('cstinfo.cstid'), nullable=False)
    styid = Column(ForeignKey('styma.styid'), nullable=False, index=True)
    karat = Column(SmallInteger, nullable=False)
    soqty = Column(MONEY, nullable=False)
    soprice = Column(MONEY, nullable=False)
    poqty = Column(MONEY, nullable=False)
    poprice = Column(MONEY, nullable=False)
    joqty = Column(MONEY, nullable=False)
    jofqty = Column(MONEY, nullable=False)
    joprice = Column(MONEY, nullable=False)
    fill_date = Column(DateTime, nullable=False)
    modi_date = Column(DateTime, nullable=False)
    tag = Column(SmallInteger, nullable=False)

    cstinfo = relationship('Cstinfo')
    styma = relationship('Styma')


class Rpapart(HKBase):
    __tablename__ = 'rpapart'

    id = Column(Integer, primary_key=True)
    jono = Column(String(50), nullable=False)
    styid = Column(ForeignKey('stockobjectma.srid'), nullable=False)
    karat = Column(Integer, nullable=False, server_default=text("0"))
    qty = Column(Float, nullable=False, server_default=text("0"))
    wgtout = Column(Float, nullable=False, server_default=text("0"))
    wgtback = Column(Float, nullable=False, server_default=text("0"))
    wgtmetal = Column(Float, nullable=False, server_default=text("0"))
    empid = Column(ForeignKey('employee.id'), nullable=False, index=True)
    pack = Column(String(20), nullable=False, server_default=text(""))
    remark = Column(String(250), nullable=False, server_default=text(""))
    fill_date = Column(DateTime, nullable=False)
    lastupdate = Column(DateTime, nullable=False)
    tag = Column(Integer, nullable=False, index=True, server_default=text("0"))
    date_back = Column(DateTime, server_default=text("getdate()"))
    brokenstone = Column(Float, nullable=False, server_default=text("0"))
    wgtst = Column(Float, nullable=False, server_default=text("0"))

    employee = relationship('Employee')
    stockobjectma = relationship('Stockobjectma')


class SecurityGrouping(HKBase):
    __tablename__ = 'security_groupings'

    grpid = Column(Integer, primary_key=True, nullable=False)
    userid = Column(ForeignKey('security_users.userid'), primary_key=True, nullable=False)

    security_user = relationship('SecurityUser')


class SecurityObject(HKBase):
    __tablename__ = 'security_objects'
    __table_args__ = (
        Index('idx_security_sysobjects', 'appid', 'control', 'window', unique=True),
    )

    objid = Column(Integer, primary_key=True)
    appid = Column(ForeignKey('security_apps.appid'), ForeignKey('security_apps.appid'), nullable=False)
    window = Column(String(64), nullable=False)
    control = Column(String(128), nullable=False)
    description = Column(String(254), nullable=False)
    objtype = Column(SmallInteger, nullable=False)
    tag = Column(SmallInteger, nullable=False)

    security_app = relationship('SecurityApp', primaryjoin='SecurityObject.appid == SecurityApp.appid')
    security_app1 = relationship('SecurityApp', primaryjoin='SecurityObject.appid == SecurityApp.appid')


class Soma(HKBase):
    __tablename__ = 'soma'

    somaid = Column(Integer, primary_key=True)
    cstid = Column(ForeignKey('cstinfo.cstid'))
    sono = Column(String(15), nullable=False, unique=True)
    fill_date = Column(DateTime, nullable=False)
    order_date = Column(DateTime, nullable=False)
    deliver_date = Column(DateTime, nullable=False)
    tag = Column(SmallInteger, nullable=False)

    cstinfo = relationship('Cstinfo')


t_stockobjectdaily = Table(
    'stockobjectdaily', metadata,
    Column('sodid', Integer, unique=True),
    Column('srid', ForeignKey('stockobjectma.srid')),
    Column('locationid', ForeignKey('locationma.locationid')),
    Column('date', Integer),
    Column('qty', Float),
    Column('tag', SmallInteger),
    Index('idx_stockdailysridlocationdate', 'srid', 'locationid', 'date')
)


class Stypropma(HKBase):
    __tablename__ = 'stypropma'
    __table_args__ = (
        Index('idx_stypropma_prop', 'propdefid', 'coden', 'codec', unique=True),
    )

    propid = Column(Integer, primary_key=True)
    propdefid = Column(ForeignKey('stypropdef.propdefid'), nullable=False, index=True)
    codec = Column(String(20), nullable=False)
    coden = Column(Numeric(7, 3), nullable=False)
    fill_date = Column(DateTime, nullable=False)

    stypropdef = relationship('Stypropdef')


class Fama(HKBase):
    __tablename__ = 'fama'

    faid = Column(Integer, primary_key=True)
    orderid = Column(ForeignKey('orderma.orderid'), nullable=False)
    fano = Column(String(10), nullable=False, index=True)
    running = Column(String(10), nullable=False, unique=True)
    stone = Column(String(2), nullable=False)
    description = Column(String(30), nullable=False)
    qty = Column(MONEY, nullable=False)
    fill_date = Column(DateTime, nullable=False)
    tag = Column(SmallInteger, nullable=False)

    orderma = relationship('Orderma')


class Invoicedtl(HKBase):
    __tablename__ = 'invoicedtl'

    invdid = Column(Integer, primary_key=True, unique=True)
    invid = Column(ForeignKey('invoicema.invid'), index=True)
    jono = Column(String(20))
    srid = Column(ForeignKey('stockobjectma.srid'), index=True)
    qty = Column(Float)
    lastuserid = Column(SmallInteger)
    lastupdate = Column(DateTime)
    tag = Column(SmallInteger)

    invoicema = relationship('Invoicema')
    stockobjectma = relationship('Stockobjectma')


class Po(HKBase):
    __tablename__ = 'po'
    __table_args__ = (
        Index('idx_po_poid_orderid', 'poid', 'orderid'),
    )

    poid = Column(Integer, primary_key=True)
    pomaid = Column(ForeignKey('poma.pomaid'), index=True)
    orderid = Column(ForeignKey('orderma.orderid'), index=True)
    qty = Column(SMALLMONEY, nullable=False)
    uprice = Column(SMALLMONEY, nullable=False)
    skuno = Column(String(50), nullable=False, index=True)
    rmk = Column(String(20), nullable=False)
    description = Column(String(50), nullable=False)
    tag = Column(SmallInteger, nullable=False, index=True)
    joqty = Column(SMALLMONEY, server_default=text("0"))

    orderma = relationship('Orderma')
    poma = relationship('Poma')


class SecurityInfo(HKBase):
    __tablename__ = 'security_info'

    objid = Column(ForeignKey('security_objects.objid'), primary_key=True, nullable=False)
    userid = Column(ForeignKey('security_users.userid'), primary_key=True, nullable=False)
    status = Column(SmallInteger, nullable=False)
    tag = Column(SmallInteger, nullable=False)

    security_object = relationship('SecurityObject')
    security_user = relationship('SecurityUser')


class So(HKBase):
    __tablename__ = 'so'
    __table_args__ = (
        Index('idx_so_soid_orderid', 'soid', 'orderid'),
    )

    soid = Column(Integer, primary_key=True)
    somaid = Column(ForeignKey('soma.somaid'), index=True)
    orderid = Column(ForeignKey('orderma.orderid'), index=True)
    skuno = Column(String(20), nullable=False)
    description = Column(String(50), nullable=False)
    qty = Column(SMALLMONEY, nullable=False)
    uprice = Column(SMALLMONEY, nullable=False)
    tag = Column(SmallInteger, nullable=False, index=True)
    joqty = Column(SMALLMONEY, server_default=text("0"))

    orderma = relationship('Orderma')
    soma = relationship('Soma')


class Jo(HKBase):
    __tablename__ = 'jo'
    __table_args__ = (
        Index('idx_jo_jono', 'alpha', 'digit', unique=True),
    )

    joid = Column(Integer, primary_key=True)
    orderid = Column(ForeignKey('orderma.orderid'), nullable=False, index=True)
    poid = Column(ForeignKey('po.poid'), nullable=False, index=True)
    soid = Column(ForeignKey('so.soid'), nullable=False, index=True)
    alpha = Column(CHAR(2), nullable=False)
    digit = Column(Integer, nullable=False)
    description = Column(String(255), nullable=False)
    running = Column(Integer, nullable=False, index=True)
    qty = Column(SMALLMONEY, nullable=False)
    qtyleft = Column(SMALLMONEY, nullable=False, index=True)
    wgt = Column(SMALLMONEY, nullable=False)
    uprice = Column(SMALLMONEY, nullable=False)
    snno = Column(String(250), nullable=False, index=True)
    ponohk = Column(String(10), nullable=False, index=True)
    remark = Column(String(250), nullable=False, index=True)
    fill_date = Column(DateTime, nullable=False)
    deadline = Column(DateTime, nullable=False)
    shipdate = Column(DateTime, nullable=False)
    opid = Column(SmallInteger, nullable=False)
    tag = Column(SmallInteger, nullable=False, index=True)
    soqty = Column(SMALLMONEY, server_default=text("0"))
    poqty = Column(SMALLMONEY, server_default=text("0"))
    edescription = Column(String(255), server_default=text("\" \""))
    createdate = Column(DateTime, index=True, server_default=text("getdate()"))
    stid = Column(ForeignKey('stone_master.stid'), server_default=text("15"))
    auxkarat = Column(SmallInteger, server_default=text("0"))
    ordertype = Column(TINYINT, server_default=text("0"))
    auxwgt = Column(SMALLMONEY, server_default=text("0"))

    orderma = relationship('Orderma')
    po = relationship('Po')
    so = relationship('So')
    stone_master = relationship('StoneMaster')


class Cstbldremark(Jo):
    __tablename__ = 'cstbldremark'

    jsid = Column(ForeignKey('jo.joid'), primary_key=True)
    remark = Column(String(100), nullable=False)
    edes = Column(String(30))
    remark1 = Column(String(30))


class Cstbldloca(HKBase):
    __tablename__ = 'cstbldloca'

    jsid = Column(ForeignKey('jo.joid'), primary_key=True, nullable=False)
    cstid = Column(SmallInteger, primary_key=True, nullable=False)
    qty = Column(SMALLMONEY, nullable=False)
    qtyleft = Column(SMALLMONEY, nullable=False)

    jo = relationship('Jo')


class Cstbllsize(HKBase):
    __tablename__ = 'cstbllsize'

    jsid = Column(ForeignKey('jo.joid'), primary_key=True, nullable=False)
    size_ = Column(SmallInteger, primary_key=True, nullable=False)
    quantity = Column(SMALLMONEY, nullable=False)

    jo = relationship('Jo')


class Cstdtl(HKBase):
    __tablename__ = 'cstdtl'

    jsid = Column(ForeignKey('jo.joid'), primary_key=True, nullable=False)
    sttype = Column(String(4), primary_key=True, nullable=False)
    stsize = Column(String(10), primary_key=True, nullable=False)
    quantity = Column(SmallInteger, nullable=False)
    wgt = Column(Float, nullable=False)
    wgt_calc = Column(Float, nullable=False)
    remark = Column(String(30))

    jo = relationship('Jo')


t_jogdstrange = Table(
    'jogdstrange', metadata,
    Column('joid', ForeignKey('jo.joid'), nullable=False, index=True),
    Column('stid', ForeignKey('stone_master.stid'), nullable=False),
    Column('minwgt', DECIMAL(5, 3), nullable=False),
    Column('maxwgt', DECIMAL(5, 3), nullable=False)
)


class Mitout(HKBase):
    __tablename__ = 'mitout'

    psaid = Column(Integer, primary_key=True)
    mitid = Column(Integer, nullable=False)
    is_out = Column(TINYINT, nullable=False)
    jsid = Column(ForeignKey('jo.joid'), nullable=False, index=True)
    qty = Column(SMALLMONEY, nullable=False)
    wgt = Column(Float, nullable=False)
    fill_date = Column(DateTime, nullable=False)
    wkid = Column(SmallInteger, nullable=False)
    tag = Column(TINYINT, server_default=text("0"))
    remark = Column(String(255), server_default=text(" "))

    jo = relationship('Jo')


class Mm(HKBase):
    __tablename__ = 'mm'

    mmid = Column(Integer, primary_key=True)
    refid = Column(ForeignKey('mmma.refid'), nullable=False, index=True)
    docno = Column(String(8), nullable=False)
    jsid = Column(ForeignKey('jo.joid'), nullable=False, index=True)
    qty = Column(SMALLMONEY, nullable=False)
    tag = Column(TINYINT, nullable=False, index=True)

    jo = relationship('Jo')
    mmma = relationship('Mmma')


class Mmmit(Mm):
    __tablename__ = 'mmmit'

    mmid = Column(ForeignKey('mm.mmid'), primary_key=True)
    wgt = Column(SMALLMONEY, nullable=False)


class Pajinv(HKBase):
    __tablename__ = 'pajinv'
    __table_args__ = (
        Index('bk_pajinv', 'invno', 'joid', unique=True),
    )

    id = Column(Numeric(18, 0), primary_key=True)
    invno = Column(String(10), nullable=False)
    joid = Column(ForeignKey('jo.joid'), nullable=False, index=True)
    qty = Column(Float, nullable=False)
    uprice = Column(SMALLMONEY, nullable=False)
    mps = Column(String(50), nullable=False)
    stspec = Column(String(200), nullable=False)
    lastmodified = Column(DateTime, nullable=False)
    china = Column(SMALLMONEY, nullable=False, server_default=text("0"))

    jo = relationship('Jo')


class Pajshp(HKBase):
    __tablename__ = 'pajshp'
    __table_args__ = (
        Index('idx_bk_pajshp', 'joid', 'pcode', 'fn', 'invno', unique=True),
    )

    id = Column(Numeric(18, 0), primary_key=True)
    fn = Column(String(100), nullable=False)
    joid = Column(ForeignKey('jo.joid'), nullable=False, index=True)
    pcode = Column(String(30), nullable=False, index=True)
    invno = Column(String(10), nullable=False)
    qty = Column(Float, nullable=False)
    orderno = Column(String(20), nullable=False)
    mtlwgt = Column(Float, nullable=False)
    stwgt = Column(Float, nullable=False)
    invdate = Column(DateTime, nullable=False)
    shpdate = Column(DateTime, nullable=False)
    filldate = Column(DateTime, nullable=False)
    lastmodified = Column(DateTime, nullable=False)

    jo = relationship('Jo')


class Stymaext(HKBase):
    __tablename__ = 'stymaext'

    styextid = Column(Integer, primary_key=True)
    styid = Column(ForeignKey('styma.styid'), index=True)
    joid = Column(ForeignKey('jo.joid'))
    styno = Column(String(10), nullable=False, index=True)
    properties = Column(String(255), nullable=False, index=True)
    remark = Column(String(255), nullable=False)

    jo = relationship('Jo')
    styma = relationship('Styma')


class Mmgd(HKBase):
    __tablename__ = 'mmgd'

    mmid = Column(ForeignKey('mm.mmid'), primary_key=True, nullable=False)
    karat = Column(SmallInteger, primary_key=True, nullable=False)
    wgt = Column(SMALLMONEY, nullable=False)

    mm = relationship('Mm')


class Mmst(HKBase):
    __tablename__ = 'mmst'

    mmid = Column(ForeignKey('mm.mmid'), primary_key=True, nullable=False)
    sttype = Column(TINYINT, primary_key=True, nullable=False)
    qty = Column(Integer, nullable=False)
    wgt = Column(DECIMAL(8, 3), nullable=False)

    mm = relationship('Mm')


class Quohistory(HKBase):
    __tablename__ = 'quohistory'

    id = Column(Integer, primary_key=True)
    styextid = Column(ForeignKey('stymaext.styextid'), nullable=False)
    cstid = Column(ForeignKey('cstinfo.cstid'), nullable=False)
    karat = Column(SmallInteger, nullable=False)
    gdprice = Column(TINYINT, nullable=False)
    price = Column(SMALLMONEY, nullable=False)
    gdwgt = Column(SMALLMONEY, nullable=False)
    price400 = Column(SMALLMONEY, nullable=False)
    diaspec = Column(TINYINT, nullable=False)
    stonespec = Column(String(255), nullable=False)
    remark = Column(String(255), nullable=False)
    tag = Column(TINYINT, nullable=False)

    cstinfo = relationship('Cstinfo')
    stymaext = relationship('Stymaext')


class Styprop(HKBase):
    __tablename__ = 'styprop'
    __table_args__ = (
        Index('idx_styprop', 'styextid', 'propid'),
    )

    id = Column(Integer, primary_key=True)
    styextid = Column(ForeignKey('stymaext.styextid'), nullable=False)
    propid = Column(ForeignKey('stypropma.propid'), nullable=False)

    stypropma = relationship('Stypropma')
    stymaext = relationship('Stymaext')
