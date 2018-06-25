# coding: utf-8
from sqlalchemy import CHAR, Column, DECIMAL, DateTime, Float, ForeignKey, Index, Integer, Numeric, SmallInteger, String, Table, text
from sqlalchemy.dialects.sybase.base import MONEY, SMALLMONEY, TINYINT
from sqlalchemy.orm import relationship
from sqlalchemy.ext.declarative import declarative_base

CNBase = declarative_base()
metadata = Base.metadata


class Bankac(CNBase):
    __tablename__ = 'bankac'

    tid = Column(Numeric(18, 0), primary_key=True)
    tmid = Column(Integer, nullable=False)
    docno = Column(String(10), nullable=False)
    wgt = Column(SMALLMONEY, nullable=False)
    nwgt = Column(SMALLMONEY, nullable=False)
    fill_date = Column(DateTime, nullable=False)
    tag = Column(TINYINT, nullable=False)


class Bankacm(CNBase):
    __tablename__ = 'bankacm'
    __table_args__ = (
        Index('idx_bankacm_tran', 'tobj', 'tpp', unique=True),
    )

    tmid = Column(Integer, primary_key=True)
    tobj = Column(TINYINT, nullable=False)
    tpp = Column(TINYINT, nullable=False)
    wgtin = Column(MONEY, nullable=False)
    wgtout = Column(MONEY, nullable=False)
    fill_date = Column(DateTime, nullable=False)
    modi_date = Column(DateTime, nullable=False)
    tag = Column(Integer, nullable=False)


class Bankcon(CNBase):
    __tablename__ = 'bankcon'

    conid = Column(SmallInteger, primary_key=True)
    conmaid = Column(Integer, nullable=False)
    amount = Column(MONEY, nullable=False)
    wgtin = Column(MONEY, nullable=False)
    wgtout = Column(MONEY, nullable=False)
    wgtloss = Column(MONEY, nullable=False)
    fill_date = Column(DateTime, nullable=False)
    modi_date = Column(DateTime, nullable=False)
    tag = Column(SmallInteger, nullable=False)
    wgtback = Column(MONEY, server_default=text("0"))
    description = Column(String(255), server_default=text("\"\""))
    alias = Column(String(50), server_default=text("\"\""))
    unit = Column(TINYINT, server_default=text("0"))
    opening = Column(MONEY, server_default=text("0"))
    uprice = Column(MONEY, nullable=False, server_default=text("0"))
    prodcode = Column(String(250), server_default=text("\"\""))


class Bankconback(CNBase):
    __tablename__ = 'bankconback'

    conbackid = Column(Integer, primary_key=True)
    conid = Column(SmallInteger, nullable=False, index=True)
    is_out = Column(TINYINT, nullable=False)
    wgt = Column(MONEY, nullable=False)
    wgtpure = Column(MONEY, nullable=False)
    fill_date = Column(DateTime, nullable=False)
    createdate = Column(DateTime, nullable=False)


class Bankconma(CNBase):
    __tablename__ = 'bankconma'

    conmaid = Column(Integer, primary_key=True)
    conno = Column(String(50), nullable=False, unique=True)
    fill_date = Column(DateTime, nullable=False)
    start_date = Column(DateTime, nullable=False)
    end_date = Column(DateTime, nullable=False)
    tag = Column(SmallInteger, nullable=False)
    calloss = Column(TINYINT, server_default=text("1"))
    calpure = Column(TINYINT, server_default=text("1"))


class Bankconspec(CNBase):
    __tablename__ = 'bankconspec'

    id = Column(SmallInteger, primary_key=True)
    srcconid = Column(SmallInteger, nullable=False)
    prodconid = Column(SmallInteger, nullable=False)
    lossrate = Column(DECIMAL(6, 3), nullable=False)
    purerate = Column(DECIMAL(6, 3), nullable=False)


class Codetable(CNBase):
    __tablename__ = 'codetable'

    id = Column(Integer, primary_key=True)
    tblname = Column(String(20), nullable=False, index=True)
    colname = Column(String(20), nullable=False)
    coden = Column(Float, nullable=False, index=True, server_default=text("0"))
    coden1 = Column(Float, nullable=False, server_default=text("0"))
    coden2 = Column(Float, nullable=False, server_default=text("0"))
    codec = Column(String(255), nullable=False, index=True, server_default=text("N/A"))
    codec1 = Column(String(255), nullable=False, server_default=text("N/A"))
    codec2 = Column(String(255), nullable=False, server_default=text("N/A"))
    coded = Column(DateTime, nullable=False, server_default=text("getdate()"))
    coded1 = Column(DateTime, nullable=False, server_default=text("getdate()"))
    coded2 = Column(DateTime, nullable=False, server_default=text("getdate()"))
    code = Column(String(30), nullable=False, server_default=text("N/A"))
    description = Column(String(60), nullable=False, server_default=text("N/A"))
    fill_date = Column(DateTime, nullable=False, server_default=text("getdate()"))
    tag = Column(SmallInteger, nullable=False, index=True, server_default=text("0"))
    pid = Column(Integer, nullable=False, index=True, server_default=text("0"))


class Cstinfo(CNBase):
    __tablename__ = 'cstinfo'

    cstid = Column(SmallInteger, primary_key=True)
    cstname = Column(CHAR(30), nullable=False, unique=True)
    description = Column(CHAR(40), nullable=False)
    range = Column(CHAR(20), nullable=False)
    groupid = Column(SmallInteger, nullable=False)
    tag = Column(TINYINT, nullable=False)
    pcstid = Column(SmallInteger, server_default=text("0"))


class Dscstinfo(CNBase):
    __tablename__ = 'dscstinfo'

    id = Column(Integer, primary_key=True)
    createdate = Column(DateTime, nullable=False, server_default=text("getdate()"))
    lastuserid = Column(SmallInteger, nullable=False)
    lastupdate = Column(DateTime, nullable=False, server_default=text("getdate()"))
    tag = Column(SmallInteger, nullable=False, server_default=text("0"))
    pid = Column(Integer, nullable=False, index=True)
    name = Column(String(250), nullable=False, unique=True)
    description = Column(String(250), nullable=False)
    contact = Column(String(250), nullable=False)
    phone = Column(String(250), nullable=False)
    address = Column(String(250), nullable=False)


class Dsmetal(CNBase):
    __tablename__ = 'dsmetal'

    id = Column(Integer, primary_key=True)
    createdate = Column(DateTime, nullable=False, server_default=text("getdate()"))
    lastuserid = Column(SmallInteger, nullable=False)
    lastupdate = Column(DateTime, nullable=False, server_default=text("getdate()"))
    tag = Column(SmallInteger, nullable=False, server_default=text("0"))
    name = Column(String(250), nullable=False, unique=True)
    alias = Column(String(250), nullable=False)
    pid = Column(Integer, nullable=False)
    factor = Column(DECIMAL(4, 3), nullable=False)
    factor1 = Column(DECIMAL(4, 3), nullable=False)


class Dsstock(CNBase):
    __tablename__ = 'dsstock'

    id = Column(Integer, primary_key=True)
    createdate = Column(DateTime, nullable=False, server_default=text("getdate()"))
    lastuserid = Column(SmallInteger, nullable=False)
    lastupdate = Column(DateTime, nullable=False, server_default=text("getdate()"))
    tag = Column(SmallInteger, nullable=False, server_default=text("0"))
    name = Column(String(250), nullable=False, unique=True)
    pid = Column(Integer, nullable=False, index=True)
    location = Column(String(250), nullable=False)
    description = Column(String(250), nullable=False)


class Dsstor(CNBase):
    __tablename__ = 'dsstor'

    id = Column(Integer, primary_key=True)
    createdate = Column(DateTime, nullable=False, server_default=text("getdate()"))
    lastuserid = Column(SmallInteger, nullable=False)
    lastupdate = Column(DateTime, nullable=False, server_default=text("getdate()"))
    tag = Column(SmallInteger, nullable=False, server_default=text("0"))
    name = Column(String(50), nullable=False, unique=True)
    docdate = Column(DateTime, nullable=False, server_default=text("getdate()"))
    type = Column(SmallInteger, nullable=False, server_default=text("0"))
    amount = Column(TINYINT, nullable=False, server_default=text("0"))
    remark = Column(String(250), nullable=False, server_default=text("N/A"))


class Dsstyma(CNBase):
    __tablename__ = 'dsstyma'

    id = Column(Integer, primary_key=True)
    createdate = Column(DateTime, nullable=False, server_default=text("getdate()"))
    lastuserid = Column(SmallInteger, nullable=False)
    lastupdate = Column(DateTime, nullable=False, server_default=text("getdate()"))
    tag = Column(SmallInteger, nullable=False, server_default=text("0"))
    name = Column(String(20), nullable=False, unique=True)
    type = Column(SmallInteger, nullable=False)
    description = Column(String(250), nullable=False)


class Dsupc(CNBase):
    __tablename__ = 'dsupc'

    id = Column(Integer, primary_key=True)
    createdate = Column(DateTime, nullable=False, server_default=text("getdate()"))
    lastuserid = Column(SmallInteger, nullable=False)
    lastupdate = Column(DateTime, nullable=False, server_default=text("getdate()"))
    tag = Column(SmallInteger, nullable=False, server_default=text("0"))
    name = Column(String(13), nullable=False, unique=True)
    tagprice = Column(TINYINT, nullable=False, server_default=text("0"))
    itemno = Column(String(250), nullable=False, server_default=text("N/A"))


class Dsvendor(CNBase):
    __tablename__ = 'dsvendor'

    id = Column(Integer, primary_key=True)
    createdate = Column(DateTime, nullable=False, server_default=text("getdate()"))
    lastuserid = Column(SmallInteger, nullable=False)
    lastupdate = Column(DateTime, nullable=False, server_default=text("getdate()"))
    tag = Column(SmallInteger, nullable=False, server_default=text("0"))
    name = Column(String(250), nullable=False, unique=True)
    description = Column(String(250), nullable=False)


class Employee(CNBase):
    __tablename__ = 'employee'

    wk_id = Column(Integer, primary_key=True)
    dept_id = Column(SmallInteger, nullable=False)
    emp_id = Column(SmallInteger, nullable=False)
    name = Column(String(12), nullable=False)
    tag = Column(SmallInteger)
    salary = Column(SMALLMONEY, server_default=text("0"))
    mark = Column(TINYINT, server_default=text("0"))
    wkidref = Column(Integer, server_default=text("0"))


class Empmark(CNBase):
    __tablename__ = 'empmark'

    year_ = Column(SmallInteger, primary_key=True, nullable=False)
    month_ = Column(TINYINT, primary_key=True, nullable=False)
    worker_id = Column(SmallInteger, primary_key=True, nullable=False)
    salary = Column(SMALLMONEY, nullable=False, server_default=text("0"))
    points = Column(Integer, nullable=False)
    tag = Column(TINYINT, nullable=False)


class Empsalary(CNBase):
    __tablename__ = 'empsalary'
    __table_args__ = (
        Index('idx_empsalary', 'wk_id', 'year', 'month'),
    )

    id = Column(Integer, primary_key=True)
    wk_id = Column(Integer, nullable=False)
    year = Column(SmallInteger, nullable=False)
    month = Column(TINYINT, nullable=False)
    salary = Column(SMALLMONEY, nullable=False)


class Expma(CNBase):
    __tablename__ = 'expma'
    __table_args__ = (
        Index('pk_expma_catalog_isrepair', 'catalog', 'isrepair', unique=True),
    )

    expmaid = Column(Integer, primary_key=True)
    karat = Column(SmallInteger, nullable=False)
    catalog = Column(SmallInteger, nullable=False)
    isrepair = Column(TINYINT, nullable=False)
    inv_date = Column(DateTime, nullable=False)
    fill_date = Column(DateTime, nullable=False)
    last_update = Column(SmallInteger, nullable=False)
    color = Column(String(250), nullable=False, server_default=text("255,255,255"))


class Expstock(CNBase):
    __tablename__ = 'expstock'

    stockid = Column(Integer, primary_key=True)
    inorout = Column(SmallInteger)
    karat = Column(SmallInteger)
    type = Column(SmallInteger)
    goldwgt = Column(MONEY)
    pgoldwgt = Column(MONEY)
    date = Column(DateTime)
    status = Column(SmallInteger)
    tag = Column(SmallInteger)


class FixRoller(CNBase):
    __tablename__ = 'fix_roller'

    id = Column(Numeric(18, 0), primary_key=True)
    romid = Column(Integer)
    jsid = Column(Integer, nullable=False, index=True)
    wgt_in = Column(MONEY, nullable=False)
    date_in = Column(DateTime, nullable=False)
    wgt_out = Column(MONEY, nullable=False)
    date_out = Column(DateTime, nullable=False)
    wgtin_st = Column(MONEY, server_default=text("0 null"))
    wgtout_st = Column(MONEY, server_default=text("0 null"))


class FixRollerm(CNBase):
    __tablename__ = 'fix_rollerm'
    __table_args__ = (
        Index('idx_fixrollerm_wk', 'worker_id', 'karat', 'tag', unique=True),
    )

    romid = Column(Integer, primary_key=True)
    worker_id = Column(SmallInteger, nullable=False)
    karat = Column(SmallInteger, nullable=False)
    rate = Column(MONEY, nullable=False)
    total_in = Column(MONEY, nullable=False)
    total_out = Column(MONEY, nullable=False)
    total_bck = Column(MONEY, nullable=False)
    fill_date = Column(DateTime, nullable=False)
    modi_date = Column(DateTime, nullable=False)
    tag = Column(Integer, nullable=False)


class FixSummary(CNBase):
    __tablename__ = 'fix_summary'
    __table_args__ = (
        Index('idx_fixsummary_pk', 'worker_id', 'karat', 'tag', unique=True),
    )

    fixsumid = Column(Integer, primary_key=True)
    worker_id = Column(SmallInteger, nullable=False)
    karat = Column(SmallInteger, nullable=False)
    tag = Column(Integer, nullable=False)
    qty = Column(SmallInteger, nullable=False)
    total_in = Column(MONEY, nullable=False)
    total_outo = Column(MONEY, nullable=False)
    total_out = Column(MONEY, nullable=False)
    total_bck = Column(MONEY, nullable=False)
    fill_date = Column(DateTime, nullable=False)
    modi_date = Column(DateTime, nullable=False)


class FsSummary(CNBase):
    __tablename__ = 'fs_summary'
    __table_args__ = (
        Index('idx_fs_summary_pk', 'worker_id', 'karat', 'tag', unique=True),
    )

    fssumid = Column(Integer, primary_key=True)
    worker_id = Column(SmallInteger, nullable=False)
    karat = Column(SmallInteger, nullable=False)
    qty = Column(Integer, nullable=False)
    total_in = Column(MONEY, nullable=False)
    total_out = Column(MONEY, nullable=False)
    total_bck = Column(MONEY, nullable=False)
    fill_date = Column(DateTime, nullable=False)
    modi_date = Column(DateTime, nullable=False)
    tag = Column(Integer, nullable=False)


class GdAuxma(CNBase):
    __tablename__ = 'gd_auxma'
    __table_args__ = (
        Index('idx_gdauxma_dptid', 'dptid', 'karat'),
    )

    btchid = Column(Integer, primary_key=True)
    pkid = Column(SmallInteger, nullable=False)
    dptid = Column(TINYINT, nullable=False)
    karat = Column(SmallInteger, nullable=False)
    qty_in = Column(Integer, nullable=False)
    wgt_in = Column(MONEY, nullable=False)
    qty_used = Column(Integer, nullable=False)
    wgt_used = Column(MONEY, nullable=False)
    qty_usedo = Column(Integer, nullable=False)
    wgt_usedo = Column(MONEY, nullable=False)
    qty_bck = Column(Integer, nullable=False)
    wgt_bck = Column(MONEY, nullable=False)
    wgt_loss = Column(MONEY, nullable=False)
    fill_date = Column(DateTime, nullable=False)
    tag = Column(Integer, nullable=False, index=True)


class GdSummary(CNBase):
    __tablename__ = 'gd_summary'
    __table_args__ = (
        Index('idx_goldmaster', 'dptid', 'karat', 'tag'),
    )

    gdsumid = Column(Integer, primary_key=True)
    dptid = Column(TINYINT, nullable=False)
    karat = Column(SmallInteger, nullable=False)
    qty = Column(Integer, nullable=False)
    total_in = Column(MONEY, nullable=False)
    total_out = Column(MONEY, nullable=False)
    fill_date = Column(DateTime, nullable=False)
    modi_date = Column(DateTime, nullable=False)
    tag = Column(Integer, nullable=False)


class Goldplatingx(CNBase):
    __tablename__ = 'goldplatingx'
    __table_args__ = (
        Index('idxgpxpk', 'styid', 'polp', 'plap', 'othp', unique=True),
    )

    id = Column(Integer, primary_key=True)
    styid = Column(Integer, nullable=False)
    polp = Column(SMALLMONEY, nullable=False)
    plap = Column(SMALLMONEY, nullable=False)
    othp = Column(SMALLMONEY, nullable=False)
    filldate = Column(DateTime, nullable=False)
    remark = Column(String(20), nullable=False)
    tag = Column(SmallInteger, nullable=False)


class Holiday(CNBase):
    __tablename__ = 'holidays'
    __table_args__ = (
        Index('idx_holiday_daterange', 'datefrom', 'dateto'),
    )

    id = Column(Integer, primary_key=True)
    datefrom = Column(DateTime, nullable=False)
    dateto = Column(DateTime, nullable=False)
    fill_date = Column(DateTime, nullable=False)
    tag = Column(Integer, nullable=False)


class InsertKey(CNBase):
    __tablename__ = 'insert_key'

    id = Column(Integer, primary_key=True)
    tblname = Column(String(20), nullable=False)
    colname = Column(String(20))
    maxvalue = Column(String(50))
    remark = Column(String(30))
    tag = Column(Integer)


class InsertKey1(CNBase):
    __tablename__ = 'insert_key1'
    __table_args__ = (
        Index('idx_insertkey_pk', 'tblname', 'colname'),
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


class Jocostma(CNBase):
    __tablename__ = 'jocostma'

    jsid = Column(Integer, primary_key=True)
    running = Column(Integer, nullable=False, unique=True)
    costrefid = Column(SmallInteger, nullable=False)
    fill_date = Column(DateTime, nullable=False)
    send_date = Column(DateTime, nullable=False)
    userid = Column(SmallInteger, nullable=False)
    tag = Column(Integer, nullable=False)
    cost = Column(MONEY, nullable=False)
    lastuserid = Column(SmallInteger, nullable=False)
    lastupdate = Column(DateTime, nullable=False, server_default=text("getdate()"))


class Jocostprinting(Jocostma):
    __tablename__ = 'jocostprinting'

    jsid = Column(ForeignKey('jocostma.jsid'), primary_key=True)
    printdate = Column(DateTime, nullable=False)


t_jocostrpt = Table(
    'jocostrpt', metadata,
    Column('id', Integer, nullable=False, unique=True),
    Column('deptid', String(10), nullable=False),
    Column('job_sheet_id', Integer, nullable=False),
    Column('date_in', DateTime, nullable=False),
    Column('date_out', DateTime),
    Column('tag', SmallInteger)
)


t_key_buffer = Table(
    'key_buffer', metadata,
    Column('tblname', String(50), nullable=False),
    Column('colname', String(50), nullable=False),
    Column('nvalue', Integer, nullable=False),
    Column('processid', Integer, nullable=False),
    Column('createdate', DateTime, nullable=False),
    Column('tag', SmallInteger, nullable=False),
    Index('idx_key_buffer', 'tblname', 'colname', 'nvalue', unique=True)
)


class LmlgDepartment(CNBase):
    __tablename__ = 'lmlg_department'

    dep_id = Column(Integer, primary_key=True)
    dep_head_id = Column(String(20), nullable=False)
    dep_name = Column(String(30), nullable=False)
    tag = Column(SmallInteger, nullable=False)


class LmlgFormua(CNBase):
    __tablename__ = 'lmlg_formua'

    formua_id = Column(Integer, primary_key=True)
    formua_name = Column(String(20), nullable=False)
    tag = Column(SmallInteger, nullable=False)


class LmlgGold(CNBase):
    __tablename__ = 'lmlg_gold'

    gold_id = Column(Integer, primary_key=True)
    gold_name = Column(String(30), nullable=False)
    tag = Column(SmallInteger, nullable=False)


class LmlgOperation(CNBase):
    __tablename__ = 'lmlg_operation'

    op_id = Column(Integer, primary_key=True)
    op_name = Column(String(25), nullable=False)
    op_type = Column(String(25), nullable=False)
    password = Column(String(25), nullable=False)
    tag = Column(SmallInteger, nullable=False)


class LmlgReturnWeigh(CNBase):
    __tablename__ = 'lmlg_return_weigh'

    return_id = Column(Integer, primary_key=True)
    pro_id = Column(Integer, nullable=False)
    return_weigh = Column(Numeric(18, 2), nullable=False)
    lost_weigh = Column(Numeric(18, 2), nullable=False)
    return_date = Column(DateTime, nullable=False)
    tag = Column(SmallInteger, nullable=False)


t_lmy_gold_io = Table(
    'lmy_gold_io', metadata,
    Column('gold_id', Integer, nullable=False, unique=True),
    Column('wk_id', Integer, nullable=False),
    Column('ratio_in', Integer, nullable=False),
    Column('ratio_out', Integer, nullable=False),
    Column('ratio_wgt', Numeric(18, 3), nullable=False),
    Column('enough_wgt', Numeric(18, 3), nullable=False),
    Column('date', DateTime, nullable=False),
    Column('tag', SmallInteger, nullable=False),
    Column('status', SmallInteger, nullable=False)
)


class Metalma(CNBase):
    __tablename__ = 'metalma'

    karat = Column(SmallInteger, primary_key=True)
    mname = Column(String(20), nullable=False)
    mtype = Column(String(10), nullable=False)
    percentage = Column(SMALLMONEY, nullable=False)
    fill_date = Column(DateTime, nullable=False)
    pkarat = Column(SmallInteger, nullable=False)
    description = Column(String(30), nullable=False)
    tag = Column(TINYINT, nullable=False)
    alias = Column(SmallInteger, server_default=text("0"))


class Mitma(CNBase):
    __tablename__ = 'mitma'

    mitid = Column(Integer, primary_key=True)
    mitname = Column(String(15), nullable=False, unique=True)
    karat = Column(SmallInteger, nullable=False)
    fill_date = Column(DateTime, nullable=False)
    wgt_in = Column(MONEY, nullable=False)
    wgt_used = Column(MONEY, nullable=False)
    wgt_bck = Column(MONEY, nullable=False)
    modi_date = Column(DateTime, nullable=False)
    tag = Column(TINYINT, nullable=False)
    labcost = Column(SMALLMONEY, nullable=False)
    cstid = Column(SmallInteger, server_default=text("-1"))
    objid = Column(SmallInteger, server_default=text("0"))


class Mm(CNBase):
    __tablename__ = 'mm'

    mmid = Column(Integer, primary_key=True)
    refid = Column(Integer, nullable=False, index=True)
    docno = Column(String(8), nullable=False)
    jsid = Column(Integer, nullable=False, index=True)
    qty = Column(SMALLMONEY, nullable=False)
    tag = Column(TINYINT, nullable=False, index=True)


class Mmma(CNBase):
    __tablename__ = 'mmma'

    refid = Column(Integer, primary_key=True)
    refno = Column(String(11), nullable=False)
    karat = Column(SmallInteger, nullable=False)
    refdate = Column(DateTime, nullable=False, index=True)
    tag = Column(Integer, nullable=False, index=True)


class Monitor(CNBase):
    __tablename__ = 'monitor'

    moid = Column(Integer, primary_key=True)
    tblname = Column(String(50), nullable=False)
    status = Column(String(50), nullable=False)
    nvalue = Column(Float, nullable=False)
    nvalue1 = Column(Float, nullable=False)
    nvalue2 = Column(Float, nullable=False)
    nvalue3 = Column(Float, nullable=False)
    cvalue = Column(String(255), nullable=False)
    cvalue1 = Column(String(255), nullable=False)
    remark = Column(String(255), nullable=False)
    userid = Column(Integer, nullable=False)
    date = Column(DateTime, nullable=False)
    tag = Column(SmallInteger, nullable=False)


class Operator(CNBase):
    __tablename__ = 'operator'

    user_id = Column(CHAR(10), primary_key=True)
    password = Column(CHAR(20))
    grade = Column(SmallInteger)
    _class = Column('class', SmallInteger)


t_pbcatcol = Table(
    'pbcatcol', metadata,
    Column('pbc_tnam', CHAR(30), nullable=False),
    Column('pbc_tid', Integer, nullable=False),
    Column('pbc_ownr', CHAR(30), nullable=False),
    Column('pbc_cnam', CHAR(30), nullable=False),
    Column('pbc_cid', SmallInteger, nullable=False),
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
    Column('pbe_edit', String(254), nullable=False),
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
    Column('pbt_tnam', CHAR(30), nullable=False),
    Column('pbt_tid', Integer, nullable=False, unique=True),
    Column('pbt_ownr', CHAR(30), nullable=False),
    Column('pbd_fhgt', SmallInteger),
    Column('pbd_fwgt', SmallInteger),
    Column('pbd_fitl', CHAR(1)),
    Column('pbd_funl', CHAR(1)),
    Column('pbd_fchr', SmallInteger),
    Column('pbd_fptc', SmallInteger),
    Column('pbd_ffce', CHAR(32)),
    Column('pbh_fhgt', SmallInteger),
    Column('pbh_fwgt', SmallInteger),
    Column('pbh_fitl', CHAR(1)),
    Column('pbh_funl', CHAR(1)),
    Column('pbh_fchr', SmallInteger),
    Column('pbh_fptc', SmallInteger),
    Column('pbh_ffce', CHAR(32)),
    Column('pbl_fhgt', SmallInteger),
    Column('pbl_fwgt', SmallInteger),
    Column('pbl_fitl', CHAR(1)),
    Column('pbl_funl', CHAR(1)),
    Column('pbl_fchr', SmallInteger),
    Column('pbl_fptc', SmallInteger),
    Column('pbl_ffce', CHAR(32)),
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


class PrPeriodInfo(CNBase):
    __tablename__ = 'prPeriodInfo'
    __table_args__ = (
        Index('idx_prdI_ym', 'appId', 'year', 'month', unique=True),
    )

    id = Column(Numeric(5, 0), primary_key=True)
    appId = Column(CHAR(20), nullable=False)
    salId = Column(Integer, nullable=False)
    year = Column(Integer, nullable=False)
    month = Column(Integer, nullable=False)
    lastModified = Column(DateTime, nullable=False)


class Prodspecphoto(CNBase):
    __tablename__ = 'prodspecphoto'

    id = Column(String(50), primary_key=True)
    parent = Column(String(50))
    created = Column(DateTime, nullable=False)
    modified = Column(DateTime, nullable=False)
    lastuserid = Column(SmallInteger, nullable=False)
    type = Column(CHAR(10), nullable=False)
    tag = Column(SmallInteger, nullable=False)


class SecurityApp(CNBase):
    __tablename__ = 'security_apps'

    appid = Column(Integer, primary_key=True)
    application = Column(String(32), nullable=False, index=True)
    description = Column(String(64), nullable=False)


class SecurityUser(CNBase):
    __tablename__ = 'security_users'

    userid = Column(Integer, primary_key=True)
    name = Column(String(16), nullable=False, unique=True)
    description = Column(String(32), nullable=False)
    priority = Column(Integer, nullable=False)
    password = Column(String(20), nullable=False)
    status = Column(TINYINT, nullable=False)
    user_type = Column(Integer, nullable=False)


class Stinvm(CNBase):
    __tablename__ = 'stinvm'

    cstid = Column(SmallInteger, primary_key=True, nullable=False)
    invid = Column(Integer, primary_key=True, nullable=False)
    tag = Column(SmallInteger)


class Stmp(CNBase):
    __tablename__ = 'stmp'

    id = Column(Integer, primary_key=True)
    bill_id = Column(Integer, nullable=False)
    cstbldid_alpha = Column(String(20), nullable=False)
    cstbldid_digit = Column(Integer, nullable=False)


class StoneMaster(CNBase):
    __tablename__ = 'stone_master'

    stid = Column(SmallInteger, primary_key=True)
    stname = Column(String(4), nullable=False, unique=True)
    edesc = Column(String(50), nullable=False)
    cdesc = Column(String(50), nullable=False)
    sttype = Column(String(20), nullable=False)
    settype = Column(TINYINT, nullable=False)
    fill_date = Column(DateTime, nullable=False)
    tag = Column(TINYINT, nullable=False)
    lastuserid = Column(SmallInteger, server_default=text("0"))
    lastupdate = Column(DateTime, server_default=text("getdate()"))


class StonePkissue(CNBase):
    __tablename__ = 'stone_pkissue'
    __table_args__ = (
        Index('idx_stonepkissue', 'pkid', 'cstid', 'fill_date'),
    )

    id = Column(Integer, primary_key=True)
    pkid = Column(Integer, nullable=False)
    cstid = Column(SmallInteger, nullable=False)
    fill_date = Column(DateTime, nullable=False)
    wgt_in = Column(Float, nullable=False)
    wgt_trans = Column(Float, nullable=False)
    wgt_bck = Column(Float, nullable=False)
    wgt_used = Column(Float, nullable=False)
    wgt_ret = Column(Float, nullable=False)
    wgt_lost = Column(Float, nullable=False)
    wgt_oin = Column(Float, nullable=False)
    wgt_oout = Column(Float, nullable=False)
    wgt_adjust = Column(Float, nullable=False)
    tag = Column(SmallInteger, nullable=False)
    wgt_laststock = Column(Float, nullable=False)


class StoneSizemaster(CNBase):
    __tablename__ = 'stone_sizemaster'

    stsizeid = Column(SmallInteger, primary_key=True)
    stsize = Column(String(20), nullable=False, unique=True)
    wgt = Column(Float, nullable=False)
    grp = Column(CHAR(2))
    pingrp = Column(SmallInteger)
    tag = Column(TINYINT, nullable=False)


class StoneSndout(CNBase):
    __tablename__ = 'stone_sndout'

    id = Column(Integer, primary_key=True)
    worker_id = Column(SmallInteger, nullable=False)
    btchid = Column(Integer, nullable=False, index=True)
    wgt = Column(Float, nullable=False)
    wgtbck = Column(Float, nullable=False)
    relatedid = Column(Integer, nullable=False, index=True)
    fill_date = Column(DateTime, nullable=False)
    tag = Column(Integer, nullable=False, index=True)


class Stoutchkinfo(CNBase):
    __tablename__ = 'stoutchkinfo'

    checkid = Column(Integer, primary_key=True)
    check_date = Column(DateTime, nullable=False)
    userid = Column(SmallInteger, nullable=False)
    tag = Column(TINYINT, nullable=False)


class Stoutchkrst(CNBase):
    __tablename__ = 'stoutchkrst'

    id = Column(Integer, primary_key=True, nullable=False)
    idx = Column(TINYINT, primary_key=True, nullable=False)
    errid = Column(SmallInteger, nullable=False)
    errmsg = Column(String(255), nullable=False)
    lastuserid = Column(SmallInteger, server_default=text("0"))


class Stsumma(CNBase):
    __tablename__ = 'stsumma'
    __table_args__ = (
        Index('idx_stsumma_pk', 'pkid', 'tag', unique=True),
    )

    stsumid = Column(Integer, primary_key=True)
    pkid = Column(Integer, nullable=False)
    qty = Column(Integer, nullable=False)
    wgt = Column(DECIMAL(11, 3), nullable=False)
    fill_date = Column(DateTime, nullable=False)
    modi_date = Column(DateTime, nullable=False)
    tag = Column(TINYINT, nullable=False)


class Sttobck(CNBase):
    __tablename__ = 'sttobck'

    tid = Column(Integer, primary_key=True)
    refno = Column(String(8), nullable=False)
    btchid = Column(Integer, nullable=False)
    qty = Column(Integer, nullable=False)
    wgt = Column(Float, nullable=False)
    filldate = Column(DateTime, nullable=False)
    wgtleft = Column(Float, nullable=False)
    tag = Column(TINYINT, nullable=False)


class Styma(CNBase):
    __tablename__ = 'styma'
    __table_args__ = (
        Index('idx_styma', 'alpha', 'digit', 'suffix', unique=True),
    )

    styid = Column(Integer, primary_key=True)
    alpha = Column(String(5), nullable=False)
    digit = Column(Integer, nullable=False)
    fixpoints = Column(TINYINT, nullable=False)
    description = Column(String(30), nullable=False)
    tag = Column(TINYINT, nullable=False)
    fill_date = Column(DateTime, nullable=False)
    suffix = Column(String(50), server_default=text("\"\""))
    name1 = Column(String(20), index=True)
    remark = Column(String(250))


class Stypropdef(CNBase):
    __tablename__ = 'stypropdef'

    propdefid = Column(Integer, primary_key=True)
    propname = Column(CHAR(50), nullable=False, unique=True)
    propgrp = Column(CHAR(50), nullable=False)
    proptype = Column(Integer, nullable=False)


class Stypropst(CNBase):
    __tablename__ = 'stypropst'

    id = Column(Integer, primary_key=True)
    stname = Column(String(10), nullable=False)
    des = Column(String(50), nullable=False)
    tag = Column(SmallInteger, nullable=False)


class SysPrintma(CNBase):
    __tablename__ = 'sys_printma'
    __table_args__ = (
        Index('idx_sys_printma', 'idx', 'objid', 'objtype'),
    )

    printid = Column(Integer, primary_key=True)
    objtype = Column(SmallInteger, nullable=False)
    objid = Column(Integer, nullable=False)
    idx = Column(SmallInteger, nullable=False)
    fill_date = Column(DateTime, nullable=False)
    tag = Column(TINYINT, nullable=False)


class Tmp(CNBase):
    __tablename__ = 'tmp'

    id = Column(Integer, primary_key=True)
    alpha = Column(String(5), nullable=False)
    digit = Column(Integer, nullable=False)
    remark = Column(String(250), nullable=False)


class Trnote(CNBase):
    __tablename__ = 'trnote'

    trmaid = Column(Integer, primary_key=True, nullable=False)
    idx = Column(TINYINT, primary_key=True, nullable=False)
    jono = Column(String(20), nullable=False)
    remark = Column(String(20), nullable=False)


class Trnotema(CNBase):
    __tablename__ = 'trnotema'

    id = Column(Integer, primary_key=True)
    name = Column(String(20), nullable=False)
    remark = Column(String(150), nullable=False)
    machineId = Column(String(50), nullable=False)
    fill_date = Column(DateTime, nullable=False)


class Tt(CNBase):
    __tablename__ = 'tt'

    id = Column(Integer, primary_key=True)
    prodcode = Column(String(20), nullable=False)


class Utilityinma(CNBase):
    __tablename__ = 'utilityinma'

    uimaid = Column(Integer, primary_key=True)
    cstid = Column(SmallInteger, nullable=False)
    docno = Column(String(250), nullable=False)
    fill_date = Column(DateTime, nullable=False)
    created_date = Column(DateTime, nullable=False)
    status = Column(SmallInteger)


class Utilityma(CNBase):
    __tablename__ = 'utilityma'

    uid = Column(Integer, primary_key=True)
    ucode = Column(String(250), nullable=False)
    catalogid = Column(SmallInteger, nullable=False)
    size = Column(String(250), nullable=False)
    spec = Column(String(250), nullable=False)
    unit = Column(SmallInteger, nullable=False)
    fill_date = Column(DateTime, nullable=False)
    lastuserid = Column(SmallInteger, nullable=False)
    lastupdate = Column(DateTime, nullable=False)
    unitsin = Column(Float, nullable=False)
    unitsout = Column(Float, nullable=False)
    unitsadjusted = Column(Float, nullable=False)
    tag = Column(SmallInteger, nullable=False)
    status = Column(SmallInteger)


t_uv_cba = Table(
    'uv_cba', metadata,
    Column('cust_bill_id', String(13)),
    Column('bill_id', Integer, nullable=False),
    Column('is_out', SmallInteger, nullable=False),
    Column('package_id', String(20), nullable=False),
    Column('fill_date', DateTime, nullable=False),
    Column('quantity', Integer, nullable=False),
    Column('weight', Float, nullable=False),
    Column('packed', TINYINT, nullable=False),
    Column('idx', TINYINT, nullable=False),
    Column('check_date', DateTime, nullable=False),
    Column('jsid', Integer, nullable=False),
    Column('running', Integer, nullable=False),
    Column('pricen', SMALLMONEY),
    Column('unit', SmallInteger, nullable=False),
    Column('styno', String(65)),
    Column('batch_id', String(15), nullable=False)
)


t_uv_cstbldst = Table(
    'uv_cstbldst', metadata,
    Column('jsid', Integer, nullable=False),
    Column('jsalpha', String(3)),
    Column('jsdigit', Integer, nullable=False),
    Column('jsdead_line', DateTime, nullable=False),
    Column('jsmodi_date', DateTime, nullable=False),
    Column('jsdocno', CHAR(7), nullable=False),
    Column('jstag', SmallInteger, nullable=False),
    Column('jsrunning', Integer, nullable=False),
    Column('jsqty', SMALLMONEY, nullable=False),
    Column('jsqtyleft', SMALLMONEY, nullable=False),
    Column('jsqtystleft', SMALLMONEY, nullable=False),
    Column('cstid', SmallInteger, nullable=False),
    Column('cstname', CHAR(30), nullable=False),
    Column('cstdesc', CHAR(40), nullable=False),
    Column('cstgrpid', SmallInteger, nullable=False),
    Column('stoutid', Integer, nullable=False),
    Column('stoutdocid', Integer, nullable=False),
    Column('stoutis_out', SmallInteger, nullable=False),
    Column('stoutfill_date', DateTime, nullable=False),
    Column('packed', TINYINT, nullable=False),
    Column('subcnt', TINYINT, nullable=False),
    Column('pkid', Integer, nullable=False),
    Column('package_id', String(20), nullable=False),
    Column('pkunit', SmallInteger, nullable=False),
    Column('pkprice', CHAR(6), nullable=False),
    Column('pkpricen', SMALLMONEY),
    Column('stinbatch_id', String(15), nullable=False),
    Column('btchid', Integer, nullable=False),
    Column('stoutidx', TINYINT, nullable=False),
    Column('stoutqty', Integer, nullable=False),
    Column('stoutwgt', Float, nullable=False),
    Column('stoutcheck_date', DateTime, nullable=False),
    Column('stoutfixedqty', SmallInteger),
    Column('worker_id', SmallInteger, nullable=False)
)


t_uv_package = Table(
    'uv_package', metadata,
    Column('bill_id', String(15), nullable=False),
    Column('pkid', Integer, nullable=False),
    Column('package_id', String(20), nullable=False),
    Column('unit', SmallInteger, nullable=False),
    Column('price', CHAR(6), nullable=False),
    Column('btchid', Integer, nullable=False),
    Column('batch_id', String(15), nullable=False),
    Column('quantity', Integer, nullable=False),
    Column('weight', Float, nullable=False),
    Column('wgt_used', Float, nullable=False),
    Column('qty_bck', Integer),
    Column('wgt_bck', Float, nullable=False),
    Column('size', String(60), nullable=False),
    Column('fill_date', DateTime, nullable=False),
    Column('is_used_up', SmallInteger, nullable=False),
    Column('adjust_wgt', Float, nullable=False),
    Column('cstid', SmallInteger, nullable=False),
    Column('qty_used', Integer, nullable=False),
    Column('cstref', String(60), nullable=False),
    Column('qtytrans', Integer, nullable=False),
    Column('wgttrans', Float, nullable=False),
    Column('wgt_tmptrans', Float, nullable=False),
    Column('wgt_prepared', Float, nullable=False),
    Column('stname', String(4), nullable=False),
    Column('cdesc', String(50), nullable=False),
    Column('sttype', String(20), nullable=False),
    Column('relpkid', Integer),
    Column('wgtunit', Float),
    Column('color', String(50))
)


t_uv_sbtch = Table(
    'uv_sbtch', metadata,
    Column('sbtchid', Integer, nullable=False),
    Column('pkid', Integer, nullable=False),
    Column('package_id', String(20), nullable=False),
    Column('unit_', SmallInteger, nullable=False),
    Column('price', CHAR(6), nullable=False),
    Column('stsizeid', SmallInteger, nullable=False),
    Column('stsize', String(20), nullable=False),
    Column('grade', TINYINT, nullable=False),
    Column('qty', Integer, nullable=False),
    Column('wgt', Float, nullable=False),
    Column('qtyadj', Integer, nullable=False),
    Column('wgtadj', Float, nullable=False),
    Column('fill_date', DateTime, nullable=False),
    Column('modi_date', DateTime, nullable=False),
    Column('qtytrans', Integer),
    Column('wgttrans', Float),
    Column('tag', Integer, nullable=False)
)


class Watchinvoicma(CNBase):
    __tablename__ = 'watchinvoicma'

    invmaid = Column(Integer, primary_key=True)
    corpid = Column(Integer, nullable=False)
    invno = Column(String(50), nullable=False)
    userid = Column(Integer, nullable=False)
    lastuserid = Column(Integer, nullable=False)
    createdate = Column(DateTime, nullable=False)
    lastupdate = Column(DateTime, nullable=False)
    status = Column(SmallInteger, nullable=False)
    tag = Column(SmallInteger, nullable=False)


class Watchobjtype(CNBase):
    __tablename__ = 'watchobjtype'

    wcobjtypeid = Column(Integer, primary_key=True)
    objname = Column(String(30), nullable=False)
    tag = Column(SmallInteger, nullable=False)


class BCustBill(CNBase):
    __tablename__ = 'b_cust_bill'
    __table_args__ = (
        Index('idx_bcstbld_cstbldid', 'cstbldid_alpha', 'cstbldid_digit', unique=True),
    )

    cstid = Column(ForeignKey('cstinfo.cstid'), nullable=False, index=True)
    jsid = Column(Integer, primary_key=True)
    remark = Column(CHAR(1), nullable=False, server_default=text("\" \""))
    cstbldid_digit = Column(Integer, nullable=False)
    styid = Column(ForeignKey('styma.styid'), nullable=False, index=True)
    dead_line = Column(DateTime, nullable=False)
    modi_date = Column(DateTime, nullable=False)
    dept_bill_id = Column(CHAR(7), nullable=False)
    tag = Column(SmallInteger, nullable=False)
    running = Column(Integer, nullable=False, index=True)
    quantity = Column(SMALLMONEY, nullable=False)
    wgt = Column(SMALLMONEY, nullable=False)
    qtyleft = Column(SMALLMONEY, nullable=False)
    qtystleft = Column(SMALLMONEY, nullable=False)
    description = Column(String(255), server_default=text("\" \""))
    shipdate = Column(DateTime)
    stid = Column(SmallInteger, server_default=text("109"))
    cstbldid_alpha = Column(String(3), server_default=text("\" \""))
    karat = Column(SmallInteger, server_default=text("0"))
    createdate = Column(DateTime, server_default=text("getdate()"))
    ponohk = Column(String(15), index=True)

    cstinfo = relationship('Cstinfo')
    styma = relationship('Styma')


class Bankcondtl(CNBase):
    __tablename__ = 'bankcondtl'

    condtlid = Column(Integer, primary_key=True)
    conid = Column(ForeignKey('bankcon.conid'), nullable=False)
    relconid = Column(SmallInteger, nullable=False)
    is_out = Column(TINYINT, nullable=False)
    wgt = Column(MONEY, nullable=False)
    wgtpure = Column(MONEY, nullable=False)
    wgtloss = Column(SMALLMONEY, nullable=False)
    fill_date = Column(DateTime, nullable=False)
    createdate = Column(DateTime, server_default=text("getdate()"))
    docno = Column(String(20), nullable=False, server_default=text("\"N/A\""))

    bankcon = relationship('Bankcon')


class Dsorder(CNBase):
    __tablename__ = 'dsorder'

    id = Column(Integer, primary_key=True)
    createdate = Column(DateTime, nullable=False, server_default=text("getdate()"))
    lastuserid = Column(SmallInteger, nullable=False)
    lastupdate = Column(DateTime, nullable=False, server_default=text("getdate()"))
    tag = Column(SmallInteger, nullable=False, server_default=text("0"))
    cstid = Column(ForeignKey('dscstinfo.id'), nullable=False)
    name = Column(String(250), nullable=False, unique=True)
    goldprice = Column(TINYINT, nullable=False, server_default=text("0"))
    remark = Column(String(250), nullable=False, server_default=text("N/A"))
    deadline = Column(DateTime, nullable=False)
    finishdate = Column(DateTime, nullable=False, server_default=text("getdate()"))
    type = Column(Integer, nullable=False)
    billno = Column(String(50), nullable=False, server_default=text("N/A"))

    dscstinfo = relationship('Dscstinfo')


class Dssto(CNBase):
    __tablename__ = 'dssto'

    id = Column(Integer, primary_key=True)
    createdate = Column(DateTime, nullable=False, server_default=text("getdate()"))
    lastuserid = Column(SmallInteger, nullable=False)
    lastupdate = Column(DateTime, nullable=False, server_default=text("getdate()"))
    tag = Column(SmallInteger, nullable=False, server_default=text("0"))
    vendorid = Column(ForeignKey('dsvendor.id'), nullable=False)
    stockid = Column(Integer, nullable=False, server_default=text("0"))
    styid = Column(ForeignKey('dsstyma.id'), nullable=False)
    running = Column(Integer, nullable=False, unique=True)
    karat = Column(ForeignKey('dsmetal.id'), nullable=False)
    type = Column(Integer, nullable=False)
    qty = Column(SMALLMONEY, nullable=False)
    qtydtl = Column(SMALLMONEY, nullable=False, server_default=text("0"))
    qtyused = Column(SMALLMONEY, nullable=False, server_default=text("0"))
    wgt = Column(TINYINT, nullable=False)
    wgtused = Column(TINYINT, nullable=False, server_default=text("0"))
    mwgt = Column(TINYINT, nullable=False, server_default=text("0"))
    mwgtused = Column(TINYINT, nullable=False, server_default=text("0"))
    laborcost = Column(TINYINT, nullable=False, server_default=text("0"))
    description = Column(String(100), nullable=False)
    cdescription = Column(String(50), nullable=False, server_default=text("N/A"))
    cdescription1 = Column(String(50), nullable=False, server_default=text("N/A"))
    refno = Column(String(30), nullable=False)
    remark = Column(String(50), nullable=False)
    cat = Column(SmallInteger, nullable=False, index=True, server_default=text("0"))

    dsmetal = relationship('Dsmetal')
    dsstyma = relationship('Dsstyma')
    dsvendor = relationship('Dsvendor')


class Expdtl(CNBase):
    __tablename__ = 'expdtl'

    expid = Column(Integer, primary_key=True)
    expmaid = Column(ForeignKey('expma.expmaid'), ForeignKey('expma.expmaid'), nullable=False, index=True)
    inv_no = Column(String(20), nullable=False, index=True)
    catalog = Column(SmallInteger, nullable=False, index=True)
    row = Column(SmallInteger, nullable=False)
    qty = Column(MONEY, nullable=False)
    ttwgt = Column(MONEY, nullable=False)
    goldwgt = Column(MONEY, nullable=False)
    pgoldwgt = Column(MONEY, nullable=False)
    fill_date = Column(DateTime, nullable=False)
    last_update = Column(SmallInteger, nullable=False)
    remark = Column(String(100))
    tag = Column(SmallInteger, nullable=False)
    bagwgt = Column(MONEY, nullable=False, server_default=text("0"))

    expma = relationship('Expma', primaryjoin='Expdtl.expmaid == Expma.expmaid')
    expma1 = relationship('Expma', primaryjoin='Expdtl.expmaid == Expma.expmaid')


class FixSndout(CNBase):
    __tablename__ = 'fix_sndout'

    id = Column(Integer, primary_key=True)
    fixsumid = Column(ForeignKey('fix_summary.fixsumid'), nullable=False)
    job_sheet_id = Column(Integer, nullable=False, index=True)
    quantity = Column(SmallInteger, nullable=False)
    wgt_in = Column(SMALLMONEY, nullable=False)
    date_in = Column(DateTime, nullable=False)
    wgt_out = Column(SMALLMONEY, nullable=False)
    date_out = Column(DateTime, nullable=False, index=True)
    rate = Column(TINYINT, nullable=False)
    gdstwgt_in = Column(SMALLMONEY, nullable=False, server_default=text("0"))
    stwgt_in = Column(SMALLMONEY, nullable=False, server_default=text("0"))
    userid = Column(SmallInteger, nullable=False, index=True, server_default=text("0"))
    lastuserid = Column(SmallInteger, nullable=False, server_default=text("0"))
    lastupdate = Column(DateTime, server_default=text("getdate()"))
    gdstwgt_out = Column(SMALLMONEY, nullable=False, server_default=text("0"))
    stwgt_out = Column(SMALLMONEY, nullable=False, server_default=text("0"))
    qtyforpc = Column(SMALLMONEY, nullable=False, server_default=text("0"))

    fix_summary = relationship('FixSummary')


class FsSndout(CNBase):
    __tablename__ = 'fs_sndout'

    id = Column(Integer, primary_key=True)
    fssumid = Column(ForeignKey('fs_summary.fssumid'), index=True)
    jsid = Column(Integer, nullable=False, index=True)
    qty = Column(TINYINT, nullable=False)
    date_in = Column(DateTime, nullable=False)
    wgt_in = Column(SMALLMONEY, nullable=False)
    date_out = Column(DateTime, nullable=False)
    wgt_out = Column(SMALLMONEY, nullable=False)
    rate = Column(TINYINT, nullable=False)
    wgtst = Column(DECIMAL(7, 3), nullable=False)

    fs_summary = relationship('FsSummary')


class GdAuxbck(CNBase):
    __tablename__ = 'gd_auxbck'

    btchid = Column(ForeignKey('gd_auxma.btchid'), primary_key=True, nullable=False)
    idx = Column(SmallInteger, primary_key=True, nullable=False)
    qty_in = Column(SmallInteger, nullable=False)
    wgt_in = Column(SMALLMONEY, nullable=False)
    wgt_loss = Column(SMALLMONEY, nullable=False)
    fill_date = Column(DateTime, nullable=False)

    gd_auxma = relationship('GdAuxma')


class GdAuxin(CNBase):
    __tablename__ = 'gd_auxin'

    btchid = Column(ForeignKey('gd_auxma.btchid'), primary_key=True, nullable=False)
    idx = Column(SmallInteger, primary_key=True, nullable=False)
    qty_in = Column(SmallInteger, nullable=False)
    wgt_in = Column(SMALLMONEY, nullable=False)
    fill_date = Column(DateTime, nullable=False)

    gd_auxma = relationship('GdAuxma')


class GdSndout(CNBase):
    __tablename__ = 'gd_sndout'

    id = Column(Integer, primary_key=True)
    gdsumid = Column(ForeignKey('gd_summary.gdsumid'), nullable=False, index=True)
    job_sheet_id = Column(Integer, nullable=False, index=True)
    qty = Column(TINYINT, nullable=False)
    wgt_in = Column(SMALLMONEY, nullable=False)
    date_in = Column(DateTime, nullable=False)
    wgt_out = Column(SMALLMONEY, nullable=False)
    date_out = Column(DateTime, nullable=False)
    session = Column(Integer, nullable=False, index=True)

    gd_summary = relationship('GdSummary')


class GdAux(GdSndout):
    __tablename__ = 'gd_aux'

    id = Column(ForeignKey('gd_sndout.id'), primary_key=True)
    btchid = Column(Integer, nullable=False)
    wgt = Column(SMALLMONEY, nullable=False)


class LmlgEmployee(CNBase):
    __tablename__ = 'lmlg_employee'

    emp_id = Column(Integer, primary_key=True)
    gonghao = Column(String(15), nullable=False)
    emp_name = Column(String(12), nullable=False)
    indate = Column(DateTime, nullable=False)
    address = Column(String(50), nullable=False)
    dep_id = Column(ForeignKey('lmlg_department.dep_id'), nullable=False)
    tag = Column(SmallInteger, nullable=False)

    dep = relationship('LmlgDepartment')


class LmlgProcessType(CNBase):
    __tablename__ = 'lmlg_process_type'

    pro_type_id = Column(Integer, primary_key=True)
    gold_id = Column(ForeignKey('lmlg_gold.gold_id'), nullable=False)
    formua_id = Column(ForeignKey('lmlg_formua.formua_id'), nullable=False)
    op_id = Column(Integer, nullable=False)
    pro_type_name = Column(String(30), nullable=False)
    tag = Column(SmallInteger, nullable=False)

    formua = relationship('LmlgFormua')
    gold = relationship('LmlgGold')


class Metalprice(CNBase):
    __tablename__ = 'metalprice'

    id = Column(SmallInteger, primary_key=True)
    karat = Column(ForeignKey('metalma.karat'), nullable=False, index=True)
    price = Column(SMALLMONEY, nullable=False)
    unit = Column(TINYINT, nullable=False)
    edate = Column(DateTime, nullable=False, index=True)
    tag = Column(TINYINT, nullable=False)

    metalma = relationship('Metalma')


t_mmgd = Table(
    'mmgd', metadata,
    Column('mmid', ForeignKey('mm.mmid'), nullable=False, index=True),
    Column('karat', SmallInteger, nullable=False),
    Column('wgt', SMALLMONEY, nullable=False),
    Index('pk_mmgd', 'mmid', 'karat', unique=True)
)


class Mmrep(CNBase):
    __tablename__ = 'mmrep'

    repid = Column(Integer, primary_key=True)
    refid = Column(ForeignKey('mmma.refid'), index=True)
    docno = Column(String(10), nullable=False)
    cstbld = Column(String(10), nullable=False)
    styno = Column(String(10), nullable=False)
    cstname = Column(String(10), nullable=False)
    karat = Column(SmallInteger, nullable=False)
    qty = Column(SMALLMONEY, nullable=False)
    wgt = Column(SMALLMONEY, nullable=False)

    mmma = relationship('Mmma')


class Mmst(CNBase):
    __tablename__ = 'mmst'

    mmid = Column(ForeignKey('mm.mmid'), primary_key=True, nullable=False)
    sttype = Column(TINYINT, primary_key=True, nullable=False)
    qty = Column(Integer, nullable=False)
    wgt = Column(DECIMAL(8, 3), nullable=False)

    mm = relationship('Mm')


t_outsourcing = Table(
    'outsourcing', metadata,
    Column('id', Integer, nullable=False, unique=True),
    Column('styid', ForeignKey('styma.styid'), nullable=False),
    Column('cstid', ForeignKey('cstinfo.cstid'), nullable=False),
    Column('venderprodno', String(20), nullable=False),
    Column('remark', String(255), nullable=False),
    Column('fill_date', DateTime, nullable=False),
    Column('qty', Integer, nullable=False),
    Column('wgt', DECIMAL(8, 2), nullable=False),
    Column('tag', SmallInteger, nullable=False),
    Column('userid', SmallInteger)
)


class PrPaymentInfo(CNBase):
    __tablename__ = 'prPaymentInfo'
    __table_args__ = (
        Index('idx_pdI_pId_wId', 'pId', 'wkId', unique=True),
    )

    id = Column(Numeric(7, 0), primary_key=True)
    pId = Column(ForeignKey('prPeriodInfo.id'), nullable=False)
    wkId = Column(Integer, nullable=False)
    sTime = Column(CHAR(254), nullable=False)
    sPay = Column(CHAR(254), nullable=False)
    base = Column(Float, nullable=False)
    bonus = Column(Float, nullable=False)
    lastModified = Column(DateTime, nullable=False)

    prPeriodInfo = relationship('PrPeriodInfo')


class Repforjo(CNBase):
    __tablename__ = 'repforjo'

    repid = Column(Integer, primary_key=True)
    styid = Column(ForeignKey('styma.styid'), nullable=False)
    pyno = Column(String(20), nullable=False)
    repno = Column(String(20), nullable=False)
    running = Column(String(20), nullable=False)

    styma = relationship('Styma')


class Sbtch(CNBase):
    __tablename__ = 'sbtch'

    sbtchid = Column(Integer, primary_key=True)
    pkid = Column(Integer, nullable=False)
    stsizeid = Column(ForeignKey('stone_sizemaster.stsizeid'), nullable=False)
    grade = Column(TINYINT, nullable=False)
    qty = Column(Integer, nullable=False)
    wgt = Column(Float, nullable=False)
    qtyadj = Column(Integer, nullable=False)
    wgtadj = Column(Float, nullable=False)
    fill_date = Column(DateTime, nullable=False)
    modi_date = Column(DateTime, nullable=False)
    tag = Column(Integer, nullable=False)
    qtytrans = Column(Integer)
    wgttrans = Column(Float)

    stone_sizemaster = relationship('StoneSizemaster')


class SecurityGrouping(CNBase):
    __tablename__ = 'security_groupings'

    grpid = Column(Integer, primary_key=True, nullable=False)
    userid = Column(ForeignKey('security_users.userid'), primary_key=True, nullable=False)

    security_user = relationship('SecurityUser')


class SecurityObject(CNBase):
    __tablename__ = 'security_objects'
    __table_args__ = (
        Index('idx_security_sysobjects', 'appid', 'control', 'window', unique=True),
    )

    objid = Column(Integer, primary_key=True)
    appid = Column(ForeignKey('security_apps.appid'), nullable=False)
    window = Column(String(64), nullable=False)
    control = Column(String(128), nullable=False)
    description = Column(String(254), nullable=False)
    objtype = Column(SmallInteger, nullable=False)
    tag = Column(SmallInteger, nullable=False)

    security_app = relationship('SecurityApp')


class Stbcktran(CNBase):
    __tablename__ = 'stbcktran'

    tid = Column(ForeignKey('sttobck.tid'), ForeignKey('sttobck.tid'), primary_key=True, nullable=False)
    idx = Column(TINYINT, primary_key=True, nullable=False)
    refno = Column(String(10), nullable=False)
    qty = Column(Integer, nullable=False)
    wgt = Column(Float, nullable=False)
    filldate = Column(DateTime, nullable=False)

    sttobck = relationship('Sttobck', primaryjoin='Stbcktran.tid == Sttobck.tid')
    sttobck1 = relationship('Sttobck', primaryjoin='Stbcktran.tid == Sttobck.tid')


class Stbtchloca(CNBase):
    __tablename__ = 'stbtchloca'

    id = Column(Integer, primary_key=True)
    btchid = Column(Integer, nullable=False)
    wk_id = Column(ForeignKey('employee.wk_id'), nullable=False)
    fill_date = Column(DateTime, nullable=False)
    lastuserid = Column(ForeignKey('security_users.userid'), nullable=False)

    security_user = relationship('SecurityUser')
    wk = relationship('Employee')


class Stinv(CNBase):
    __tablename__ = 'stinv'

    cstid = Column(ForeignKey('cstinfo.cstid'), primary_key=True, nullable=False)
    invid = Column(Integer, primary_key=True, nullable=False)
    idx = Column(SmallInteger, primary_key=True, nullable=False)
    classid = Column(SmallInteger, nullable=False)
    typeid = Column(SmallInteger, nullable=False)
    shapeid = Column(SmallInteger, nullable=False)
    size = Column(String(10), nullable=False)
    inqty = Column(Integer, nullable=False)
    inwgt = Column(Float, nullable=False)
    price = Column(Float, nullable=False)
    unit = Column(SmallInteger, nullable=False)
    indate = Column(DateTime, nullable=False)
    bckqty = Column(Integer)
    bckwgt = Column(Float)
    bckdate = Column(DateTime)
    usedqty = Column(Integer)
    usedwgt = Column(Float)
    payment = Column(Float)
    tag = Column(SmallInteger)

    cstinfo = relationship('Cstinfo')


class StoneGetbck(CNBase):
    __tablename__ = 'stone_getbck'

    id = Column(ForeignKey('stone_sndout.id'), primary_key=True, nullable=False)
    idx = Column(TINYINT, primary_key=True, nullable=False)
    qty = Column(Integer, nullable=False)
    wgt = Column(Float, nullable=False)
    cust_bill_id = Column(CHAR(6), nullable=False)
    checker_id = Column(SmallInteger, nullable=False)
    fill_date = Column(DateTime, nullable=False)

    stone_sndout = relationship('StoneSndout')


class StonePkma(CNBase):
    __tablename__ = 'stone_pkma'

    pkid = Column(Integer, primary_key=True)
    package_id = Column(String(20), nullable=False, unique=True)
    unit = Column(SmallInteger, nullable=False)
    price = Column(CHAR(6), nullable=False)
    stid = Column(ForeignKey('stone_master.stid'))
    stshpid = Column(SmallInteger)
    pricen = Column(SMALLMONEY)
    tag = Column(TINYINT)
    fill_date = Column(DateTime)
    relpkid = Column(Integer)
    stsizeidf = Column(SmallInteger)
    stsizeidt = Column(SmallInteger)
    lastuserid = Column(SmallInteger, server_default=text("0"))
    lastupdate = Column(DateTime, server_default=text("getdate()"))
    wgtunit = Column(Float, server_default=text("0.2"))
    color = Column(String(50), server_default=text("N/A"))

    stone_master = relationship('StoneMaster')


class Stoutma(CNBase):
    __tablename__ = 'stoutma'
    __table_args__ = (
        Index('idx_stoutma_pk', 'bill_id', 'is_out', unique=True),
    )

    id = Column(Integer, primary_key=True)
    bill_id = Column(Integer, nullable=False)
    is_out = Column(SmallInteger, nullable=False)
    jsid = Column(Integer, nullable=False)
    fill_date = Column(DateTime, nullable=False)
    checkid = Column(ForeignKey('stoutchkinfo.checkid'), nullable=False, index=True)
    create_date = Column(DateTime, nullable=False)
    subcnt = Column(TINYINT, nullable=False)
    orderid = Column(SmallInteger)

    stoutchkinfo = relationship('Stoutchkinfo')


class Stsumdtl(CNBase):
    __tablename__ = 'stsumdtl'

    id = Column(Integer, primary_key=True)
    refno = Column(String(10), nullable=False)
    dptid = Column(TINYINT, nullable=False)
    jsid = Column(Integer, nullable=False)
    stsumid = Column(ForeignKey('stsumma.stsumid'), nullable=False)
    qty = Column(Integer, nullable=False)
    wgt = Column(DECIMAL(10, 3), nullable=False)
    fill_date = Column(DateTime, nullable=False)

    stsumma = relationship('Stsumma')


class Stymaext(CNBase):
    __tablename__ = 'stymaext'

    styextid = Column(Integer, primary_key=True)
    styid = Column(ForeignKey('styma.styid'), nullable=False, index=True)
    styno = Column(String(10), nullable=False, index=True)
    properties = Column(String(100), nullable=False)

    styma = relationship('Styma')


class Stypropma(CNBase):
    __tablename__ = 'stypropma'

    propid = Column(Integer, primary_key=True)
    propdefid = Column(ForeignKey('stypropdef.propdefid'), nullable=False, index=True)
    codec = Column(CHAR(20), nullable=False)
    coden = Column(DECIMAL(10, 3), nullable=False)
    fill_date = Column(DateTime, nullable=False)

    stypropdef = relationship('Stypropdef')


class Utilityin(CNBase):
    __tablename__ = 'utilityin'

    uiid = Column(Integer, primary_key=True)
    uimaid = Column(Integer, nullable=False)
    uid = Column(ForeignKey('utilityma.uid'), nullable=False)
    units = Column(Float, nullable=False)
    lastuserid = Column(SmallInteger, nullable=False)

    utilityma = relationship('Utilityma')


class Utilityout(CNBase):
    __tablename__ = 'utilityout'

    uoid = Column(Integer, primary_key=True)
    uid = Column(ForeignKey('utilityma.uid'), nullable=False)
    units = Column(Float, nullable=False)
    worker_id = Column(SmallInteger, nullable=False)
    fill_date = Column(DateTime, nullable=False)
    lastuserid = Column(SmallInteger, nullable=False)
    lastupdate = Column(DateTime, nullable=False)

    utilityma = relationship('Utilityma')


class Watchobjma(CNBase):
    __tablename__ = 'watchobjma'

    wcobjid = Column(Integer, primary_key=True)
    wcobjtypeid = Column(ForeignKey('watchobjtype.wcobjtypeid'), nullable=False)
    wcobjno = Column(String(30), nullable=False)
    size = Column(String(30), nullable=False)
    specs = Column(String(30), nullable=False)
    remark1 = Column(String(50), nullable=False)
    remark2 = Column(String(50), nullable=False)
    remark3 = Column(String(50), nullable=False)
    userid = Column(Integer, nullable=False)
    createdate = Column(DateTime, nullable=False)
    tag = Column(SmallInteger, nullable=False)

    watchobjtype = relationship('Watchobjtype')


class BCstbldLabcost(CNBase):
    __tablename__ = 'b_cstbld_labcost'

    jsid = Column(ForeignKey('b_cust_bill.jsid'), primary_key=True, nullable=False)
    costid = Column(TINYINT, primary_key=True, nullable=False)
    cost = Column(SMALLMONEY, nullable=False)

    b_cust_bill = relationship('BCustBill')


class BCstbldMit(CNBase):
    __tablename__ = 'b_cstbld_mit'

    jsid = Column(ForeignKey('b_cust_bill.jsid'), primary_key=True, nullable=False)
    mitid = Column(ForeignKey('mitma.mitid'), primary_key=True, nullable=False)
    wgt = Column(SMALLMONEY, nullable=False)

    b_cust_bill = relationship('BCustBill')
    mitma = relationship('Mitma')


class BJobSheet(CNBase):
    __tablename__ = 'b_job_sheet'

    jsid = Column(ForeignKey('b_cust_bill.jsid'), nullable=False)
    job_sheet_id = Column(Integer, primary_key=True)
    quantity = Column(SMALLMONEY, nullable=False)
    recent_dept = Column(CHAR(8), nullable=False, index=True)
    dept_date = Column(DateTime, nullable=False)
    tag = Column(TINYINT)

    b_cust_bill = relationship('BCustBill')


class Cstdtl(CNBase):
    __tablename__ = 'cstdtl'

    jsid = Column(ForeignKey('b_cust_bill.jsid'), primary_key=True, nullable=False)
    sbtchid = Column(ForeignKey('sbtch.sbtchid'), primary_key=True, nullable=False)
    qty = Column(SmallInteger, nullable=False)
    wgt = Column(Float, nullable=False)
    actuqty = Column(SMALLMONEY)
    cutqty = Column(SMALLMONEY)
    tag = Column(TINYINT)
    marked = Column(TINYINT)
    settype = Column(TINYINT, server_default=text("0"))

    b_cust_bill = relationship('BCustBill')
    sbtch = relationship('Sbtch')


class Dsstodtl(CNBase):
    __tablename__ = 'dsstodtl'

    id = Column(Integer, primary_key=True)
    createdate = Column(DateTime, nullable=False, server_default=text("getdate()"))
    lastuserid = Column(SmallInteger, nullable=False)
    lastupdate = Column(DateTime, nullable=False, server_default=text("getdate()"))
    tag = Column(SmallInteger, nullable=False, server_default=text("0"))
    stoid = Column(ForeignKey('dssto.id'), nullable=False, index=True)
    stockid = Column(ForeignKey('dsstock.id'), nullable=False, server_default=text("0"))
    qty = Column(SMALLMONEY, nullable=False)
    qtyused = Column(SMALLMONEY, nullable=False, server_default=text("0"))
    qtyback = Column(SMALLMONEY, nullable=False, server_default=text("0"))
    unitwgt = Column(TINYINT, nullable=False)
    unitmwgt = Column(TINYINT, nullable=False)
    factorycost = Column(TINYINT, nullable=False, server_default=text("0"))
    goldprice = Column(TINYINT, nullable=False, server_default=text("0"))
    upcid = Column(ForeignKey('dsupc.id'), nullable=False, server_default=text("-1"))
    location = Column(String(50), nullable=False, server_default=text("N/A"))
    size = Column(String(10), nullable=False, server_default=text(""))
    clarity = Column(String(20), nullable=False, server_default=text(""))
    color = Column(String(20), nullable=False, server_default=text(""))
    label = Column(String(50), nullable=False, server_default=text(""))
    remark = Column(String(50), nullable=False, server_default=text(""))

    dsstock = relationship('Dsstock')
    dssto = relationship('Dssto')
    dsupc = relationship('Dsupc')


class Expstuff(CNBase):
    __tablename__ = 'expstuff'

    esid = Column(Integer, primary_key=True)
    expid = Column(ForeignKey('expdtl.expid'), ForeignKey('expdtl.expid'))
    type = Column(Integer)
    value = Column(DECIMAL(8, 2))
    tag = Column(SmallInteger)

    expdtl = relationship('Expdtl', primaryjoin='Expstuff.expid == Expdtl.expid')
    expdtl1 = relationship('Expdtl', primaryjoin='Expstuff.expid == Expdtl.expid')


class FsSt(CNBase):
    __tablename__ = 'fs_st'

    id = Column(ForeignKey('fs_sndout.id'), primary_key=True, nullable=False)
    idx = Column(TINYINT, primary_key=True, nullable=False)
    stid = Column(TINYINT, nullable=False)
    qty = Column(SmallInteger, nullable=False)
    wgt = Column(DECIMAL(7, 3), nullable=False)
    tag = Column(TINYINT, nullable=False)

    fs_sndout = relationship('FsSndout')


class Goldplating(CNBase):
    __tablename__ = 'goldplating'
    __table_args__ = (
        Index('idx_gp_key', 'jsid', 'vendor', 'tag', 'height', 'uprice', 'remark', unique=True),
    )

    id = Column(Integer, primary_key=True)
    jsid = Column(ForeignKey('b_cust_bill.jsid'), nullable=False, index=True)
    vendor = Column(Integer, nullable=False, server_default=text("1"))
    height = Column(String(30), nullable=False)
    uprice = Column(Float, nullable=False, server_default=text("0"))
    qty = Column(Float, nullable=False, server_default=text("0"))
    remark = Column(String(50), nullable=False, server_default=text(" "))
    fill_date = Column(DateTime, nullable=False, server_default=text("getdate()"))
    tag = Column(SmallInteger, nullable=False, server_default=text("0"))

    b_cust_bill = relationship('BCustBill')


class Jobypc(CNBase):
    __tablename__ = 'jobypc'

    id = Column(Integer, primary_key=True)
    jsid = Column(ForeignKey('b_cust_bill.jsid'), nullable=False)
    pid = Column(SmallInteger, nullable=False)
    uprice = Column(SMALLMONEY, nullable=False)
    fill_date = Column(DateTime, nullable=False)
    userid = Column(SmallInteger, nullable=False)
    lastupdate = Column(DateTime, nullable=False)
    refid = Column(Integer)

    b_cust_bill = relationship('BCustBill')


class Jocost(CNBase):
    __tablename__ = 'jocost'

    jocostid = Column(Integer, primary_key=True)
    jsid = Column(ForeignKey('b_cust_bill.jsid'), nullable=False, index=True)
    costid = Column(SmallInteger, nullable=False)
    cost = Column(MONEY, nullable=False)

    b_cust_bill = relationship('BCustBill')


class Joprop(CNBase):
    __tablename__ = 'joprop'

    jsid = Column(ForeignKey('b_cust_bill.jsid'), primary_key=True, nullable=False)
    sample = Column(TINYINT, primary_key=True, nullable=False)

    b_cust_bill = relationship('BCustBill')


class Jostset(CNBase):
    __tablename__ = 'jostset'

    jsid = Column(ForeignKey('b_cust_bill.jsid'), primary_key=True, nullable=False)
    settype = Column(TINYINT, primary_key=True, nullable=False)
    qty = Column(Integer, nullable=False)
    fqty = Column(Integer, nullable=False)

    b_cust_bill = relationship('BCustBill')


class LmlgHistory(CNBase):
    __tablename__ = 'lmlg_history'

    his_id = Column(Integer, primary_key=True)
    pro_type_id = Column(ForeignKey('lmlg_process_type.pro_type_id'), nullable=False)
    inweigh = Column(Numeric(18, 2), nullable=False)
    shweigh = Column(Numeric(18, 2), nullable=False)
    residue = Column(Numeric(18, 2), nullable=False)
    sdate = Column(DateTime, nullable=False)
    rdate = Column(DateTime, nullable=False)
    op_id = Column(ForeignKey('lmlg_operation.op_id'), nullable=False)
    lost_weigh = Column(Numeric(18, 2), nullable=False)

    op = relationship('LmlgOperation')
    pro_type = relationship('LmlgProcessType')


class LmlgLibrary(CNBase):
    __tablename__ = 'lmlg_library'

    lib_id = Column(Integer, primary_key=True)
    pro_type_id = Column(ForeignKey('lmlg_process_type.pro_type_id'), nullable=False)
    inweigh = Column(Numeric(18, 2), nullable=False)
    shuweigh = Column(Numeric(18, 2), nullable=False)
    noweigh = Column(Numeric(18, 2), nullable=False)
    indate = Column(DateTime, nullable=False)
    tag = Column(SmallInteger, nullable=False)
    op_id = Column(ForeignKey('lmlg_operation.op_id'), nullable=False)
    pro_id = Column(Integer, nullable=False)
    pro_type_id2 = Column(Integer, nullable=False)

    op = relationship('LmlgOperation')
    pro_type = relationship('LmlgProcessType')


class LmlgProces(CNBase):
    __tablename__ = 'lmlg_process'

    pro_id = Column(Integer, primary_key=True)
    dep_id = Column(ForeignKey('lmlg_department.dep_id'), nullable=False)
    emp_id = Column(ForeignKey('lmlg_employee.emp_id'), nullable=False)
    pro_type_id = Column(Integer, nullable=False)
    op_id = Column(ForeignKey('lmlg_operation.op_id'), nullable=False)
    bdate = Column(DateTime, nullable=False)
    rdate = Column(DateTime, nullable=False)
    bweigh = Column(Numeric(18, 2), nullable=False)
    rweigh = Column(Numeric(18, 2), nullable=False)
    lweigh = Column(Numeric(18, 2), nullable=False)
    highgold = Column(Numeric(18, 2), nullable=False)
    patchgold = Column(Numeric(18, 2), nullable=False)
    pro_type_id2 = Column(Integer, nullable=False)
    tag = Column(SmallInteger, nullable=False)

    dep = relationship('LmlgDepartment')
    emp = relationship('LmlgEmployee')
    op = relationship('LmlgOperation')


class LmlgRecieve(CNBase):
    __tablename__ = 'lmlg_recieve'

    rec_id = Column(Integer, primary_key=True)
    pro_type_id = Column(ForeignKey('lmlg_process_type.pro_type_id'), nullable=False)
    lib_id = Column(Integer, nullable=False)
    rec_out = Column(Numeric(18, 2), nullable=False)
    rec_in = Column(Numeric(18, 2), nullable=False)
    rec_clase = Column(Numeric(18, 2), nullable=False)
    rec_balance = Column(Numeric(18, 2), nullable=False)
    rec_date = Column(DateTime, nullable=False)
    tag = Column(SmallInteger, nullable=False)

    pro_type = relationship('LmlgProcessType')


class LmyStoneWktopk(CNBase):
    __tablename__ = 'lmy_stone_wktopk'
    __table_args__ = (
        Index('index_pk_id', 'wk_id', 'pk_id', unique=True),
    )

    wkpkid = Column(Integer, primary_key=True)
    wk_id = Column(ForeignKey('employee.wk_id'), nullable=False)
    pk_id = Column(ForeignKey('stone_pkma.pkid'), nullable=False)
    checker_id = Column(Integer)
    qty = Column(Integer)
    wgt = Column(Float)
    sum_qty = Column(Integer)
    sum_wgt = Column(Float)
    tag = Column(SmallInteger)
    status = Column(SmallInteger)
    receivc_date = Column(DateTime)
    check_date = Column(DateTime)

    pk = relationship('StonePkma')
    wk = relationship('Employee')


class Sbtchadj(CNBase):
    __tablename__ = 'sbtchadj'

    id = Column(Integer, primary_key=True)
    sbtchid = Column(ForeignKey('sbtch.sbtchid'), nullable=False)
    qty = Column(Integer, nullable=False)
    wgt = Column(Float, nullable=False)
    fill_date = Column(DateTime, nullable=False)
    tag = Column(TINYINT, nullable=False)

    sbtch = relationship('Sbtch')


class Sbtchin(CNBase):
    __tablename__ = 'sbtchin'

    id = Column(Integer, primary_key=True)
    btchid = Column(Integer, nullable=False)
    sbtchid = Column(ForeignKey('sbtch.sbtchid'), nullable=False)
    qty = Column(Integer, nullable=False)
    wgt = Column(Float, nullable=False)
    fill_date = Column(DateTime, nullable=False)

    sbtch = relationship('Sbtch')


class SecurityInfo(CNBase):
    __tablename__ = 'security_info'

    objid = Column(ForeignKey('security_objects.objid'), primary_key=True, nullable=False)
    userid = Column(ForeignKey('security_users.userid'), primary_key=True, nullable=False)
    status = Column(SmallInteger, nullable=False)
    tag = Column(SmallInteger, nullable=False)

    security_object = relationship('SecurityObject')
    security_user = relationship('SecurityUser')


t_st_dpt_bld = Table(
    'st_dpt_bld', metadata,
    Column('dept_bill_id', CHAR(8), nullable=False),
    Column('idx', SmallInteger, nullable=False),
    Column('jsid', ForeignKey('b_cust_bill.jsid'), nullable=False),
    Column('st_type', String(15)),
    Column('quantity', SmallInteger, nullable=False),
    Column('fill_date', DateTime, nullable=False),
    Column('settype', SmallInteger, nullable=False),
    Index('idx_st_dpt_bld_pk', 'jsid', 'idx', unique=True)
)


class StoneIn(CNBase):
    __tablename__ = 'stone_in'

    pkid = Column(ForeignKey('stone_pkma.pkid'), nullable=False)
    btchid = Column(Integer, primary_key=True)
    batch_id = Column(String(15), nullable=False, unique=True)
    bill_id = Column(String(15), nullable=False)
    quantity = Column(Integer, nullable=False)
    weight = Column(Float, nullable=False)
    wgt_used = Column(Float, nullable=False, server_default=text("0"))
    wgt_bck = Column(Float, nullable=False, server_default=text("0"))
    size = Column(String(60), nullable=False)
    fill_date = Column(DateTime, nullable=False)
    is_used_up = Column(SmallInteger, nullable=False, index=True)
    adjust_wgt = Column(Float, nullable=False, server_default=text("0"))
    cstid = Column(ForeignKey('cstinfo.cstid'), nullable=False)
    qty_used = Column(Integer, nullable=False, server_default=text("0"))
    cstref = Column(String(60), nullable=False)
    qtytrans = Column(Integer, nullable=False, server_default=text("0"))
    wgttrans = Column(Float, nullable=False, server_default=text("0"))
    wgt_tmptrans = Column(Float, nullable=False, server_default=text("0"))
    wgt_prepared = Column(Float, nullable=False, server_default=text("0"))
    lastuserid = Column(SmallInteger, server_default=text("0"))
    lastupdate = Column(DateTime, server_default=text("getdate()"))
    qty_bck = Column(Integer, server_default=text("0"))

    cstinfo = relationship('Cstinfo')
    stone_pkma = relationship('StonePkma')


class Stbtchlocama(StoneIn):
    __tablename__ = 'stbtchlocama'

    btchid = Column(ForeignKey('stone_in.btchid'), primary_key=True)
    wk_id = Column(ForeignKey('employee.wk_id'))
    fill_date = Column(DateTime, nullable=False)
    lastuserid = Column(ForeignKey('security_users.userid'), nullable=False)
    lastupdate = Column(DateTime, nullable=False)

    security_user = relationship('SecurityUser')
    wk = relationship('Employee')


class StoneOutMaster(CNBase):
    __tablename__ = 'stone_out_master'
    __table_args__ = (
        Index('idx_stom_key', 'bill_id', 'is_out', unique=True),
    )

    id = Column(Integer, primary_key=True)
    bill_id = Column(Integer, nullable=False)
    is_out = Column(SmallInteger, nullable=False)
    jsid = Column(ForeignKey('b_cust_bill.jsid'), nullable=False)
    qty = Column(SMALLMONEY, nullable=False)
    fill_date = Column(DateTime, nullable=False)
    packed = Column(TINYINT, nullable=False, index=True)
    subcnt = Column(TINYINT, nullable=False)
    worker_id = Column(SmallInteger, server_default=text("0"))
    lastuserid = Column(SmallInteger, server_default=text("0"))
    lastupdate = Column(DateTime, server_default=text("getdate()"))

    b_cust_bill = relationship('BCustBill')


class Stout(CNBase):
    __tablename__ = 'stout'

    id = Column(ForeignKey('stoutma.id'), primary_key=True, nullable=False, index=True)
    idx = Column(TINYINT, primary_key=True, nullable=False)
    btchid = Column(Integer, nullable=False)
    qty = Column(Integer, nullable=False)
    wgt = Column(Float, nullable=False)
    check_date = Column(DateTime, nullable=False)
    pageid = Column(TINYINT, nullable=False)
    rowid = Column(TINYINT, nullable=False)

    stoutma = relationship('Stoutma')


class Styprop(CNBase):
    __tablename__ = 'styprop'

    styextid = Column(ForeignKey('stymaext.styextid'), primary_key=True, nullable=False)
    propid = Column(Integer, primary_key=True, nullable=False)

    stymaext = relationship('Stymaext')


class Warestock(CNBase):
    __tablename__ = 'warestock'

    wfstockid = Column(Integer, primary_key=True)
    wcobjid = Column(ForeignKey('watchobjma.wcobjid'), nullable=False)
    inqty = Column(Integer, nullable=False)
    outqty = Column(Integer, nullable=False)
    machqty = Column(Integer, nullable=False)
    qty = Column(Integer, nullable=False)
    tag = Column(SmallInteger, nullable=False)

    watchobjma = relationship('Watchobjma')


class Watchspec(CNBase):
    __tablename__ = 'watchspecs'

    wcspid = Column(Integer, primary_key=True)
    wcobjid = Column(ForeignKey('watchobjma.wcobjid'), nullable=False)
    jsid = Column(Integer, nullable=False)
    carapace = Column(String(30), nullable=False)
    machqty = Column(Integer, nullable=False)
    outqty = Column(Integer, nullable=False)
    userid = Column(Integer, nullable=False)
    lastuserid = Column(Integer, nullable=False)
    createdate = Column(DateTime, nullable=False)
    status = Column(SmallInteger, nullable=False)
    tag = Column(SmallInteger, nullable=False)

    watchobjma = relationship('Watchobjma')


t_b_dept_bill = Table(
    'b_dept_bill', metadata,
    Column('dept_bill_id', CHAR(8), nullable=False),
    Column('job_sheet_id', ForeignKey('b_job_sheet.job_sheet_id'), nullable=False),
    Column('dept_date', DateTime, nullable=False),
    Index('pk_b_dept_bill', 'dept_bill_id', 'job_sheet_id', unique=True)
)


class Dsmp(CNBase):
    __tablename__ = 'dsmps'

    id = Column(Integer, primary_key=True)
    createdate = Column(DateTime, nullable=False, server_default=text("getdate()"))
    lastuserid = Column(SmallInteger, nullable=False)
    lastupdate = Column(DateTime, nullable=False, server_default=text("getdate()"))
    tag = Column(SmallInteger, nullable=False, server_default=text("0"))
    orderid = Column(ForeignKey('dsorder.id'), nullable=False)
    stodtlid = Column(ForeignKey('dsstodtl.id'), nullable=False)
    cstid = Column(ForeignKey('dscstinfo.id'), nullable=False)
    qty = Column(SMALLMONEY, nullable=False)
    qtyback = Column(SMALLMONEY, nullable=False, server_default=text("0"))
    goldprice = Column(TINYINT, nullable=False, server_default=text("0"))
    unitprice = Column(TINYINT, nullable=False, server_default=text("0"))
    upcid = Column(ForeignKey('dsupc.id'), nullable=False)

    dscstinfo = relationship('Dscstinfo')
    dsorder = relationship('Dsorder')
    dsstodtl = relationship('Dsstodtl')
    dsupc = relationship('Dsupc')


class Jomit(CNBase):
    __tablename__ = 'jomit'

    jocostid = Column(ForeignKey('jocost.jocostid'), primary_key=True, nullable=False)
    mitid = Column(ForeignKey('mitma.mitid'), primary_key=True, nullable=False)
    qty = Column(MONEY, nullable=False)
    wgt = Column(MONEY, nullable=False)
    uprice = Column(Float, server_default=text("0 null"))

    jocost = relationship('Jocost')
    mitma = relationship('Mitma')


class Josettingcost(CNBase):
    __tablename__ = 'josettingcost'

    jocostid = Column(ForeignKey('jocost.jocostid'), primary_key=True, nullable=False)
    stid = Column(TINYINT, primary_key=True, nullable=False)
    setid = Column(TINYINT, primary_key=True, nullable=False)
    uprice = Column(SMALLMONEY, nullable=False)
    qty = Column(Integer, nullable=False)
    fill_date = Column(DateTime, nullable=False)
    tag = Column(TINYINT, nullable=False)

    jocost = relationship('Jocost')


class LmyStone(CNBase):
    __tablename__ = 'lmy_stone'

    id = Column(Integer, primary_key=True)
    wkpkid = Column(Integer, nullable=False)
    btchid = Column(ForeignKey('stone_in.btchid'), nullable=False)
    qty = Column(DECIMAL(8, 3), nullable=False)
    wgt = Column(DECIMAL(8, 3), nullable=False)
    date = Column(DateTime, nullable=False)
    status = Column(SmallInteger, nullable=False)
    tag = Column(SmallInteger, nullable=False)
    jsid = Column(ForeignKey('b_cust_bill.jsid'))
    docno = Column(CHAR(10), server_default=text("\"000000\""))
    remark = Column(String(20), nullable=False, server_default=text("\"\""))

    stone_in = relationship('StoneIn')
    b_cust_bill = relationship('BCustBill')


class StSize(CNBase):
    __tablename__ = 'st_size'

    btchid = Column(ForeignKey('stone_in.btchid'), primary_key=True, nullable=False)
    size = Column(String(6), primary_key=True, nullable=False)
    loca = Column(TINYINT, primary_key=True, nullable=False)
    wgt = Column(Float, nullable=False)
    fill_date = Column(DateTime, primary_key=True, nullable=False)

    stone_in = relationship('StoneIn')


class StoneBck(CNBase):
    __tablename__ = 'stone_bck'

    btchid = Column(ForeignKey('stone_in.btchid'), primary_key=True, nullable=False)
    idx = Column(SmallInteger, primary_key=True, nullable=False)
    weight = Column(Float, nullable=False)
    bill_id = Column(String(8), nullable=False)
    fill_date = Column(DateTime, nullable=False)
    lastuserid = Column(SmallInteger, server_default=text("0"))
    lastupdate = Column(DateTime, server_default=text("getdate()"))
    quantity = Column(Integer, server_default=text("0"))

    stone_in = relationship('StoneIn')


class StoneOut(CNBase):
    __tablename__ = 'stone_out'

    id = Column(Integer, primary_key=True, nullable=False)
    idx = Column(TINYINT, primary_key=True, nullable=False)
    btchid = Column(ForeignKey('stone_in.btchid'), nullable=False, index=True)
    worker_id = Column(SmallInteger, nullable=False)
    quantity = Column(Integer, nullable=False)
    weight = Column(Float, nullable=False)
    checker_id = Column(SmallInteger, nullable=False)
    check_date = Column(DateTime, nullable=False)
    qty = Column(SmallInteger)
    printid = Column(Integer)
    lastuserid = Column(SmallInteger, server_default=text("0"))
    lastupdate = Column(DateTime, server_default=text("getdate()"))

    stone_in = relationship('StoneIn')


class StoneTransfer(CNBase):
    __tablename__ = 'stone_transfer'
    __table_args__ = (
        Index('idx_stone_transfer_pk', 'btchidf', 'btchidt', 'reltable'),
    )

    tid = Column(Integer, primary_key=True)
    docno = Column(String(20), nullable=False)
    btchidf = Column(ForeignKey('stone_in.btchid'), nullable=False)
    btchidt = Column(ForeignKey('stone_in.btchid'), nullable=False)
    fill_date = Column(DateTime, nullable=False)
    qty = Column(Integer, nullable=False)
    wgt = Column(Float, nullable=False)
    reltable = Column(TINYINT, nullable=False)
    tag = Column(TINYINT)
    lastuserid = Column(SmallInteger, server_default=text("0"))
    lastupdate = Column(DateTime, server_default=text("getdate()"))

    stone_in = relationship('StoneIn', primaryjoin='StoneTransfer.btchidf == StoneIn.btchid')
    stone_in1 = relationship('StoneIn', primaryjoin='StoneTransfer.btchidt == StoneIn.btchid')


class Stoutmalst(CNBase):
    __tablename__ = 'stoutmalst'

    id = Column(ForeignKey('stone_out_master.id'), primary_key=True, nullable=False)
    lstno = Column(String(20), primary_key=True, nullable=False)
    fill_date = Column(DateTime, nullable=False)
    lastuserid = Column(SmallInteger, nullable=False)
    lastupdate = Column(DateTime, nullable=False)

    stone_out_master = relationship('StoneOutMaster')


class Watchinvoic(CNBase):
    __tablename__ = 'watchinvoic'

    invid = Column(Integer, primary_key=True)
    wfstockid = Column(ForeignKey('warestock.wfstockid'), nullable=False)
    invmaid = Column(Integer, nullable=False)
    qty = Column(Integer, nullable=False)
    userid = Column(Integer, nullable=False)
    lastuserid = Column(Integer, nullable=False)
    tag = Column(SmallInteger, nullable=False)

    warestock = relationship('Warestock')


class Watchtakeout(CNBase):
    __tablename__ = 'watchtakeout'

    outid = Column(Integer, primary_key=True)
    wcspid = Column(ForeignKey('watchspecs.wcspid'), nullable=False)
    outqty = Column(Integer, nullable=False)
    outdate = Column(DateTime, nullable=False)
    userid = Column(Integer, nullable=False)
    tag = Column(SmallInteger, nullable=False)

    watchspec = relationship('Watchspec')


class Dsstordtl(CNBase):
    __tablename__ = 'dsstordtl'

    id = Column(Integer, primary_key=True)
    createdate = Column(DateTime, nullable=False, server_default=text("getdate()"))
    lastuserid = Column(SmallInteger, nullable=False)
    lastupdate = Column(DateTime, nullable=False, server_default=text("getdate()"))
    tag = Column(SmallInteger, nullable=False, server_default=text("0"))
    storid = Column(ForeignKey('dsstor.id'), nullable=False, index=True)
    mpsid = Column(ForeignKey('dsmps.id'), nullable=False, index=True)
    qty = Column(SMALLMONEY, nullable=False)

    dsmp = relationship('Dsmp')
    dsstor = relationship('Dsstor')
