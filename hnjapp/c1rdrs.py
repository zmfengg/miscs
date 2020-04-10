# coding=utf-8
'''
Created on 2018-04-28
classes to read data from C1's monthly invoices
need to be able to read the 2 kinds of files: C1's original and calculator file
@author: zmFeng
'''

from numbers import Number
from tempfile import gettempdir
from re import compile as cpl
from time import time
from datetime import datetime
from collections import namedtuple
from os import path

from sqlalchemy import and_, func
from sqlalchemy.orm import Query
from xlwings import Sheet, constants

from hnjapp.svcs.db import jesin
from hnjapp.pajcc import WgtInfo, addwgt
from hnjcore import JOElement, karatsvc
from hnjcore.models.cn import (JO, MM, Codetable, Customer, MMgd, MMMa,
                               StoneBck, StoneIn, StoneOut, StoneOutMaster,
                               StonePk, Style)
from hnjcore.utils.consts import NA
from utilz import (NamedList, NamedLists, daterange, getfiles, isnumeric,
                   splitarray, trimu, xwu, tofloat, stsizefmt, list2dict)

from .common import _date_short, _logger as logger, config, _getdefkarat

_ptnbtno = cpl(r"(\d+)([A-Z]{3})(\d+)")

''' items for c1inv and c1stone '''
C1InvItem = namedtuple("C1InvItem", "source,jono,qty,labor,setting,remarks,stones,mtlwgt,styno")
C1InvStone = namedtuple("C1InvStone", "shape,stone,qty,wgt,remark")
def _nf(part, cnt):
    try:
        i = int(part)
        part = ("%%0%sd" % ("%d" % cnt)) % i
    except:
        pass
    return part


def _fmtbtno(btno):
    """ in C1's STIO excel file, the batch# is quite malformed
    this method format it to standard ones
    """
    if not btno:
        return None
    flag = False
    if isinstance(btno, Number):
        btno = "%08d" % int(btno)
    else:
        if isinstance(btno, datetime):
            btno = btno.strftime("%d-%b-%y")
            flag = True
        btno = btno.replace("‘", "")
        if btno.find("-") > 0:
            cnts = (2, 2, 3, 2, 2, 2)
            ss, pts = btno.split("-"), []
            for i in zip(ss, cnts):
                pts.append(_nf(i[0], i[1]))
            btno = trimu("".join(pts))
        else:
            #'10AUG311A01' or alike
            if not (len(btno) > 8 and btno[-3] == 'A'):
                mt = _ptnbtno.search(btno)
                if mt:
                    btno = btno[mt.start(1):mt.end(2)] + ("%03d" % int(mt.group(3)))
    return ("-" if flag else "") + btno


def _fmtpkno(pkno):
    """ in C1's STIO excel file, the package# is quite malformed
    this method format it to standard ones
    """
    if not pkno:
        return None
    #contain invalid character, not a PK#
    pkno = trimu(pkno)
    if sum([1 for x in pkno if ord(x) <= 31 or ord(x) >= 127]) > 0:
        return None
    pkno0 = pkno
    if pkno.find("-") >= 0:
        pkno = pkno.replace("-", "")
    pfx, pkno, sfx, idx = pkno[:3], pkno[3:], "", 0
    for idx in range(len(pkno) - 1, -1, -1):
        ch = pkno[idx]
        if "A" <= ch <= "Z":
            sfx = ch + sfx
        else:
            if sfx:
                idx += 1
                break
            sfx = ch + sfx
    pkno = pkno[:idx]
    if isnumeric(pkno):
        pkno = ("%0" + str(8 - len(pfx) - len(sfx)) + "d") % (int(float(pkno)))
        special = False
    else:
        special = True
        rpm = {"O": "0", "*": "X", "S": "5"}
        for x in rpm.items():
            if pkno.find(x[0]) >= 0:
                logger.debug("PK#(%s)'s %s -> %s in it's numeric part" %
                             (pkno0, x[0], x[1]))
                pkno = pkno.replace(x[0], x[1])
                special = True
    pkno = pfx + pkno + sfx
    return pkno, special


class C1InvRdr(object):
    '''
        read the daily or monthly shipment data out from C1's file
        @param labor_ttl(boolean): read from labor's total(总工费金额) instead of labor + setting. This apply to C3 because their setting fields contains charges like
        stone purchasing and palting
    '''
    def __init__(self, labor_ttl=False):
        self._vdr = 'C1'
        self._st_mp = {x[0]: x[1:] for x in config.get("c1shp.stone.alias")}
        self._labor_ttl = labor_ttl
        self._rdr_km = {
            "styno": "图片,",
            "jono": u"工单,",
            "setting": u"镶工,",
            "labor": u"胚底,",
            "remark": u"备注,",
            "joqty": u"数量,",
            "stname": u"石名,",
            "stqty": "粒数,",
            "stwgt": u"石重,",
            "karat": "成色,",
            "swgt": "净银重,",
            "gwgt": "净金重,",
            "pwgt": "配件重,",
            "netwgt": "货重,",
            "labor_ttl": "总工费,"
        }

    @staticmethod
    def getReader(wb):
        '''
        return a reader(due to 2019/05/06, it's C1InvRdr/C3InvRdr)
        '''
        rdrmap = {
            "诚艺,胤雅": "c1",
            "帝宝": "c3"
        }
        for sht in wb.sheets:
            for x in rdrmap:
                if tuple(y for y in x.split(",") if xwu.find(sht, y)):
                    rdr = rdrmap[x]
                    return C1InvRdr() if rdr == "c1" else C3InvRdr()
        return None

    def read(self, wb, read_method="shipment", pwgt_check=True):
        """
        perform the read action
        @param wb: a Book or Sheet instance or a string referring an excel file
        @read_method: can be one of:
            "shipment": read latest shipment data in the latest sheet
            "first": the first shipment data in latest sheet
            "last": the same as "shipment"
            "c1cc": for the calculator, get all data in latest month
            other: all shipment data in all sheet
        @param pwgt_check:
            True: pwgt will be treated as parts'wgt if is pendant.
            False: pwgt will always be treated as parts' weight of whatever style,  only used for jocost handling
        @return: a tuple of tuple where the inner tuple contains:
            tuple(C1InvItem), invDate
        """

        byme, sht0 = False, None
        if not wb or (isinstance(wb, str) and not path.exists(wb)):
            wb = config.get('default.file.c1aio')
        elif isinstance(wb, Sheet):
            sht0, wb = wb, wb.book
        if isinstance(wb, str):
            byme = True
            if not read_method:
                read_method = "aio"
            app, killxw = xwu.appmgr.acq()
        try:
            if byme:
                wb = app.books.open(wb)
            items = list()
            shts = self._select_sheet(wb)
            if not shts:
                return None
            # given sheet is not the latest one, give up
            if sht0 and sht0 != shts[0]:
                return None
            done = False
            for sht in shts:
                if done:
                    break
                if read_method == 'aio':
                    # inside a aio file, there might not be 日期 field(s), so just skip finding it
                    nl = xwu.find(sht, "合计", find_all=True)
                    if nl:
                        for rng in nl:
                            rng.value = ""
                    rng = xwu.find(sht, "图片")
                    nl = [(sht.cells(1, 1), None, ), ] if rng else []
                else:
                    nl = xwu.find(sht, "日期", find_all=True)
                    if not nl:
                        continue
                    # sometimes the latest one contains nothing except date, an example is 20190221, so read from buttom
                    nl = [(x, xwu.nextcell(x).value.date()) for x in nl]
                    nl = sorted([(x[0], x[1]) for x in nl if x[1]], key=lambda x: x[1], reverse=read_method != 'first')
                for x in nl:
                    rng = xwu.find(sht, "图片", After=x[0])
                    sn = self._read_from(rng, pwgt_check)
                    if not sn:
                        continue
                    if read_method in ('first', 'last', 'shipment'):
                        if read_method == 'shipment':
                            for y in [x for x in shts if x != sht]:
                                y.delete()
                        done = True
                    items.append((sn[0], x[1], ))
                    if done:
                        break
                if items and read_method == "c1cc":
                    break
            if byme:
                wb.close()
                wb = None
        finally:
            if byme:
                if wb:
                    wb.close()
                xwu.appmgr.ret(killxw)
        return items

    @classmethod
    def _select_sheet(cls, wb):
        # mp, cnsc1 = {}, u"工单号 镶工 胚底 备注".split()
        mp, cnsc1 = {}, u"工单号 镶工 胚底".split()
        for sht in wb.sheets:
            sn = sht.name
            nl = sn.find("月")
            if nl <= 0:
                continue
            nl = sn[:nl]
            if nl.isnumeric():
                rngs = [x for x in (xwu.find(sht, var, lookat=constants.LookAt.xlPart) for var in cnsc1) if x]
                if len(cnsc1) != len(rngs):
                    continue
                mp[int(nl)] = sht
        if mp:
            sn = sorted([x for x in mp], reverse=True)
            # Jan contains Dec case
            if len(sn) > 1 and sn[0] - sn[-1] > 10:
                sn[0], sn[-1] = sn[-1], sn[0]
            mp = [mp[x] for x in sn]
        return mp

    def _read_from(self, rng, pwgt_check=True):
        '''
        @param rng: the upper-left of the data region, but in the case of C3, column where rng resident was not merged. how?
        '''
        s0 = rng.sheet
        nl = xwu.find(s0, '合计', After=rng)
        if nl:
            rng = s0.range(rng, s0.cells(nl.row - 1, rng.current_region.last_cell.column))
        nls = [x for x in xwu.NamedRanges(rng, self._rdr_km)]
        if not nls:
            return None
        nl, items = nls[0], "jono gwgt swgt".split()
        kns = [1 for x in items if nl.getcol(x)]
        if len(kns) != len(items):
            logger.debug("sheet(%s), range(%s) does not contain necessary key columns, exps(%s) but is (%s)" % (rng.sheet.name, rng.address, items, kns))
            return None
        # last_act states:
        # 1 -> JO#
        # 2 -> blank but prior is &le; 2
        # 3 -> blank but prior is 2, void
        items, c1, netwgt_c1, jo_lidx, last_act = [], None, {}, {}, 3
        _cnstqnw = "stqty,stwgt".split(",")
        _cnsnl = ("labor_ttl" if self._labor_ttl else "setting,labor").split(",")
        for idx, nl in enumerate(nls):
            s0 = nl.jono
            if s0:
                if s0 == '466167':
                    print('x')
                # print("reading jo(%s)" % s0)
                je = JOElement(s0)
                if je.isvalid():
                    if idx - jo_lidx.get(je.value, idx) > 0:
                        logger.debug("Duplicated JO#(%s) found, the prior is near row(%d)" % (je.value, jo_lidx[je.value]))
                        last_act = 3
                        continue
                    snl = [tofloat(nl.get(s0, 0)) for s0 in _cnsnl]
                    if not any(snl):
                        logger.debug("JO(%s) does not contains any labor cost", je.value)
                        snl = (0, 0)
                    elif self._labor_ttl:
                        snl = [round(snl[0] / nl.joqty, 2), 0]
                    c1 = C1InvItem(self._vdr, je.value, nl.joqty, snl[1], snl[0], nl.remark, [], None, nl.styno)
                    items.append(c1)
                    last_act = 1
                else:
                    # avoid an invalid JO#(for example, 耳迫配件) following a JO will send the weights to prior JO
                    last_act = 3
                    continue
            else:
                if last_act > 2:
                    continue
                last_act = 2
            s0 = self._extract_st_mtl(c1, nl, _cnstqnw, pwgt_check)
            if s0:
                c1 = items[-1] = s0
                if c1.jono not in netwgt_c1:
                    netwgt_c1[c1.jono] = [0, 0]
                netwgt_c1[c1.jono][0] += tofloat(nl['netwgt'] or 0)
                netwgt_c1[c1.jono][1] += tofloat(nl['pwgt'] or 0)
                jo_lidx[c1.jono] = idx
            else:
                last_act = 3
        # now calculate the net weight for each c1
        s0 = []
        for c1 in items:
            # stone wgt in karat
            if not c1.qty:
                continue
            kt = sum((x.wgt for x in c1.stones)) / 5 if c1.stones else 0
            if c1.mtlwgt:
                wgt = sum((x.wgt for x in c1.mtlwgt.wgts if x)) + kt
                s0.append(c1._replace(mtlwgt=c1.mtlwgt._replace(netwgt=round(wgt, 2))))
                nl = (wgt - netwgt_c1[c1.jono][1] / c1.qty, netwgt_c1[c1.jono][0] / c1.qty)
                if abs(nl[0] - nl[1]) > 0.01:
                    logger.debug("JO#(%s)'s netwgt contains error, it should be %4.2f but was %4.2f" % (c1.jono, nl[0], nl[1]))
            else:
                logger.debug("JO(%s) does not contains metal wgt", c1.jono)
        return s0, rng.last_cell.row + len(nls)

    @classmethod
    def verify(cls, c1):
        '''
        verify if c1's weight is valid. a valid c1 weight should fulfill below
        criteria:
            .netwgt == c1.mtlwgt + c1.stwgt. that is, the netwgt of mtlwgt
        JO#463625, violate this rule, but seems ok
        '''
        rc = not bool(c1.netwgt) ^ bool(c1.mtlwgt)
        if rc and c1.netwgt:
            rc = abs(c1.netwgt - c1.mtlwgt.netwgt) < 0.01
        return rc

    @classmethod
    def _is_pendant(cls, styno):
        return styno and styno[:2].upper().find("P") >= 0

    def _new_stone(self, stname, qty, wgt):
        '''
        descent override this for stone parsing
        '''
        snn = self._st_mp.get(stname, ('_R', '_' + stname))
        return C1InvStone(snn[0], snn[1], qty, wgt, None)

    def _extract_st_mtl(self, c1, nl, cnstqnw, pwgt_check=True):
        '''
        extract the stone and metal into c1
        @param pwgt_check:
            True: pwgt will be treated as parts'wgt if is pendant.
            False: pwgt will always be treated as parts' weight of whatever style,  only used for jocost handling
        '''
        #stone data
        if not c1.qty:
            return c1
        hc = 0
        qnw = [tofloat(nl[x]) for x in cnstqnw]
        if all(qnw):
            s0 = nl.stname
            if s0 and isinstance(s0, str):
                joqty = c1.qty
                c1.stones.append(self._new_stone(nl.stname, qnw[0] / joqty, round(qnw[1] / joqty, 4)))
                hc += 1
        #wgt data
        kt, gw, sw, pwgt = nl.karat, nl.gwgt, nl.swgt, nl.pwgt
        if kt and isnumeric(kt):
            joqty = c1.qty
            if not joqty:
                logger.debug("JO(%s) without qty, skipped" % nl.jono)
                return None
            hc += 1
            kt, wgt = self._tokarat(kt), gw or sw
            #only pendant's pwgt is pwgt, else to mainpart
            if pwgt and pwgt_check and not self._is_pendant(c1.styno):
                wgt += pwgt
                pwgt = 0
            mtlwgt = addwgt(c1.mtlwgt, WgtInfo(kt, wgt / joqty, 4)).follows(_getdefkarat(c1.jono))
            c1 = c1._replace(mtlwgt=mtlwgt)
            if pwgt:
                c1 = self._adjust_pwgt(c1, kt, pwgt / joqty)
        return c1 if hc else None

    def _adjust_pwgt(self, c1, kt, pwgt):
        return c1._replace(mtlwgt=addwgt(c1.mtlwgt, WgtInfo(kt, pwgt, 4), True))

    @classmethod
    def _tokarat(cls, kt):
        if kt < 1:
            kt = int(kt * 1000)
        if 924 <= kt <= 935:
            rc = 925
        elif 330 <= kt <= 340:
            rc = 8
        elif 370 <= kt <= 385:
            rc = 9
        elif 410 <= kt <= 420:
            rc = 10
        elif 580 <= kt <= 590:
            rc = 14
        elif 745 <= kt <= 755:
            rc = 18
        else:
            rc = kt
        return rc

class C3InvRdr(C1InvRdr):
    '''
    c3 reader, the only different between them is the stone parser
    '''
    def __init__(self):
        super().__init__(True)
        self._vdr = 'C3'

    def _new_stone(self, stname, qty, wgt):
        '''
        descent override this for stone parsing
        '''
        snn = self._parse_stname(stname)
        return C1InvStone(snn[0], snn[1], qty, wgt, snn[2])

    def _parse_stname(self, stname):
        pfx, sfx = "", None
        for idx, ch in enumerate(stname):
            if '0' <= ch <= '9':
                sfx = stname[idx:]
                break
            pfx += ch
        sfx = stsizefmt(sfx, True) if sfx else ' '
        snn = self._st_mp.get(pfx, ('_R', '_' + pfx))
        return snn[0], snn[1], sfx


class C1JCMkr(object):
    r"""
    C1 JOCost maker, First, Invoke C1STHdlr to create Stone Usage , then generate the jocost report to given folder(default is p:\aa\)
    """

    def __init__(self, cnsvc, bcsvc, invfldr=None, rmbtohk=1.25):
        r"""
        @param cnsvc: the CNSvc instance
        @param dbsvc: the BCSvc instance
        @param invfldr: folder contains all C1's invoices(at least of the given month). A default one inside conf.json:default.file.cx_aio. Can be several files separated by ;
        """
        self._cnsvc = cnsvc
        self._bcsvc = bcsvc
        self._invfldr = invfldr or config.get("default.file.cx_aio")
        self._rmbtohk = rmbtohk or 1.25

    #return refid by running, from existing list or db#
    def _getrefid(self, runn, mpss):
        refid = None
        if mpss:
            for x in mpss:
                if x[0][0] <= runn <= x[0][1]:
                    refid = x[1]
                    break
        if not refid:
            x = self._cnsvc.getjcrefid(runn)
            logger.debug(
                "fetch refid %s from db trigger by running %d" % (x, runn))
            if x:
                mpss.append((x[1], x[0]))
                refid = x[0]
        return refid

    #return mps of given refid #
    def _getmps(self, refid, mpsmp):
        if refid not in mpsmp:
            mp = self._cnsvc.getjcmps(refid)
            logger.debug("using MPS(%s) based on refid(%d)" % (mp, refid))
            mpsmp[refid] = mp
        if refid in mpsmp:
            return mpsmp[refid]

    def get_st_costs(self, runns):
        """
        return the stone costs by map, running as key and cost as value
        """
        return self._cnsvc.getjostcosts(runns)

    def get_st_of_jos(self, runns):
        '''
        return stone usage of given JO
        @param runns: a collection of running(integer)
        '''
        ttl = "jobn,styno,running,package_id,quantity,weight,pricen,unit,is_out,bill_id,fill_date,check_date".split(
            ",")
        lst = []
        with self._cnsvc.sessionctx() as cur:
            q = Query([
                JO.name.label("jono"), JO.deadline,
                Style.name.label("styno"), JO.running,
                StonePk.name.label("pkno"), StoneOut.qty, StoneOut.wgt,
                StonePk.pricen, StonePk.unit, StoneOutMaster.isout,
                StoneOutMaster.name.label("billid"), StoneOutMaster.filldate,
                StoneOut.checkdate
            ]).join(Style).join(StoneOutMaster).join(StoneOut).join(
                StoneIn).join(StonePk)
            for arr in splitarray(runns, 50):
                try:
                    lst1 = q.filter(JO.running.in_(arr)).with_session(cur).all()
                    lst.extend(lst1)
                except:
                    pass
        lst1, lst = lst, [ttl]
        lst.extend(
            [("'" + x.jono.value, x.styno.value, x.running, x.pkno, x.qty,
              round(float(x.wgt), 3), x.pricen, x.unit, x.isout, x.billid,
              x.filldate, x.checkdate) for x in lst1])
        return lst

    def get_st_broken(self, df, dt):
        '''
        return the broken stone in given date range
        @param df: the starting date(included)
        @param dt: the end date(excluded)
        '''
        lst = None
        with self._cnsvc.sessionctx() as cur:
            q = Query([
                JO.name.label("jono"), JO.deadline,
                Style.name.label("styno"), JO.running,
                StonePk.name.label("pkno"), StoneOut.qty, StoneOut.wgt,
                StonePk.pricen, StonePk.unit, StoneOutMaster.isout,
                StoneOutMaster.name.label("billid"), StoneOut.idx,
                StoneOutMaster.filldate, StoneOut.checkdate
            ]).join(Style).join(StoneOutMaster).join(StoneOut).join(
                StoneIn).join(StonePk).filter(
                    and_(StoneOutMaster.filldate >= df,
                         StoneOutMaster.filldate < dt)).filter(
                             and_(StoneOutMaster.isout >= -10,
                                  StoneOutMaster.isout <= 10))
            lst = q.with_session(cur).all()
        if not lst:
            return None
        ttl = "jobn,styno,running,package_id,quantity,weight,pricen,unit,is_out,bill_id,idx,fill_date,check_date".split(
            ",")
        lst1, lst = lst, [ttl]
        lst.extend(
            [("'" + x.jono.value, x.styno.value, x.running, x.pkno, x.qty,
              round(float(x.wgt), 3), x.pricen, x.unit, x.isout, x.billid,
              x.idx, x.filldate, x.checkdate) for x in lst1])
        return lst

    def _read_shp(self, cur, df, dt):
        ptncx = cpl(r"C(\d)$")
        mmids, runns, jcs = set(), set(), {}
        ttls = ("mmid," + config.get("jocost.cost_sheet_fields")).split(",")
        nl = NamedList(list2dict(ttls))
        q = Query([
            JO.name.label("jono"),
            Customer.name.label("cstname"),
            Style.name.label("styno"), JO.running,
            JO.karat.label("jokarat"), MMgd.karat, MM.id,
            MM.name.label("docno"), MM.qty,
            func.sum(MMgd.wgt).label("wgt"),
            func.max(MMMa.refdate).label("refdate")
        ]).join(Customer).join(
            MM).join(MMMa).join(MMgd).join(Style).group_by(
                JO.name, Customer.name, Style.name, JO.running, JO.karat,
                MMgd.karat, MM.id, MM.name, MM.qty).filter(
                    and_(
                        and_(MMMa.refdate >= df, MMMa.refdate < dt),
                        MM.name.like("%C[0-9]")))
        lst = q.with_session(cur).all()
        for x in lst:
            if x.id in mmids:
                continue
            jn = x.jono.value
            mmids.add(x.id)
            if jn not in jcs:
                nl.setdata([0] * len(ttls))
                nl.mmid, nl.lastmmdate, nl.jobno = x.id, "'" + x.refdate.strftime(
                    _date_short), "'" + x.jono.value
                nl.cstname, nl.styno, nl.running = x.cstname.strip(
                ), x.styno.value, x.running
                nl.mstone, nl.description, nl.karat = "_ST", "_EDESC", karatsvc.getfamily(
                    x.jokarat).karat
                nl.goldwgt, nl.cflag, nl.rmb2hk = [], "NA", self._rmbtohk
                mt = ptncx.search(x.docno)
                if mt:
                    nl.cflag = "'" + mt.group(1)
                jcs[jn] = nl.data
                runns.add(int(x.running))
            nl.setdata(jcs[jn])["joqty"] += float(x.qty)
        bcs = self._bcsvc.getbcsforjc(runns)
        bcs = {x.runn: (x.desc, x.ston) for x in bcs} if bcs else {}
        return nl, jcs, runns, bcs

    def _validate(self, **kwds):
        bcs, runns, actname, invs = (kwds.get(x) for x in 'bcs runns actname invs'.split())
        if not bcs or len(bcs) < len(runns):
            logger.debug("%s:Not all records found in BCSystem" % actname)
        stcosts = self.get_st_costs(runns)
        if not stcosts or len(stcosts) < len(runns) / 2:
            logger.debug(
                "%s:No stone or less than 1/2 has stone, make sure you've prepared stone data with C1STIOData"
                % actname)
        if not invs or len(invs) < len(runns):
            logger.debug("%s:No or not enough C1 invoice data from file(%s)"
                            % (actname, self._invfldr))
        return stcosts

    def _calc_dtos(self, gccols, stcosts, **kwds):
        nl = kwds["nl"]
        var = nl.running
        nl.stonecost = stcosts.get(var, None)
        var = kwds["bcs"].get(str(var))
        if var:
            nl.description, nl.mstone = var[:2]
        var = kwds["invs"].get(nl.jobno[1:])
        if not var:
            logger.debug("%s:No invoice data for JO(%s)" % (kwds["actname"], var))
            nl.goldwgt = None # set the [] to None to avoid excel filling deadloop
            return
        #unitwgt to total wgt
        for idx, wi in enumerate(var.mtlwgt.follows(nl.karat).wgts):
            if not (wi and wi.wgt):
                continue
            nl[gccols[idx][0]] = wi._replace(wgt=round(wi.wgt * nl["joqty"], 2))
        nl.laborcost = round(
            (var.setting + var.labor) * self._rmbtohk * nl["joqty"], 2)


    def _calc(self, **kwds):
        stcosts = self._validate(**kwds)
        nl = kwds.get("nl")
        refs, mpsmp = [], {}
        cstlst = config.get("jocost.cost_fields").split(",")
        gccols = [
            x.split(",") for x in config.get("jocost.mtlwgt_to_cost_fields")
            .split(";")
        ]
        for x in kwds.get("jcs").values():
            nl.setdata(x)
            self._calc_dtos(gccols, stcosts, **kwds)
            # override the karat field for the JO#, an example is 466167
            var = self._getrefid(nl.running, refs)
            if not var:
                logger.critical((
                    "No refid found for running(%d),"
                    " Pls. create one in codetable with (jocostma/costrefid) "
                ) % nl.running)
                return None
            var = self._getmps(var, mpsmp)
            for wn, wc in gccols: # wgt name and wgt cost
                wi = nl[wn]
                if not wi:
                    continue
                if wi.karat not in var:
                    logger.critical(
                        "No MPS found for running(%d)'s karat(%d)" %
                        (nl.running, wi.karat))
                    cost = -1000
                else:
                    cost = round(float(var[wi.karat]) * float(wi.wgt), 2)
                nl[wn], nl[wc] = wi.wgt, cost
                # cost of the parts
                if wn == "extgoldwgt2" and wi.wgt:
                    nl.extlaborcost = round(
                        wi.wgt * (2.5 if wi.karat in (925, 200) else 30), 2)
            for var in cstlst:
                nl["totalcost"] += nl[var]
            nl.unitcost = round(nl["totalcost"] / nl["joqty"], 2)
        return True


    def read(self, year, month, day=1):
        '''
        class to create the C1 JOCost file for HK accountant
        return a tuple with 3 tuple as element, first item of each element is the title row:
            .tuple for jocost
            .tuple for jo stone
            .tuple for broken stone
        '''
        with self._cnsvc.sessionctx() as cur:
            app, tk = xwu.appmgr.acq()
            invs = []
            for var in self._invfldr.split(";"):
                var = app.books.open(var)
                y = C1InvRdr.getReader(var).read(var, read_method="aio", pwgt_check=False)
                var.close()
                if y:
                    invs.extend(y[0][0])
            xwu.appmgr.ret(tk)
            invs = {x.jono: x for x in invs} if invs else {}
            df, dt = daterange(year, month, day)
            var = self._read_shp(cur, df, dt) # nl, jcs, runns, bcs
            # locals()
            kwds = {
                "actname": "C1JOCost of (%04d%02d)" % (year, month),
                "nl": var[0],
                "jcs": var[1],
                "runns": var[2],
                "bcs": var[-1],
                "invs": invs
            }
            var = self._calc(**kwds)
        if var:
            var = kwds["nl"].getcol("running") - 1
            jcs = list([x[1:] for x in kwds["jcs"].values()]) # remove the mmid field
            jcs = sorted(jcs, key=lambda x: x[var])
            jcs.insert(0, config.get("jocost.cost_sheet_fields").split(","))
            return jcs, self.get_st_of_jos(kwds["runns"]), self.get_st_broken(df, dt)

    @staticmethod
    def write(jcs, wb, remove_unused=True):
        '''
        read and write, for example,  C1JCMkr.write(C1JCMkr().read(), wb)
        '''
        sns = "JOCost,JOStone,Broken".split(",")
        for idx, sn in enumerate(sns):
            sht = wb.sheets.add() if idx >= len(
                wb.sheets) else wb.sheets[idx]
            sht.name = sn
            if jcs[idx]:
                sht.range("A1").value = jcs[idx]
            sht.autofit("c")
        if remove_unused:
            idx = len(sns)
            while idx < len(wb.sheets):
                wb.sheets(idx).delete()
        return wb


class C1STHdlr(object):
    r"""
    Read C1Stone's IO from newest file in folder(default \\172.16.8.46\pb\dptfile\quotation\2017外发工单工费明细
    \CostForPatrick\StReadLog\) and save directly to heng_ngai db
    """

    def __init__(self, cnsvc):
        self._cnsvc = cnsvc

    def _remove_done(self, usgs, ionmp):
        '''
        remove the usage records that has already been done
        '''
        def _is_done(cur, q0, u, ionmp):
            """ check if the given usage record(stone_out) has been imported """
            done = False
            try:
                if u.type in ionmp:
                    lst = q0.filter(
                        and_(JO.name == u.jono, StoneOutMaster.isout == ionmp[
                            u.type][0][0])).with_session(cur).all()
                else:
                    lst = Query(
                        [StoneBck.qty, StoneBck.wgt]).join(StoneIn).filter(
                            StoneIn.name == u.btchno).with_session(cur).all()
                for x in lst:
                    done = x.qty == u.qty and abs(x.wgt - u.wgt) < 0.001
                    if done:
                        break
            except:
                pass
            return done

        pflen, pfts, pfcnt = len(usgs), time(), 0
        lb, ub, idx, ipt = 0, pflen - 1, -1, False
        ptr = (lb + ub) // 2
        q0 = Query([StoneOut.qty, StoneOut.wgt]).join(StoneOutMaster).join(JO)
        with self._cnsvc.sessionctx() as cur:
            while idx < 0:
                if ptr == lb:
                    ipt = _is_done(cur, q0, usgs[ptr], ionmp)
                    if not ipt:
                        idx = lb
                    else:
                        if ub != ptr:
                            # save one calculation time if ub == lb
                            ipt = _is_done(cur, q0, usgs[ub], ionmp)
                        if not ipt:
                            idx = ub
                        elif ub < len(usgs) - 1:
                            idx = ub + 1
                    break
                ipt = _is_done(cur, q0, usgs[ptr], ionmp)
                if ipt:
                    lb = ptr + 1
                else:
                    ub = ptr - 1
                ptr = (lb + ub) // 2
                pfcnt += 1
        logger.debug(
            "use %d seconds and %d loops to find the new usage in %d items" %
            (int(time() - pfts), pfcnt, pflen))
        if 0 <= idx < len(usgs):
            logger.debug("the new item's id is %d" % usgs[idx].id)
            return usgs[idx:]
        logger.debug("no new item is found")
        return None

    def _read_from_file(self, fn):
        """
        read the batch/usage data out from the excel
        return a tuple with:
        btnos: a set of well-formatted batch#
        src_pk_mp: a map with well-formatted PK# as key and the last row of batch data as data
        usgs :  a list of usage's row data
        src_bt_mp: a map with well-formatted Bt# as key and the row of batch data as data
        pk_fmted: a tuple of pks that's formatted as (seqid,newpk#,oldpk#,remark)

        return btnos,src_pk_mp,usgs,src_bt_mp,pk_fmted
        """
        if not path.exists(fn):
            return None
        fns = (fn, )
        app, kxl = xwu.appmgr.acq()
        kwds = {
            "src_pk_mp": {},
            "src_bt_mp": {},
            "usgs": [],
            "pk_fmted": []
        }
        try:
            for x in fns:
                wb = app.books.open(x)
                self._read_in(wb, **kwds)
                self._read_out(wb, **kwds)
                wb.close()
        finally:
            if kxl:
                xwu.appmgr.ret(kxl)
        # return src_pk_mp, src_bt_mp, usgs, pk_fmted
        return kwds

    @staticmethod
    def _read_in(wb, **kwds):
        sht, hc = xwu.findsheet(wb, "进"), 0
        vvs = sht.range("A1").expand("table").value
        km = config.get("stio.in.rdr.colmap")
        nls = NamedLists(vvs, km)
        if len(nls.namemap) < len(km):
            logger.debug("not enough key column provided")
            return hc
        for nl in nls:
            if nl.karat and nl.karat != '无':
                continue
            if not nl.btchno:
                break
            pkno = _fmtpkno(nl.pkno)
            if not pkno:
                continue
            flag, pkno = pkno[1], pkno[0]
            if pkno != nl.pkno or flag:
                kwds["pk_fmted"].append((int(nl.id), nl.pkno, pkno,
                                "Special" if flag else "Normal"))
                nl.pkno = pkno
            nl.btchno = _fmtbtno(nl.btchno)
            kwds["src_pk_mp"][nl.pkno] = kwds["src_bt_mp"][nl.btchno] = nl
            hc += 1
        return hc

    @staticmethod
    def _read_out(wb, **kwds):
        sht = xwu.findsheet(wb, u"用")
        vvs, hc = sht.range("A1").expand("table").value, 0
        src_bt_mp, usgs = [kwds[x] for x in ('src_bt_mp', 'usgs')]
        nls, skipcnt = NamedLists(vvs, config.get("stio.out.rdr.colmap")), 0
        for nl in nls:
            btchno = nl.btchno
            if not (btchno or nl.qty):
                skipcnt += 1
                if skipcnt > 3:
                    break
                else:
                    continue
            skipcnt = 0
            #has batch, but qty is empty, sth. blank, but not so blank as above criteria
            if not nl.qty:
                continue
            btchno = _fmtbtno(btchno)
            if btchno not in src_bt_mp:
                continue
            je = JOElement(nl.jono)
            if not je.isvalid():
                #logger.debug("invalid JO#(%s) found in usage seqid(%d),batch(%s)" % (nl.jono,int(nl.id),nl.btchno))
                continue
            nl.btchno, nl.jono = btchno, je
            usgs.append(nl)
            hc += 1
        return hc

    def _get_shp_dates(self, jes):
        """
        return the max shipment date of given JOElement collection as a dict of
        tuple(JOElement, maxRefdate)
        """
        if not jes:
            return
        q0 = Query([JO.name, func.max(MMMa.refdate)]).join(MM).join(MMMa)
        d0 = []
        with self._cnsvc.sessionctx() as cur:
            for arr in splitarray(list(jes)):
                try:
                    d0.extend(
                        q0.filter(jesin(arr, JO)).group_by(
                            JO.name).with_session(cur).all())
                except:
                    pass
        return {x[0]: x[1] for x in d0}

    @staticmethod
    def _normalize_log(fh, pknos, btnos, jonos, **dc):
        crterr = False
        print("stio log of %s" % datetime.today().strftime("%Y%m%d %H:%M"), file=fh)
        if pknos[1]:
            print(
                "Below PK# does not exist, Pls. acquire them from HK first",
                file=fh)
            var = dc["src_pk_mp"]
            lst = sorted([(var[x].id, x) for x in pknos[1]])
            for x in lst:
                print("%d,%s" % x, file=fh)
            print(
                "#use below sql to fetch pk info from hk's pms system",
                file=fh)
            print("use hnj_hk", file=fh)
            print(
                "select package, unit_price, case when unit = 1 then 'PC' when unit = 2 then 'PC' when unit = 3 then 'CT' when unit = 4 then 'GM' when unit = 5 then 'PR' when unit = 6 then 'TL' end from package_ma where package in ('%s')"
                % "','".join([x[1] for x in lst]),
                file=fh)
            crterr = True
        if btnos[1]:
            print(
                "Below BT# does not exists, Pls. get confirm from Kary",
                file=fh)
            var = dc["src_bt_mp"]
            for x in sorted([(var[x].id, x, var[x].pkno) for x in btnos[1]]):
                print("%d,%s,%s" % x, file=fh)
        if jonos and jonos[1]:
            print("Below JO(s) does not exist", file=fh)
            for x in jonos[1]:
                print(x.name, file=fh)
            crterr = True
        var = dc["pk_fmted"]
        if var:
            print("---the converted PK#---", file=fh)
            for x in var:
                print("%d,%s,%s,%s" % x, file=fh)
        var = dc["usgs"]
        if var:
            print("---usage data---", file=fh)
            for y in sorted([(int(x.id), x.type, x.btchno, x.jono.value,
                                x.qty, x.wgt) for x in var]):
                print("%d,%s,%s,%s,%d,%f" % y, file=fh)
        return crterr

    def _normalize(self, **kwds):
        '''
        check if pk/bt was created
        '''
        usgs = kwds.get('usgs')
        with self._cnsvc.sessionctx() as cur:
            var = cur.query(Codetable.codec0, Codetable.coden0).filter(
                and_(Codetable.tblname == "stone_out_master",
                     Codetable.colname == "is_out")).all()
            ionmp = {x.codec0.strip(): int(x.coden0) for x in var}
            msomid = cur.query(func.max(
                StoneOutMaster.id.label("id"))).first()[0]
            var = cur.query(StoneOutMaster.isout,
                            func.max(StoneOutMaster.name).label("bid")).filter(
                                StoneOutMaster.isout.in_(
                                    list(ionmp.values()))).group_by(
                                        StoneOutMaster.isout).all()
            var = {x.isout: x.bid for x in var}
            #make it a isoutname -> (isout,maxid) tuple
            ionmp = {x[0]: [x[1], var[x[1]]]
                            for x in ionmp.items()
                            if x[1] in var}
            mbtid = cur.query(func.max(StoneIn.id)).first()[0]
        ionmp = {x[0]: [ionmp[y] for y in x[1].split(",")] for x in {"补烂": "补石,*退烂石", "补失": "补石,*退失石", "配出": "配出", '加退': '加退'}.items()}
        #print this out and ask for pkdata, or I can not create any further
        var = path.join(gettempdir(), "c1readst.log")
        with open(var, "a") as fh:
            usgs = self._remove_done(usgs, ionmp)
            if not usgs:
                print("stio log of %s (NO USAGE IS FOUND)" % datetime.today().strftime("%Y%m%d %H:%M"), file=fh)
                return True
            jonos = {x.jono for x in usgs if x.jono.isvalid()} if usgs else set()
            # first stage result map in-stack
            mp = {
                "ionmp": ionmp,
                "msomid": msomid,
                "mbtid": mbtid,
                "usgs": usgs,
                "jes": jonos
            }
            btnos = self._cnsvc.getstins(kwds["src_bt_mp"].keys())
            pknos = self._cnsvc.getstpks(kwds["src_pk_mp"].keys())
            jonos = self._cnsvc.getjos(jonos)
            crterr = self._normalize_log(fh, pknos, btnos, jonos, **kwds)
        logger.info("log were saved to %s" % var)
        if crterr:
            logger.critical(
                "There are Package or JO does not exist, Pls. correct them first"
            )
            return None
        mp.update({
            "btnos": btnos,
            "pknos": pknos,
            "jonos": jonos,
        })
        return mp

    def _build_nbt(self, rtmp, **kwds):
        src_bt_mp = kwds.get('src_bt_mp')
        lnm = lambda cl: {x.name: x for x in cl}
        btmp, pkbyns = (lnm(rtmp[x][0]) if rtmp[x] and rtmp[x][0] else {} for x in ("btnos", "pknos"))
        #new batch,stoneoutmaster and stoneout,newstoneback, newclosebatch
        nbtmp, td, mbtid = {}, datetime.today(), rtmp["mbtid"]
        for x in (x for x in src_bt_mp.items() if x[0] not in btmp):
            si = StoneIn()
            mbtid, si.filldate = mbtid + 1, x[1].date
            si.docno = "AG" + x[1].date.strftime("%y%m%d")[1:]
            si.id, si.cstref, si.lastupdate, si.lastuserid = mbtid, NA, td, 1
            si.name, si.qty, si.qtytrans, si.qtyused, si.cstid = x[0], x[
                1].qty, 0, 0, 1
            si.size, si.tag, si.wgt, si.wgtadj, si.wgtbck = NA, 0, x[
                1].wgt, 0, 0
            si.wgtprep = si.wgttrans = si.wgtused = si.qtybck = si.wgttmptrans = 0
            if x[1].pkno in pkbyns:
                si.pkid = pkbyns[x[1].pkno].id
            else:
                si.pkid = x[1].pkno
            nbtmp[si.name] = si
        return nbtmp

    def _build_bck_sos(self, rtmp, **kwds):
        joshd = self._get_shp_dates(rtmp["jes"])
        msomid = rtmp["msomid"]
        nsos, nbck, ncbt, td = {}, [], set(), datetime.today()
        var = lambda cl: {x.name: x for x in cl}
        btmp, jobyns = (var(rtmp[x][0]) if rtmp[x] and rtmp[x][0] else {} for x in ("btnos", "jonos"))
        for x in kwds["usgs"]:
            if x.type not in rtmp["ionmp"]:
                var = StoneBck()
                nbck.append(var)
                ncbt.add(x.btchno)
                var.btchid = btmp[x.btchno].id if x.btchno in btmp else x.btchno
                var.idx, var.lastuserid, var.qty, var.wgt = 1, 1, x.qty, x.wgt
                var.filldate = var.lastupdate = td
                var.docno = "AG" + td.strftime("%y%m%d")[1:]
            else:
                for iof in rtmp["ionmp"][x.type]:
                    var = nsos.setdefault(x.jono.value + "," + str(iof[0]), {})
                    if not var:
                        som = StoneOutMaster()
                        som.joid = jobyns[x.jono].id
                        var["som"] = som
                        msomid += 1
                        iof[1] += 1
                        som.id, som.isout, som.name = msomid, iof[0], iof[1]
                        som.packed, som.qty, som.subcnt, som.workerid = 0, 0, 0, 1393
                        som.filldate, som.lastupdate, som.lastuserid = joshd.get(
                            x.jono, td), td, 1
                    else:
                        som = var["som"]
                    var, so = var.setdefault("sos", []), StoneOut()
                    var.append(so)
                    so.id, so.idx, so.joqty, so.lastupdate, so.lastuserid = som.id, len(var), 0, td, 1
                    so.printid, so.qty, so.wgt, so.workerid = 0, x.qty, x.wgt, 1393
                    so.checkerid, so.checkdate = 0, som.filldate
                    so.btchid = btmp[x.btchno].id if x.btchno in btmp else x.btchno
        return {
            "nsos": nsos,
            "nbck": nbck,
            "ncbt": ncbt,
            "btmp": btmp
        }

    def _build(self, **kwds):
        '''
        return None or a map contains:
            nbtmp: a map as key = btchno and value = StoneIn
            nsos: a map as key = (JO,isOut) and map("som", "sos")
            nbck: a tuple of StoneIn that is closed
            ncbt: the closed BT, a tuple of StoneIn
            btmp: a map as key = btchno and value = StoneIn from the existing db, for batch# -> StoneIn lookup
        '''
        rtmp = self._normalize(**kwds)
        if not rtmp:
            # error
            return None
        if isinstance(rtmp, bool):
            return True
        kwds.update(rtmp)
        nbtmp = self._build_nbt(rtmp, **kwds)
        mp = self._build_bck_sos(rtmp, **kwds)
        mp.update(nbtmp=nbtmp)
        return mp

    def _persist(self, **kwds):
        err = True
        uCnt = 0
        #it's quite strange that even if I by monitor one by one, some committed records can not be found in the db, so generate SQL instead
        # reason found: when there is locked batches, the transaction should have been rollbacked, but the ORM still don't throw the exception!
        cmds = []
        with self._cnsvc.sessionctx() as cur:
            # new batches
            nbt, nsos, btmp = (kwds.get(x) for x in 'nbtmp nsos btmp'.split())
            lst = []
            for x in nbt.items():
                x[1].qty = int(x[1].qty) if x[1].qty else 0
                lst.append(x[1])
                uCnt += 1
            if uCnt > 0:
                cur.add_all(lst)
                cur.flush()
                cur.commit()
            uCnt = 0
            fmtDt = lambda x: "'" + x.strftime('%Y-%m-%d %H:%M') + "'"
            std = fmtDt(datetime.today())
            for x in nsos.values():
                som = x["som"]
                som.subcnt = len(x["sos"])
                theSQL = 'insert into stone_out_master(id, bill_id, is_out, jsid, qty, fill_date, packed, subcnt, worker_id, lastuserid, lastupdate) values(%d, %d, %d, %d, %d, %s, %d, %d, %d, %d, %s)'
                theSQL = theSQL % (som.id, som.name, som.isout, som.joid, 0, std, 0, som.subcnt, som.workerid, 1, std)
                # cur.add(som)
                cmds.append(theSQL)
                for so in x['sos']:
                    if isinstance(so.btchid, str):
                        so.btchid = nbt[so.btchid].id
                    # cur.add(so)
                    theSQL = 'insert into stone_out(id, idx, btchid, worker_id, quantity, weight, checker_id, check_date, qty, printid, lastuserid, lastupdate) values (%d, %d, %d, %d, %d, %f, %d, %s, %d, %d, %d, %s)'
                    theSQL = theSQL % (so.id, so.idx, so.btchid, 0, so.qty, so.wgt, 0, fmtDt(so.checkdate), 0, 0, 1, std)
                    cmds.append(theSQL)
                uCnt += 1
            if cmds:
                for theSQL in cmds:
                    cur.execute(theSQL)
                cur.commit()
            self._persist_bck(cur, kwds.get("nbck"), nbt)
            uCnt = 0
            ctag = int(datetime.today().strftime("%m%d"))
            for x in kwds["ncbt"]:
                btno = btmp[x] if x in btmp else nbt[x]
                btno.tag = ctag
                cur.add(btno)
                uCnt += 1
            if uCnt > 0:
                cur.flush()
            cur.commit()
            err = False
        return not err

    @staticmethod
    def _persist_bck(cur, nbck, nbt):
        ''' persist the stone_bck items '''
        if not nbck:
            return
        uCnt = 0
        lst = [x.btchid for x in nbck if not isinstance(x.btchid, str)]
        if lst:
            try:
                y = []
                for k in splitarray(lst, 20):
                    lst = Query([
                        StoneBck.btchid,
                        func.max(StoneBck.idx).label("idx")
                    ]).filter(StoneBck.btchid.in_(k)).group_by(
                        StoneBck.btchid).with_session(cur).all()
                    y.extend(lst)
                lst = {x.btchid: x.idx for x in y}
            except:
                pass
        else:
            lst = {}
        for x in nbck:
            if isinstance(x.btchid, str):
                x.btchid = nbt[x.btchid].id
            idx = lst[x.btchid] if x.btchid in lst else 0
            #very rare case, check if it's been imported
            if idx > 0:
                dup = False
                try:
                    y = Query([StoneBck.qty, StoneBck.wgt]).filter(StoneBck.btchid == x.btchid).with_session(cur).all()
                    for yy in y:
                        dup = yy.qty == x.qty and abs(yy.wgt - x.wgt) < 0.001
                        if dup:
                            break
                except:
                    pass
                if dup:
                    logger.debug("trying to return duplicated item")
                    continue
            idx += 1
            lst[x.btchid], x.idx = idx, idx
            cur.add(x)
            uCnt += 1
        if uCnt > 0:
            cur.flush()

    def readst(self, fn, persist=True):
        """
        read and create the stone usage record from C1, input files only
        """
        #check if one usage item has been inputted. the rule is:
        # if jo+iotype+qty+closeWgt found, treated as dup. Once one item is found
        # not imported , all item behind it was think of not imported
        if path.isdir(fn):
            #only read the most-recent file
            maxd = datetime(1980, 1, 1)
            fns, fx = getfiles(fn, "xls"), None
            for x in fns:
                d0 = datetime.fromtimestamp(path.getmtime(x))
                if maxd < d0:
                    fx = x
                    maxd = d0
            logger.debug(
                "%d files in folder(%s), most-updated file(%s) is selected" %
                (len(fns), fn, path.basename(fx)))
            fn = fx
        rst = self._build(**self._read_from_file(fn))
        if not rst:
            # some invalid items found by _buildrst, throw exception
            return None
        if isinstance(rst, bool):
            # no usage record found, just return OK
            return True
        return self._persist(**rst) if persist else rst
