# coding=utf-8
'''
Created on 2018-04-28
classes to read data from C1's monthly invoices
need to be able to read the 2 kinds of files: C1's original and calculator file
@author: zmFeng
'''

from datetime import datetime, date
from numbers import Number
from tempfile import gettempdir
from re import compile as cpl
from time import time
from collections import namedtuple
from os import path

from sqlalchemy import and_, func
from sqlalchemy.orm import Query
from xlwings import Sheet, constants

from hnjapp.dbsvcs import jesin
from hnjapp.pajcc import WgtInfo, addwgt
from hnjcore import JOElement, karatsvc
from hnjcore.models.cn import (JO, MM, Codetable, Customer, MMgd, MMMa,
                               StoneBck, StoneIn, StoneOut, StoneOutMaster,
                               StonePk, Style)
from hnjcore.utils.consts import NA
from utilz import (NamedList, NamedLists, daterange, getfiles, isnumeric,
                   splitarray, trimu, xwu, tofloat)

from .common import _date_short
from .common import _logger as logger

_ptnbtno = cpl(r"(\d+)([A-Z]{3})(\d+)")

''' items for c1inv and c1stone '''
C1InvItem = namedtuple("C1InvItem", "source,jono,qty,labor,setting,remarks,stones,mtlwgt,styno")
C1InvStone = namedtuple("C1InvStone", "stone,qty,wgt,remark")
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
    if not btno: return
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
            mt = _ptnbtno.search(btno)
            if mt:
                btno = btno[mt.start(1):mt.end(2)] + ("%03d" % int(mt.group(3)))
    return ("-" if flag else "") + btno


def _fmtpkno(pkno):
    """ in C1's STIO excel file, the package# is quite malformed
    this method format it to standard ones
    """
    if not pkno: return
    #contain invalid character, not a PK#
    pkno = trimu(pkno)
    if sum([1 for x in pkno if ord(x) <= 31 or ord(x) >= 127]) > 0:
        return
    pkno0 = pkno
    if pkno.find("-") >= 0: pkno = pkno.replace("-", "")
    pfx, pkno, sfx = pkno[:3], pkno[3:], ""
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
    """
        read the monthly invoice files from both C1 & CC version
    """
    _rdr_km = {
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
        "netwgt": "货重,"
    }

    def __init__(self, c1log=None, cclog=None):
        self._c1log = c1log
        self._cclog = cclog

    def read(self, fldr=None, pwgt_check=True):
        """
        perform the read action
        @param fldr: the folder contains the invoice files
        @param pwgt_check:
            True: pwgt will be treated as parts'wgt if is pendant.
            False: pwgt will always be treated as parts' weight of whatever style,  only used for jocost handling
        @return: a list of C1InvItem
        """

        if not fldr or not path.exists(fldr):
            fldr = r"\\172.16.8.46\pb\dptfile\quotation\2017外发工单工费明细\CostForPatrick\AIO_F.xlsx"
        if not path.exists(fldr):
            return None
        if path.isfile(fldr):
            fns = [fldr]
        else:
            fns = getfiles(fldr)
        if not fns:
            return None
        killxw, app = xwu.app(False)
        wb = None
        try:
            cnsc1 = u"工单号,镶工,胚底,备注".split(",")
            for fn in fns:
                wb = app.books.open(fn)
                items = list()
                for sht in wb.sheets:
                    rngs = [
                        x for x in (xwu.find(
                            sht, var, lookat=constants.LookAt.xlPart)
                                    for var in cnsc1) if x
                    ]
                    if len(cnsc1) != len(rngs):
                        continue
                    var = self.read_c1(sht, skip_checking=True, pwgt_check=pwgt_check)
                    if var:
                        items.extend(var[0])
                wb.close()
        finally:
            if killxw:
                app.quit()
        return items

    @classmethod
    def read_c1(cls, sht, skip_checking=False, pwgt_check=True):
        """
        read c1 invoice file
        @param   sht: the sheet that is verified to be the C1 format
        @param hdrow: the row of the header
        @param pwgt_check:
            True: pwgt will be treated as parts'wgt if is pendant.
            False: pwgt will always be treated as parts' weight of whatever style,  only used for jocost handling

        @return: a list of C1InvItem with source = "C1"
        """

        if not skip_checking:
            # should not check given sheet only, because at the begining of month, data in prior month, return a reversed sorted-by-month list and check
            nl = cls._select_sheet(sht)
            if not nl:
                return
            # don't use "is", might be false even if 2 sheets pointing to the same
            if sht != nl[0]:
                if sht in nl:
                    sht.delete()
                return None
            #there might be several date, get the biggest one
            nl = xwu.find(sht, "日期", find_all=True)
            if not nl:
                return None
            nl = [(x, x.offset(0, 1).value) for x in nl if x.offset(0, 1).value]
            # sometimes the latest one contains nothing except date, an example is 20190221, so read from buttom
            nl = sorted(nl, key=lambda x: x[1], reverse=True)
            for x in nl:
                rng = xwu.find(sht, "图片", After=x[0])
                sn = cls._read_from(rng, pwgt_check)
                if sn:
                    return sn[0], x[1].date()
            rng = None
        else:
            rng, invdate = xwu.find(sht, "图片"), date.today()
        if not rng:
            return None
        sn = cls._read_from(rng, pwgt_check)
        return sn[0], invdate if sn else None

    @classmethod
    def _select_sheet(cls, sht):
        mp = {}
        for x in sht.book.sheets:
            sn = x.name
            nl = sn.find("月")
            if nl <= 0:
                continue
            nl = sn[:nl]
            if nl.isnumeric():
                mp[int(nl)] = x
        if mp:
            sn = sorted([x for x in mp], reverse=True)
            # Jan contains Dec case
            if len(sn) > 1 and sn[0] - sn[-1] > 10:
                sn[0], sn[-1] = sn[-1], sn[0]
            mp = [mp[x] for x in sn]
        return mp

    @classmethod
    def read_c1_all(cls, sht):
        '''
        read all the occurences inside given sheet
        @param sht: a sheet instance. If a string is provided, will be treated as file, then only the sheet with "\d{1,2}月" will be processed
        '''
        tk = wb = data = None
        if not isinstance(sht, Sheet):
            if isinstance(sht, str):
                app, tk = xwu.appmgr.acq()
                wb = app.books.open(sht)
            else:
                wb = sht
            rng = [x for x in wb.sheets if x.name.find('月') >= 0]
            sht = rng[0] if rng else None
        if sht:
            rng, rngs, data, kw = None, set(), list(), "图片"
            while True:
                if rng:
                    rng = xwu.find(sht, kw, After=rng)
                else:
                    rng = xwu.find(sht, kw)
                if not rng or rng in rngs:
                    break
                rngs.add(rng)
            rngs, lidx = sorted(rngs, key=lambda x: x.row), 0
            for rng in rngs:
                if rng.row < lidx:
                    # in the AIO mode, some title was covered by the top title, just skip it
                    continue
                x = cls._read_from(rng)
                if x:
                    data.extend(x[0])
                    lidx = x[1]
        if tk:
            wb.close()
            xwu.appmgr.ret(tk)
        return data

    @classmethod
    def _read_from(cls, rng, pwgt_check=True):
        nls = [x for x in xwu.NamedRanges(rng, name_map=cls._rdr_km)]
        if not nls:
            return None
        nl = nls[0]
        kns = [1 for x in "jono,gwgt,swgt".split(",") if nl.getcol(x)]
        if len(kns) != 3:
            logger.debug("sheet(%s), range(%s) does not contain necessary key columns" % (rng.sheet.name, rng.address))
            return None
        # last_act states:
        # 1 -> JO#
        # 2 -> blank but prior is &le; 2
        # 3 -> blank but prior is 2, void
        items, c1, netwgt_c1, jo_lidx, last_act = [], None, {}, {}, 3
        _cnstqnw = "stqty,stwgt".split(",")
        _cnsnl = "setting,labor".split(",")
        for idx, nl in enumerate(nls):
            s0 = nl.jono
            if s0:
                je = JOElement(s0)
                if je.isvalid():
                    if idx - jo_lidx.get(je.value, idx) > 0:
                        logger.debug("Duplicated JO#(%s) found, the prior is near row(%d)" % (je.value, jo_lidx[je.value]))
                        last_act = 3
                        continue
                    snl = [tofloat(nl[s0]) for s0 in _cnsnl]
                    if not any(snl):
                        logger.debug("JO(%s) does not contains any labor cost", je.value)
                        snl = (0, 0)
                    c1 = C1InvItem("C1", je.value, nl.joqty, snl[1], snl[0], nl.remark, [], None, nl.styno)
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
            s0 = cls._extract_st_mtl(c1, nl, _cnstqnw, pwgt_check)
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
            if c1.jono in ('B70981'):
                print('x')
            kt = sum((x.wgt for x in c1.stones)) / 5 if c1.stones else 0
            if c1.mtlwgt:
                wgt = sum((x.wgt for x in c1.mtlwgt.wgts if x)) + kt
                s0.append(
                    c1._replace(
                        mtlwgt=c1.mtlwgt._replace(netwgt=round(wgt, 2))))
                nl = (wgt - netwgt_c1[c1.jono][1] / c1.qty, netwgt_c1[c1.jono][0] / c1.qty)
                if abs(nl[0] - nl[1]) > 0.01:
                    logger.debug("JO#(%s)'s netwgt contains error, it should be %4.2f but was %4.2f" % (c1.jono, nl[0], nl[1]))
            else:
                logger.debug("JO(%s) does not contains metal wgt", c1.jono)
        return s0, rng.last_cell.row + len(nls)

    @classmethod
    def verify(cls, c1):
        ''' verify if c1's weight is valid. a valid c1 weight should fulfill below
        criteria:
            .netwgt == c1.mtlwgt + c1.stwgt. that is, the netwgt of mtlwgt
        JO#463625, violate this rule, but seems ok
        '''
        rc = not (bool(c1.netwgt) ^ bool(c1.mtlwgt))
        if rc and c1.netwgt:
            rc = abs(c1.netwgt - c1.mtlwgt.netwgt) < 0.01
        return rc

    @classmethod
    def _is_pendant(cls, styno):
        return styno and styno[:2].upper().find("P") >= 0

    @classmethod
    def _extract_st_mtl(cls, c1, nl, cnstqnw, pwgt_check=True):
        '''
        extract the stone and metal into c1
        @param pwgt_check:
            True: pwgt will be treated as parts'wgt if is pendant.
            False: pwgt will always be treated as parts' weight of whatever style,  only used for jocost handling
        '''
        #stone data
        hc = 0
        qnw = [tofloat(nl[x]) for x in cnstqnw]
        if all(qnw):
            s0 = nl.stname
            if s0 and isinstance(s0, str):
                joqty = c1.qty
                c1.stones.append(C1InvStone(nl.stname, qnw[0] / joqty, round(qnw[1] / joqty, 4), "N/A"))
                hc += 1
        #wgt data
        kt, gw, sw, pwgt = nl.karat, nl.gwgt, nl.swgt, nl.pwgt
        if kt and isnumeric(kt):
            joqty = c1.qty
            if not joqty:
                logger.debug("JO(%s) without qty, skipped" % nl.jono)
                return None
            hc += 1
            kt, wgt = cls._tokarat(kt), gw or sw
            #only pendant's pwgt is pwgt, else to mainpart
            if pwgt and pwgt_check and not cls._is_pendant(c1.styno):
                wgt += pwgt
                pwgt = 0
            c1 = c1._replace(mtlwgt=addwgt(c1.mtlwgt, WgtInfo(kt, wgt / joqty, 4)))
            if pwgt:
                c1 = c1._replace(mtlwgt=addwgt(c1.mtlwgt, WgtInfo(kt, pwgt / joqty, 4), True))
        return c1 if hc else None

    @classmethod
    def _tokarat(cls, kt):
        if kt < 1:
            kt = int(kt * 1000)
        if kt >= 924 and kt <= 926:
            rc = 925
        elif kt >= 330 and kt <= 340:
            rc = 8
        elif kt >= 370 and kt <= 385:
            rc = 9
        elif kt >= 580 and kt <= 590:
            rc = 14
        elif kt >= 745 and kt <= 755:
            rc = 18
        else:
            rc = kt
        return rc


class C1JCMkr(object):
    r"""
    C1 JOCost maker, First, Invoke C1STHdlr to create Stone Usage , then generate the jocost report to given folder(default is p:\aa\)
    """

    def __init__(self, cnsvc, bcsvc, invfldr):
        r"""
        @param cnsvc: the CNSvc instance
        @param dbsvc: the BCSvc instance
        @param invfldr: folder contains C1's invoices, one example is
            \\172.16.8.46\pb\dptfile\quotation\2017外发工单工费明细\CostForPatrick\AIO_F.xlsx
        """
        self._cnsvc = cnsvc
        self._bcsvc = bcsvc
        self._invfldr = invfldr

    #return refid by running, from existing list or db#
    def _getrefid(self, runn, mpss):
        refid = None
        if mpss:
            for x in mpss:
                if x[0][0] <= runn and x[0][1] >= runn:
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

    def read(self, year, month, day=1, rmbtohk=1.25, tplfn=None, tarfldr=None):
        """class to create the C1 JOCost file for HK accountant"""
        df, dt = daterange(year, month, day)
        refs, mpsmp, runns = [], {}, set()
        actname = "C1JOCost of (%04d%02d)" % (year, month)
        ptncx = cpl(r"C(\d)$")
        with self._cnsvc.sessionctx() as cur:
            mmids, vvs, refs = set(), {}, []
            gccols = [
                x.split(",") for x in
                "goldwgt,goldcost;extgoldwgt,extgoldcost;extgoldwgt2,extgoldcost2"
                .split(";")
            ]
            ttls = (
                "mmid,lastmmdate,jobno,cstname,styno,running,mstone,description,joqty,karat,goldwgt,goldcost,extgoldcost,extgoldcost2,stonecost,laborcost,extlaborcost,extcost,totalcost,unitcost,extgoldwgt,extgoldwgt2,cflag,rmb2hk"
            ).split(",")
            nl = NamedList(xwu.list2dict(ttls))
            invs = C1InvRdr().read(self._invfldr, pwgt_check=False)
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
            #vvs["_TITLE_"] = ttls
            for x in lst:
                jn = x.jono.value
                if x.id not in mmids:
                    mmids.add(x.id)
                    if jn not in vvs:
                        nl.setdata([0] * len(ttls))
                        nl.mmid, nl.lastmmdate, nl.jobno = x.id, "'" + x.refdate.strftime(
                            _date_short), "'" + x.jono.value
                        nl.cstname, nl.styno, nl.running = x.cstname.strip(
                        ), x.styno.value, x.running
                        nl.mstone, nl.description, nl.karat = "_ST", "_EDESC", karatsvc.getfamily(
                            x.jokarat).karat
                        nl.goldwgt, nl.cflag, nl.rmb2hk = [], "NA", rmbtohk
                        mt = ptncx.search(x.docno)
                        if mt:
                            nl.cflag = "'" + mt.group(1)
                        vvs[jn] = nl.data
                        runns.add(int(x.running))
                    nl.setdata(vvs[jn])["joqty"] += float(x.qty)
                #nl.setdata(vvs[jn]).goldwgt.append((karatsvc.getfamily(x.karat).karat, x.wgt))
            bcs = self._bcsvc.getbcsforjc(runns)
            if not bcs or len(bcs) < len(runns):
                logger.debug("%s:Not all records found in BCSystem" % actname)
            bcs = {x.runn: (x.desc, x.ston) for x in bcs}
            stcosts = self.get_st_costs(runns)
            if not stcosts or len(stcosts) < len(runns) / 2:
                logger.debug(
                    "%s:No stone or less than 1/2 has stone, make sure you've prepared stone data with C1STIOData"
                    % actname)
            x = {x.jono for x in invs if x.jono in vvs} if invs else set()
            if not invs or len(x) < len(runns):
                logger.debug("%s:No or not enough C1 invoice data from file(%s)"
                             % (actname, self._invfldr))
            invs = {x.jono: x for x in invs} if invs else {}
            cstlst = "goldcost,extgoldcost,stonecost,laborcost,extlaborcost,extcost,extgoldcost2".split(
                ",")
            for x in vvs.values():
                nl.setdata(x)
                runn = nl.running
                if runn in stcosts:
                    nl.stonecost = stcosts[runn]
                runn = str(runn)
                if runn in bcs:
                    dns = bcs[runn]
                    nl.description, nl.mstone = dns[0], dns[1]

                runn = nl.jobno[1:]
                if runn in invs:
                    inv = invs[runn]
                    nl.laborcost = round(
                        (inv.setting + inv.labor) * rmbtohk * nl["joqty"], 2)
                else:
                    logger.debug(
                        "%s:No invoice data for JO(%s)" % (actname, runn))
                    continue
                prdwgt = invs.get(
                    nl.jobno[1:]).mtlwgt  # A "'" should be skipped
                prdwgt = (prdwgt.main, prdwgt.aux, prdwgt.part)
                #unitwgt to total wgt
                for idx, wi in enumerate(prdwgt):
                    if not wi:
                        continue
                    wi = wi._replace(wgt=round(wi.wgt * nl["joqty"], 2))
                    nl[gccols[idx][0]] = wi
                refid = self._getrefid(nl.running, refs)
                if not refid:
                    logger.critical((
                        "No refid found for running(%d),"
                        " Pls. create one in codetable with (jocostma/costrefid) "
                    ) % nl.running)
                    vvs = None
                    break
                else:
                    mp = self._getmps(refid, mpsmp)
                    for vv in gccols:
                        wi = nl[vv[0]]
                        if not wi:
                            continue
                        if wi.karat not in mp:
                            logger.critical(
                                "No MPS found for running(%d)'s karat(%d)" %
                                (nl.running, wi.karat))
                            cost = -1000
                        else:
                            cost = round(float(mp[wi.karat]) * float(wi.wgt), 2)
                        nl[vv[0]], nl[vv[1]] = wi.wgt, cost
                        if vv[0] == "extgoldwgt2" and wi.wgt:
                            nl.extlaborcost = round(
                                wi.wgt * (2.5 if wi.karat == 925 or
                                          wi.karat == 200 else 30), 2)
                if vvs is None:
                    break
                for cx in cstlst:
                    nl["totalcost"] += nl[cx]
                nl.unitcost = round(nl["totalcost"] / nl["joqty"], 2)
        if vvs:
            ll = list([x[1:] for x in vvs.values()])
            refid = nl.getcol("running") - 1
            ll = sorted(ll, key=lambda x: x[refid])
            ll.insert(0, ttls[1:])
            return ll, self.get_st_of_jos(runns), self.get_st_broken(df, dt)


class C1STHdlr(object):
    r"""
    Read C1Stone's IO from newest file in folder(default \\172.16.8.46\pb\dptfile\quotation\2017外发工单工费明细
    \CostForPatrick\StReadLog\) and save directly to heng_ngai db
    """

    def __init__(self, cnsvc):
        self._cnsvc = cnsvc

    def _rviptusg(self, usgs, ionmap):
        """ remove the imported usage records """

        def _ckript(cur, q0, u, ionmap):
            """ check if the given usage record(stone_out) has been imported """
            ipt = False
            try:
                if u.type in ionmap:
                    lst = q0.filter(
                        and_(JO.name == u.jono, StoneOutMaster.isout == ionmap[
                            u.type][0][0])).with_session(cur).all()
                else:
                    lst = Query(
                        [StoneBck.qty, StoneBck.wgt]).join(StoneIn).filter(
                            StoneIn.name == u.btchno).with_session(cur).all()
                for x in lst:
                    ipt = x.qty == u.qty and abs(x.wgt - u.wgt) < 0.001
                    if ipt: break
            except:
                pass
            return ipt

        pflen, pfts, pfcnt = len(usgs), time(), 0
        lb, ub, idx, ipt = 0, pflen - 1, -1, False
        ptr = (lb + ub) // 2
        q0 = Query([StoneOut.qty, StoneOut.wgt]).join(StoneOutMaster).join(JO)
        with self._cnsvc.sessionctx() as cur:
            while idx < 0:
                if ptr == lb:
                    if not _ckript(cur, q0, usgs[lb], ionmap):
                        idx = lb
                    else:
                        if not _ckript(cur, q0, usgs[ub], ionmap):
                            idx = ub
                        elif ub < len(usgs):
                            idx = ub + 1
                    break
                ipt = _ckript(cur, q0, usgs[ptr], ionmap)
                if ipt:
                    lb = ptr + 1
                else:
                    ub = ptr - 1
                ptr = (lb + ub) // 2
                pfcnt += 1
        logger.debug(
            "use %d seconds and %d loops to find the new usage in %d items" %
            (int(time() - pfts), pfcnt, pflen))
        if idx >= 0:
            return usgs[idx:]

    def _readfrmfile(self, fn):
        """
        read the batch/usage data out from the excel
        return a tuple with:
        btnos: a set of well-formatted batch#
        pkmap: a map with well-formatted PK# as key and the last row of batch data as data
        usgs :  a list of usage's row data
        btmap: a map with well-formatted Bt# as key and the row of batch data as data
        pkfmted: a tuple of pks that's formatted as (seqid,newpk#,oldpk#,remark)

        return btnos,pkmap,usgs,btmap,pkfmted
        """
        if not path.exists(fn): return
        fns = [fn]
        kxl, app = xwu.app(False)

        btmap, pkfmted, usgs = {}, [], []
        pkmap = {}
        try:
            for fn in fns:
                wb = app.books.open(fn)
                shts = {}
                for sht in wb.sheets:
                    shts[sht.name] = sht
                sht = shts[u"进"]
                vvs = sht.range("A1").expand("table").value
                km = {
                    "id": u"序号",
                    "btchno": u"水号",
                    "pkno": u"包头",
                    "date": u"日期,",
                    "type": u"类别",
                    "karat": u"成色",
                    "qty": u"数量,",
                    "wgt": u"重量,",
                    "qtyunit": u"数量单位",
                    "unit": u"重量单位",
                    "remark": u"备注"
                }
                nls = NamedLists(vvs, km)
                if len(nls.namemap) < len(km):
                    logger.debug("not enough key column provided")
                    break
                for nl in nls:
                    if nl.karat: continue
                    if not nl.btchno: break
                    pkno = _fmtpkno(nl.pkno)
                    if not pkno: continue
                    flag = pkno[1]
                    pkno = pkno[0]
                    if pkno != nl.pkno or flag:
                        pkfmted.append((int(nl.id), nl.pkno, pkno,
                                        "Special" if flag else "Normal"))
                        nl.pkno = pkno
                    nl.btchno = _fmtbtno(nl.btchno)
                    pkmap[nl.pkno], btmap[nl.btchno] = nl, nl
                sht = shts[u"用"]
                vvs = sht.range("A1").expand("table").value
                km = {
                    "id": u"序号",
                    "btchno": u"水号",
                    "jono": u"工单",
                    "qty": u"数量",
                    "wgt": u"重量",
                    "type": u"记录,",
                    "btchid": u"备注"
                }
                nls = NamedLists(vvs, km)
                skipcnt = 0
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
                    if not nl.qty: continue
                    btchno = _fmtbtno(btchno)
                    if btchno not in btmap:
                        continue
                    je = JOElement(nl.jono)
                    if not je.isvalid():
                        #logger.debug("invalid JO#(%s) found in usage seqid(%d),batch(%s)" % (nl.jono,int(nl.id),nl.btchno))
                        continue
                    nl.btchno, nl.jono = btchno, je
                    usgs.append(nl)
                wb.close()
        finally:
            if kxl: app.quit()
        return pkmap, btmap, usgs, pkfmted

    def _getjoshpdate(self, jes):
        """
        return the max shipment data of given JOElement collection as a dict of
        (JOElement,maxRefdate)
        """
        if not jes: return
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
        return dict([(x[0], x[1]) for x in d0])

    def _buildrst(self, pkmap, btmap, usgs, pkfmted):
        with self._cnsvc.sessionctx() as cur:
            lst = cur.query(Codetable.codec0, Codetable.coden0).filter(
                and_(Codetable.tblname == "stone_out_master",
                     Codetable.colname == "is_out")).all()
            msomids = dict([(x.codec0.strip(), int(x.coden0)) for x in lst])
            msomid = cur.query(func.max(
                StoneOutMaster.id.label("id"))).first()[0]
            lst = cur.query(StoneOutMaster.isout,
                            func.max(StoneOutMaster.name).label("bid")).filter(
                                StoneOutMaster.isout.in_(
                                    list(msomids.values()))).group_by(
                                        StoneOutMaster.isout).all()
            lst = dict([(x.isout, x.bid) for x in lst])
            #make it a isoutname -> (isout,maxid) tuple
            msomids = dict([(x[0], [x[1], lst[x[1]]])
                            for x in msomids.items()
                            if x[1] in lst])
            mbtid = cur.query(func.max(StoneIn.id)).first()[0]
        ionmap = {}
        for x in {"补烂": "补石,*退烂石", "补失": "补石,*退失石", "配出": "配出"}.items():
            ionmap[x[0]] = [msomids[y] for y in x[1].split(",")]
        usgs = self._rviptusg(usgs, ionmap)
        jonos = set()
        if usgs:
            for nl in usgs:
                if nl.jono.isvalid():
                    jonos.add(nl.jono)
        btnos = self._cnsvc.getstins(btmap.keys())
        pknos = self._cnsvc.getstpks(pkmap.keys())
        jes = jonos
        jonos = self._cnsvc.getjos(jonos)
        tmpf = gettempdir() + path.sep
        #print this out and ask for pkdata, or I can not create any further
        fn, crterr = tmpf + "c1readst.log", False
        with open(fn, "w") as fh:
            if pknos[1]:
                print(
                    "Below PK# does not exist, Pls. acquire them from HK first",
                    file=fh)
                lst = sorted([(pkmap[x].id, x) for x in pknos[1]])
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
                for x in sorted(
                    [(btmap[x].id, x, btmap[x].pkno) for x in btnos[1]]):
                    print("%d,%s,%s" % x, file=fh)
            if jonos and jonos[1]:
                print("Below JO(s) does not exist", file=fh)
                for x in jonos[1]:
                    print(x.name, file=fh)
                crterr = True
            if pkfmted:
                print("---the converted PK#---", file=fh)
                for x in pkfmted:
                    print("%d,%s,%s,%s" % x, file=fh)
            if usgs:
                print("---usage data---", file=fh)
                for y in sorted([(int(x.id), x.type, x.btchno, x.jono.value,
                                  x.qty, x.wgt) for x in usgs]):
                    print("%d,%s,%s,%s,%d,%f" % y, file=fh)

        logger.info("log were saved to %s" % fn)
        if crterr:
            logger.critical(
                "There are Package or JO does not exist, Pls. correct them first"
            )
            return
        lnm = lambda cl: dict([(x.name, x) for x in cl])

        btbyns = lnm(btnos[0])
        pkbyns = lnm(pknos[0])
        jobyns = lnm(jonos[0]) if jonos and jonos[0] else {}
        #new batch,stoneoutmaster and stoneout,newstoneback, newclosebatch
        nbtmap, sos, nbck, ncbt = {}, {}, [], set()
        td = datetime.today()
        for x in btmap.items():
            if x[0] not in btbyns:
                si = StoneIn()
                mbtid, si.filldate = mbtid + 1, x[1].date
                si.docno = "AG" + x[1].date.strftime("%y%m%d")[1:]
                si.id, si.cstref, si.lastupdate, si.lastuserid = mbtid, NA, td, 1
                si.name, si.qty, si.qtytrans, si.qtyused, si.cstid = x[0], x[
                    1].qty, 0, 0, 1
                si.size, si.tag, si.wgt, si.wgtadj, si.wgtbck = NA, 0, x[
                    1].wgt, 0, 0
                si.wgtprep, si.wgttrans, si.wgtused, si.qtybck, si.wgttmptrans = 0, 0, 0, 0, 0
                if x[1].pkno in pkbyns:
                    si.pkid = pkbyns[x[1].pkno].id
                else:
                    si.pkid = x[1].pkno
                nbtmap[si.name] = si
        if usgs:
            joshd = self._getjoshpdate(jes)
            for x in usgs:
                s0 = x.type
                if s0 not in ionmap:
                    nb = StoneBck()
                    nbck.append(nb)
                    ncbt.add(x.btchno)
                    if x.btchno in btbyns:
                        nb.btchid = btbyns[x.btchno].id
                    else:
                        nb.btchid = x.btchno
                    nb.idx, nb.filldate, nb.lastupdate = 1, td, td
                    nb.lastuserid, nb.qty, nb.wgt = 1, x.qty, x.wgt
                    nb.docno = "AG" + td.strftime("%y%m%d")[1:]
                else:
                    for iof in ionmap[s0]:
                        key = x.jono.value + "," + str(iof[0])
                        somso = sos.setdefault(key, {})
                        if len(somso) == 0:
                            som = StoneOutMaster()
                            som.joid = jobyns[x.jono].id
                            somso["som"] = som
                            msomid += 1
                            iof[1] += 1
                            som.id, som.isout, som.name = msomid, iof[0], iof[1]
                            som.packed, som.qty, som.subcnt, som.workerid = 0, 0, 0, 1393
                            som.filldate, som.lastupdate, som.lastuserid = joshd.get(
                                x.jono, td), td, 1
                        else:
                            som = somso["som"]
                        lst1 = somso.setdefault("sos", [])
                        so = StoneOut()
                        lst1.append(so)
                        so.id, so.idx, so.joqty, so.lastupdate, so.lastuserid = som.id, len(
                            lst1), 0, td, 1
                        so.printid, so.qty, so.wgt, so.workerid = 0, x.qty, x.wgt, 1393
                        so.checkerid, so.checkdate = 0, som.filldate
                        if x.btchno in btbyns:
                            so.btchid = btbyns[x.btchno].id
                        else:
                            so.btchid = x.btchno
        return nbtmap, sos, nbck, ncbt, btbyns

    def _persist(self, nbt, sos, nbck, ncbt, btbyns):
        err = True
        with self._cnsvc.sessionctx() as cur:
            for x in nbt.items():
                x[1].qty = int(x[1].qty) if x[1].qty else 0
                cur.add(x[1])
            cur.flush()
            self._persist_bck(nbck, cur, nbt)
            for x in sos.items():
                cur.add(x[1]["som"])
                for y in x[1]["sos"]:
                    if isinstance(y.btchid, str):
                        y.btchid = nbt[y.btchid].id
                    cur.add(y)
            cur.flush()
            ctag = int(datetime.today().strftime("%m%d"))
            for x in ncbt:
                btno = btbyns[x] if x in btbyns else nbt[x]
                btno.tag = ctag
                cur.add(btno)
            cur.flush()
            cur.commit()
            err = False
        return not err
    
    def _persist_bck(self, nbck, cur, nbt):
        ''' persist the stone_bck items '''
        if not nbck:
            return
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
        cur.flush()

    def readst(self, fn):
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
        rst = self._buildrst(*self._readfrmfile(fn))
        if not rst:
            #some invalid items found by _buildrst, throw exception
            return None
        return self._persist(*rst)
