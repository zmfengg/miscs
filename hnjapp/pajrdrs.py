# coding=utf-8
'''
Created on Apr 17, 2018

the replacement of the Paj Shipment Invoice Reader, which was implmented
in PAJQuickCost.xls#InvMatcher

@author: zmFeng
'''

from collections import namedtuple
from datetime import datetime, date
import datetime as dtm
import numbers
import os
import sys
import re
from decimal import Decimal

from xlwings.constants import LookAt

import logging as logging
from hnjcore import JOElement
from hnjcore import xwu, p17u, appathsep, deepget
from hnjcore.models.hk import PajShp, PajInv
import xlwings.constants as const
from quordrs import DAO
from dbsvcs import HKSvc, CNSvc

_accdfmt = "%Y-%m-%d %H:%M:%S"


def _accdstr(dt):
    """ make a date into an access date """
    return dt.strftime(_accdfmt) if dt and isinstance(dt, date) else dt


def _removenonascii(s0):
    """remove thos non ascii characters from given string"""
    if isinstance(s0, basestring):
        return "".join([x for x in s0 if ord(x) > 31 and ord(x) < 127])
    return s0


def _getjoids(jonos, hnjhkdb):
    """get the joIds by the provided jonos
    @param jonos: a list of JOElement  
    """

    rc = None
    s0 = "or".join([" (alpha = '%s' and digit = %d) " %
                    (x.alpha, x.digit) for x in jonos])
    s0 = "select alpha,digit,joid from jo where (%s)" % s0
    cur = hnjhkdb.cursor()
    try:
        cur.execute(s0)
        rows = cur.fetchall()
        if rows:
            rc = dict((JOElement(x.alpha, x.digit).value, x.joid)
                      for x in rows)
    finally:
        if cur:
            cur.close()
    return rc


class ShpReader:

    def __init__(self, accdb, hksvc):
        self._accessdb = accdb
        self._hksvc = hksvc
        self._ptnfd = re.compile(r"\d+")
        self._ptngwt = re.compile(r"(\d*\.\d*)\s*[gG]+")
        self._ptnswt = re.compile(r"\d*\.\d*")

    def _hasread(self, fn):
        """
            check if given file has been read, by default, from AccessDb
            @param fn: the full-path filename 
        """
        rc = 0
        cur = self._accessdb.cursor()
        try:
            rfn = os.path.basename(fn).replace("_", "")
            cur.execute(
                r"SELECT max(lastModified) from jotop17 where fn='%s'" % rfn)
            row = cur.fetchone()
            if row and row[0]:
                rc = 2 if row[0] < datetime.fromtimestamp(os.path.getmtime(fn)).replace(microsecond=0)\
                    else 1
        except BaseException as e:
            rc = -1
            print(e)
        finally:
            if cur:
                cur.close()
        return rc

    def _getshpdate(self, fn, isfile=True):
        """extract the shipdate from file name"""
        import datetime as dt
        parts = self._ptnfd.findall(os.path.basename(fn))
        if not parts:
            return

        try:
            d0 = None
            parts = [int(x) for x in parts]
            lg = len(parts)
            if lg >= 3:
                d0 = dt.date(parts[0], parts[1], parts[2])
                # mm-ddxxxx_01.xlsx case, the first not year
                if d0.year < 2017:
                    parts = (parts[0], parts[1])
                    lg = 2
                    d0 = None
            if lg == 2 and isfile:
                # only month/date,guess the year
                d1 = dt.date.fromtimestamp(os.path.getmtime(fn))
                d0 = dt.date(d1.year, parts[0], parts[1])
                df = d1 - d0
                if df.days < -3:
                    d0 = dt.date(d0.year - 1, d0.month, d0.day)
        except:
            d0 = None
        return d0

    def _readquodata(self, sht, qmap):
        """extract gold/stone weight data from the QUOXX sheet
        @param sht:  the DL_QUOTATION sheet that need to read data from
        @param qmap: the dict with p17 as key and (goldwgt,stwgt) as value
        """
        rng = xwu.find(sht, "Item*No", lookat=LookAt.xlPart)
        if not rng:
            return
        # because there is merged cells rng.expand('table').value
        # or sht.range(rng.end('right'),rng.end('down')).value failed
        vvs = sht.range(rng, rng.current_region.last_cell).value
        nmap = {"p17": 0}
        tr = vvs[0]
        for it in {"stone": re.compile(r"stone\s*weight", re.I),
                   "metal": re.compile(r"metal\s*weight", re.I)}.iteritems():
            cns = [x for x in range(len(tr)) if it[1].search(tr[x])]
            if len(cns) > 0:
                nmap[it[0]] = cns[0]
        for x in range(1, len(vvs)):
            tr = vvs[x]
            p17 = tr[nmap["p17"]]
            if not p17:
                continue
            if p17u.isvalidp17(p17) and not p17 in qmap:
                sw = 0 if not tr[nmap["stone"]] else \
                    sum([float(x)
                         for x in self._ptnswt.findall(tr[nmap['stone']])])
                mtls = tr[nmap['metal']]
                mw = mtls if isinstance(mtls, numbers.Number) else sum(
                    [float(x) for x in self._ptngwt.findall(mtls)])
                qmap[p17] = (mw, sw)

    def _persist(self, dups, items):
        """save the data to db
        @param dups: a list contains file names that need to be removed
        @param items: all the ShipItems that need to be persisted
        """

        if len(dups) + len(items) <= 0:
            return 0, None
        cur = self._accessdb.cursor()
        sess = self._hksvc.session()
        e = None
        try:
            if dups:
                cur.execute("delete from jotoP17 where fn in ('%s')" %
                            "','".join([os.path.basename(x) for x in dups]))
                # maybe a little stupid: find and delete
                sess.query(PajShp).filter(PajShp.fn.in_([_removenonascii(os.path.basename(x)) for x in dups]))\
                    .delete(synchronize_session=False)
            if items:
                dcts = list([x._asdict() for x in items.values()])
                jns = [JOElement(x.jono) for x in items.values()]
                jns = self._hksvc.getjos(jns)[0]
                jns = dict([(x.name,x) for x in jns])                
                for dct in dcts:
                    dc1 = dict(dct)
                    dc1["fn"] = _removenonascii(dct["fn"])
                    dc1["joid"] = jns[JOElement(dct["jono"])].id
                    dc1["pcode"] = dc1["p17"]
                    dc1["orderno"] = dc1["ordno"]
                    for x in "fillDate,lastModified,invDate,shpDate".split(","):
                        dct[x] = _accdstr(dct[x])
                    shp = PajShp()
                    for x in dc1.iteritems():
                        k = x[0]
                        lk = k.lower()
                        if hasattr(shp, lk):
                            shp.__setattr__(lk, dc1[k])
                    cur.execute(("insert into jotop17 (fn,jono,p17,fillDate,lastModified,invno,qty"
                                 ",InvDate,ShpDate,OrdNo,MtlWgt,StWgt) values('%(fn)s','%(jono)s','%(p17)s'"
                                 ",#%(fillDate)s#,#%(lastModified)s#,'%(invno)s',%(qty)f,#%(invDate)s#"
                                 ",#%(shpDate)s#,'%(ordno)s',%(mtlWgt)f,%(stWgt)f)") % dct)
                    sess.add(shp)
        except Exception as e:
            logging.debug("Error occur in ShpRdr:%s" % e)
        finally:
            if e:
                sess.rollback()
                self._accessdb.rollback()
            else:
                sess.commit()
                self._accessdb.commit()
            if cur:
                cur.close()
            if sess:
                sess.close()
        return -1 if e else 1, e

    def read(self, fldr):
        """
        read the shipment file and send to 2dbs
        @param fldr: the folder contains the files. sub-folders will be ignored 
        """

        ptn = re.compile(r"HNJ\s+\d*-", re.IGNORECASE)
        fns = [appathsep(fldr) + unicode(x, sys.getfilesystemencoding())
               for x in os.listdir(fldr) if ptn.match(x)]
        if not fns:
            return
        errors = list()
        killxls, app = xwu.app(False)
        PajShpItem = namedtuple("PajShpItem", "fn,ordno,jono,qty,p17,invno,invDate" +
                                ",mtlWgt,stWgt,shpDate,lastModified,fillDate")
        try:
            td0 = datetime.today()
            for fn in fns:
                idx = self._hasread(fn)
                toRv = list()
                items = {}
                if idx == 1:
                    # logging.debug("%s has been read" % fn)
                    continue
                elif idx == 2:
                    logging.debug("%s is expired" % fn)
                    toRv.append(fn)
                lmd = datetime.fromtimestamp(os.path.getmtime(fn))
                shd1 = self._getshpdate(fn)
                shd0 = shd1
                logging.debug("processing file(%s) of date(%s)" % (fn, shd0))
                wb = app.books.open(fn)
                try:
                    # in new sample case, use DL_QUOXXX sheet's weight, so prepare it if there is
                    qmap = {}
                    for sht in wb.sheets:
                        if sht.name.find(u"返修") >= 0:
                            continue
                        rng = xwu.find(sht, u"十七*", lookat=LookAt.xlPart)
                        if not rng:
                            rng = xwu.find(sht, u"物料*", lookat=LookAt.xlPart)
                        if not rng:
                            continue
                        # don't use this, sometimes the stupid user skip some table header
                        # vvs = rng.end('left').expand("table").value
                        vvs = xwu.usedrange(sht).value
                        th = vvs[0]
                        tm = xwu.listodict(th, {u"订单号": "odx", u"发票日期": "invdate", u"订单序号": "odseq",
                                                u"平均单件石头,XXX": "stwgt", u"发票号": "invno", u"订单号序号": "ordno", u"十七位,十七,物料": "p17",
                                                u"平均单件金,XX": "mtlwgt", u"工单,job": "jono", u"数量": "qty", u"cost": "cost"})
                        x = [x for x in "invno,p17,jono,qty,invdate".split(
                            ",") if x not in tm]
                        if x:
                            logging.debug(
                                "failed to find key columns(%s) in sheet(%s)" % (x, sht.name))
                            continue
                        bfn = os.path.basename(fn).replace("_", "")
                        shd = self._getshpdate(sht.name, False)
                        if shd:
                            df = shd - shd1
                            shd = shd if abs(df.days) <= 7 else shd0
                        else:
                            shd = shd0
                        # finally I give up, don't use the shipdate, use invdate as shipdate
                        if "mtlwgt" in tm:
                            for ridx in range(1, len(vvs)):
                                tr = vvs[ridx]
                                if not tr[tm['p17']]:
                                    break
                                mwgt = tr[tm["mtlwgt"]]
                                if not (isinstance(mwgt, numbers.Number) and mwgt > 0):
                                    continue
                                invno = tr[tm["invno"]] if [
                                    tm["invno"]] else "N/A"
                                if 'ordno' in tm:
                                    odno = tr[tm['ordno']]
                                elif len([1 for x in "odx,odseq".split(",") if x in tm]) == 2:
                                    odno = tr[tm['odx']] + \
                                        "-" + tr[tm["odseq"]]
                                else:
                                    odno = "N/A"
                                if not tr[tm["stwgt"]]:
                                    tr[tm["stwgt"]] = 0
                                jono = JOElement(tr[tm["jono"]]).value
                                thekey = jono + "," + \
                                    tr[tm['p17']] + "," + invno
                                if items.has_key(thekey):
                                    si = items[thekey]
                                    items[thekey] = si._replace(
                                        qty=si.qty + tr[tm["qty"]])
                                else:
                                    ivd = tr[tm['invdate']]
                                    si = PajShpItem(bfn, odno, jono, tr[tm["qty"]], tr[tm['p17']],
                                                    invno, ivd, mwgt, tr[tm['stwgt']], ivd, lmd, td0)
                                    items[thekey] = si
                        else:
                            # new sample case, extract weight data from the quo sheet
                            if not qmap:
                                for x in [xx for xx in wb.sheets if xx.api.Visible == -1 and xx.name.lower().find('dl_quotation') >= 0]:
                                    self._readquodata(x, qmap)
                            if qmap:
                                import random
                                for x in range(1, len(vvs)):
                                    tr = vvs[x]
                                    # no cost item means repairing
                                    if "cost" in tm and not tr[tm["cost"]]:
                                        continue
                                    odno = tr[tm['ordno']] if tm.has_key(
                                        'ordno') else "N/A"
                                    p17 = tr[tm['p17']]
                                    if not p17:
                                        break
                                    if p17 in qmap:
                                        ivd = tr[tm['invdate']]
                                        si = PajShpItem(bfn, odno, JOElement(tr[tm["jono"]]).value, tr[tm["qty"]], p17,
                                                        tr[tm["invno"]], ivd, qmap[p17][0], qmap[p17][1], ivd, lmd, td0)
                                    else:
                                        logging.critical(
                                            "failed to get quoinfo for pcode(%s)" % p17)
                                    # new sample won't have duplicated items
                                    items[random.random()] = si
                            else:
                                qmap["_SIGN_"] = 0
                finally:
                    if wb:
                        wb.close()
                x = self._persist(toRv, items)
                if x[0] <> 1:
                    errors.append(x[1])
                    logging.critical("file(%s) contains errors" %
                                     os.path.basename(fn))
                    logging.critical(x[1])
        finally:
            if killxls:
                app.quit()
        return -1 if errors > 0 else 1, errors


class InvReader(object):
    """
    read the invoices(17PMXXX or alike) from given folder, generate data for hk
    """

    def __init__(self, accdb, hksvc):
        self._accessdb = accdb
        self._hksvc = hksvc

    def _getinv(self, fn):
        """
            extract the inv# from the file name
            @param fn: the full path just the file name 
        """
        s0 = os.path.basename(fn)
        return s0[:s0.find(".")].strip()

    def _hasread(self, fn):
        """
            check if given file has been read, by default, from AccessDb
            @param fn: the full-path filename 
        """
        rc = 0
        cur = self._accessdb.cursor()
        try:
            cur.execute(
                r"select lastModified from PajInv where invno = '%s'" % self._getinv(fn))
            row = cur.fetchone()
            if row and row[0]:
                rc = 2 if row[0] < \
                    datetime.fromtimestamp(os.path.getmtime(fn)).replace(microsecond=0) else 1
        finally:
            if cur:
                cur.close()
        return rc

    def _getjonos(self, p17, invno):
        """
        return the JO#s of the given p17 and inv#
        @param p17:     a valid p17 code
        @param invno:   the invoice# 
        """
        cur = self._accessdb.cursor()
        try:
            cur.execute(
                "select jono from jotop17 where p17 = '%s' and invno = '%s'" % (p17, invno))
            rows = cur.fetchall()
            rc = set(x.jono.trim() for x in rows) if rows else None
        finally:
            if cur:
                cur.close()
        return rc

    def _persist(self, dups, items):
        """persist the data
        @param dups:  a list of invnos
        @param items: the InvItems that need to be inserted
        """
        x = (len(dups), len(items))
        e = None
        if any(x):
            try:
                # maybe the insert/delete should be implemented by the hksvr
                cur = self._accessdb.cursor()
                sess = self._hksvc.session()
                if x[0]:
                    lst = list(dups)
                    cur.execute(
                        "delete from PajInv where invno in ('%s')" % "','".join(lst))
                    sess.query(PajInv).filter(PajInv.invno.in_(lst))\
                        .delete(synchronize_session=False)
                if x[1]:
                    dcts = list([x0._asdict() for x0 in items.values()])
                    jns = [JOElement(x0.jono) for x0 in items.values()]
                    jns = self._hksvc.getjos(jns,psess = sess)[0]
                    jns = dict([(x.name,x) for x in jns])
                    for dct in dcts:
                        # todo::make the china value for the user
                        dct["china"] = 0
                        dc1 = dict(dct)
                        dc1["joid"] = jns[JOElement(dc1["jono"])].id
                        dc1["stspec"] = dc1["stone"]
                        dc1["uprice"] = dc1["price"]
                        for y in "lastmodified".split(","):
                            dct[y] = _accdstr(dct[y])
                        iv = PajInv()
                        for it in dc1.iteritems():
                            k = it[0]
                            lk = it[0].lower()
                            if hasattr(iv, lk):
                                iv.__setattr__(lk, dc1[k])
                        cur.execute(("insert into PajInv(Invno,Jono,StSpec,Qty,UPrice,Mps,LastModified) values"
                                     "('%(invno)s','%(jono)s','%(stone)s',%(qty)f,%(price)f,'%(mps)s',#%(lastmodified)s#)") % dct)
                        iv = sess.add(iv)
            except Exception as e:  # there might be pyodbc.IntegrityError if dup found
                logging.debug("Error occur in InvRdr:%s" % e.message)
            finally:
                if not e:
                    sess.commit()
                    self._accessdb.commit()
                else:
                    sess.rollback()
                    self._accessdb.rollback()
                cur.close()
                sess.close()

        return (0 if sum(x) == 0 else -1 if e else 1), e

    def read(self, invfldr, writeJOBack=True):
        """
        read files back, instead of using os.walk(root), use os.listdir()
        @param invfldr: the folder contains the invoices
        @param writeJOBack: write the JO# back to the source sheet 
        """

        if not os.path.exists(invfldr):
            return
        invs = {}
        errs = list()
        PajInvItem = namedtuple(
            "PajInvItem", "invno,p17,jono,qty,price,mps,stone,lastmodified")
        fns = [appathsep(invfldr) + unicode(x, sys.getfilesystemencoding()) for x in os.listdir(invfldr)
               if x.lower()[2:4] == "pm" and x.lower().find(".xls")]
        if not fns:
            return
        killexcel, app = xwu.app(False)

        for fn in fns:
            invno = self._getinv(fn).upper()
            if invno in invs:
                continue
            dups = []
            idx = self._hasread(fn)
            if idx == 1:
                # logging.debug("%s has been read" % fn)
                continue
            elif idx == 2:
                logging.debug("%s is expired" % fn)
                dups.append(invno)
            lmd = datetime.fromtimestamp(os.path.getmtime(fn))
            wb = app.books.open(fn)
            updcnt = 0
            items = {}
            try:
                for sh in wb.sheets:
                    rng = xwu.find(sh, "Invo No:")
                    if not rng:
                        continue
                    rng = xwu.find(sh, "Item*No", lookat=const.LookAt.xlWhole)
                    if not rng:
                        continue
                    rng = rng.expand("table")
                    vals = rng.value
                    # table header map
                    tm = {}
                    tr = [xx.lower() for xx in vals[0]]
                    tm = xwu.listodict(tr, {"gold,": "gold", "silver,": "silver", u"job#,工单": "jono",
                                            "price,": "price", "unit,": "qty", "stone,": "stone"})
                    tm["p17"] = 0
                    x = [x for x in "price,qty,stone".split(
                        ",") if not tm.has_key(x)]
                    if x:
                        logging.info(
                            "key columns(%s) missing in sheet('%s') of file (%s)" % (x, sh.name, fn))
                        continue
                    for jj in range(1, len(vals)):
                        tr = vals[jj]
                        if not tr[tm["price"]]:
                            continue
                        p17 = tr[tm["p17"]]
                        if not p17u.isvalidp17(p17):
                            logging.debug(
                                "invalid p17 code(%s) in %s" % (p17, fn))
                            continue
                        jn = JOElement(tr[tm["jono"]]).value if tm.has_key(
                            "jono") else None
                        if not jn:
                            jns = self._getjonos(p17, invno)
                            if jns:
                                jn = jns[0]
                            if jn and writeJOBack:
                                sh.range(rng.row + jj, rng.column +
                                         tm["jono"]).value = jn
                                updcnt += 1
                        else:
                            jn = "%d" % jn if isinstance(
                                jn, numbers.Number) else jn.strip()
                        if not jn:
                            logging.debug(
                                "No JO# found for p17(%s) in file %s" % (tr[tm["p17"]], fn))
                            continue
                        key = invno + "," + jn
                        if items.has_key(key):
                            it = items[key]
                            items[key] = it._replace(
                                qty=it.qty + tr[tm["qty"]])
                        else:
                            mps = "S=%3.2f;G=%3.2f" % (tr[tm["silver"]], tr[tm["gold"]]) \
                                if tm.has_key("gold") and tm.has_key("silver") else "S=0;G=0"
                            it = PajInvItem(
                                invno, p17, jn, tr[tm["qty"]], tr[tm["price"]], mps, tr[tm["stone"]], lmd)
                            items[key] = it
            finally:
                if updcnt > 0:
                    wb.save()
                wb.close()
            x = (0, None) if len(dups) + \
                len(items) == 0 else self._persist(dups, items)
            if(x[0] == -1):
                errs.append(x[1])
            else:
                logging.debug("invoice (%s) processed" %
                              fn + ("" if x[0] else " but all are repairing"))
        if killexcel:
            app.quit()
        return x[0], items if x[0] == 1 else () if x[0] == 0 else x[1]


class PAJCReader(object):
    """class to create the PAJ JOCost file for HK accountant"""

    def __init__(self, hksvc=None, cnsvc=None, bcdb=None):
        self._hksvc = hksvc
        self._cnsvc = cnsvc
        self._bcdb = bcdb

    def run(self, year, month, tplfn=None, tarfldr=None):
        """ create report file of given year/month"""

        def _makemap(sht=None):
            coldefs = (u"invoice date=invdate;invoice no.=invno;order no.=orderno;customer=cstname;"
                       u"job no.=jono;style no.=styno;running no.=running;paj item no.=pcode;karat=karat;"
                       u"描述=cdesc;in english=edesc;job quantity=joqty;quantity received=shpqty;"
                       u"total cost=ttlcost;cost=uprice;平均单件金银重g=umtlwgt;平均单件石头重g=ustwgt;石头=stspec;"
                       u"mm program in#=iono;jmp#=jmpno;date=shpdate;remark=rmk;has dia=hasdia")
            vvs = sht.range("A1").expand("right").value
            vvs = [x.lower() if isinstance(x, basestring) else x for x in vvs]
            imap = {}
            nmap = {}
            for s0 in coldefs.split(";"):
                ss0 = s0.split("=")
                x = [x for x in range(
                    len(vvs)) if x not in imap and vvs[x].find(ss0[0]) >= 0]
                if x:
                    imap[x[0]] = ss0[1]
                    nmap[ss0[1]] = x[0]
                else:
                    print("failed to get colname %s" % s0)

            return nmap, imap

        dfmt = "%Y/%m/%d"
        df = dtm.date(year, month, 1)
        if False:
            #test purpose, get only 2 day
            dt = dtm.date(year, month, 3)
        else:
            month += 1
            if month > 12:
                year += 1
                month = 1        
            dt = dtm.date(year, month, 1)

        runns = set()
        jes = set()
        dao = DAO(self._bcdb)

        mms = self._cnsvc.getshpforjc(df, dt)
        for x in mms:
            jo = x.JO
            rn = str(jo.running)
            if rn not in runns:
                runns.add(rn)
            jn = jo.name
            if jn not in jes:
                jes.add(jn)
        runns = tuple(runns)
        bcs = dict([(x.runn.strip(), x.desc.strip()) for x
                    in dao.getbcsforjc(runns)])
        lst = self._hksvc.getpajshpinv(jes)
        pajs = {}
        pajsjn = {}
        for x in lst:
            jn = x.JO.name
            pajs["%s,%s" % (jn, x.PajShp.invdate.strftime(dfmt))] = x
            jn = jn.value
            if jn not in pajsjn:
                pajsjn[jn] = []
            lst1 = pajsjn[jn]
            lst1.append(x)
        ios = dict([("%s,%s,%s" % (x.running, x.jmp, x.shpdate.strftime(dfmt)), x) for x
                    in self._hksvc.getmmioforjc(df, dt, runns)])
        killxls, app = xwu.app(False)
        lst = []
        fn = None
        try:
            wb = xwu.fromtemplate(tplfn, app)
            sht = wb.sheets("Data")
            nmps = _makemap(sht)
            ss = ("cstname,Customer.name,karat,JO.karat,cdesc,JO.description,joqty"
                    ",JO.qty,jmpno,MM.name,shpdate,MMMa.refdate,shpqty,MM.qty").split(",")
            dtmap0 = dict(zip(ss[0:len(ss) - 1:2], ss[1:len(ss):2]))
            ss = ("invdate,PajShp.invdate,invno,PajShp.invno,orderno,PajShp.orderno"
                    ",pcode,PajShp.pcode,uprice,PajInv.uprice,umtlwgt,PajShp.mtlwgt"
                    ",ustwgt,PajShp.stwgt,stspec,PajInv.stspec").split(",")
            dtmap1 = dict(zip(ss[0:len(ss) - 1:2], ss[1:len(ss):2]))

            for row in mms:
                mp = {}
                rn = str(row.JO.running)
                jn = row.JO.name.value
                for x in dtmap0.iteritems():
                    mp[x[0]] = deepget(row,x[1])
                mp["running"] = rn
                mp["jono"] = "'" + jn
                mp["styno"] = row.Style.name.value
                mp["edesc"] = bcs[rn] if rn in bcs else "N/A"

                key1 = jn
                key = "%s,%s" % (jn, mp["shpdate"].strftime(dfmt))
                fnd = False
                if key in pajs:
                    x = pajs[key]
                    for y in dtmap1.iteritems():
                        mp[y[0]] = deepget(x,y[1])
                    fnd = True
                elif key1 in pajsjn:
                    lst1 = pajsjn[key1]
                    if lst1:
                        hts = []
                        shpd = mp["shpdate"]
                        for x in lst1:
                            ddiff = x.PajShp.shpdate - shpd
                            if (abs(ddiff.days) <= 5 and abs(x.PajShp.qty - float(mp["shpqty"])) < 0.1):
                                hts.append(x)
                                for y in dtmap1.iteritems():
                                    mp[y[0]] = deepget(x,y[1])
                                fnd = True
                                break
                        if hts:
                            for x in hts:
                                lst1.remove(x)
                        if not lst1:
                            del pajsjn[key1]
                if fnd:
                    mp["ttlcost"] = mp["uprice"] * mp["shpqty"]
                if not fnd:
                    for x in dtmap1.keys():
                        mp[x] = None
                    mp["ttlcost"] = None
                key = "%s,%s,%s" % (
                    rn, mp["jmpno"], mp["shpdate"].strftime(dfmt))
                mp["rmk"] = ("QtyError" if not (mp["joqty"] and mp["shpqty"]) else
                             "" if mp["joqty"] == mp["shpqty"] else "Partial")
                mp["iono"] = ios[key].inoutno if key in ios else "N/A"
                hasdia = (mp["cdesc"].find(u"钻") >= 0 or mp["cdesc"].find(u"占") >= 0 or
                          (mp["edesc"] and mp["edesc"].lower().find("dia") >= 0))
                mp["hasdia"] = "D" if hasdia else "N"

                x = [mp[nmps[1][x]] for x in range(len(nmps[1]))]
                lst.append(["" if not y else y.strip() if isinstance(y, basestring) else
                            y.strftime(dfmt) if isinstance(y, datetime) else float(y) if isinstance(y, Decimal) else
                            y for y in x])
            sht.range("A2").value = lst
            for x in [x for x in wb.sheets if x != sht]:
                x.delete()
            if tarfldr:
                fn = appathsep(tarfldr) + df.strftime("%Y%m")
                wb.save(fn)
        finally:
            if killxls and not wb:
                app.quit()
            else:
                app.visible = True
        return lst, fn


class PriceTracker(object):
    """ class to keep track of Pcode price changes
    to use this method, put a dat file inside a folder which should contains sty#
    then I will read and show the price trends
    """

    def __init__(self, hkdb):
        self._hkdb = hkdb

    def read(self, fldr):
        if not fldr:
            return
        fldr = appathsep(fldr)
        fns = [unicode(x, sys.getfilesystemencoding()) for x in os.listdir(fldr)
               if x.lower().find("dat") >= 0]
        if not fns:
            return
        stynos = set()
        for x in fns:
            with open(fldr + x, "wb") as fh:
                for ln in fh:
                    je = JOElement(ln)
                    if je.isvalid and not je in stynos:
                        stynos.add(je)
        dao = DAO(self._hkdb)
        #lst = dao.getpajprices(stynos)
