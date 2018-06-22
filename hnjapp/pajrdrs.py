# coding=utf-8
'''
Created on Apr 17, 2018

the replacement of the Paj Shipment Invoice Reader, which was implmented
in PAJQuickCost.xls#InvMatcher

@author: zmFeng
'''

import datetime as dtm
import numbers
import os
import re
import sys
from collections import namedtuple
from datetime import date, datetime
from decimal import Decimal

import xlwings.constants as const
from xlwings.constants import LookAt

from hnjcore import JOElement, appathsep, deepget, karatsvc, p17u, xwu
from hnjcore.models.hk import PajInv, PajShp, Orderma,Style as Styhk, JO as JOhk
from hnjcore.utils import getfiles, daterange, p17u, isnumeric
from hnjcore.utils.consts import NA

from .common import _logger as logger
from .dbsvcs import CNSvc, HKSvc, BCSvc
from .pajcc import P17Decoder, PrdWgt, WgtInfo
from sqlalchemy.orm import Query

_accdfmt = "%Y-%m-%d %H:%M:%S"

def _accdstr(dt):
    """ make a date into an access date """
    return dt.strftime(_accdfmt) if dt and isinstance(dt, date) else dt


def _removenonascii(s0):
    """remove thos non ascii characters from given string"""
    if isinstance(s0, str):
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

def readbom(fldr):
    """ read BOM from given folder
    @param fldr: the folder contains the BOM file(s)
    return a dict with "pcode" as key and dict as items
        the item dict has keys("pcode","wgts")
    """
    _ptnoz = re.compile(r"\(\$\d*/OZ\)")
    _ptnsil = re.compile(r"(925)")
    _ptngol = re.compile(r"^(\d*)K")    
    _ptdst = re.compile(r"[\(（](\d*)[\)）]")
    #the parts must have karat, if not, follow the parent
    _mtlpts = u"金,银,耳勾,线圈,耳针,Chain".lower().split(",")

    _pcdec = P17Decoder()
    
    def _parsekarat(mat,wis = None ,ispol =True):
        """ return karat from material string """
        kt = None
        if ispol:
            mt = _ptnoz.search(mat)
            if not mt: return        
        if _ptnsil.search(mat):
            kt = 925
        else:
            mt = _ptngol.search(mat)
            if mt:
                kt = int(mt.group(1))
            else:
                mt = _ptdst.search(mat)
                if mt:
                    kt = int(mt.group(1))
                    if not karatsvc.getkarat(kt):
                        karat = karatsvc.getbyfiness(kt)
                        if karat: kt = karat.karat
        if not kt:
            #not found, has must have keyword? if yes, follow master
            if wis and any(wis):
                s0 = mat.lower()
                for x in _mtlpts:
                    if s0.find(x) >= 0:
                        if s0.find(u"银") >= 0:
                            kt = 925
                        elif s0.find(u"金") >= 0:
                            for wi in wis:
                                if not wi: continue
                                karat = karatsvc.getkarat(wi.karat)
                                if karat and karat.category == karatsvc.CATEGORY_GOLD:
                                    kt = wi.karat
                                    break
                        #finally no one is found, follow master
                        if not kt: kt = wis[0].karat
                    if kt: break                
            else:
                logger.error("No karat found for (%s) and no default provided" % mat)
                kt = -1
        if kt and kt > 0:
            karat = karatsvc.getkarat(kt)
            if not karat: kt = -1
        return kt

    def _ispendant(pcode):       
        return _pcdec.decode(pcode,"PRODTYPE") == '3'
    
    def _mergewgt(wgts,wgt,maxidx,tryins = False):
        sltid = -1
        for ii in range(maxidx + 1):
            if wgts[ii]:
                if wgts[ii].karat == wgt.karat:
                    wis[ii] = WgtInfo(wgt.karat, wgt.wgt + wgts[ii].wgt)
                    return ii
            else:
                if tryins and sltid < 0:
                    sltid = ii
                    break
        if sltid >= 0:
            wgts[sltid] = wgt            
        return sltid

    if os.path.isdir(fldr):
        fns = getfiles(fldr,"xls")
    else:
        fns = [fldr]
    if not fns: return

    kxl,app = xwu.app(False)
    try:
        pmap = {}
        for fn in fns:
            wb = app.books.open(fn)
            shts = [0,0]
            for sht in wb.sheets:
                rng = xwu.find(sht, u"十七位")
                if not rng: continue
                if xwu.find(sht, u"抛光后"):
                    shts[0] = (sht,rng)
                elif xwu.find(sht, u"物料"):
                    shts[1] = (sht,rng)
            if not all(shts): break
            for jj in range(len(shts)):
                vvs = shts[jj][1].end("left").expand("table").value
                if jj == 0:                    
                    wcn = xwu.list2dict(vvs[0],{u"十七位,":"pcode",u"材质,":"mat",u"抛光,":"mtlwgt"})
                    for ii in range(1,len(vvs)):
                        pcode = vvs[ii][wcn["pcode"]]
                        if not p17u.isvalidp17(pcode): break
                        kt = _parsekarat(vvs[ii][wcn["mat"]])
                        if not kt: continue
                        it = pmap.setdefault(pcode,{"pcode":pcode})
                        it.setdefault("wgts",[]).append((kt,vvs[ii][wcn["mtlwgt"]]))
                elif jj == 1:
                    nmap = {u"十七位,":"pcode",u"物料名称":"name", \
                        u"物料特征":"spec",u"数量":"qty",u"重量":"wgt",u"单位":"unit"}
                    wcn = xwu.list2dict(vvs[0],nmap)
                    nmap = [x for x in nmap.values() if x.find("pcode") < 0]
                    for ii in range(1,len(vvs)):
                        pcode = vvs[ii][wcn["pcode"]]
                        if not p17u.isvalidp17(pcode): break
                        it = pmap.setdefault(pcode,{"pcode":pcode})
                        mats, it = it.setdefault("parts",[]), {}
                        mats.append(it)
                        for cn in nmap:
                            it[cn] = vvs[ii][wcn[cn]]                        
            wb.close()
        for x in pmap.items():
            lst = x[1]["wgts"]
            sltid, wis = -1, [None,None,None]            
            for y in lst:
                sltid = _mergewgt(wis,WgtInfo(y[0],y[1]),1,True)
                if sltid < 0:
                    logger.error("failed to get slot to store prodwgt for pcode %s" % x[0])
                    x[1]["wgts"] = None
                    continue
            if "parts" in x[1]:                
                if wis[1]:
                    if wis[0].wgt < wis[1].wgt:
                        wis[0],wis[1] = wis[1], wis[0]
                for y in x[1]["parts"]:
                    kt = _parsekarat(y["name"],wis,False)
                    y["karat"] = kt
                    if not kt: continue                
                    done = False
                    if _ispendant(x[0]) and y["name"].lower().find("chain") >= 0:
                        wgt0 = wis[2]
                        if not wgt0 or y["karat"] == wgt0.karat:
                            wgt = y["wgt"] + (wgt0.wgt if wgt0 else 0)
                            wis[2] = WgtInfo(y["karat"],wgt)
                            if wgt0:
                                logger.debug("Multi chain found for pcode(%s)" % x[0])
                            done = True
                        else:
                            logger.debug("No wgt pos for chain(%s) in pcode(%s),merged to main" % (y["name"],x[0]))                        
                    if not done:
                        _mergewgt(wis,WgtInfo(y["karat"],y["wgt"]),1)
            x[1]["wgts"] = PrdWgt(wis[0],wis[1],wis[2])        
    finally:
        if kxl and app: app.quit()
    return pmap

def readbom2jos(fldr,hksvc,fn = None,mindt = None):
    """ build a jo collection list based on the BOM file provided
        @param fldr: the folder contains the BOM file(s)
        @param hksvc: the HK db service
        @param fn: save the file to
        @param mindt: the minimum datetime the query fetch until
          if None is provided, it will be 2017/01/01
        return a workbook contains the result
    """
    def _fmtwgt(prdwgt):
        wgt = (prdwgt.main,prdwgt.aux,prdwgt.part)
        lst = []
        [lst.extend((x.karat,x.wgt) if x else (0,0)) for x in wgt]
        return lst
    def _samewgt(wgt0,wgt1):
        wis = []
        for x in (wgt0,wgt1):
            wis.append((x.main,x.aux,x.part))
        for i in range(3):
            wts = (wis[0][i],wis[1][i])
            eq = all(wts) or not(any(wts))
            if not eq: break
            if not all(wts): continue
            eq = wts[0].karat == wts[0].karat or \
                karatsvc.getfamily(wts[0].karat) == karatsvc.getfamily(wts[1].karat)
            if not eq: break
            eq = abs(round(wis[0][i].wgt - wis[1][i].wgt,2)) <= 0.02
        return eq

    pmap = readbom(fldr)
    ffn = None
    if not pmap: return ffn
    vvs = ["pcode,m.karat,m.wgt,p.karat,p.wgt,c.karat,c.wgt".split(",")]
    jos = ["Ref.pcode,JO#,Sty#,Run#,m.karat,m.wgt,p.karat,p.wgt,c.karat,c.wgt,rm.wgt,rp.wgt,rc.wgt".split(",")]
    if not mindt:
        mindt = dtm.datetime(2017,1,1)
    qp = Query(Styhk.id).join(Orderma, Orderma.styid == Styhk.id) \
        .join(JOhk,Orderma.id == JOhk.orderid).join(PajShp,PajShp.joid == JOhk.id)
    qj = Query([JOhk.name.label("jono"),Styhk.name.label("styno"),JOhk.running]) \
        .join(Orderma, Orderma.id == JOhk.orderid).join(Styhk).filter(JOhk.createdate >= mindt) \
        .order_by(JOhk.createdate)
    with hksvc.sessionctx() as sess:
        cnt = 0;ln = len(pmap)
        for x in pmap.values():
            lst, wgt = [x["pcode"]], x["wgts"]
            if isinstance(wgt,PrdWgt):
                lst.extend(_fmtwgt((wgt)))
            else:
                lst.extend((0,0,0,0,0,0))
            vvs.append(lst)

            pcode = x["pcode"]
            q = qp.filter(PajShp.pcode == pcode).limit(1).with_session(sess)
            try:
                sid = q.one().id
                q = qj.filter(Orderma.styid == sid).with_session(sess)
                lst1 = q.all()
                for jn in lst1:
                    jowgt = hksvc.getjowgts(jn.jono)
                    if not _samewgt(jowgt,wgt):
                        lst = [pcode,jn.jono.value,jn.styno.value,jn.running]
                        lst.extend(_fmtwgt(jowgt))
                        lst.extend(_fmtwgt(wgt)[1::2])
                        jos.append(lst)
                    else:
                        logger.debug("JO(%s) has same weight as pcode(%s)"\
                            % (jn.jono.value,pcode))
            except:
                pass
            
            cnt += 1
            if cnt % 20 == 0:
                print("%d of %d done" % (cnt,ln))

        app = xwu.app(False)[1]
        wb = app.books.add()
        sns, data = "BOMData,JOs".split(","), (vvs,jos)
        for idx in range(len(sns)):
            sht = wb.sheets[idx]
            sht.name = sns[idx]
            sht.range(1,1).value = data[idx]
            sht.autofit("c")
        wb.save(fn)
        ffn = wb.fullname
        app.quit()
    return ffn

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
                   "metal": re.compile(r"metal\s*weight", re.I)}.items():
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
        with self._hksvc.sessionctx() as sess:
            err = False
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
                        je = JOElement(dct["jono"])
                        if je not in jns or not p17u.isvalidp17(dct["p17"]):
                            logger.info("Item(%s) does not contains valid JO# or pcode" % dct)
                            continue
                        dc1 = dict(dct)
                        dc1["fn"] = _removenonascii(dct["fn"])
                        dc1["joid"] = jns[je].id
                        dc1["pcode"] = dc1["p17"]
                        dc1["orderno"] = dc1["ordno"]
                        for x in "fillDate,lastModified,invDate,shpDate".split(","):
                            dct[x] = _accdstr(dct[x])
                        shp = PajShp()
                        for x in dc1.items():
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
                logger.debug("Error occur in ShpRdr:%s" % e.args)
                err = True
            finally:
                if err:
                    sess.rollback()
                    self._accessdb.rollback()
                else:
                    sess.commit()
                    self._accessdb.commit()
                if cur:
                    cur.close()
        return -1 if err else 1, err

    def read(self, fldr):
        """
        read the shipment file and send to 2dbs
        @param fldr: the folder contains the files. sub-folders will be ignored 
        """

        ptn = re.compile(r"HNJ\s+\d*-", re.IGNORECASE)
        fns = getfiles(fldr,"xls",True)
        if fns:
            p = appathsep(fldr)
            fns = [p + x for x in fns if ptn.match(x)]
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
                    # logger.debug("%s has been read" % fn)
                    continue
                elif idx == 2:
                    logger.debug("%s is expired" % fn)
                    toRv.append(fn)
                lmd = datetime.fromtimestamp(os.path.getmtime(fn))
                shd1 = self._getshpdate(fn)
                shd0 = shd1
                logger.debug("processing file(%s) of date(%s)" % (os.path.basename(fn), shd0))
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
                        tm = xwu.list2dict(th, {u"订单号": "odx", u"发票日期": "invdate", u"订单序号": "odseq",
                                                u"平均单件石头,XXX": "stwgt", u"发票号": "invno", u"订单号序号": "ordno", u"十七位,十七,物料": "p17",
                                                u"平均单件金,XX": "mtlwgt", u"工单,job": "jono", u"数量": "qty", u"cost": "cost"})
                        x = [x for x in "invno,p17,jono,qty,invdate".split(
                            ",") if x not in tm]
                        if x:
                            logger.debug(
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
                                if thekey in items:
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
                                    odno = tr[tm['ordno']] if 'ordno' in tm else "N/A"
                                    p17 = tr[tm['p17']]
                                    if not p17:
                                        break
                                    if p17 in qmap:
                                        ivd = tr[tm['invdate']]
                                        si = PajShpItem(bfn, odno, JOElement(tr[tm["jono"]]).value, tr[tm["qty"]], p17,
                                                        tr[tm["invno"]], ivd, qmap[p17][0], qmap[p17][1], ivd, lmd, td0)
                                    else:
                                        logger.critical(
                                            "failed to get quoinfo for pcode(%s)" % p17)
                                    # new sample won't have duplicated items
                                    items[random.random()] = si
                            else:
                                qmap["_SIGN_"] = 0
                finally:
                    if wb:
                        wb.close()
                x = self._persist(toRv, items)
                if x[0] != 1:
                    errors.append(x[1])
                    logger.critical("file(%s) contains errors" %
                                     os.path.basename(fn))
                    logger.critical(x[1])
        finally:
            if killxls:
                app.quit()
        return -1 if len(errors) > 0 else 1, errors


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
            cur = self._accessdb.cursor()
            with self._hksvc.sessionctx() as sess:
                try:
                    # maybe the insert/delete should be implemented by the hksvr                
                    if x[0]:
                        lst = list(dups)
                        cur.execute(
                            "delete from PajInv where invno in ('%s')" % "','".join(lst))
                        sess.query(PajInv).filter(PajInv.invno.in_(lst))\
                            .delete(synchronize_session=False)
                    if x[1]:
                        dcts = list([x0._asdict() for x0 in items.values()])
                        jns = [JOElement(x0.jono) for x0 in items.values()]
                        jns = self._hksvc.getjos(jns)[0]
                        jns = dict([(x0.name,x0) for x0 in jns])
                        for dct in dcts:
                            # todo::make the china value for the user
                            if not dct["stone"]:
                                dct["stone"] = NA
                            else:
                                dct["stone"] = _removenonascii(dct["stone"])
                            dct["china"] = 0
                            dc1 = dict(dct)
                            dc1["joid"] = jns[JOElement(dc1["jono"])].id
                            dc1["stspec"] = _removenonascii(dc1["stone"])
                            dc1["uprice"] = dc1["price"]
                            for y in "lastmodified".split(","):
                                dct[y] = _accdstr(dct[y])
                            iv = PajInv()
                            for it in dc1.items():
                                k = it[0]
                                lk = it[0].lower()
                                if hasattr(iv, lk):
                                    iv.__setattr__(lk, dc1[k])
                            cur.execute(("insert into PajInv(Invno,Jono,StSpec,Qty,UPrice,Mps,LastModified) values"
                                        "('%(invno)s','%(jono)s','%(stone)s',%(qty)f,%(price)f,'%(mps)s',#%(lastmodified)s#)") % dct)
                            iv = sess.add(iv)
                except Exception as e:  # there might be pyodbc.IntegrityError if dup found
                    logger.debug("Error occur in InvRdr:%s" % e)
                finally:
                    if not e:
                        sess.commit()
                        self._accessdb.commit()
                    else:
                        sess.rollback()
                        self._accessdb.rollback()
                    cur.close()
        return (0 if sum(x) == 0 else -1 if e else 1), e

    def read(self, invfldr, writeJOBack=True):
        """
        read files back, instead of using os.walk(root), use os.listdir()
        @param invfldr: the folder contains the invoices
        @param writeJOBack: write the JO# back to the source sheet 
        """

        if not os.path.exists(invfldr):
            return
        PajInvItem = namedtuple(
            "PajInvItem", "invno,p17,jono,qty,price,mps,stone,lastmodified")
        fns = getfiles(invfldr,"xls",True)
        if fns:
            p = appathsep(invfldr)
            fns = [p + x for x in fns if x[2:4].lower() == "pm"]
        if not fns:
            return
        killexcel, app = xwu.app(False)

        items, errs, invs = {}, [], {}
        for fn in fns:
            invno = self._getinv(fn).upper()
            if invno in invs:
                continue
            dups = []
            idx = self._hasread(fn)
            if idx == 1:
                # logger.debug("%s has been read" % fn)
                continue
            elif idx == 2:
                logger.debug("%s is expired" % fn)
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
                    tm = xwu.list2dict(tr, {"gold,": "gold", "silver,": "silver", u"job#,工单": "jono",
                                            "price,": "price", "unit,": "qty", "stone,": "stone"})
                    tm["p17"] = 0
                    x = [x for x in "price,qty,stone".split(",") if not x in tm]
                    if x:
                        logger.info(
                            "key columns(%s) missing in sheet('%s') of file (%s)" % (x, sh.name, fn))
                        continue
                    for jj in range(1, len(vals)):
                        tr = vals[jj]
                        if not tr[tm["price"]]:
                            continue
                        p17 = tr[tm["p17"]]
                        if not (p17u.isvalidp17(p17) and 
                                len([1 for y in [x for x in "qty,price,silver,gold".split(",")]\
                                if not isnumeric(tr[tm[y]])]) == 0):
                            logger.debug(
                                "invalid p17 code(%s) or wgt/qty/price data in %s" % (p17, fn))
                            continue
                        jn = JOElement(tr[tm["jono"]]).value if "jono" in tm else None
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
                            logger.debug(
                                "No JO# found for p17(%s) in file %s" % (tr[tm["p17"]], fn))
                            continue
                        key = invno + "," + jn
                        if key in items.keys():
                            it = items[key]
                            items[key] = it._replace(
                                qty=it.qty + tr[tm["qty"]])
                        else:
                            mps = "S=%3.2f;G=%3.2f" % (tr[tm["silver"]], tr[tm["gold"]]) \
                                if "gold" in tm and "silver" in tm else "S=0;G=0"
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
                logger.debug("invoice (%s) processed" %
                              fn + ("" if x[0] else " but all are repairing"))
        if killexcel:
            app.quit()
        #return x[0], items if x[0] == 1 else () if x[0] == 0 else errs
        return -1, errs if errs else 1, items if items else 0, None


class PAJCReader(object):
    """class to create the PAJ JOCost file for HK accountant"""

    def __init__(self, hksvc=None, cnsvc=None, bcdb=None):
        self._hksvc = hksvc
        self._cnsvc = cnsvc
        self._bcdb = bcdb

    def run(self, year, month, day = 1, tplfn=None, tarfldr=None):
        """ create report file of given year/month"""

        def _makemap(sht=None):
            coldefs = (u"invoice date=invdate;invoice no.=invno;order no.=orderno;customer=cstname;"
                       u"job no.=jono;style no.=styno;running no.=running;paj item no.=pcode;karat=karat;"
                       u"描述=cdesc;in english=edesc;job quantity=joqty;quantity received=shpqty;"
                       u"total cost=ttlcost;cost=uprice;平均单件金银重g=umtlwgt;平均单件石头重g=ustwgt;石头=stspec;"
                       u"mm program in#=iono;jmp#=jmpno;date=shpdate;remark=rmk;has dia=hasdia")
            vvs = sht.range("A1").expand("right").value
            vvs = [x.lower() if isinstance(x, str) else x for x in vvs]
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
        df, dt = daterange(year,month,day)

        runns, jes = set(), set()
        bcsvc = BCSvc(self._bcdb)

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
        bcs = dict([(x.runn, x.desc) for x in bcsvc.getbcsforjc(runns)])
        lst = self._hksvc.getpajinvbyjes(jes)
        pajs, pajsjn = {}, {}
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
        lst, fn = [], None
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
                for x in dtmap0.items():
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
                    for y in dtmap1.items():
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
                                for y in dtmap1.items():
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
                lst.append(["" if not y else y.strip() if isinstance(y, str) else
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
        fns = getfiles(fldr,"dat")
        if not fns:
            return
        stynos = set()
        for x in fns:
            with open(fldr + x, "wb") as fh:
                for ln in fh:
                    je = JOElement(ln)
                    if je.isvalid and not je in stynos:
                        stynos.add(je)
        dao = BCSvc(self._hkdb)
        #lst = dao.getpajinvbyse(stynos)
