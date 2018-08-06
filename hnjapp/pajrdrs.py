# coding=utf-8
'''
Created on Apr 17, 2018

the replacement of the Paj Shipment Invoice Reader, which was implmented
in PAJQuickCost.xls#InvMatcher

@author: zmFeng
'''

import datetime as dtm
from collections import OrderedDict
import numbers
import os
import random
from os import path
import re
import sys
from collections import namedtuple
from datetime import date, datetime
from decimal import Decimal
from tkinter import filedialog
import tkinter as tk

import xlwings.constants as const
from xlwings.utils import col_name
from sqlalchemy import and_, func
from sqlalchemy.engine import create_engine
from sqlalchemy.orm import Query
from xlwings.constants import LookAt

from hnjcore import JOElement, appathsep, deepget, karatsvc, p17u, xwu
from hnjcore.models.hk import JO as JOhk
from hnjcore.models.hk import Orderma, PajInv, PajShp, PajCnRev
from hnjcore.models.hk import Style as Styhk
from hnjcore.utils import daterange, getfiles, isnumeric, p17u
from hnjcore.utils.consts import NA
from utilz import NamedList, NamedLists, ResourceCtx, SessionMgr, list2dict, splitarray, triml, trimu

from .common import _logger as logger
from .dbsvcs import BCSvc, CNSvc, HKSvc
from .localstore import PajInv as PajInvSt
from .localstore import PajItem, PajCnRev as PajCnRevSt
from .localstore import PajWgt as PrdWgtSt
from .pajcc import PAJCHINAMPS, P17Decoder, PajCalc, PrdWgt, WgtInfo, MPS

_accdfmt = "%Y-%m-%d %H:%M:%S"
_dfkt = {"4":925,"5":925,"M":8,"B":9,"G":10,"Y":14,"P":18}

def _accdstr(dt):
    """ make a date into an access date """
    return dt.strftime(_accdfmt) if dt and isinstance(dt, date) else dt


def _removenonascii(s0):
    """remove thos non ascii characters from given string"""
    if isinstance(s0, str):
        return "".join([x for x in s0 if ord(x) > 31 and ord(x) < 127 and x != "?"])
    return s0

def _getdefkarat(jn):
    return _dfkt.get(jn[0])

class PajBomHhdlr(object):
    """ methods to read BOMs from PAJ """

    @classmethod
    def readbom(self,fldr):
        """ read BOM from given folder
        @param fldr: the folder contains the BOM file(s)
        return a dict with "pcode" as key and dict as items
            the item dict has keys("pcode","wgts")
        """
        _ptnoz = re.compile(r"\(\$\d*/OZ\)")
        _ptnsil = re.compile(r"(925)")
        _ptngol = re.compile(r"^(\d*)K")    
        _ptdst = re.compile(r"[\(（](\d*)[\)）]")
        _ptfrcchain = re.compile(r"(弹簧扣)|(龙虾扣)|(狗仔头扣)")
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
                            karat = karatsvc.getbyfineness(kt)
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
            return _pcdec.decode(pcode,"PRODTYPE").find("吊") >= 0
        
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

        if path.isdir(fldr):
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
                        wcn = xwu.list2dict(vvs[0],{u"pcode":"十七位,","mat":u"材质,","mtlwgt":u"抛光,"})
                        for ii in range(1,len(vvs)):
                            pcode = vvs[ii][wcn["pcode"]]
                            if not p17u.isvalidp17(pcode): break
                            kt = _parsekarat(vvs[ii][wcn["mat"]])
                            if not kt: continue
                            it = pmap.setdefault(pcode,{"pcode":pcode})
                            it.setdefault("wgts",[]).append((kt,vvs[ii][wcn["mtlwgt"]]))
                    elif jj == 1:
                        nmap = {"pcode":u"十七位,","name":u"物料名称", \
                            "spec":u"物料特征","qty":u"数量","wgt":u"重量","unit":u"单位","length":u"长度"}
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
                    ispendant,haschain,haskou,chlenerr = _ispendant(x[0]), False,False, False
                    if ispendant:
                        for y in x[1]["parts"]:
                            nm = y["name"]
                            if triml(nm).find("chain") >= 0:
                                haschain = True
                            if _ptfrcchain.search(nm):
                                haskou = True
                    for y in x[1]["parts"]:
                        nm = y["name"]
                        kt = _parsekarat(nm,wis,False)
                        if not kt: continue                
                        y["karat"] = kt
                        done = False
                        if ispendant:
                            if haschain:
                                isch = triml(nm).find("chain") >= 0
                                done = isch or (haskou and (_ptfrcchain.search(nm) or nm.find("圈") >= 0))
                                if isch:
                                    lc = y["length"]
                                    if not lc is None:
                                        try:
                                            lc = float(lc)
                                        except:
                                            lc = 0
                                        if lc > 0: chlenerr = True
                                if done:
                                    wgt0 = wis[2]
                                    if not wgt0 or y["karat"] == wgt0.karat:
                                        wgt = y["wgt"] + (wgt0.wgt if wgt0 else 0)
                                        wis[2] = WgtInfo(y["karat"],wgt)
                                        if wgt0:
                                            logger.debug("Multi chain found for pcode(%s)" % x[0])
                                    else:
                                        done = False
                                else:
                                    logger.debug("No wgt pos for chain(%s) in pcode(%s),merged to main" % (y["name"],x[0]))
                        if not done:
                            _mergewgt(wis,WgtInfo(y["karat"],y["wgt"]),1)
                if chlenerr and wis[2]:
                    wis[2] = WgtInfo(wis[2].karat,-wis[2].wgt)
                x[1]["wgts"] = PrdWgt(wis[0],wis[1],wis[2])        
        finally:
            if kxl and app: app.quit()
        return pmap

    @classmethod
    def readbom2jos(self,fldr,hksvc,fn = None,mindt = None):
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

        pmap = self.readbom(fldr)
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

class PajShpHdlr(object):
    """
    Tasks:
        .integrating PajShpRdr/PajInvRdr into one to generate data from HKdb
        .genereate shipment data for py.mm/py.bc
    """
    def __init__(self,hksvc):
        self._hksvc = hksvc

    @classmethod
    def _getshpdate(self, fn, isfile=True):
        """extract the shipdate from file name"""
        import datetime as dt
        ptnfd = re.compile(r"\d+")
        parts = ptnfd.findall(path.basename(fn))
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
                d1 = dt.date.fromtimestamp(path.getmtime(fn))
                d0 = dt.date(d1.year, parts[0], parts[1])
                df = d1 - d0
                if df.days < -3:
                    d0 = dt.date(d0.year - 1, d0.month, d0.day)
        except:
            d0 = None
        return d0
    
    @classmethod
    def _getfmd(self,fn):
        return datetime.fromtimestamp(path.getmtime(fn)).replace(microsecond=0).replace(second=0)
    
    @classmethod
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
        _ptngwt = re.compile("[\d.]+")
        vvs = sht.range(rng, rng.current_region.last_cell).value
        nls = NamedLists(vvs,{"pcode":"Item,","stone":"stone,","metal":"metal ,"},False)
        for tr in nls:
            p17 = tr.pcode
            if not p17:
                continue
            if p17u.isvalidp17(p17) and not p17 in qmap:
                sw = 0 if not tr.stone else \
                    sum([float(x)
                         for x in _ptngwt.findall(tr.stone)])
                mtls = tr.metal
                if isinstance(mtls, numbers.Number):
                    mw = (WgtInfo(0,mtls),)
                else:
                    s0, mw = tr.metal.replace("：",":"), []
                    if s0.find(":") > 0:
                        for x in s0.split("\n"):
                            ss= x.split(":")
                            mt = _ptngwt.search(ss[0])
                            karat = 925 if not mt else int(mt.group())
                            mt = _ptngwt.search(ss[1])
                            wgt = float(mt.group()) if mt else 0
                            mw.append(WgtInfo(karat,wgt))
                    else:
                        mt = _ptngwt.search(s0)
                        mw.append(WgtInfo(0,float(mt.group()) if mt else 0))
                qmap[p17] = (mw, sw)

    def _hasread(self,fmd, fn, invno = None):
        """
            check if given file(in shpment case) or inv#(in invoice case) has been read
            @param fn: the full-path filename 
            return:
                1 if hasread and up to date
                2 if expired
                0 if not read
        """
        rc= 0
        if not invno:
            fn = _removenonascii(path.basename(fn))
            with self._hksvc.sessionctx() as cur:
                shp = Query([PajShp.fn,func.min(PajShp.lastmodified)]).group_by(PajShp.fn).filter(PajShp.fn == fn).with_session(cur).first()
                if shp:
                    rc = 2 if shp[1] < fmd else 1
        else:
            with self._hksvc.sessionctx() as cur:
                inv = Query([PajInv.invno,func.min(PajInv.lastmodified)]).group_by(PajInv.invno).filter(PajInv.invno == invno).with_session(cur).first()
                if inv:
                    rc = 2 if inv[1] < fmd else 1
        return rc

    @classmethod
    def _rawreadinv(self, sht, invno = None, fmd = None):
        PajInvItem = namedtuple(
            "PajInvItem", "invno,pcode,jono,qty,uprice,mps,stspec,lastmodified")
        items = {}
        rng = xwu.find(sht, "Item*No", lookat=const.LookAt.xlWhole)
        if not rng:
            return
        if not invno: invno = self._readinvno(sht)
        rng = rng.expand("table")
        nls = tuple(NamedLists(rng.value,{"pcode":"item,","gold":"gold,", "silver":"silver,", "jono": u"job#,工单", "uprice": "price,", "qty": "unit,", "stspec": "stone,"}))
        if not nls: return
        th = nls[0]
        x = [x for x in "uprice,qty,stspec".split(",") if not th.getcol(x)]
        if x:
            logger.debug(
                "key columns(%s) missing in invoice sheet('%s')" % (x, sht.name))
            return
        for tr in nls: 
            if not tr.uprice:
                continue
            p17 = tr.pcode
            if not (p17u.isvalidp17(p17) and 
                    len([1 for y in [x for x in "qty,uprice,silver,gold".split(",")]\
                    if not isnumeric(tr[y])]) == 0):
                logger.debug(
                    "invalid p17 code(%s) or wgt/qty/uprice data in invoice sheet(%s)" % (p17, invno))
                continue
            jn = JOElement(tr.jono).value if th.getcol("jono") else None
            if jn:
                jn = "%d" % int(jn) if isinstance(jn, numbers.Number) else jn.strip()
            else:
                #todo:: get from local map, because the read process can have jo -> inv map
                #jns = self._getjonos(p17, invno)
                pass
            if not jn:
                logger.debug("No JO# found for p17(%s)" % p17)
                continue
            key = invno + "," + jn
            if key in items.keys():
                it = items[key]
                items[key] = it._replace(qty=it.qty + tr.qty)
            else:
                mps = MPS("S=%3.2f;G=%3.2f" % (tr.silver, tr.gold)).value \
                    if th.getcol("gold") and th.getcol("silver") else "S=0;G=0"
                it = PajInvItem(invno, p17, jn, tr.qty, tr.uprice, mps, tr.stspec, fmd)
                items[key] = it
        return items

    @classmethod
    def _readinvno(self, sht):
        rng = xwu.find(sht, "Inv*No:")
        if not rng: return
        return rng.offset(0, 1).value

    def _readinv(self, fn, sht, fmd):
        """
        read files back, instead of using os.walk(root), use os.listdir()
        @param invfldr: the folder contains the invoices
        """

        invno, dups = self._readinvno(sht) , []
        idx = self._hasread(fmd, fn, invno)
        if idx == 1:
            return 
        elif idx == 2:
            dups.append(invno)            
        items = self._rawreadinv(sht, invno, fmd)
        return items,dups

    @classmethod
    def _readshp(self,fn,fshd,fmd,sht):
        """ 
        @param fshd: the shipdate extracted by the file name
        @param fmd: the last-modified date
        @param fn: the full-path filename
        """

        vvs = xwu.usedrange(sht).value
        if not vvs: return
        PajShpItem = namedtuple("PajShpItem", "fn,orderno,jono,qty,pcode,invno,invdate" +
                                ",mtlwgt,stwgt,shpdate,lastmodified,filldate")
        items, td0, qmap = {}, datetime.today(), None
        nls = tuple(NamedLists(vvs,{"odx": u"订单号", "invdate": u"发票日期", "odseq": u"订单序号","stwgt": u"平均单件石头,XXX", "invno": u"发票号", "orderno": u"订单号序号", "pcode": u"十七位,十七,物料","mtlwgt": u"平均单件金,XX", "jono": u"工单,job", "qty": u"数量", "cost": u"cost"}))
        th = nls[0]
        x = [x for x in "invno,pcode,jono,qty,invdate".split(
            ",") if th.getcol(x) is None]
        if x:
            return
        bfn = path.basename(fn).replace("_", "")
        shd = PajShpHdlr._getshpdate(sht.name, False)
        if shd:
            df = shd - fshd
            shd = shd if abs(df.days) <= 7 else fshd 
        else:
            shd = fshd 
        # finally I give up, don't use the shipdate, use invdate as shipdate
        if th.getcol("mtlwgt"):
            for tr in nls:
                if not tr.pcode:
                    break
                mwgt = tr.mtlwgt
                if not (isinstance(mwgt, numbers.Number) and mwgt > 0):
                    continue
                invno = tr.invno
                if not invno: invno = "N/A"
                if th.getcol('orderno'):
                    odno = tr.orderno
                elif len([1 for x in "odx,odseq".split(",") if th.getcol(x)]) == 2:
                    odno = tr.odx + "-" + tr.odseq
                else:
                    odno = "N/A"
                if not tr.stwgt:
                    tr.stwgt = 0
                jono = JOElement(tr.jono).value
                thekey = "%s,%s,%s" % (jono,tr.pcode,invno)
                if thekey in items:
                    si = items[thekey]
                    items[thekey] = si._replace(qty = si.qty + tr.qty, mtlwgt = si.mtlwgt + mwgt)
                else:
                    ivd = tr.invdate
                    si = PajShpItem(bfn, odno, jono, tr.qty, tr.pcode, invno, ivd, (WgtInfo(_getdefkarat(jono),mwgt),), tr.stwgt, ivd, fmd, td0)
                    items[thekey] = si
        else:
            # new sample case, extract weight data from the quo sheet
            if not qmap:
                wb, qmap = sht.book, {}
                for x in [xx for xx in wb.sheets if xx.api.Visible == -1 and xx.name.lower().find('dl_quotation') >= 0]:
                    PajShpHdlr._readquodata(x, qmap)
            if qmap:
                for tr in nls:
                    # no cost item means repairing
                    if not tr.get("cost"): continue
                    p17 = tr.pcode
                    if not p17:
                        break
                    ivd, odno = tr.invdate, tr.get("orderno",NA)
                    if p17 in qmap:
                        snm = qmap[p17]
                    else:
                        logger.info("failed to get quo info for pcode(%s)" % p17)
                        snm = (0,0)
                    si = PajShpItem(bfn, odno, JOElement(tr.jono).value, tr.qty, p17,
                            tr.invno, ivd, snm[0], snm[1], ivd, fmd, td0)
                    # new sample won't have duplicated items
                    items[random.random()] = si
            else:
                qmap["_SIGN_"] = 0
        return items

    def _persist(self, shps, invs):
        """save the data to db
        @param dups: a list contains file names that need to be removed
        @param items: all the ShipItems that need to be persisted
        """

        err = True
        with self._hksvc.sessionctx() as sess:            
            if shps[0]:
                sess.query(PajShp).filter(PajShp.fn.in_([_removenonascii(path.basename(x)) for x in shps[0]])).delete(synchronize_session=False)
            if invs[0]:
                sess.query(PajInv).filter(PajInv.invno.in_(invs[0])).delete(synchronize_session=False)
            jns = set()
            if shps[1]:
                jns.update([JOElement(x.jono) for x in shps[1].values()])
            if invs[1]:
                jns.update([JOElement(x.jono) for x in invs[1].values()])
            if jns:
                jns = self._hksvc.getjos(jns)[0]
                jns = dict([(x.name,x) for x in jns])
                if shps[1]:
                    for dct in [x._asdict() for x in shps[1].values()]:
                        je = JOElement(dct["jono"])
                        if je not in jns or not p17u.isvalidp17(dct["pcode"]):
                            logger.info("Item(%s) does not contains valid JO# or pcode" % dct)
                            continue
                        dct["fn"] = _removenonascii(dct["fn"])
                        dct["joid"] = jns[je].id
                        #"mtlwgt" is a list of WgtInfo Object
                        dct["mtlwgt"] = sum([x.wgt for x in dct["mtlwgt"]])
                        shp = PajShp()
                        for x in dct.items():
                            k = x[0]
                            lk = k.lower()
                            if hasattr(shp, lk):
                                setattr(shp,lk,dct[k])
                        sess.add(shp)
                if invs[1]:
                    for dct in [x._asdict() for x in invs[1].values()]:
                        if not dct["stspec"]:
                            dct["stspec"] = NA
                        else:
                            dct["stspec"] = _removenonascii(dct["stspec"])
                        dct["china"] = 0
                        dct["joid"] = jns[JOElement(dct["jono"])].id
                        iv = PajInv()
                        for it in dct.items():
                            k = it[0]
                            lk = it[0].lower()
                            if hasattr(iv, lk):
                                iv.__setattr__(lk, dct[k])
                        iv = sess.add(iv)
            sess.commit()
            err = False
        return -1 if err else 1, err


    def process(self,fldr):
        """
        read the shipment file and send shipment/invoice to hkdb
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
        try:
            #when excel open a file, the file's modified date will be changed, so, in
            #order to get the actual modified date, get it first
            fmds = dict([(x,self._getfmd(x)) for x in fns])
            fns = sorted([(x,self._getshpdate(x)) for x in fns], key = lambda x: x[1])
            fns = [x[0] for x in fns]
            for fn in fns:
                rflag = self._hasread(fmds[fn],fn)
                shptorv, invtorv = [], []
                shps, invs = {}, {}
                shtshps, shtinvs = [], []
                if rflag == 2:
                    shptorv.append(fn)
                shd0, fmd, wb = self._getshpdate(fn), fmds[fn], app.books.open(fn)
                try:
                    for sht in wb.sheets:
                        if sht.name.find(u"返修") >= 0:
                            continue
                        rng = xwu.find(sht, u"十七*", lookat=LookAt.xlPart)
                        if not rng:
                            rng = xwu.find(sht, u"物料*", lookat=LookAt.xlPart)
                        if not rng:
                            if xwu.find(sht,"PAJ"):
                                shtinvs.append(sht)
                        else:
                            shtshps.append(sht)
                    if shtshps and shtinvs:
                        if rflag != 1:
                            for sht in shtshps:
                                its = PajShpHdlr._readshp(fn, shd0, fmd, sht)
                                if its: shps.update(its)
                        for sht in shtinvs:
                            its = self._readinv(fn, sht, fmd)
                            if its:
                                if its[0]: invs.update(its[0])
                                if its[1]: invtorv.extend(its[1])
                    elif bool(shtshps) ^ bool(shtinvs):
                        logger.info("Error::Not both shipment and invoice in file(%s), No data updated" % path.basename(fn))
                finally:
                    if wb:
                        wb.close()
                if sum((len(x) for x in (shptorv,shps,invtorv,invs))) == 0:
                    if rflag != 1:
                        logger.debug("no valid data returned from file(%s)" % path.basename(fn))
                    else:
                        logger.debug("data in file(%s) is up-to-date" % path.basename(fn))
                    continue
                logger.debug("counts of file(%s) are: Shp2Rv=%d, Shps=%d, Inv2Rv=%d, Invs=%d" % (path.basename(fn), len(shptorv),len(shps),len(invtorv),len(invs)))
                if True:
                    x = self._persist((shptorv, shps),(invtorv,invs))
                    if x[0] != 1:
                        errors.append(x[1])
                        logger.info("file(%s) contains errors" % path.basename(fn))
                        logger.info(x[1])
                    else:
                        logger.debug("data in file(%s) were updated" % (path.basename(fn)))
        finally:
            if killxls:
                app.quit()
        return -1 if len(errors) > 0 else 1, errors

class PajJCMkr(object):
    """
    the JOCost maker of Paj for HK accountant, the twin brother of C1JCMkr
    """

    def __init__(self, hksvc=None, cnsvc=None, bcsvc=None):
        self._hksvc = hksvc
        self._cnsvc = cnsvc
        self._bcsvc = bcsvc

    def run(self, year, month, day = 1, tplfn=None, tarfn=None):
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
        bcsvc = self._bcsvc

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
        lst = []
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
            if tarfn:
                wb.save(tarfn)
        finally:
            if killxls and not wb:
                app.quit()
            else:
                app.visible = True
        return lst, tarfn

class PajUPTcr(object):
    """ 
    Paj unit-price tracer
    to use this method, put a dat file inside a folder which should contains sty#
    then I will read and show the price trends
    
    to speed up the process of fetching data from hk, the key data(wgt/poprices) were localized by a sqlitedb.

    the original purpose is to track the stamping products, but in fact, can be use for any Paj items.

    """

    def __init__(self, hksvc,localeng, srcfn):
        self._hksvc = hksvc
        self._srcfn = srcfn
        self._localsm = SessionMgr(localeng)
        if False:
            sess = ResourceCtx(self._localsm)
            with sess as cur:
                try:
                    ext = cur.query(PajInvSt).limit(1).first()
                except:
                    ext = None
                if not ext:
                    PajInvSt.metadata.create_all(localeng)
        else:
            PajInvSt.metadata.create_all(localeng)

    def readpcodes(self):
        with open(self._srcfn,"r+t") as fh:
            return set([x[:-1] for x in fh.readlines() if x[0] != "#"])

    def createcache(self,pcodes):
        cc = PajCalc()
        td = datetime.today()
        for pcode in pcodes:
            wgts = None
            with ResourceCtx((self._localsm,self._hksvc.sessmgr())) as curs:
                try:
                    ivca = Query(PajItem).filter(PajItem.pcode == pcode).with_session(curs[0]).first()
                    if ivca:
                        logger.debug("pcode(%s) already localized" % pcode)
                        continue
                    q0 =Query([PajShp.pcode,JOhk.name.label("jono"),Styhk.name.label("styno"),JOhk.createdate,PajShp.invdate,PajInv.uprice,PajInv.mps]).join(JOhk).join(Orderma).join(Styhk).join(PajInv,and_(PajShp.joid == PajInv.joid, PajShp.invno == PajInv.invno)).filter(PajShp.pcode == pcode)
                    lst = q0.with_session(curs[1]).all()
                    if not lst: continue

                    pi = PajItem()
                    pi.pcode, pi.createdate, pi.tag = pcode, td, 0
                    curs[0].add(pi)
                    curs[0].flush()
                    jeset = set()
                    for jnv in lst:
                        je = jnv.jono
                        if je in jeset: continue
                        jeset.add(je)
                        if not wgts:
                            wgts = self._hksvc.getjowgts(je)
                            if not wgts: continue
                            wgtarr = wgts.wgts
                            for idx in range(len(wgtarr)):                        
                                if not wgtarr[idx]: continue
                                pw = PrdWgtSt()
                                pw.pid, pw.karat, pw.wgt, pw.remark,pw.tag = pi.id, wgtarr[idx].karat, wgtarr[idx].wgt, je.value, 0
                                pw.createdate, pw.lastmodified = td, td
                                pw.wtype = 0 if idx == 0 else 100 if idx == 2 else 10
                                curs[0].add(pw)
                        up,mps = jnv.uprice, jnv.mps
                        cn = cc.calchina(wgts,up,mps,PAJCHINAMPS)
                        if cn:
                            ic = PajInvSt()
                            ic.pid, ic.uprice,ic.mps = pi.id, up, mps
                            ic.cn = cn.china
                            ic.jono, ic.styno = je.value,jnv.styno.value
                            ic.jodate, ic.createdate, ic.invdate, ic.lastmodified = jnv.createdate, td, jnv.invdate, td
                            ic.mtlcost, ic.otcost = cn.metalcost, cn.china - cn.metalcost
                            curs[0].add(ic)                  
                    curs[0].commit()
                    logger.debug("pcode(%s) localized" % pcode)
                except Exception as e:
                    logger.debug("Error occur while persisting localize result %s" % e)
                    curs[0].rollback()

    def localize(self):
        self.localizerev()
        pcodes = list(self.readpcodes())
        if not pcodes: return
        logger.debug("totally %d pcodes send for localize" % len(pcodes))
        cnt = 0
        for arr in splitarray(pcodes,50):
            self.createcache(arr)
            cnt += 1
    
    def localizerev(self):
        """ localize the rev history
        """
        affdate = datetime(2018,4,4)
        q0 = Query((func.max(PajCnRevSt.createdate),))
        with ResourceCtx(self._localsm) as cur:            
            lastcrdate = q0.with_session(cur).first()[0]
        if not lastcrdate: lastcrdate = affdate
        q0 = Query(PajCnRev).filter(and_(PajCnRev.filldate > lastcrdate,and_(PajCnRev.tag == 0, PajCnRev.revdate >= affdate)))
        with ResourceCtx((self._localsm,self._hksvc.sessmgr())) as curs:
            srcs = q0.with_session(curs[1]).all()
            pcodes = set([x.pcode for x in srcs])
            pcs = {}
            for arr in splitarray(list(pcodes)):
                q0 = Query(PajItem).filter(PajItem.pcode.in_(arr))
                try:
                    pcs.update(dict([(x.pcode,x) for x in q0.with_session(curs[0]).all()]))
                except:
                    pass
            if pcs:
                prs = []
                for arr in splitarray([x.id for x in pcs.values()]):
                    try:
                        prs.extend(curs[0].query(PajCnRevSt).filter(PajCnRevSt.id.in_(arr)).all())
                    except:
                        pass
            if prs:
                tag = curs[0].query((func.max(PajCnRevSt.tag),)).first()[0]
                if not tag: tag = 0
                tag = int(tag) + 1
                for x in prs:
                    x.tag = tag
                    curs[0].add(x)
            td = datetime.today()
            npis = []
            for x in [y for y in srcs if y.pcode not in pcs]:                
                pi = PajItem()
                pi.pcode, pi.createdate, pi.tag = x.pcode, td, 0
                curs[0].add(pi)
                npis.append(pi)
            if npis:
                curs[0].flush()
                pcs.update(dict([(x.pcode,x) for x in npis]))
            for x in srcs: 
                pi = pcs[x.pcode]
                rev = PajCnRevSt()
                rev.pid, rev.uprice = pi.id, x.uprice
                rev.revdate, rev.createdate, rev.tag = x.revdate, td, 0
                curs[0].add(rev)
            curs[0].commit()
    
    def analyse(self,cutdate = None):
        if not cutdate: cutdate = datetime(2018,5,1)
        
        mixcols = "oc,cn,invdate,jono".split(",")        
        gelmix = NamedList(mixcols)
        def _minmax(arr):
            """ return a 3 element tuple, each element contains mixcols data
            first   -> min
            second  -> max
            third   -> last
            """
            fill =lambda ar: [float(ar.otcost),float(ar.cn),ar.invdate,ar.jono]
            li, lx = 9999, -9999
            mi, mx = None, None
            for ar in arr:
                lb = float(ar.otcost)
                if lb > lx:
                    mx = fill(ar)
                    lx = lb
                if lb < li:
                    mi = fill(ar)
                    li = lb
            df = lx - li
            if df < 0.1 or df / li < 0.01:
                df = (lx + li) / 2.0
                mi[0],mx[0] = df, df
            return mi,mx,fill(arr[-1])

        def getonly(cns,arr):
            if isinstance(cns,str):
                cns = cns.split(",")
            lst = []
            for ar in arr:
                gelmix.setdata(ar)
                lst.extend([gelmix[cn] for cn in cns])
            return lst
        almosteq = lambda x,y: abs(x - y) <= 0.1 or abs(x - y) / x < 0.01
        gelq = NamedList("pcode,jono,styno,invdate,cn,otcost")
        with ResourceCtx(self._localsm) as cur:
            q0 = Query([PajItem.pcode,PajInvSt.jono,PajInvSt.styno,PajInvSt.invdate,PajInvSt.cn,PajInvSt.otcost]).join(PajInvSt).order_by(PajItem.pcode).order_by(PajInvSt.invdate)
            #q0 = q0.limit(50)
            lst = q0.with_session(cur).all()            
            mp = {}
            [mp.setdefault(it.pcode,[]).append(it) for it in lst]
            q0 = Query([PajItem.pcode,PajCnRevSt.revdate,PajCnRevSt.uprice]).join(PajCnRevSt)
            revdates = {}
            for arr in splitarray(list(mp.keys())):
                try:
                    revdates.update(dict([(y.pcode,(y.revdate,float(y.uprice))) for y in \
                    [x for x in q0.filter(PajItem.pcode.in_(arr)).with_session(cur).all
                ()]]))
                except Exception as e:
                    print(e)
                    pass

            noaff, mixture, noeng, drp, pum, nochg = [], [], [], [],[], []
            for it in mp.items():
                lst = it[1]
                gelq.setdata(lst[0])
                flag = len(lst) > 1
                acutdate = revdates.get(it[0],cutdate)
                revcn = 0
                if isinstance(acutdate,tuple):
                    revcn = acutdate[1]
                    acutdate = acutdate[0]
                if flag:
                    for idx in range(len(lst)):
                        flag = lst[idx].invdate >= acutdate
                        if flag: break
                if not flag:
                    noaff.append((gelq.pcode,gelq.styno,acutdate,revcn,_minmax(lst)))
                else:
                    mix0,mix1 = _minmax(lst[:idx]), _minmax(lst[idx:])
                    iot = gelmix.getcol("oc")
                    val = (gelq.pcode,gelq.styno,acutdate,revcn,mix0,mix1)                    
                    if mix0[0][iot] * 2.0 / 3.0 + 0.05 >= mix1[1][iot]:
                        drp.append(val)
                    elif almosteq(mix0[0][iot],mix1[1][iot]):
                        nochg.append(val)
                    elif mix0[0][iot] > mix1[1][iot]:
                        noeng.append(val)
                    elif mix0[0][iot] < mix1[1][iot]:
                        #old's max under new's min
                        pum.append(val)
                    else:
                        mixture.append(val)
            mp = {"NotAffected":noaff,"NoChanges":nochg,"Mixture":mixture,"NoEnough":noeng, "PriceDrop1of3":drp,"PriceUp":pum}
            app = xwu.app(True)[1]            
            grp0 = (("",),"Before,After".split(","))
            grp1 = "Min.,Max.,Last".split(",")
            grp2 = "pcode,styno,revdate,cn,karat".split(",")
            ctss = ("cn,invdate".split(","),"oc,cn,jono,invdate".split(","))
            shts,pd = [], P17Decoder()
            wb = app.books.add()
            for x in mp.items():
                if not x[1]: continue
                shts.append(wb.sheets.add())
                sht = shts[-1]
                sht.name, vvs = x[0], []
                gidx = 0 if x[0] == "NotAffected" else 1
                ttl0,ttl1 = ["","","","",""],["","","","",""]
                for z in grp0[gidx]:
                    ttl0.append(z)
                    for ii in range(len(ctss[gidx]) * len(grp1) - 1):
                        ttl0.append(" ")
                    for xx in grp1:
                        ttl1.append(xx)
                        for ii in range(len(ctss[gidx])-1):
                            ttl1.extend(" ")
                if len(grp0[gidx]) > 1:vvs.append(ttl0)
                vvs.append(ttl1)

                ttl = grp2.copy()
                ttlen = len(ttl) - 1
                cnt = 0
                while cnt < len(grp1)*len(grp0[gidx]):
                    ttl.extend(ctss[gidx])
                    cnt += 1
                vvs.append(ttl)                
                for it in x[1]:
                    ttl = list(it[:ttlen])
                    ttl.append(pd.decode(ttl[0],"karat"))
                    [ttl.extend(getonly(ctss[gidx],kk)) for kk in it[ttlen:]]
                    vvs.append(ttl)
                sht.range(1,1).value = vvs
                sht.autofit("c")
                #let the karat column smaller
                sht[1,grp2.index("karat")].column_width = 10
                xwu.freeze(sht.range(3 + (1 if len(grp0[gidx]) > 1 else 0),ttlen+2))

            for sht in wb.sheets:
                if sht not in shts:
                    sht.delete()

def easydlg(dlg):
    """ open a tk dialog and return sth. easily """
    rt = tk.Tk()
    rt.withdraw()
    dlg.master = rt
    rc = dlg.show()
    rt.quit()
    return rc

class PajNSOFRdr(object):
    """
    class to read a NewSampleOrderForm's data out
    """
    _tplfn = r"\\172.16.8.46\pb\dptfile\pajForms\PAJSKUSpecTemplate.xlt"

    def readsettings(self, fn = None):
        usetpl, mp = False, None
        if not fn:
            fn, usetpl = self._tplfn, True
        kxl, app = xwu.app()
        try:
            wb = app.books.open(fn) if not usetpl else xwu.fromtemplate(fn,app)
            shts = [x for x in wb.sheets if triml(x.name).find("setting") >= 0]
            if shts:
                rng = xwu.find(shts[0],"name")
                nls = NamedLists(rng.expand("table").value)
                mp = dict([(triml(nl.name),nl) for nl in nls])
        except:
            pass
        finally:
            if wb: wb.close()
            if kxl: app.quit()
        return mp if mp else None

class ShpMkr(object):
    """ class to make the daily shipment, include below functions
    .build the report if there is not and maintain the runnings
    .build the bc data
    .make the import
    .do invoice comparision
    Technique that I don't know:: UI under python, use tkinter, and it's simle messages
    """
    _mergeshpjo = False

    def __init__(self, cnsvc):
        self._cnsvc = cnsvc
        self._snrpt = "Rpt"
    
    def _newerr(self,loc,etype,msg):
        return {"location":loc,"type":etype,"msg":msg}

    def merge(self,fldr = None,tarfn = None):
        """ merge the files in given folder into one file """
        if not fldr:
            fldr = easydlg(filedialog.Directory(title="Choose folder contains one-time-shipment files"))
            if not path.exists(fldr): return
        fns = getfiles(fldr,".xls")
        dffn = [x for x in fns if x.find("_") >= 0]
        if len(dffn) > 0:
            logger.debug("target has already an result file(%s)" % os.path.basename(dffn[0]) )
            return dffn[0]
        ptn = re.compile(r"^HNJ \d+")
        kxl, app = xwu.app(False)
        wb = app.books.add()
        nshts = [x for x in wb.sheets]
        bfsht = wb.sheets[0]
        for fn in fns:
            if fn.find("_") >= 0: continue
            if not dffn and ptn.search(os.path.basename(fn)):
                sd = PajShpHdlr._getshpdate(fn)
                if sd:
                    dffn = "HNJ %s 出货明细_" % sd.strftime("%Y-%m-%d")
            wbx = xwu.safeopen(app, fn)
            try:
                for sht in wbx.sheets:
                    if sht.api.visible == -1:
                        sht.api.Copy(Before = bfsht.api)
            finally:
                wbx.close()
        for x in nshts:
            x.delete()
        if dffn:
            dffn = os.path.join(fldr,dffn)
            wb.save(dffn)
            logger.debug("merged file saved to %s" % dffn)
        elif kxl:
            app.quit()
        return dffn
    
    def readc1(self, sht, args):
        pass
    
    def readc2(self, sht, args):
        pass

    def readpaj(self, sht, args):
        """ return tuple(map,errlist)
        where errlist contain err locations
        """
        shps = PajShpHdlr._readshp(args["fn"],args["shpdate"],args["fmd"],sht)
        if not shps: return (None, None)
        mp, errs, shn = {}, [], sht.name
        for shp in shps.values():
            jn = shp.jono
            key = jn if self._mergeshpjo else jn + str(random.random())
            it = mp.setdefault(key,{"jono":jn,"qty":0,"location":(shn,jn)})
            qty, dfkarat, wgts = shp.qty, _getdefkarat(jn), shp.mtlwgt
            for wi in wgts:
                karat, wgt = wi.karat if wi.karat else dfkarat, wi.wgt
                if not wgt:
                    errs.append(self._newerr(jn,"Error","failed to get metal wgt"))
                else:
                    idx = it.get("karat1")
                    if idx:
                        idx = 1 if idx == karat else 2
                        if self._mergeshpjo:
                            wgt = (wgt * shp.qty + it["wgt%d" % idx] * it["qty"]) / (shp.qty + it["qty"])
                    else:
                        idx = 1
                it["karat%d" % idx] = karat
                it["wgt%d" % idx] = wgt
            it["qty"] += qty
        if mp:
            mp["invdate"] = shp.invdate
        return (mp, errs)
            
    def _genrpt(self, fldr):
        if os.path.isdir(fldr):
            fn = self.merge(fldr)
        else:
            fn = fldr
        if not fn: return
        pajopts = {"fn":fn,"shpdate":PajShpHdlr._getshpdate(fn),"fmd": datetime.fromtimestamp(os.path.getmtime(fn))}
        kxl, app = xwu.app()
        wb = app.books.open(fn)
        invs, its, errs, vt, vn = [], [], [], None, None
        rdrmap = {"长兴珠宝":("c2",self.readc2),"诚艺,胤雅":("c1",self.readc1) ,"十七,物料编号":("paj",self.readpaj)}
        rsht =[x for x in wb.sheets if triml(x.name) == "rpt"]
        if rsht:
            wb.close()
            return
        
        rdr = None
        for sht in wb.sheets:
            if not vt:
                for x in rdrmap.keys():
                    for y in x.split(","):
                        if xwu.find(sht, y):
                            vt, vn = x, rdrmap[x][0]
                            break
                    if vt: break
            rdr = rdrmap.get(vt)
            if rdr:
                print("Processing sheet(%s)" % sht.name)
                lst = rdr[1](sht, pajopts)
                if lst and len([x for x in lst if x]):
                    mp = lst[0]
                    if mp: 
                        if "invdate" in mp:
                            ivd = mp["invdate"]
                            if isinstance(ivd,str):
                                ivd = datetime.strptime(ivd, "%Y-%m-%d")
                            td = datetime.today() - ivd
                            if td.days > 2 or td.days < 0:
                                errs.append(self._newerr("_Date_","warning","Maybe invalid file or invoice date"))
                            del mp["invdate"]
                        its.extend(mp.values())
                    if lst[1]: errs.extend(lst[1])
                else:
                    logger.debug("sheet does not contain standard shipment data")
                    if rdr[0] == "paj":
                        invno = PajShpHdlr._readinvno(sht)
                        if not invno: continue
                        lst = PajShpHdlr._rawreadinv(sht, invno)
                        if not lst:
                            invs.extend(lst)
            else:
                logger.debug("no suitable reader for sheet(%s)" % sht.name)
        if its:
            jns = set([x["jono"] for x in its])
            with self._cnsvc.sessionctx():
                jos = self._cnsvc.getjos(jns)
                jos = dict([(x.name.value,x) for x in jos[0]])
                nmap = {"cstname":"customer.name","styno":"style.name.value","running":"running","description":"description","qtyleft":"qtyleft"}
                for mp in its:
                    jo = jos.get(mp["jono"])
                    if not jo:
                        errs.append(self._newerr(mp["location"],"Error","Invalid JO#(%s)" % mp["jono"]))
                        continue
                    for y in nmap.items():
                        sx = deepget(jo,y[1])
                        if sx and isinstance(sx,str): sx = sx.strip()
                        mp[y[0]] = sx
                    jo.qtyleft = jo.qtyleft - mp["qty"]
                    if jo.qtyleft < 0:
                        errs.append(self._newerr(mp["location"],"Error","Qty not enough"))
            its = sorted(its,key = lambda d0: "%s,%6.1f" % (d0["jono"],d0["qtyleft"]))
            self._write(wb,its,errs,ivd,vn)
        return fn

    def _write(self,wb,its,errs,invdate,vdrname):
        app = wb.app
        sts = PajNSOFRdr().readsettings()
        fn = sts.get(triml("Shipment.IO")).value
        wbio, iorst = app.books.open(fn), {}
        shtio = wbio.sheets["master"]
        nls = [x for x in NamedLists(xwu.usedrange(shtio).value)]
        itio, ridx = nls[-1], len(nls) + 2
        je = JOElement(itio["n#"])
        iorst["n#"], iorst["date"] = "%s%d" % (je.alpha,je.digit + 1), invdate        
        pfx = invdate.strftime("%y%m%d")
        if vdrname != "paj": pfx = pfx[1:]
        pfx = 'J' + pfx
        existing = [x["jmp#"] for x in nls[-20:] if x["jmp#"].find(pfx) == 0]
        if existing:
            if vdrname != "paj" :
                logger.debug("%s should not have more than one shipment in one date" % vdrname)
                return
            sfx = "%01d" % (int(max(existing)[-1])+1)
        else:
            sfx = "01" if vdrname == "paj" else trimu(vdrname)
        iorst["jmp#"] = pfx + sfx
        for idx in range(len(nls) - 1,0,-1):
            jn = nls[idx]["jmp#"]
            flag = (jn.find("C") >= 0) ^ (vdrname == "paj")
            if flag: break
        iorst["maxrun#"] = int(nls[idx]["maxrun#"])

        s0 = sts.get("shipment.rptmgns.%s" % vdrname)
        if not s0: s0 = sts.get("shipment.rptmgns")
        sht = wb.sheets.add(name = self._snrpt)
        pfx = "sht.api.pagesetup"
        shtcmds = [ pfx + ".%smargin=%s" % tuple(y.split("=")) for y in triml(s0.value).split(";")]
        shtcmds.append(pfx + ".printtitlerows='$1:$1'")
        shtcmds.append(pfx + ".leftheader='%s'" % ("%s年%s月%s日落货资料" % tuple(invdate.strftime("%Y-%m-%d").split("-"))))
        shtcmds.append(pfx + ".centerheader='%s'" % iorst["jmp#"])
        shtcmds.append(pfx + ".rightheader='%s'" % iorst["n#"])
        shtcmds.append(pfx + ".rightfooter='&P of &N'")
        shtcmds.append(pfx + ".fittopageswide=1")
        for x in shtcmds:
            exec(x)
        
        s0 = sts.get("shipment.hdrs." + vdrname)
        if not s0: s0 = sts.get("shipment.hdrs")
        ttl, ns = [], {}
        for x in s0.value.replace("\\n","\n").split(";"):
            y = x.split("=")
            y1 = y[1].split(",")
            ttl.append(y[0])
            if len(y1) > 1:
                ns[y1[0]] = y[0]
            sht.range(1,len(ttl)).column_width = float(y1[len(y1) - 1])
        nl, maxr = NamedList(list2dict(ttl,ns)), iorst["maxrun#"]
        lsts, ns, hls = [ttl], "jono,running,qty,cstname,styno,description,qtyleft,karat1,wgt1".split(","), []
        for it in its:
            if not it["running"]:
                maxr += 1
                it["running"], nl["running"] = maxr, maxr
                hls.append(len(lsts))
            ttl = ["" for x in range(len(ttl))]
            nl.setdata(ttl)
            for col in ns:
                nl[col] = it[col]
            nl.jono = "'" + nl.jono
            lsts.append(ttl)
            if "karat2" in it:
                nl.setdata(["" for x in range(len(ttl))])
                nl.karat, nl.wgt = it["karat2"], it["wgt2"]
                lsts.append(ttl)
        iorst["maxrun#"] = maxr
        sht.range(1,1).value = lsts

        #write sum formula at the bottom
        s0 = int(nl.getcol("qty")) + 1
        sht.range(len(lsts) + 1, s0).formula = "=sum(%s1:%s%d)" % (col_name(s0), col_name(s0), len(lsts))

        #high-light the new runnings
        #final step
        for knv in iorst.items():
            shtio.range(ridx,itio.getcol(knv[0])+1).value = knv[1]