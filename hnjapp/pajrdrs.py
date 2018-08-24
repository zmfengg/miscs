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
from datetime import date, datetime, timedelta
import time
from decimal import Decimal
from tkinter import filedialog

import xlwings.constants as const
from xlwings.utils import col_name
from xlwings import Book
from sqlalchemy import and_, func
from sqlalchemy.engine import create_engine
from sqlalchemy.orm import Query, aliased
from xlwings.constants import LookAt, BorderWeight, LineStyle, FormatConditionOperator, FormatConditionType, Constants
from hnjapp.c1rdrs import C1InvRdr

from hnjcore import JOElement, appathsep, deepget, karatsvc, p17u, xwu, samekarat
from hnjcore.models.hk import JO as JOhk, PajAck
from hnjcore.models.hk import Orderma, PajInv, PajShp, PajCnRev, POItem
from hnjcore.models.hk import Style as Styhk
from hnjcore.utils import daterange, getfiles, isnumeric, p17u
from hnjcore.utils.consts import NA
from utilz import NamedList, NamedLists, ResourceCtx, SessionMgr, list2dict, splitarray, triml, trimu, easydialog

from .common import _logger as logger, _getdefkarat
from .dbsvcs import BCSvc, CNSvc, HKSvc, jesin
from .localstore import PajInv as PajInvSt
from .localstore import PajItem, PajCnRev as PajCnRevSt
from .localstore import PajWgt as PrdWgtSt
from .pajcc import PAJCHINAMPS, P17Decoder, PajCalc, PrdWgt, WgtInfo, MPS, _tofloat, cmpwgt, addwgt

_accdfmt = "%Y-%m-%d %H:%M:%S"
_appmgr = xwu.appmgr()

def _accdstr(dt):
    """ make a date into an access date """
    return dt.strftime(_accdfmt) if dt and isinstance(dt, date) else dt


def _removenonascii(s0):
    """remove thos non ascii characters from given string"""
    if isinstance(s0, str):
        return "".join([x for x in s0 if ord(x) > 31 and ord(x) < 127 and x != "?"])
    return s0

class PajBomHhdlr(object):
    """ methods to read BOMs from PAJ """

    @classmethod
    def readbom(self,fldr):
        """ read BOM from given folder
        @param fldr: the folder contains the BOM file(s)
        return a dict with "pcode" as key and dict as items
            the item dict has keys("pcode","mtlwgt")
        """
        _ptnoz = re.compile(r"\(\$\d*/OZ\)")
        _ptnsil = re.compile(r"(925)|(银)")
        _ptncop = re.compile(r"(BRONZE)|(铜)")
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
                #bronze does not have oz
                if not mt and not _ptncop.search(mat): return        
            if _ptnsil.search(mat):
                kt = 925
            elif _ptncop.search(mat):
                kt = 200
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
        
        def _isring(pcode):
            return _pcdec.decode(pcode,"PRODTYPE").find("戒") >= 0
        
        def _readwb(wb, pmap):
            """ read bom in the given wb to pmap
            """
            shts = [0,0]
            for sht in wb.sheets:
                rng = xwu.find(sht, u"十七位")
                if not rng: continue
                if xwu.find(sht, u"抛光后"):
                    shts[0] = (sht,rng)
                elif xwu.find(sht, u"物料特征"):
                    shts[1] = (sht,rng)
                if all(shts): break
            if not all(shts): return
            #duplicated item detection
            mstrs, pts = set(), set()
            shts[0][0].name, shts[1][0].name = "BOM_mstr", "BOM_part"
            nmps = {0:{u"pcode":"十七位,","mat":u"材质,","mtlwgt":u"抛光,","up":"单价","fwgt":"成品重"},1:{"pcode":u"十七位,","matid":"物料ID,","name":u"物料名称", "spec":u"物料特征","qty":u"数量","wgt":u"重量","unit":u"单位","length":u"长度"}}
            nis0 = lambda x: x if x else 0
            for jj in range(len(shts)):
                vvs = shts[jj][1].end("left").expand("table").value
                nls = NamedLists(vvs,nmps[jj])
                if jj == 0:
                    for nl in nls:
                        pcode = nl.pcode                        
                        if not p17u.isvalidp17(pcode): break
                        fpt = tuple(nis0(x) for x in (nl.pcode, nl.mat, nl.up, nl.mtlwgt, nl.fwgt))
                        key = ("%s" * len(fpt)) % fpt
                        if key in mstrs:
                            logger.debug("duplicated bom_mstr found(%s, %s)" % (nl.pcode, nl.mat))
                            continue
                        mstrs.add(key)
                        kt = _parsekarat(nl.mat)
                        if not kt: continue
                        it = pmap.setdefault(pcode,{"pcode":pcode})
                        it.setdefault("mtlwgt",[]).append((kt,nl.mtlwgt))
                elif jj == 1:
                    nmp = [x for x in nmps[jj].keys() if x.find("pcode") < 0]
                    for nl in nls:
                        pcode = nl.pcode
                        if not p17u.isvalidp17(pcode): break
                        fpt =  tuple(nis0(x) for x in (nl.pcode, nl.matid, nl.name, nl.spec, nl.qty, nl.wgt, nl.unit, nl.length))
                        key = ("%s" * len(fpt)) % fpt
                        if key in pts:
                            logger.debug("duplicated bom_part found(%s, %s)" % (nl.pcode, nl.name))
                            continue
                        pts.add(key)
                        it = pmap.setdefault(pcode,{"pcode":pcode})
                        mats, it = it.setdefault("parts",[]), {}
                        mats.append(it)
                        for cn in nmp:
                            it[cn] = nl[cn]        
        pmap = {}
        if isinstance(fldr, Book):
            _readwb(fldr, pmap)
        else:
            fns = getfiles(fldr,"xls") if path.isdir(fldr) else (fldr, )
            if not fns: return
            app, kxl = _appmgr.acq()
            try:            
                for fn in fns:
                    wb = app.books.open(fn)
                    _readwb(wb, pmap)                                        
                    wb.close()
            finally:
                if kxl and app: _appmgr.ret(kxl)
        
        for x in pmap.items():
            lst = x[1]["mtlwgt"]
            prdwgt = None
            for y in lst:
                prdwgt = addwgt(prdwgt,WgtInfo(y[0],y[1]))
            if "parts" in x[1]:
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
                    kt = _parsekarat(nm,prdwgt.wgts,False)
                    if not kt: continue                
                    y["karat"] = kt
                    ispart = False
                    if ispendant:
                        if haschain:
                            isch = triml(nm).find("chain") >= 0
                            ispart = isch or (haskou and (_ptfrcchain.search(nm) or nm.find("圈") >= 0))
                            if isch and not chlenerr:
                                lc = y["length"]
                                if not lc is None:
                                    try:
                                        lc = float(lc)
                                    except:
                                        lc = 0
                                    if lc > 0: chlenerr = True
                            if ispart:
                                wgt0 = prdwgt.part
                                ispart = (not wgt0 or y["karat"] == wgt0.karat)
                            if not ispart:
                                if isch:
                                    logger.debug("No wgt slot for chain(%s) in pcode(%s),merged to main" % (y["name"],x[0]))
                                else:
                                    logger.debug("parts(%s) in pcode(%s) merged to main" % (y["name"],x[0]))
                    #turn autoswap off in parts appending procedure to avoid main karat being modified
                    prdwgt = addwgt(prdwgt,WgtInfo(y["karat"],y["wgt"]), ispart, autoswap = False)
                if False and chlenerr:
                    #in common  case, chain should not have length, when this happen
                    #make the weight negative. Skipped
                    prdwgt = prdwgt._replace(part = WgtInfo(prdwgt.part.karat, -prdwgt.part.wgt))
            x[1]["mtlwgt"] = prdwgt
        
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
                lst, wgt = [x["pcode"]], x["mtlwgt"]
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

            app, kxl = _appmgr.acq()
            wb = app.books.add()
            sns, data = "BOMData,JOs".split(","), (vvs,jos)
            for idx in range(len(sns)):
                sht = wb.sheets[idx]
                sht.name = sns[idx]
                sht.range(1,1).value = data[idx]
                sht.autofit("c")
            wb.save(fn)
            ffn = wb.fullname
            _appmgr.ret(kxl)
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
        _ptngwt = re.compile(r"[\d.]+")
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
                        mw.append(WgtInfo(0,float(mt.group()) if mt else None))
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
        """
        read the invoice, return a map with inv#+jo# as key and PajInvItem as item
        """
        PajInvItem = namedtuple(
            "PajInvItem", "invno,pcode,jono,qty,uprice,mps,stspec,lastmodified")
        mp = {}
        rng = xwu.find(sht, "Item*No", lookat=const.LookAt.xlWhole)
        if not rng:
            return
        if not invno: invno = self._readinvno(sht)
        if sht.name != invno: sht.name = invno
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
            if not jn:
                logger.debug("No JO# found for p17(%s)" % p17)
                continue
            key = invno + "," + jn
            if key in mp.keys():
                it = mp[key]
                mp[key] = it._replace(qty=it.qty + tr.qty)
            else:
                mps = MPS("S=%3.2f;G=%3.2f" % (tr.silver, tr.gold)).value \
                    if th.getcol("gold") and th.getcol("silver") else "S=0;G=0"
                it = PajInvItem(invno, p17, jn, tr.qty, tr.uprice, mps, tr.stspec, fmd)
                mp[key] = it
        return mp

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
        def _extring(x):
            return x[:9] + x[11:]
        items, td0, qmap = {}, datetime.today(), None
        nls = tuple(NamedLists(vvs,{"odx": u"订单号", "invdate": u"发票日期", "odseq": u"订单序号","stwgt": u"平均单件石头,XXX", "invno": u"发票号", "orderno": u"订单号序号", "pcode": u"十七位,十七,物料","mtlwgt": u"平均单件金,XX", "jono": u"工单,job", "qty": u"数量", "cost": u"cost"}))
        th = nls[0]
        x = [x for x in "invno,pcode,jono,qty,invdate".split(
            ",") if th.getcol(x) is None]
        if x:
            return

        def _getbomwgt(bomap, bomapring, pcode):            
            """ in the case of ring, there is only one code there
            """
            if not (bomap and pcode): return
            prdwgt = bomap.get(pcode)
            if not prdwgt: # and is ring                
                if pcode[1] == "4" and bomapring:
                    pcode = _extring(pcode)
                    prdwgt = bomapring.get(pcode)
            if not prdwgt:
                print("failed to get bom wgt for pcode(%s)" % pcode)
            return prdwgt
        bfn = path.basename(fn).replace("_", "")
        shd = PajShpHdlr._getshpdate(sht.name, False)
        if shd:
            df = shd - fshd
            shd = shd if abs(df.days) <= 7 else fshd 
        else:
            shd = fshd 
        bomwgts = PajBomHhdlr.readbom(sht.book)
        if bomwgts: 
            bomwgtsrng = dict([(_extring(x[0]), x[1]["mtlwgt"]) for x in bomwgts.items() if x[0][1] == "4" ])
            bomwgts = dict([(x[0],x[1]["mtlwgt"]) for x in bomwgts.items()])
        else:
            bomwgtsrng, bomwgts = (None,) * 2
        if not th.getcol("cost"):
            for tr in nls:
                if not tr.pcode:
                    break
                jono = JOElement(tr.jono).value
                mwgt = _getbomwgt(bomwgts, bomwgtsrng, tr.pcode)
                if not mwgt:
                    mwgt = tr.get("mtlwgt")
                    if not mwgt: mwgt = 0
                    mwgt = PrdWgt(WgtInfo(_getdefkarat(jono),mwgt,4))
                    bomsrc = False
                else:
                    bomsrc = True
                invno = tr.invno
                if not invno: invno = "N/A"
                if th.getcol('orderno'):
                    odno = tr.orderno
                elif len([1 for x in "odx,odseq".split(",") if th.getcol(x)]) == 2:
                    odno = tr.odx + "-" + tr.odseq
                else:
                    odno = "N/A"
                stwgt = tr.get("stwgt")
                if stwgt is None: stwgt = 0
                thekey = "%s,%s,%s" % (jono,tr.pcode,invno)
                if thekey in items:
                    #order item's weight does not have karat event, so append it to 
                    #the main, but bom source case, no weight-recalculation is neeeded
                    si = items[thekey]
                    wi = si.mtlwgt
                    if not bomsrc:
                        wi = wi._replace(main = wi.main._replace(wgt = _tofloat((wi.main.wgt * si.qty + mwgt.main.wgt * tr.qty)/(si.qty + tr.qty),4)))
                    items[thekey] = si._replace(qty = si.qty + tr.qty, mtlwgt = wi)
                else:
                    ivd = tr.invdate
                    si = PajShpItem(bfn, odno, jono, tr.qty, tr.pcode, invno, ivd, mwgt, stwgt, ivd, fmd, td0)
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
                    prdwgt = _getbomwgt(bomwgts, bomwgtsrng, p17)
                    if p17 in qmap:
                        #metal and stone weights
                        mns = qmap[p17]
                    else:
                        logger.info("failed to get quo info for pcode(%s)" % p17)
                        mns = ((None,),0)
                    if not prdwgt:                        
                        wis = list(mns[0])
                        for idx in range(len(wis)):
                            wi = wis[idx]
                            if not wi: continue
                            if not wi.karat: wis[idx] = wi._replace(karat = _getdefkarat(tr.jono))
                        prdwgt = PrdWgt(*wis)
                    si = PajShpItem(bfn, odno, JOElement(tr.jono).value, tr.qty, p17, tr.invno, ivd, prdwgt, mns[1],ivd, fmd, td0)
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
                        dct["mtlwgt"] = sum([x.wgt for x in dct["mtlwgt"].wgts if x])
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
        app, kxl = _appmgr.acq()
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
            _appmgr.ret(kxl)
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
        app, kxl = _appmgr.acq()
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
            _appmgr.ret(kxl)
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
        q0 = Query(PajCnRev).filter(and_(PajCnRev.filldate > lastcrdate,PajCnRev.tag == 0, PajCnRev.revdate >= affdate))
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
            

class PajNSOFRdr(object):
    """
    class to read a NewSampleOrderForm's data out
    """
    _tplfn = r"\\172.16.8.46\pb\dptfile\pajForms\PAJSKUSpecTemplate.xlt"

    def readsettings(self, fn = None):
        usetpl, mp = False, None
        if not fn:
            fn, usetpl = self._tplfn, True
        app, kxl = _appmgr.acq()
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
            _appmgr.ret(kxl)
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
    #warning and error category
    _ec_qty = "ERROR.QTYLEFT"
    _ec_jn = "ERROR.JO#"
    _ec_wgt = "ERROR.WGTNA"
    _wc_wgt = "WARN.WGTDIFF"
    _wc_ack = "WARN.ACK"
    _wc_date = "WARN.DATE"
    _wc_qty = "WARN.QTYSHP&INV"

    def __init__(self, cnsvc, hksvc, bcsvc):
        self._cnsvc, self._hksvc, self._bcsvc = cnsvc, hksvc, bcsvc
        self._snrpt, self._snerr, self._snwarn, self._snbc = "Rpt", "错误", "警告", "BCData"
    
    def _newerr(self,loc,etype,msg):
        return {"location":loc,"type":etype,"msg":msg}

    def _pajfldr2file(self, fldr):
        if not fldr:
            fldr = easydialog(filedialog.Directory(title="Choose folder contains all raw files from PAJ"))
            if not path.exists(fldr): return
        sts = self._nsofsettings
        tarfldr, tarfn = path.dirname(sts.get("shp.template").value), None
        fns = getfiles(fldr,".xls")        
        ptn = re.compile(r"^HNJ \d+")
        for fn in fns:
            if ptn.search(path.basename(fn)):
                sd = PajShpHdlr._getshpdate(fn)
                if sd:
                    tarfn = "HNJ %s 出货明细"  % sd.strftime("%Y-%m-%d")
                    break
        if not tarfn:
            return
        sts = getfiles(tarfldr, tarfn)
        if sts:
            tarfn = sts[0]
            logger.debug("result file(%s) already there" % tarfn)
            return tarfn
        
        if len(fns) == 1:
            return fns[0]
        
        app, kxl = _appmgr.acq()
        wb = app.books.add()
        nshts = [x for x in wb.sheets]
        bfsht = wb.sheets[0]
        for fn in fns:
            wbx = xwu.safeopen(app, fn)
            try:
                for sht in wbx.sheets:
                    if sht.api.visible == -1:
                        sht.api.Copy(Before = bfsht.api)
            finally:
                wbx.close()
        for x in nshts:
            x.delete()
        if tarfn:
            wb.save(path.join(tarfldr,tarfn))
            tarfn = wb.fullname
            logger.debug("merged file saved to %s" % tarfn)
            wb.close()
        _appmgr.ret(kxl)
        return tarfn
    
    def _readc1(self, sht, args):
        """ determine the header row """
        for shp in sht.shapes:
            shp.delete()
        ridx, flag = -1, False
        for row in xwu.usedrange(sht).rows:
            if not row.api.entirerow.hidden:
                ridx = row.row
                break
            else:
                flag = True
        if flag and ridx >= 0:
            sht.range("1:%d" % ridx).api.entirerow.delete
        rng = xwu.find(sht, "日期")
        if not rng:
            logger.debug("not valid data in sheet(%s)" % sht.name)
            return
        if rng: args["shpdate"] = rng.offset(0, 1).value
        mp, errs, shn, its ={}, [], sht.name, C1InvRdr._readc1(sht)
        for shp in its:
            jn = shp.jono
            key = jn if self._mergeshpjo else jn + str(random.random())
            it = mp.setdefault(key,{"jono":jn,"qty":0,"location": "%s,%s" % (shn,jn)})
            it["mtlwgt"] = shp.mtlwgt
            it["qty"] += shp.qty
        if mp:
            mp["invdate"] = args.get("shpdate")
        return (mp, errs)
    
    def _readc2(self, sht, args):
        pass

    def _readpaj(self, sht, args):
        """ return tuple(map,errlist)
        where errlist contain err locations
        """
        shps = PajShpHdlr._readshp(args["fn"],args["shpdate"],args["fmd"],sht)
        if not shps: return (None, None)
        
        mp, errs, shn = {}, [], sht.name
        for shp in shps.values():
            jn = shp.jono
            key = jn if self._mergeshpjo else jn + str(random.random())
            it = mp.setdefault(key,{"jono":jn,"qty":0,"location": "%s,%s" % (shn,jn)})
            it["mtlwgt"] = shp.mtlwgt
            it["qty"] += shp.qty
        if mp:
            mp["invdate"] = shp.invdate
        return (mp, errs)
            
    def _shpcheck(self, shpmp, invmp, errlst):
        """ check the source data about weight/pajinv
        @param shpmp: the shipment map with JO# as key and map as value
        """
        jns = set([x["jono"] for x in shpmp])
        #pajinv verify
        if invmp:
            logger.debug("begin to fetch ack/inv data")
            t0 = time.clock()
            with self._hksvc.sessionctx() as cur:
                q = Query([JOhk,PajAck]).join(PajAck).filter(jesin(set([JOElement(x) for x in jns]),JOhk))
                q = q.with_session(cur).all()
                logger.debug("using %fs to fetch %d JOs for above action" % (time.clock() - t0, len(jns)))
                acks = dict([(x[0].name.value,(x[1].uprice,x[1].mps,x[1].ackdate,x[1].docno,x[1].mps, x[1].pcode))for x in q]) if q else {}
                if acks: nlack = NamedList(list2dict("uprice,mps,date,docno,mps,pcode".split(",")))
            tmp = {}
            for x in invmp.values():
                tmp1 = tmp.setdefault(x.jono,{"jono":x.jono})
                if "inv" not in tmp1:
                    tmp1["inv"], tmp1["invqty"] = x, 0
                else:
                    x0 = tmp1["inv"]
                    if abs(x0.uprice - x.uprice) > 0.001:
                        errlst.append(self._newerr(x.jono, self._wc_ack, "工单(%s)对应的发票单价不一致" % x.jono))
                tmp1["invqty"] += x.qty
            for x in shpmp:
                jn = x["jono"]
                tmp1 = tmp.get(jn)
                if tmp1:
                    tmp1["qty"] = tmp1.get("qty",0) + x["qty"]
            for x in tmp.values():
                if x.get("invqty") != x.get("qty"):
                    if not x.get("qty"):
                        errlst.append(self._newerr("Inv(%s),JO#(%s)" % (x["inv"].invno, x["jono"]), self._wc_qty , "未发现发票号(%s)对应工单号(%s)的落货记录" % (x["inv"].invno, x["jono"])))
                    else:
                        errlst.append(self._newerr(x["jono"], self._wc_qty , "落货数量(%s)与发票数量(%s)不一致" % (str(x.get("qty", 0)), str(x.get("invqty", 0)))))
                if acks:
                    ack = acks.get(x["jono"])
                    if not ack: continue
                    inv = x["inv"]
                    nlack.setdata(ack)
                    if abs(inv.uprice - float(nlack.uprice)) > 0.01:
                        errlst.append(self._newerr(x["jono"], self._wc_ack , "%s发票单价(%s@%s@%s)与\r\nAck.单价(%s@%s@%s)不一致.\r\nAck文件（%s）,日期(%s)" % ("+" if inv.uprice > nlack.uprice else "",  str(inv.uprice), inv.mps, inv.pcode, str(nlack.uprice), nlack.mps, nlack.pcode, nlack.docno,nlack.date.strftime("%Y-%m-%d"))))
            tmp1 = jns.difference(tmp.keys())
            if tmp1:
                for x in tmp1:
                    errlst.append(self._newerr(jn, self._wc_qty , "未发现工单号(%s)对应的发票" % x))
        t0 = time.clock()
        logger.debug("Begin to verify shipment qty&wgt")
        with self._cnsvc.sessionctx():
            jos = self._cnsvc.getjos(jns)
            jos = dict([(x.name.value,x) for x in jos[0]])
            jwgtmp = {}
            nmap = {"cstname":"customer.name","styno":"style.name.value","running":"running","description":"description","qtyleft":"qtyleft"}
            for mp in shpmp:
                jn = mp["jono"]
                jo = jos.get(jn)
                if not jo:
                    errlst.append(self._newerr(mp["location"], self._ec_jn ,"错误的工单 JO#(%s)" % jn))
                    continue
                for y in nmap.items():
                    sx = deepget(jo,y[1])
                    if sx and isinstance(sx,str): sx = sx.strip()
                    mp[y[0]] = sx
                jo.qtyleft = jo.qtyleft - mp["qty"]
                if jo.qtyleft < 0:
                    s0 = "数量不足"
                    errlst.append(self._newerr(mp["location"], self._ec_qty,s0))
                    mp["errmsg"] = s0
                elif jo.qtyleft > 0:
                    s0 = "数量有余"
                    errlst.append(self._newerr(mp["location"], self._wc_qty,s0))
                    mp["errmsg"] = s0
                else:
                    mp["errmsg"] = ""
                jwgt = jwgtmp.get(jn)
                if not jwgt and jn not in jwgtmp:
                    jwgt = self._hksvc.getjowgts(jn)
                    jwgtmp[jn] = jwgt
                if jwgt:
                    if not cmpwgt(jwgt,mp["mtlwgt"]):
                        errlst.append(self._newerr(mp["location"], self._wc_wgt if [x for x in mp["mtlwgt"].wgts if x and x.wgt > 0] else self._ec_wgt , "落货(%s),金控(%s)" % (mp["mtlwgt"], jwgt)))
        logger.debug("using %fs for above action" % (time.clock() - t0))

    def _writebc(self, wb, shpmp, newrunmp, invdate):
        """ create a bc template
        """
        dmp, lsts, rcols = {}, [], "lymd,lcod,styn,mmon,mmo2,runn,detl,quan,gwgt,gmas,jobn,ston,descn,desc,rem1,rem2,rem3,rem4,rem5,rem6,rem7,rem8".split(",")
        refjo, refpo, refodma = aliased(JOhk), aliased(POItem), aliased(Orderma)
        rems, nl =[999,0], NamedList(list2dict(rcols))
        for x in nl._colnames:
            if x.find("rem") == 0:
                idx = int(x[len("rem"):])
                if idx < rems[0]: rems[0] = idx
                if idx > rems[1]: rems[1] = idx
        rems[1] += 1
        with self._hksvc.sessionctx() as cur:
            dt = datetime.today() - timedelta(days = 365)
            jes = set(JOElement(x["jono"]) for x in shpmp)
            logger.debug("begin to select same sku items for BC")
            t0 = time.clock()
            q = Query([JOhk.name,func.max(refjo.running)]).join((refjo,JOhk.id != refjo.id), (POItem, JOhk.poid == POItem.id), (Orderma, JOhk.orderid == Orderma.id), (refodma, refjo.orderid == refodma.id), (refpo, and_(POItem.skuno != '', refpo.id == refjo.poid, refpo.skuno == POItem.skuno))).filter(and_(POItem.id >0, refjo.createdate > dt, Orderma.cstid == refodma.cstid)).group_by(JOhk.name)
            lst = []
            for arr in splitarray(jes,20):
                qx = q.filter(jesin(arr,JOhk))
                lst0 = qx.with_session(cur).all()
                if lst0: lst.extend(lst0)
            logger.debug("using %fs to fetch %d records for above action" % (time.clock() - t0, len(lst)))
            josku = dict([(x[1], x[0].value) for x in lst if x[1] > 0]) if lst else {}
            
        joskubcs = self._bcsvc.getbcs([x for x in josku.keys()])
        joskubcs = dict([(josku[int(x.runn)],x) for x in joskubcs]) if joskubcs else {}
        
        stynos = set([x["styno"] for x in shpmp if x["jono"] not in joskubcs])
        bcs = self._bcsvc.getbcs(stynos, True)
        for it in bcs:
            dmp.setdefault(it.styn,[]).append(it)
        for x in dmp.keys():
            dmp[x] = sorted(dmp[x], key = lambda x: x.runn, reverse = True)
        bcmp, dmp, lymd = dmp, {}, invdate.strftime("%Y%m%d %H:%M%S")

        lsts.append(rcols)
        for it in shpmp:
            jn = it["jono"]
            if jn not in newrunmp or jn in dmp: continue
            dmp[jn], styno = 1, it["styno"]
            bc, rmks = joskubcs.get(jn), []
            if not bc:
                bcs = bcmp.get(styno)
                if bcs: 
                    #find the same karat and longest remarks as template
                    for bcx in bcs[:10]:
                        if not samekarat(jn,bcx.jobn): continue
                        mc0 = [x for x in [getattr(bcx,"rem%d" %y).strip() for y in range(*rems)] if x]
                        if len(mc0) > len(rmks):
                            rmks, bc = mc0, bcx
                    if not bc:
                        bc = bcs[0]
                        rmks = [x for x in [getattr(bc,"rem%d" %y).strip() for y in range(*rems)] if x]
                flag = False
            else:
                flag = True
            nl.setdata([None]*len(rcols))
            nl.lymd, nl.lcod, nl.styn, nl.mmon = lymd, styno, styno, "'" + lymd[2:4]
            nl.mmo2, nl.runn, nl.detl = lymd[4:6], "'%d" % it["running"], it["cstname"]
            nl.quan, nl.jobn = it["qty"], "'" + jn
            nl.descn = ("---" if flag else "") + it["description"]
            prdwgt = it["mtlwgt"]
            nl.gmas, nl.gwgt = prdwgt.main.karat, "'" + str(prdwgt.main.wgt)
            if not bc:
                nl.ston, nl.desc = "--", "TODO"
            else:
                nl.ston, nl.desc = bc.ston, bc.desc
                rmks = [x for x in [getattr(bc,"rem%d" % y).strip() for y in range(*rems)] if x]
            nrmks = []
            for x in ((prdwgt.aux,"*%s %4.2f"),(prdwgt.part,"*%sPTS %4.2f")):
                if x[0]: nrmks.append(x[1] % (karatsvc.getkarat(x[0].karat).name,x[0].wgt))                
            cn = len(nrmks) + len(rmks) - rems[1] + 1
            if cn > 0:
                rmks[-cn-1] = ";".join(rmks[-cn-1:])
                nrmks.extend(rmks[:-cn])
            else:
                nrmks.extend(rmks)
            for idx in range(len(nrmks)):
                nl["rem%d" % (idx + 1)] = nrmks[idx]
            lsts.append(nl.data)
        sht = wb.sheets.add(self._snbc)
        sht.range(1,1).value = lsts
        sht.autofit()

    @property
    def _nsofsettings(self):
        if not hasattr(self, "_nsofsts"):
            self._nsofsts = PajNSOFRdr().readsettings()
        return self._nsofsts

    def _writerpts(self,wb,shpmp,errlst,newrunmp,invdate,vdrname):
        """ send the shipment and errs to related sheets(Rpt/Err)
        """
        app = wb.app
        sts = self._nsofsettings

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
            sfx = "%d" % (int(max(existing)[-1])+1)
        else:
            sfx = "1" if vdrname == "paj" else trimu(vdrname)
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
        ns["thisleft"] = "此次,"
        nl, maxr = NamedList(list2dict(ttl,ns)), iorst["maxrun#"]
        lsts, ns, hls = [ttl], "jono,running,qty,cstname,styno,description,qtyleft,errmsg".split(","), []
        for it in shpmp:
            ttl = ["" for x in range(len(ttl))]
            nl.setdata(ttl)
            if not it["running"]:
                if it["jono"] not in newrunmp:
                    maxr += 1
                    it["running"], nl["running"] = maxr, maxr
                    hls.append(len(lsts) + 1)
                    newrunmp[it["jono"]] = maxr
            else:
                newrunmp[it["jono"]] = it["running"]
            for col in ns:
                nl[col] = it[col]
            nl.jono = "'" + nl.jono
            karats = {}
            for wi in it["mtlwgt"].wgts:
                if not wi: continue
                if wi.karat not in karats:
                    karats[wi.karat] = wi.wgt
                else:
                    karats[wi.karat] += wi.wgt
            karats = [(x[0],x[1]) for x in karats.items()]
            nl.karat1, nl.wgt1 = karats[0][0], karats[0][1]
            lsts.append(ttl)
            if len(karats) > 1:
                jn = nl.jono
                for idx in range(1,len(karats)):
                    nl.setdata(["" for x in range(len(ttl))])
                    nl.jono = jn
                    nl.karat1, nl.wgt1 = karats[idx][0], karats[idx][1]
                lsts.append(nl.data)
        sht.range(1,1).value = lsts
        #the qtyleft formula
        s0, s1, s2 = col_name(nl.getcol("qty") + 1), col_name(nl.getcol("qtyleft") + 1), col_name(nl.getcol("thisleft") + 1)
        for idx in range(2,len(lsts) + 1):
            rng = sht.range("%s%d" % (s2,idx))
            rng.formula = "=%s%d-%s%d" % (s1,idx,s0,idx)
            rng.api.numberformatlocal = "_ * #,##0_ ;_ * -#,##0_ ;_ * ""-""_ ;_ @_ "
            rng.api.formatconditions.add(FormatConditionType.xlCellValue , FormatConditionOperator.xlLess, "0")
            rng.api.formatconditions(1).interior.colorindex = 3
        
        rng = sht.range(sht.range(1,1),sht.range(len(lsts),len(nl._colnames)))
        rng.api.borders.linestyle = LineStyle.xlContinuous
        rng.api.borders.weight = BorderWeight.xlThin

        #write sum formula at the bottom
        s0 = int(nl.getcol("qty")) + 1
        rng = sht.range(len(lsts) + 1, s0)
        rng.formula = "=sum(%s1:%s%d)" % (col_name(s0), col_name(s0), len(lsts))
        rng.api.font.bold = True
        rng.api.borders.linestyle = LineStyle.xlContinuous
        rng.api.borders.weight = BorderWeight.xlThin
        sht.range("A2:A%d" % (len(lsts) + 1)).row_height = 18
        xwu.usedrange(sht).api.VerticalAlignment = Constants. xlCenter
        if hls:
            s0 = nl.getcol("running") + 1
            for idx in hls:
                sht.range(idx,s0).api.interior.colorindex = 6
                sht.range(idx,s0).api.interior.pattern = 1

        #write IOs back
        iorst["maxrun#"] = maxr
        for knv in iorst.items():
            shtio.range(ridx,itio.getcol(knv[0])+1).value = knv[1]
        
        if errlst:
            sns = (self._snerr, self._snwarn)
            data = self._errandwarn(errlst)
            for idx in range(len(sns)):
                shpmp = [[x for x in errlst[0].keys()]]
                ttl = shpmp[0]            
                for mp in data[idx]:
                    shpmp.append([("%s" % mp.get(x)) for x in ttl])
                sht = wb.sheets.add(sns[idx])
                sht.range(1,1).value = shpmp
                sht.autofit("c")
        return fn
    
    def _errandwarn(self, errlst):
        return ([x for x in errlst if x["type"].find("ERR") >= 0], [x for x in errlst if x["type"].find("ERR") < 0])

    def buildrpts(self, fldr = None):
        """ create the rpt/err/bc sheets if they're not available
        """
        sts = self._nsofsettings
        getrf = lambda x: triml(path.basename(path.dirname(x)))
        rfs = [getrf(sts[x].value) for x in "shp.template,shpc1.template".split(",")]

        if not fldr:
            fldr = easydialog(filedialog.Open("Choose a file to create shipment"))
        if not path.exists(fldr): return
        if not path.isdir(fldr):
            #if the file's parent folder not in rfs, treate it as pajraw files
            if getrf(fldr) not in rfs:
                fldr = path.dirname(fldr)
        if path.isdir(fldr):
            fn = self._pajfldr2file(fldr)
        else:
            fn = fldr
        if not fn: return            
        pajopts = {"fn":fn,"shpdate":PajShpHdlr._getshpdate(fn),"fmd": datetime.fromtimestamp(path.getmtime(fn))}
        app, kxl = _appmgr.acq()
        wb, flag = app.books.open(fn), False
        for sn in (self._snrpt,self._snerr):
            shts = [x for x in wb.sheets if triml(x.name).find(triml(sn)) >= 0]
            if not shts: continue
            for sht in shts:
                rng = xwu.usedrange(sht)
                if not rng: continue
                flag = [x for x in rng.value if x]
                if flag: break
            if flag: break
        if flag:
            logger.debug("target file(%s) already contains one of sheet%s" % (path.basename(fn),(self._snrpt,self._snerr)))
            return
        invmp, shpmp, errlst, vt, vn = {}, [], [], None, None
        rdrmap = {"长兴珠宝":("c2",self._readc2),"诚艺,胤雅":("c1",self._readc1) ,"十七,物料编号,paj,diamondlite":("paj",self._readpaj)}
        rsht =[x for x in wb.sheets if triml(x.name) == "rpt"]
        if rsht:
            wb.close()
            _appmgr.ret(kxl)
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
            rdr, flag = rdrmap.get(vt), False
            if rdr:
                lst = rdr[1](sht, pajopts)
                if lst and len([x for x in lst if x]):
                    flag = True
                    mp = lst[0]
                    if mp: 
                        if "invdate" in mp:
                            ivd = mp["invdate"]
                            if isinstance(ivd,str):
                                ivd = datetime.strptime(ivd, "%Y-%m-%d")
                            td = datetime.today() - ivd
                            if td.days > 2 or td.days < 0:
                                errlst.append(self._newerr("_日期_", self._wc_date, "日期%s可能错误" % ivd.strftime("%Y-%m-%d")))
                            del mp["invdate"]
                        shpmp.extend(mp.values())
                    if lst[1]: errlst.extend(lst[1])
            if not flag:
                if vn == "paj":
                    invno = PajShpHdlr._readinvno(sht)
                    if not invno: continue
                    mp = PajShpHdlr._rawreadinv(sht, invno)
                    if mp:
                        invmp.update(mp)
                        flag = True
            if not flag:
                logger.debug("sheet(%s) does not contain any valid data" % sht.name)
        if shpmp:
            self._shpcheck(shpmp, invmp, errlst)
            shpmp = sorted(shpmp,key = lambda d0: "%s,%6.1f" % (d0["jono"],d0["qtyleft"]))
            newrunmp = {}
            self._writerpts(wb,shpmp,errlst,newrunmp,ivd,vn)
            self._writebc(wb, shpmp, newrunmp, ivd)
        return fn

    def doimport(self, fn):
        if not fn: return
        app = _appmgr.acq()[0]
        wb = app.books.open(fn)
        sht = wb.sheets[self._snrpt]
               

class ShpImptr(object):
    
    def __init__(self, cnsvc, hksvc, bcsvc):
        self._cnsvc, self._hksvc, self._bcsvc = cnsvc, hksvc, bcsvc
        
    def doimport(self, fn):
        if not fn: return
        with self._cnsvc.sessionctx() as cur:
            pass