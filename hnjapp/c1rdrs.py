# coding=utf-8
'''
Created on 2018-04-28
classes to read data from C1's monthly invoices
need to be able to read the 2 kinds of files: C1's original and calculator file
@author: zmFeng
'''

import numbers
import os
import re
import sys
import tempfile
from collections import namedtuple, OrderedDict
from os import path
import datetime

from sqlalchemy import and_, func
from sqlalchemy.orm import Query
from xlwings import constants

from hnjcore import JOElement, karatsvc
from hnjcore.models.cn import JO, MM, Customer, MMgd, MMMa, Style,StoneOutMaster,StoneOut,Codetable,StoneIn,StoneBck,StonePk
from hnjcore.utils import appathsep, daterange, getfiles, isnumeric, xwu, splitarray
from hnjcore.utils.consts import NA
from utilz import NamedList, NamedLists, list2dict, trimu
from hnjapp.dbsvcs import jesin

from .common import _date_short
from .common import _logger as logger


class InvRdr():
    """
        read the monthly invoices from both C1 version and CC version
    """

    def __init__(self, c1log=None, cclog=None):
        self._c1log = c1log
        self._cclog = cclog
        self._cnstqnw = "stqty,stwgt".split(",")
        self._cnsnl = "setting,labor".split(",")

    def read(self, fldr):
        """
        perform the read action 
        @param fldr: the folder contains the invoice files
        @return: a list of C1InvItem
        """

        if not os.path.exists(fldr):
            return
        if os.path.isfile(fldr):
            fns = [fldr]
        else:
            root = appathsep(fldr)
            fns = getfiles(root)
        if not fns:
            return
        killxw, app = xwu.app(False)
        wb = None
        try:
            cnsc1 = u"工单号,镶工,胚底,备注".split(",")
            cnscc = u"镶石费$,胚底费$,工单,参数,备注".split(",")
            for fn in fns:
                wb = app.books.open(fn)
                items = list()
                for sht in wb.sheets:
                    rngs = list()
                    for s0 in cnsc1:
                        rng = xwu.find(sht, s0, lookat=constants.LookAt.xlPart)
                        if rng:
                            rngs.append(rng)
                    if len(cnsc1) == len(rngs):
                        items.extend(self._readc1(sht, rngs[0].row))
                    else:
                        for s0 in cnscc:
                            rng = xwu.find(
                                sht, s0, lookat=constants.LookAt.xlWhole)
                            if rng:
                                rngs.append(rng)
                        if len(cnsc1) == len(rngs):
                            items.extend(self._readcalc(sht))
                wb.close()
        finally:
            if killxw:
                app.quit()
        return items

    def _readc1(self, sht, hdrow):
        """
        read c1 invoice file
        @param   sht: the sheet that is verified to be the C1 format
        @param hdrow: the row of the header 
        @return: a list of C1InvItem with source = "C1"
        """
        rng = sht.range("A%d" % hdrow).end("left")
        rng = sht.range(sht.range(rng.row, rng.column),
                        xwu.usedrange(sht).last_cell)
        vvs = rng.value
        C1InvItem = namedtuple(
            "C1InvItem", "source,jono,qty,labor,setting,remarks,stones,parts")
        C1InvStone = namedtuple("C1InvStone", "stone,qty,wgt,remark")

        km = {u"工单号": "jono", u"镶工": "setting", u"胚底,": "labor", u"备注,": "remark",
              u"数量": "joqty", u"石名称": "stname", u"粒数": "stqty", u"石重,": "stwgt"}
        nls = NamedLists(vvs,km,False)
        if len(nls.namemap) < len(km):
            logger.debug("key columns(%s) not found in sheet(%s)" %
                         (nls.namemap, sht.name))
            return None

        items = list()
        for nl in nls: 
            s0 = nl.jono
            if isinstance(s0, numbers.Number):
                s0 = str(int(s0))
            je = JOElement(s0)
            if je.isvalid():
                snl = []
                for x in self._cnsnl:
                    a0 = nl[x]
                    snl.append(float(a0) if isnumeric(a0) else 0)
                if any(snl):
                    c1 = C1InvItem(
                        "C1", je.value, nl.joqty, snl[1], snl[0], nl.remark, [], "N/A")
                    items.append(c1)
            qnw = []
            for x in self._cnstqnw:
                if not isnumeric(nl[x]):
                    break
                qnw.append(float(nl[x]))
            if len(qnw) == 2:
                s0 = nl.stname
                if s0 and isinstance(s0, str):
                    joqty = c1.qty
                    c1.stones.append(C1InvStone(
                        nl.stname, qnw[0] / joqty, qnw[1] / joqty, "N/A"))
        return items

    def _readcalc(self, sht):
        """
        read cc file
        @param   sht: the sheet that is verified to be the CC format
        @return: a list of C1InvItem with source = "CC"
        """
        # todo::missing
        cns = u"镶石费$,胚底费$,工单,参数,配件,笔电,链尾,分色,电咪,其它,银夹金,石料,形状,尺寸,粒数,重量,镶法,备注".split(
            ",")
        rng = xwu.find(sht, cns[0], lookat=constants.LookAt.xlWhole)
        x = xwu.usedrange(sht)
        rng = sht.range((rng.row, x.columns.count),
                        (x.last_cell().row, x.last_cell().column))
        vvs = rng.value


class C1JCReader(object):
    def __init__(self, cnsvc, bcsvc, invfldr):
        """
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
            if x:
                mpss.append((x[1], x[0]))
                refid = x[0]
        return refid

    #return mps of given refid #
    def _getmps(self, refid, mpsmp):
        if refid not in mpsmp:
            mp = self._cnsvc.getjcmps(refid)
            mpsmp[refid] = mp
        if refid in mpsmp:
            return mpsmp[refid]

    def _getstcosts(self,runns):
        """
        return the stone costs by map, running as key and cost as value
        """
        lst, cdmap  = [], None
        sign = lambda x: 0 if x == 0 else 1 if x > 0 else -1
        q0 = Query([JO.running,StoneOutMaster.isout,StonePk.pricen,StonePk.unit,func.sum(StoneOut.qty).label("qty"),func.sum(StoneOut.wgt).label("wgt")]).join(StoneOutMaster).join(StoneOut).join(StoneIn).join(StonePk).group_by(JO.running,StoneOutMaster.isout,StonePk.pricen,StonePk.unit)
        with self._cnsvc.sessionctx() as cur:
            for arr in splitarray(runns,50):
                try:
                    lst1 = q0.filter(JO.running.in_(arr)).with_session(cur).all()
                    if lst1: lst.extend(lst1)
                except:
                    pass
            if lst:
                lst1 = Query([Codetable.coden0,Codetable.tag]).filter(and_(Codetable.tblname == "stone_pkma",Codetable.colname == "unit")).with_session(cur).all()
                cdmap = dict([(int(x.coden0),x.tag) for x in lst1])
        if lst and cdmap:
            costs = dict(zip(runns,[0 for x in range(len(runns))]))
            for x in lst:
                costs[int(x.running)] += round(sign(float(x.isout)) * float(x.pricen) * (float(x.qty) if cdmap[x.unit] == 0 else float(x.wgt)),2)
        return costs
 
    def _getjostone(self,runns):
        ttl = "jobn,styno,running,package_id,quantity,weight,pricen,unit,is_out,bill_id,fill_date,check_date".split(",")
        lst = []
        with self._cnsvc.sessionctx() as cur:
            q = Query([JO.name.label("jono"),JO.deadline,Style.name.label("styno"),JO.running,\
            StonePk.name.label("pkno"),StoneOut.qty,StoneOut.wgt,StonePk.pricen,StonePk.unit,\
            StoneOutMaster.isout,StoneOutMaster.name.label("billid"),StoneOutMaster.filldate,\
            StoneOut.checkdate]).join(Style).join(StoneOutMaster).join(StoneOut).\
            join(StoneIn).join(StonePk)
            for arr in splitarray(runns,50):
                try:
                    lst1 = q.filter(JO.running.in_(arr)).with_session(cur).all()
                    lst.extend(lst1)
                except:
                    pass
        lst1, lst = lst, [ttl]
        lst.extend([("'" + x.jono.value,x.styno.value,x.running,x.pkno,x.qty,round(float(x.wgt),3)\
        ,x.pricen,x.unit,x.isout,x.billid,x.filldate,x.checkdate) for x in lst1])
        return lst

    def _getbroken(self,df,dt):
        lst = None
        with self._cnsvc.sessionctx() as cur:
            q = Query([JO.name.label("jono"),JO.deadline,Style.name.label("styno"),JO.running,\
            StonePk.name.label("pkno"),StoneOut.qty,StoneOut.wgt,StonePk.pricen,StonePk.unit,\
            StoneOutMaster.isout,StoneOutMaster.name.label("billid"),StoneOut.idx,StoneOutMaster.filldate,\
            StoneOut.checkdate]).join(Style).join(StoneOutMaster).join(StoneOut).\
            join(StoneIn).join(StonePk).filter(and_(StoneOutMaster.filldate >= df,StoneOutMaster.filldate< dt)).\
            filter(and_(StoneOutMaster.isout >= -10,StoneOutMaster.isout <= 10))            
            lst = q.with_session(cur).all()
        if not lst:return
        ttl = "jobn,styno,running,package_id,quantity,weight,pricen,unit,is_out,bill_id,idx,fill_date,check_date".split(",")
        lst1, lst = lst, [ttl]
        lst.extend([("'" + x.jono.value,x.styno.value,x.running,x.pkno,x.qty,round(float(x.wgt),3)\
        ,x.pricen,x.unit,x.isout,x.billid,x.idx,x.filldate,x.checkdate) for x in lst1])
        return lst

    def read(self, year, month, day=1, rmbtohk = 1.25, tplfn=None, tarfldr=None):
        """class to create the C1 JOCost file for HK accountant"""
        df, dt = daterange(year, month, day)
        refs, mpsmp, runns = [], {}, set()
        actname = "C1JOCost of (%04d%02d)" % (year,month)
        ptncx = re.compile(r"C(\d)$")
        with self._cnsvc.sessionctx() as cur:
            mmids, vvs, refs = set(), {}, []
            gccols = [
                x.split(",") for x in "goldwgt,goldcost;extgoldwgt,extgoldcost".split(";")]
            ttls = ("mmid,lastmmdate,jobno,cstname,styno,running,mstone,description,joqty"
                    ",karat,goldwgt,goldcost,extgoldcost,stonecost,laborcost,extlaborcost,extcost,"
                    "totalcost,unitcost,extgoldwgt,cflag").split(",")
            cnmap= xwu.list2dict(ttls)
            nl = NamedList(cnmap)            
            q = Query([JO.name.label("jono"), Customer.name.label("cstname"),
                       Style.name.label("styno"), JO.running, JO.karat.label(
                           "jokarat"), MMgd.karat,
                       MM.id, MM.name.label("docno"), MM.qty, func.sum(MMgd.wgt).label("wgt"), func.max(MMMa.refdate).label("refdate")]).\
                join(Customer).join(MM).join(MMMa).join(MMgd).join(Style).\
                group_by(JO.name, Customer.name, Style.name, JO.running, JO.karat, MMgd.karat, MM.id, MM.name, MM.qty).\
                filter(and_(and_(MMMa.refdate >= df, MMMa.refdate < dt),
                            MM.name.like("%C[0-9]")))
            lst = q.with_session(cur).all()
            vvs["_TITLE_"] = ttls
            for x in lst:
                jn = x.jono.value
                # if jn != "580356": continue
                if x.id not in mmids:
                    mmids.add(x.id)
                    if jn not in vvs:
                        ll = [x.id, "'" + x.refdate.strftime(_date_short), "'" + x.jono.value, x.cstname.strip(),
                              x.styno.value, x.running, "_ST", "_EDESC", 0, karatsvc.getfamily(x.jokarat).karat, [],
                              0, 0, 0, 0, 0, 0, 0, 0, 0, "NA"]
                        mt = ptncx.search(x.docno)
                        if mt:
                            ll[cnmap["cflag"]] = "'" + mt.group(1)
                        vvs[jn] = ll
                        runns.add(int(x.running))
                    vvs[jn][cnmap["joqty"]] += float(x.qty)
                vvs[jn][cnmap["goldwgt"]].append((karatsvc.getfamily(x.karat).karat, x.wgt))
            bcs = self._bcsvc.getbcsforjc(runns)
            if not bcs or len(bcs) < len(runns):
                logger.debug("%s:Not all records found in BCSystem" % actname)
            bcs = dict([(x.runn, (x.desc, x.ston)) for x in bcs])            
            stcosts = self._getstcosts(runns)
            if not stcosts or len(stcosts) < len(runns) / 2:
                logger.debug("%s:No stone or less than 1/2 has stone, make sure you've prepared stone data with C1STIOData" % actname)
            invs = InvRdr().read(self._invfldr)
            x = set([x.jono for x in invs if x.jono in vvs]) if invs else set()
            if not invs or len(x) < len(runns):
                logger.debug(
                    "%s:No or not enough C1 invoice data from file(%s)" % (actname,self._invfldr))
            invs = dict([(x.jono, x) for x in invs]) if invs else {}
            cstlst = "goldcost,extgoldcost,stonecost,laborcost".split(",")
            for x in vvs.values():
                # the title
                if x[0] == ttls[0]:
                    continue
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
                    nl.laborcost = round((inv.setting + inv.labor) * rmbtohk,2)
                else:
                    logger.debug("%s:No invoice data for JO(%s)" %
                                (actname,runn))
                lst1 = nl.goldwgt
                if len(lst1) > 1:
                    mmids = {}
                    for knw in lst1:
                        if knw[0] in mmids:
                            mmids[knw[0]] += knw[1]
                        else:
                            mmids[knw[0]] = knw[1]
                    kt = nl.karat
                    nl.goldwgt = [(kt, mmids[kt])]
                    del mmids[kt]
                    if len(mmids) > 0:
                        nl.extgoldwgt = list(mmids.items())
                lst1 = nl.goldwgt
                refid = self._getrefid(nl.running, refs)
                if not refid:
                    logger.critical(("No refid found for running(%d),"
                                     " Pls. create one in codetable with (jocostma/costrefid) ") % nl.running)
                    vvs = None
                    break
                else:
                    mp = self._getmps(refid, mpsmp)
                    for vv in gccols:
                        lst1 = x[cnmap[vv[0]]]
                        if lst1:
                            ttlwgt = 0.0
                            for knw in lst1:
                                kt = knw[0]
                                # in extgold case, if different karat exists, the weight is merged
                                ttlwgt += float(knw[1])
                                if kt not in mp:
                                    logger.critical("No MPS found for running(%d)'s karat(%d)" %
                                                    (nl.running, kt))
                                    x[cnmap[vv[1]]] = -1000
                                else:
                                    x[cnmap[vv[1]]
                                    ] += round(float(mp[kt]) * float(knw[1]), 2)
                            x[cnmap[vv[0]]] = ttlwgt
                if vvs == None:
                    break
                for cx in cstlst:
                    nl["totalcost"] += nl[cx]
                nl.unitcost = round(nl["totalcost"] / nl["joqty"],2)
        if vvs:
            return [x[1:] for x in vvs.values()],self._getjostone(runns),self._getbroken(df,dt)

class C1STIOReader(object):
    """
    class to read C1Stone's IO from files like \\172.16.8.46\pb\dptfile\quotation\2017外发工单工费明细
    \CostForPatrick\StReadLog\C1IO20180619.xlsx
    and save directly to db
    """
    def __init__(self,cnsvc):
        self._cnsvc = cnsvc

    def _rviptusg(self,usgs,ionmap):
        def _ckript(cur,q0,u,ionmap):
            ipt = False
            try:
                if u.type in ionmap:                
                    lst = q0.filter(and_(JO.name == u.jono,StoneOutMaster.isout == ionmap[u.type][0][0])).with_session(cur).all()                    
                else:
                    lst = Query([StoneBck.qty,StoneBck.wgt]).join(StoneIn).filter(StoneIn.name == u.btchno).all()
                for x in lst:
                    ipt = x.qty == u.qty and abs(x.wgt - u.wgt) < 0.001
                    if ipt: break
            except:
                pass
            return ipt
        lb, ub, idx ,ipt = 0,len(usgs) - 1, -1, False
        ptr = (lb+ub) // 2
        q0 = Query([StoneOut.qty,StoneOut.wgt]).join(StoneOutMaster).join(JO)
        with self._cnsvc.sessionctx() as cur:
            while idx < 0:
                if ptr == lb:
                    if not _ckript(cur,q0,usgs[lb],ionmap):
                        idx = lb
                    else:
                        if not _ckript(cur,q0,usgs[ub],ionmap):
                            idx = ub
                        elif ub < len(usgs):
                            idx = ub + 1                            
                    break
                ipt = _ckript(cur,q0,usgs[ptr],ionmap)
                if ipt:
                    lb = ptr + 1
                else:
                    ub = ptr - 1                
                ptr = (lb+ub)//2
        if idx >= 0:
            return usgs[idx:]

    def _readfrmfile(self,fn):
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
        ptnbtno = re.compile(r"(\d+)([A-Z]{3})(\d+)")

        def _fmtbtno(btno):
            if isinstance(btno,numbers.Number):
                btno = "%08d" % int(btno)
            else:
                mt = ptnbtno.search(btno)
                if mt:
                    btno = btno[mt.start(1):mt.end(2)] + ("%03d" % int(mt.group(3)))
            return btno

        def _fmtpkno(pkno):
            if not pkno: return
            #contain invalid character, not a PK#
            pkno = trimu(pkno)
            if sum([1 for x in pkno if ord(x) <= 31 or ord(x) >= 127]) > 0:
                return
            pkno0 = pkno
            if pkno.find("-") >= 0: pkno = pkno.replace("-","")
            pfx, pkno, sfx = pkno[:3], pkno[3:], ""
            for idx in range(len(pkno) - 1,-1,-1):
                ch = pkno[idx]
                if ch >= "A" and ch <= "Z":
                    sfx = ch + sfx
                else:
                    if len(sfx) > 0:
                        idx += 1
                        break
                    sfx = ch + sfx
            pkno = pkno[:idx]
            if isnumeric(pkno):
                pkno = ("%0" + str(8 - len(pfx) - len(sfx)) + "d") % (int(float(pkno)))
                special = False 
            else:
                special =True
                rpm = {"O":"0","*":"X","S":"5"}
                for x in rpm.items():
                    if pkno.find(x[0]) >= 0:
                        logger.debug("PK#(%s)'s %s -> %s in it's numeric part" % (pkno0,x[0],x[1]))
                        pkno = pkno.replace(x[0],x[1])
                        special = True
            pkno = pfx + pkno + sfx
            return pkno,special

        btmap, pkfmted , usgs = {},[], []
        pkmap  =  {}
        try:
            for fn in fns:
                wb = app.books.open(fn)
                shts = {}
                for sht in wb.sheets:
                    shts[sht.name] = sht
                sht = shts[u"进"]
                vvs = sht.range("A1").expand("table").value
                km = {u"序号":"id",u"水号":"btchno",u"包头":"pkno",u"日期,":"date",u"类别":"type",u"成色":"karat",u"数量,":"qty",u"重量,":"wgt",u"数量单位":"qtyunit",u"重量单位":"unit",u"备注":"remark"}
                nls = NamedLists(vvs,km)
                if len(nls.namemap) < len(km):
                    logger.debug("not enough key column provided")
                    break 
                for nl in nls:
                    if nl.karat: continue
                    if not nl.btchno: break
                    pkno = _fmtpkno(nl.pkno)
                    if not pkno: continue
                    flag = pkno[1]; pkno = pkno[0]
                    if pkno != nl.pkno or flag:
                        pkfmted.append((int(nl.id),nl.pkno,pkno, "Special" if flag else "Normal"))
                        nl.pkno = pkno
                    nl.btchno = _fmtbtno(nl.btchno)
                    pkmap[nl.pkno], btmap[nl.btchno]= nl,nl
                sht = shts[u"用"]
                vvs = sht.range("A1").expand("table").value
                km = {u"序号":"id",u"水号":"btchno",u"工单":"jono",u"数量":"qty",u"重量":"wgt",u"记录,":"type",u"备注":"btchid"}
                nls = NamedLists(vvs,km)
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
                    nl.btchno, nl.jono = btchno, JOElement(nl.jono)
                    usgs.append(nl) 
                wb.close()
        finally:
            if kxl: app.quit()
        return pkmap,btmap,usgs,pkfmted

    def _getjoshpdate(self,jes):
        """
        return the max shipment data of given JOElement collection as a dict of
        (JOElement,maxRefdate)
        """
        if not jes: return
        q0 = Query([JO.name,func.max(MMMa.refdate)]).join(MM).join(MMMa)
        d0 = []
        with self._cnsvc.sessionctx() as cur:
            for arr in splitarray(list(jes)):
                try:
                    d0.extend(q0.filter(jesin(arr,JO)).group_by(JO.name).with_session(cur).all())
                except:
                    pass
        return dict([(x[0],x[1]) for x in d0])
            
    def _buildrst(self, pkmap,btmap,usgs,pkfmted):
        with self._cnsvc.sessionctx() as cur:
            lst = cur.query(Codetable.codec0,Codetable.coden0).filter(and_(Codetable.tblname == "stone_out_master",Codetable.colname == "is_out")).all()
            msomids = dict([(x.codec0.strip(),int(x.coden0)) for x in lst])
            msomid = cur.query(func.max(StoneOutMaster.id.label("id"))).first()[0]
            lst =cur.query(StoneOutMaster.isout,func.max(StoneOutMaster.name).label("bid")).filter(StoneOutMaster.isout.in_(list(msomids.values()))).group_by(StoneOutMaster.isout).all()
            lst = dict([(x.isout,x.bid) for x in lst])
            #make it a isoutname -> (isout,maxid) tuple
            msomids = dict([(x[0],[x[1],lst[x[1]]]) for x in msomids.items() if x[1] in lst])
            mbtid =cur.query(func.max(StoneIn.id)).first()[0]
        ionmap = {}
        for x in {"补烂":"补石,*退烂石","补失":"补石,*退失石","配出":"配出"}.items():
            ionmap[x[0]] = [msomids[y] for y in x[1].split(",")]
        usgs = self._rviptusg(usgs,ionmap)
        jonos = set()
        if usgs:
            for nl in usgs:
                if nl.jono.isvalid():
                    jonos.add(nl.jono)
        btnos = self._cnsvc.getstins(btmap.keys())
        pknos = self._cnsvc.getstpks(pkmap.keys())
        jes = jonos
        jonos = self._cnsvc.getjos(jonos)
        tmpf = tempfile.gettempdir() + path.sep
        #print this out and ask for pkdata, or I can not create any further
        fn, crterr = tmpf + "c1readst.log", False
        with open(fn,"w") as fh:
            if pknos[1]:
                print("Below PK# does not exist, Pls. acquire them from HK first",file = fh)
                for x in sorted([(pkmap[x].id,x) for x in pknos[1]]):
                    print("%d,%s" % x,file = fh)
                crterr = True
            if btnos[1]:                
                print("Below BT# does not exists, Pls. get confirm from Kary",file = fh)
                for x in sorted([(btmap[x].id,x,btmap[x].pkno) for x in btnos[1]]):
                    print("%d,%s,%s" % x,file = fh)
            if jonos and jonos[1]:
                print("Below JOs does not exists",file = fh)
                for x in jonos[1]:
                    print(x.name,file  = fh)
                crterr = True                    
            if pkfmted:
                print("---the converted PK#---",file = fh)
                for x in pkfmted:
                    print("%d,%s,%s,%s" % x,file = fh)
            if usgs:
                print("---usage data---",file=fh)
                for y in sorted([(int(x.id),x.type,x.btchno,x.jono.value,x.qty,x.wgt) for x in usgs]):
                    print("%d,%s,%s,%s,%d,%f" % y,file=fh)

        logger.info("log were saved to %s" % fn)            
        if crterr:
            logger.critical("There are Package or JO does not exist, Pls. correct them first")
            return None
        lnm = lambda cl: dict([(x.name,x) for x in cl])

        btbyns = lnm(btnos[0])
        pkbyns = lnm(pknos[0])
        jobyns = lnm(jonos[0]) if jonos and jonos[0] else {}
        #new batch,stoneoutmaster and stoneout,newstoneback, newclosebatch
        nbtmap, sos, nbck,ncbt = {},{},[],set()
        td = datetime.datetime.today()
        for x in btmap.items():
            if x[0] not in btbyns:
                si = StoneIn()
                mbtid, si.filldate = mbtid + 1, x[1].date
                si.docno = "AG" + x[1].date.strftime("%y%m%d")[1:]
                si.id,si.cstref,si.lastupdate,si.lastuserid = mbtid,NA,td,1
                si.name, si.qty, si.qtytrans,si.qtyused,si.cstid = x[0],x[1].qty,0,0,1
                si.size,si.tag,si.wgt,si.wgtadj,si.wgtbck = NA, 0,x[1].wgt,0,0
                si.wgtprep, si.wgttrans, si.wgtused, si.qtybck, si.wgttmptrans = 0,0,0,0,0
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
                    nb.idx, nb.filldate, nb.lastupdate = 1, td,td
                    nb.lastuserid, nb.qty, nb.wgt = 1, x.qty, x.wgt
                    nb.docno = "AG" + td.strftime("%y%m%d")[1:]
                else:
                    for iof in ionmap[s0]:
                        if not x.jono.isvalid():
                            logger.debug("invalid JO# found in usage seqid(%d),batch(%s)" % (int(x.id),x.btchno))
                            continue
                        key = x.jono.value + "," + str(iof[0])
                        somso = sos.setdefault(key,{})                        
                        if len(somso) == 0:
                            som = StoneOutMaster()
                            som.joid = jobyns[x.jono].id 
                            somso["som"] = som
                            msomid += 1
                            iof[1] += 1                            
                            som.id, som.isout, som.name = msomid, iof[0], iof[1]
                            som.packed, som.qty, som.subcnt, som.workerid = 0,0,0,1393
                            som.filldate, som.lastupdate, som.lastuserid = joshd[x.jono], td,1
                        else:
                            som = somso["som"]
                        lst1 = somso.setdefault("sos",[])
                        so = StoneOut()
                        lst1.append(so)
                        so.id, so.idx, so.joqty,so.lastupdate,so.lastuserid  = som.id,len(lst1),0,td,1
                        so.printid,so.qty,so.wgt, so.workerid = 0,x.qty,x.wgt,1393
                        so.checkerid, so.checkdate = 0, som.filldate 
                        if x.btchno in btbyns:
                            so.btchid = btbyns[x.btchno].id
                        else:
                            so.btchid = x.btchno
        return nbtmap,sos,nbck,ncbt,btbyns

    def _persist(self,nbt,sos,nbck,ncbt,btbyns):
        with self._cnsvc.sessionctx() as cur:
            for x in nbt.items():
                x[1].qty = int(x[1].qty) if x[1].qty else 0
                cur.add(x[1])
            cur.flush()
            for x in sos.items():
                cur.add(x[1]["som"])
                for y in x[1]["sos"]:
                    if isinstance(y.btchid,str):
                        y.btchid = nbt[y.btchid].id
                    cur.add(y)
            cur.flush()
            if nbck:
                lst = [x.btchid for x in nbck if not isinstance(x.btchid,str)]
                if lst:
                    try:
                        y = []
                        for k in splitarray(lst,20):
                            lst = Query([StoneBck.btchid,func.max(StoneBck.idx).label("idx")]).filter(StoneBck.btchid.in_(k)).group_by(StoneBck.btchid).with_session(cur).all()
                            y.extend(lst)
                        lst =dict([(x.btchid,x.idx) for x in y])
                    except:
                        pass
                else:
                    lst = {}
                for x in nbck:
                    if isinstance(x.btchid,str):
                        x.btchid = nbt[x.btchid].id
                    idx = lst[x.btchid] if x.btchid in lst else 0
                    #very rare case, check if it's been imported
                    if idx > 0:
                        dup = False
                        try:
                            y = Query([StoneBck.qty,StoneBck.wgt]).filter(StoneBck.btchid == x.btchid).with_session(cur).all()
                            for yy in y:
                                dup = yy.qty == x.qty and abs(yy.wgt - x.wgt) < 0.001
                                if dup: break
                        except:
                            pass
                        if dup:
                            logger.debug("trying to return duplicated item")
                            continue
                    idx += 1
                    lst[x.btchid],x.idx = idx,idx
                    cur.add(x)
            cur.flush()
            ctag = int(datetime.datetime.today().strftime("%m%d"))
            for x in ncbt:
                btno = btbyns[x] if x in btbyns else nbt[x]
                btno.tag = ctag
                cur.add(btno)
            cur.flush()
            cur.commit()
        #return btmap,usgs

    def readst(self, fn):
        """
        read and create the stone usage record from C1, input files only
        """
               #check if one usage item has been inputted. the rule is:
        # if jo+iotype+qty+closeWgt found, treated as dup. Once one item is found
        # not imported , all item behind it was think of not imported
        pkmap,btmap,usgs,pkfmted = self._readfrmfile(fn)
        nbt,sos,nbck,ncbt,btbyns = self._buildrst(pkmap,btmap,usgs,pkfmted)
        return self._persist(nbt,sos,nbck,ncbt,btbyns)
        
