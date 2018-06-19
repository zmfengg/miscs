# coding=utf-8
'''
Created on 2018-04-28
classes to read data from C1's monthly invoices
need to be able to read the 2 kinds of files: C1's original and calculator file
@author: zmFeng
'''

from collections import namedtuple
from hnjcore.utils import xwu, daterange
from xlwings import constants
from hnjcore import JOElement
from hnjcore.models.cn import JO,Customer, Customer, Style, MM, MMMa, MMgd
from hnjcore import karatsvc
import os, sys
from .common import _logger as logger, _date_short
import numbers
from sqlalchemy.orm import Query
from sqlalchemy import and_, func


class InvRdr():
    """
        read the monthly invoices from both C1 version and CC version
    """
    
    C1InvItem = namedtuple("C1InvItem", "source,jono,labor,setting,remarks,stones,parts")
    C1InvStone = namedtuple("C1InvStone", "stone,qty,wgt,remark")
    
    def __init__(self, c1log=None, cclog=None):
        self._c1log = c1log
        self._cclog = cclog
    
    def read(self, fldr):
        """
        perform the read action 
        @param fldr: the folder contains the invoice files
        @return: a list of C1InvItem
        """
         
        if not os.path.exists(fldr): return
        root = fldr + os.path.sep if fldr[len(fldr) - 1] != os.path.sep else ""
        fns = [root + x for x in \
            os.listdir(fldr) if x.lower().find("_f") > 0]
        if not fns: return
        killxw, app = xwu.app(False);wb = None
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
                        if rng: rngs.append(rng)
                    if len(cnsc1) == len(rngs):
                        items.append(self._readc1(sht, rngs[0].row))
                    else:
                        for s0 in cnscc:
                            rng = xwu.find(sht, s0, lookat=constants.LookAt.xlWhole)
                            if rng: rngs.append(rng)
                        if len(cnsc1) == len(rngs): items.append(self._readcalc(sht))                           
        finally:            
            if killxw: app.quit()
    
    def _makec1stone(self, tr, tm, joqty):
        """ make a c1 stone based on the row data and mapping """
        if not tr[tm["stqty"]]: return None
        return self.C1InvStone(tr[tm["stname"]], tr[tm["stqty"]] / joqty, tr[tm["stwgt"]] / joqty, "N/A")
    
    def _makec1item(self,em):
        sc = em["setting"] if isinstance(em["setting"], numbers.Number) else 0
        return self.C1InvItem("C1", em["jono"], em["labor"], sc, em["remark"], em["stones"], "N/A")
    
    def _readc1(self, sht, hdrow):
        """
        read c1 invoice file
        @param   sht: the sheet that is verified to be the C1 format
        @param hdrow: the row of the header 
        @return: a list of C1InvItem with source = "C1"
        """        
        rng = sht.range("A%d" % hdrow).end("left")
        rng = sht.range(sht.range(rng.row, rng.column), xwu.usedrange(sht).last_cell)
        vvs = rng.value
        tr = vvs[0]
        km = {u"工单号":"jono", u"镶工":"setting", u"胚底,":"labor", u"备注,":"remark", u"数量":"joqty" \
            , u"石名称":"stname", u"粒数":"stqty", u"石重,":"stwgt"}
        tm = xwu.list2dict(tr, km)
        if len(tm) < len(km):
            logger.debug("key columns(%s) not found in sheet(%s)" % (tm, sht.name))
            return None
        
        items = list()
        em = None
        for ridx in range(1, len(vvs)):
            tr = vvs[ridx]
            s0 = tr[tm["jono"]]
            if isinstance(s0, str):
                s0 = s0.strip()
                if not s0: s0 = None
            jn = JOElement(tr[tm["jono"]]).value
            if jn:
                if em:
                    items.append(self._makec1item(em))
                em = {"jono":jn, "setting":tr[tm["setting"]], "labor":tr[tm["labor"]], \
                    "remark":tr[tm["remark"]], "qty":tr[tm["joqty"]], "stones":list()}
            else:
                if s0 and not jn:
                    em = None
            if em:
                st = self._makec1stone(tr, tm, em["qty"])
                if st: em["stones"].append(st)
        if em: items.append(self._makec1item(em))
        
        if self._c1log and len(items) > 0:        
            import csv, codecs, datetime
            #fc = "utf-8"
            #with codecs.open(self._c1log, "a+b", encoding=fc) as f: #failed
            fc = sys.getfilesystemencoding()
            with open(self._c1log, "a+b") as f:
                ec = codecs.getincrementalencoder(fc)()
                f.write(ec.encode(u"# --- begin to 你国dump C1 log of file (%s) at %s  --- #\n" % \
                    (sht.book.name, datetime.datetime.now())))
                try:
                    wtr = csv.writer(f, delimiter="\t")
                    for it in items:
                        r = [ec.encode(x) if isinstance(x,str) else x for x in \
                             [it.source, it.jono, it.setting, it.labor,it.remarks if it.remarks else "" ]]
                        wtr.writerow(r)
                except Exception as e:
                    print(e)
        return items
            
    def _readcalc(self, sht):
        """
        read cc file
        @param   sht: the sheet that is verified to be the CC format
        @return: a list of C1InvItem with source = "CC"
        """
        #todo::missing
        cns = u"镶石费$,胚底费$,工单,参数,配件,笔电,链尾,分色,电咪,其它,银夹金,石料,形状,尺寸,粒数,重量,镶法,备注".split(",")
        rng = xwu.find(sht, cns[0], lookat=constants.LookAt.xlWhole)
        x = xwu.usedrange(sht)
        rng = sht.range((rng.row, x.columns.count), (x.last_cell().row, x.last_cell().column))
        vvs = rng.value

class C1JCReader(object):
    
    def __init__(self,cnsvc):
        self._cnsvc = cnsvc

    """class to create the C1 JOCost file for HK accountant"""
    def read(self, year, month, day = 1, tplfn=None, tarfldr=None):
        df, dt = daterange(year,month,day)
        with self._cnsvc.sessionctx() as cur:
            mmids, vvs = set(), {}
            cnmap = xwu.list2dict(("mmid,lastmmdate,jobno,cstname,styno,running,mstone,description,joqty"
            ",karat,goldwgt,goldcost,extgoldcost,stonecost,laborcost,extlaborcost,extcost,"
            "totalcost,unitcost,extgoldwgt,cflag").split(","))
            q = Query([JO.name.label("jono"),Customer.name.label("cstname"),\
            Style.name.label("styno"),JO.running,JO.karat.label("jokarat"),MMgd.karat, \
            MM.id,MM.qty,func.sum(MMgd.wgt).label("wgt"),func.max(MMMa.refdate).label("refdate")]).\
            join(Customer).join(MM).join(MMMa).join(MMgd).join(Style).\
            group_by(JO.name,Customer.name,Style.name,JO.running,JO.karat,MMgd.karat,MM.id,MM.qty).\
            filter(and_(MMMa.refdate >= df,MMMa.refdate < dt))
            lst = q.with_session(cur).all()
            for x in lst:
                jn = x.jono.value
                if jn != "580356": continue
                if x.id not in mmids:
                    mmids.add(x.id)
                    if jn not in vvs:
                        vvs[jn] = [x.id,x.refdate.strftime(_date_short),x.jono.value, x.cstname.strip(),\
                            x.styno.value,x.running,"_ST", "_EDESC", x.qty,karatsvc.getfamily(x.jokarat).karat,[],\
                            0,0,0,0,0,0,0,0,0,"NA"]
                    else:
                        vvs[jn][cnmap["joqty"]] += x.qty
                vvs[jn][cnmap["goldwgt"]].append((karatsvc.getfamily(x.karat).karat,x.wgt))
        for x in vvs.values():
            lst1 = x[cnmap["goldwgt"]]
            if len(lst1) > 1:
                mmids = {}
                for y in lst1:
                    if y[0] in mmids:
                        mmids[y[0]] += y[1]
                    else:
                        mmids[y[0]] = y[1]
                x[cnmap["goldwgt"]] = mmids[x[cnmap["karat"]]]
                del mmids[x[cnmap["karat"]]]
                if len(mmids) > 0:
                    x[cnmap["extgoldcost"]] = list(mmids.items())
        print(vvs)
        return vvs