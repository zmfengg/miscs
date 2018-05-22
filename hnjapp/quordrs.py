# coding=utf-8
'''
Created on Apr 23, 2018
module try to read data from quo catalog
@author: zmFeng
'''

from collections import namedtuple
import logging
from os import path
import os
import sys
import re
import pajcc as pc
import csv
import numbers
import math

from hnjcore import xwu
from hnjcore import utils as hnju
from hnjcore.utils import appathsep
from hnjcore import JOElement
from hnjapp.pajcc import MPS


def _checkfile(fn, fns):
    """ check file exists and file expiration
    @param fn: the log file
    @param fns: the files that need to generate log from
    """
    flag = path.exists(fn)
    if flag:
        if fns:
            ld = path.getmtime(fn)
            flag = ld > max([path.getmtime(x) for x in fns])
    return flag


def _getexcels(fldr):
    if fldr:
        return [fldr + unicode(x, sys.getfilesystemencoding()) 
                    for x in os.listdir(fldr) if x.lower().find("xls") >= 0]

def readsignetcn(fldr):
    """ read file format like \\172.16.8.46\pb\dptfile\quotation\date\Date2018\0521\123
    for signet's "RETURN TO VENDOR" sheet
    """
    if not os.path.exists(fldr): return
    fldr = hnju.appathsep(fldr)
    fns = [unicode(x, sys.getfilesystemencoding()) 
        for x in os.listdir(fldr) if x.lower().find("txt") >= 0]
    ptncn = re.compile("CN\d{3,}")
    ptndec = re.compile("\d*\.\d+")
    ptnsty = re.compile("[\w/\.]+\s+")
    ptndscerr = re.compile("\(?[\w/\.]+\)?(\s{4,})")
    styPFX = "YOUR STYLE: ";lnPFX = len(styPFX)
    ttlPFX = "STYLE TOT"; lnTtlPFX = len(ttlPFX)
    lstall = []
    cnt = 0
    for fn in fns:
        cnt = cnt + 1
        with open(fldr + fn,"rb") as fh:
            cn = None;stage = 0
            lstfn = [];dct = None
            for ln in fh:
                if not cn:
                    mt = ptncn.search(ln)
                    if mt: cn = mt.group()
                else:
                    if stage >= 3 or stage <= 0:
                        idx = ln.find(styPFX)                            
                        if idx >= 0:
                            ln = ln[idx + lnPFX:]
                            idx = ln.find(" ")
                            mt = ptnsty.search(ln)
                            mt1 = ptndscerr.search(ln[mt.end():])                                
                            if mt1:
                                idx1 = mt1.end() + mt.end()
                            else:
                                idx1 = mt.end()                                
                            dct = {"cn":cn,"styno": ln[:idx1].strip(),"fn":fn}
                            dct["description"] = ln[idx1:].strip()
                            lstfn.append(dct)
                            stage = 1
                    elif stage == 1:
                        if ptndec.search(ln): stage += 1
                    elif stage == 2:
                        idx = ln.find(ttlPFX)
                        mt = ptndec.search(ln)
                        if idx >= 0 and mt:
                            dct["ttl"] = float(mt.group())
                            dct["qty"] = float(ln[idx + lnTtlPFX + 1:mt.start()].strip())
                            stage += 1
                    else:
                        pass
            if lstfn: lstall.extend(lstfn)                    
    if lstall:
        app = xwu.app(True)[1]
        wb = app.books.add()
        lstfn = []
        for x in lstall:
            if "qty" in x:
                lstfn.append((x["styno"],x["qty"],x["cn"],x["description"],"",x["ttl"],x["fn"]))
            else:
                print("data error:(%s)" % x)
        rng = wb.sheets(1).range("A1")
        rng.value = lstfn


def readagq(fn):
    """
    read AGQ reference prices
    @param fn: the file to read data from  
    """
    
    if not path.exists(fn): return
    
    kxl, app = xwu.app(False)    
    wb = app.books.open(fn)
    try:
        rng = xwu.usedrange(wb.sheets(r'Running lines'))
        cidxs = list()
        vals = rng.value
        Point = namedtuple("Point", "x,y")
        val = "styno,idx,type,mps,uprice,remarks"
        Item = namedtuple("Item", val)
        pt = re.compile("cost\s*:", re.IGNORECASE)        
        items = list()
        items.append(Item._make(val.split(",")))        
        hts = list()
                
        ccnt = 3
        for ridx in range(len(vals)):
            tr = vals[ridx]
            if len(cidxs) < ccnt:
                for cidx in range(len(tr)):
                    val = tr[cidx]
                    if isinstance(val, basestring) and pt.match(val):
                        if len(cidxs) < ccnt and not (cidx in cidxs):
                            cidxs.append(cidx)
            if len(cidxs) < ccnt: continue
            val = tr[cidxs[0]] if isinstance(tr[cidxs[0]], basestring) else None
            if not(val and  pt.match(val)): continue
            for ii in range(0, ccnt):
                hts.append(Point(ridx, cidxs[ii]))
        
        # hardcode, 4 prices, in the 16th columns
        mpss = [vals[x][16] for x in range(4)]
        for pt in hts:
            stynos = list()
            # RG + 5% is special case, treat it as a new item
            rgridx = 0
            # 10 rows up, discard if not found
            for x in range(1, 10):
                ridx = pt.x - x
                if ridx < 0: break
                val = vals[ridx][pt.y]
                if isinstance(val, basestring):
                    if val.lower().find("style#") == 0:
                        for x in val[len("style#"):].split(","):
                            je = JOElement(x.strip())
                            if len(je.alpha) == 1 and je.digit > 0: stynos.append(str(je))
                        break
                    else:
                        if len(val) < 5: continue
                        if val.lower()[:2] == 'rg': rgridx = ridx                            
                        for x in val.split(","):
                            je = JOElement(x.strip())
                            if len(je.alpha) == 1 and je.digit > 0: stynos.append(str(je))
            if not stynos:
                logging.getLogger(__name__).debug("failed to get sty# for pt %s" % (pt,))                
            else:
                # 4 rows down, must have
                rxs = [x + pt.x for x in range(1, 5)]
                if rgridx: rxs.append(rgridx)
                for x in rxs:
                    v0 = vals[pt.x][pt.y + 2]
                    v0 = "" if not v0 else v0.lower()
                    # some items with stone, extend the columns if necessary
                    ccnt = 2 if v0 == "labour" else 3
                    tr = vals[x]
                    for jj in range(1, ccnt):
                        val = tr[pt.y + jj]
                        if not isinstance(val, numbers.Number): continue
                        # remark/type
                        rmk = tr[pt.y]
                        tp = "SS" if rmk.lower() == "silver" else "RG+" if x == rgridx else "S+"                        
                        v0 = vals[pt.x][pt.y + jj]
                        if v0 : rmk += ";" + v0
                        if x == rgridx:
                            mpsidx = 1
                        else:
                            mpsidx = (x - pt.x - 1) % 2
                        mps = "S=%3.2f;G=%3.2f" % (mpss[mpsidx + 2], mpss[mpsidx])                        
                        for s0 in stynos:
                            items.append(Item(s0, mpsidx if x <> rgridx else 2 , \
                                tp, mps, round(val, 2), rmk.upper()))
        wb1 = app.books.add()
        sht = wb1.sheets[0]
        vals = list(items)
        v0 = sht.range((1, 1), (len(items), len(items[0])))
        v0.value = vals
    finally:
        wb.close()
        if not wb1 and kxl:
            app.quit()
        else:
            if wb1: app.visible = True

class DAO(object):
    """a handy Hnjhk dao for data access in this tests
    return a map with keys(id,name,styid,skuno,wgts)
    """
    _ptnmit = re.compile("^M[A-Z]T")
    _silveralphas = set(("4", "5"))
    _querysize = 20  # batch query's batch, don't be too large

    def __init__(self, hkdb, pydb=None, bcdb=None):
        """ dbs won't be closed by me, the caller do it """
        self._hkdb = hkdb
        if pydb: self._pydb = pydb
        if bcdb: self._bcdb = bcdb        
    
    def _samekarat(self, srcje, tarje):
        return srcje.alpha == tarje.alpha or (srcje.alpha in self._silveralphas and tarje.alpha in self._silveralphas)
    
    def getjo(self, je):        
        knws = [None, None, None];jo = None
        cur = self._hkdb.cursor()
        try:            
            # wgt info including mit
            if isinstance(je, basestring): je = JOElement(je)
            cur.execute("select jo.joid,od.styid,c.cstname,sty.alpha,sty.digit,po.skuno,od.karat,jo.wgt,auxkarat"
                ",jo.auxwgt,d.sttype,d.wgt as partwgt from jo inner join orderma od on jo.orderid = od.orderid"
                " inner join po on jo.poid = po.poid inner join cstinfo c on c.cstid = od.cstid inner join styma sty"
                " on od.styid = sty.styid left join cstdtl d on jo.joid = d.jsid where jo.alpha = '%s' and"
                " jo.digit = %d" % (je.alpha, je.digit))
            rows = cur.fetchall()
            if rows:
                for row in rows:
                    if(not knws[0]):
                        knws[0] = pc.WgtInfo(row.karat, float(row.wgt))
                        rk = knws[0]
                        joid = row.joid;skuno = row.skuno;styid = row.styid
                        cstname = row.cstname.strip();styno = JOElement(row.alpha, row.digit)
                        if(skuno):
                            skuno = skuno.strip()
                            if skuno in ("", "N/A"): skuno = None
                            if skuno and [x for x in skuno if ord(x) <= 31 or ord(x) > 127]:
                                skuno = None
                        if(row.auxwgt and row.auxwgt > 0): 
                            knws[1] = pc.WgtInfo(row.auxkarat, float(row.auxwgt))
                            if(knws[1].karat == 925): rk = knws[1]
                    if(not row.sttype): break
                    if(row.partwgt > 0 and self._ptnmit.search(row.sttype)):
                        knws[2] = pc.WgtInfo(rk.karat, float(row.partwgt))
                        break
                jo = {"id":joid, "name":je, "styid":styid, "skuno":skuno, \
                    "wgts":pc.PrdWgt(knws[0], knws[1], knws[2]), "cstname":cstname, "styno":styno} if any(knws) else None
        finally:
            cur.close()
        return jo
    
    def getpaj(self, jo):
        """ return the je,pcode,uprice,mps
        @param jo: a dict contains jo data, the dict can be returned by this.getjo(je)
        @return:  a map with keys(jono,pcode,uprice,mps)  
        """
        
        def _mapx(x, je):
            return dict(zip("jono,pcode,uprice,mps".split(","), (je, x.pcode.strip(), x.uprice, x.mps)))
        
        ups = None
        cur = self._hkdb.cursor()
        try:
            je = jo["name"]
            cur.execute(("select pcode,inv.uprice,inv.mps from jo inner join pajshp shp on jo.joid = shp.joid"
                " and jo.alpha = '%s' and jo.digit = %d left join pajinv inv on shp.joid = inv.joid"
                " and shp.invno = inv.invno order by shp.shpdate desc") % (je.alpha, je.digit))
            rows = cur.fetchall()
            ups = [_mapx(x, je) for x in rows if x.uprice and x.mps] if(rows) else None
        finally:
            cur.close()
        return ups
    
    def getrevcn(self, pcode):
        """return the revised for given pcode"""
        revcn = 0
        cur = self._hkdb.cursor()
        try:
            cur.execute("select top 2 uprice from pajcnrev where pcode = '%s' order by tag" % pcode)
            rows = cur.fetchall()
            if rows: revcn = float(rows[0][0])
        finally:
            cur.close()
        return revcn
    
    def extsearch(self, jo, level=1):
        """find the JOs with the same sty# of given jo, which can be obtained by this.getjo(je). 
        return same karat if found, else return all karats
        @param level:   0 for extract SKU match
                        1 for extract karat match
                        1+ for extract style match
        """ 
        cur = self._hkdb.cursor()  
        rc = None;level = 0 if level < 0 else level
        try:
            s0 = ("select distinct jo.alpha,jo.digit,po.skuno from jo inner join orderma od on" 
                " jo.orderid = od.orderid and od.styid = %d inner join pajinv inv on jo.joid = inv.joid"
                " inner join po on jo.poid = po.poid") % jo["styid"]
            cur.execute(s0)
            rows = cur.fetchall()
            if(rows): 
                jns = dict((JOElement(x.alpha, x.digit), x.skuno) for x in rows)
                sks = [x[0] for x in jns.iteritems() if x[1] == jo["skuno"]]
                if not sks and level > 0: 
                    sks = [x[0] for x in jns.iteritems() if self._samekarat(jo["name"], x[0])]
                    if not sks and level > 1: sks = jns.keys
                rc = sks
        finally:
            cur.close()
        return [self.getjo(x) for x in rc] if rc else None
    
    def getjosbyrunns(self, runns):
        logging.debug("begin to fetch JO#s for running, count = %d" % len(runns))
        cur = self._hkdb.cursor()
        mp = {}        
        try:            
            for x in hnju.splitarray(runns, self._querysize):
                cur.execute("select running,alpha,digit from jo where running in (%s)" % (",".join(x)))
                rows = cur.fetchall()
                if(rows):
                    for pr in [(JOElement(row.alpha, row.digit), str(row.running)) for row in rows]:
                        if(not mp.has_key(pr[1])): mp[pr[1]] = pr[0]
        finally:
            cur.close()
        runns = set(runns)
        df = runns.difference(mp.keys()) if len(runns) > len(mp) else None
        logging.debug("Running -> JO done")
        return mp if len(mp) > 0 else None, df
    
    def getjoandchina(self, je):
        """ get the weight of given JO# and calc the china
            return a map with keys (jo,china,paj)
         """        
        if(isinstance(je, basestring)): je = JOElement(je)
        rmap = {"jo":None, "china":None, "paj":None}
        if(not je.isvalid): return rmap
        
        jo = self.getjo(je)
        if not jo: return rmap
        rmap["jo"] = jo
        ups = self.getpaj(jo) if jo else None
        if not ups: return rmap
        rmap["paj"] = ups[0]        
        revcn = self.getrevcn(ups[0]["pcode"])    
        
        rmap["china"] = pc.newchina(revcn, jo["wgts"]) if revcn else \
            pc.PajCalc.calchina(jo["wgts"], float(ups[0]["uprice"]), MPS(ups[0]["mps"]))        
        return rmap
    
    def getshpforjc(self, df, dt):
        """return py shipment data for PAJCReader
        @param df: start date(include) a date ot datetime object
        @param dt: end date(exclude) a date or datetime object 
        """
        s0 = ("select jo.jsid,ma.refdate,c.cstname,cstbldid_alpha as joalpha"
            ",jo.cstbldid_digit as jodigit,sty.alpha,sty.digit,jo.running,jo.karat"
            ",jo.description,jo.quantity,mm.qty as shpqty,mm.docno"
            " from mm inner join mmma ma on mm.refid = ma.refid inner join b_cust_bill jo"
            " on mm.jsid = jo.jsid inner join cstinfo c on c.cstid = jo.cstid"
            " inner join styma sty on jo.styid = sty.styid"
            " where ma.refdate >= '%s' and ma.refdate < '%s'")
        s0 = s0 % tuple(x.strftime("%Y/%m/%d") for x in (df, dt))
        lst = None
        cur = self._pydb.cursor()
        try:
            cur.execute(s0)
            lst = cur.fetchall()
        finally:
            if cur: cur.close()
        return lst
    
    def getmmioforjc(self, df, dt, runns):
        """return the mmstock's I/O# for PAJCReader"""
        s0 = ("select sm.running,remark1 as jmp,iv.docdate as shpdate,iv.inoutno"
            " from stockobjectma sm inner join invoicedtl dt on dt.srid = sm.srid"
            " inner join invoicema iv on dt.invid = iv.invid and iv.inoutno like 'N%%'"
            " and remark1 <> '' where iv.docdate >= '%s' and iv.docdate < '%s' and sm.running in ('%s')")
        lst = list();dft = [x.strftime("%Y/%m/%d") for x in (df, dt)]
        if not isinstance(runns[0], basestring): runns = [str(x) for x in runns]
        cur = self._hkdb.cursor()
        try:
            for x in hnju.splitarray(runns, self._querysize):
                rns = "','".join(x)
                cur.execute(s0 % (dft[0], dft[1], rns))
                rows = cur.fetchall()
                if rows: lst.extend(rows)
        finally:
            if cur: cur.close()
        return lst
    
    def getbcsforjc(self, runns):
        """return running and description from bc with given runnings """
        if not runns: return
        if not isinstance(runns[0], basestring): runns = [str(x) for x in runns]
        s0 = "select runn,desc from stocks where runn in (%s)";lst = []
        cur = self._bcdb.cursor()
        try:
            for x in hnju.splitarray(runns, self._querysize):
                cur.execute(s0 % ("'" + "','".join(x) + "'"))
                rows = cur.fetchall()
                if rows: lst.extend(rows)
        finally:
            if cur: cur.close()
        return lst
    
    def getpjforjc(self, jes):
        """ get the paj data for jocost
        @param jes: a set of JOElement
        """
        
        if not jes: return
        lst = []
        s0 = ("select jo.alpha,jo.digit,s.invdate, s.invno, s.orderno, s.pcode, s.qty, i.uprice"
            ",s.mtlwgt, s.stwgt, i.stspec,s.shpdate from jo inner join pajshp s"
            " on jo.joid = s.joid inner join pajinv i on"
            " (s.joid = i.joid) and (s.invno = i.invno) where (%s)")
        jns = ["alpha = '%s' and digit = %d" % (x.alpha, x.digit) for x in jes]
        cur = self._hkdb.cursor()        
        try:
            for ii in hnju.splitarray(jns, self._querysize):
                cur.execute(s0 % (") or (".join([str(x) for x in ii])))
                rows = cur.fetchall()
                if rows: lst.extend(rows)
        finally:
            if cur: cur.close()
        return lst


class PajDataByRunn(object):
    """
    class to read such file as \\172.16.8.46\pb\dptfile\quotation\date\
    Date2018\0502\(2) quote to customer\*.xls which has:
      .A field contains \d{6}, which will be treated as a running
      .Sth. like Silver@20.00/oz in the first 10 rows as MPS, if no, use the 
         caller's MPS as default MPS
         
    If inside the folder, there is file named "runOKs.dat", excel files
    won't be processed, this class use it as source running.
    
    return a csv file with below fields:
        Runn,OrgJO#,Sty#,Cstname,JO#,SKU#,PCode,ttcost,mtlcost,mps,rmks,discount
    
    An example to read data from folder "d:\temp"
        PajDataByRunn(hkdb).run(r"d:\temp", defaultmps=pajcc.MPS("S=20;G=1350"))
    """
    _ptnrunn = re.compile("\d{6}")
    _duprunn = False

    def __init__(self, db):
        self.dao = DAO(db)
    
    '''
    def _write(self, runn, srcjono, wnc, tarmps, f, rmk=None):
        flag = wnc["jo"] and wnc["china"]
        if flag:
            cost = pc.PajCalc.calctarget(wnc["china"], tarmps)
            jo = wnc["jo"]
            paj = wnc["paj"]
            print(runn)
            f.write(",".join([str(x) for x in (runn, srcjono, jo["styno"], jo["cstname"], jo["name"], \
                "N/A" if not jo["skuno"] else jo["skuno"], paj["pcode"], cost.china, cost.metalcost, \
                cost.mps.value, rmk) if x]) + "\n")
        elif wnc["jo"]:
            rmk = rmk + ";" if rmk else "" + "Failed" 
            f.write(",".join((runn, rmk, str(wnc["jo"]["wgts"]))) + "\n")
        return flag
    '''

    def _readMPS(self, rows):
        """try to parse out the MPS from the first 5 rows of an excel
        the rows should be obtained by xwu.usedranged(sht).value
        @return: A MPS() object or None
        """
        if not rows: return None
        vvs = [0, 0, "S", "G"]
        for row in [rows[x] for x in range(min(5, len(rows)))]:
            for s0 in row:
                if not isinstance(s0, basestring): continue
                s0 = s0.lower()
                xx = (s0.find("@"), s0.find("/oz"))
                if len([x for x in xx if x >= 0]):
                    pr = s0[xx[0] + 1:xx[1]].strip().replace(",", "")
                    fv = None
                    try:
                        fv = float(pr)
                    finally:
                        pass                                                                                
                    if fv:
                        kt = s0[:xx[0]]
                        idx = 0 if kt.find("si") >= 0 or kt.find("925") >= 0 else 1
                        if not vvs[idx]: vvs[idx] = float(pr)
        if any(vvs[:2]):
            return MPS(";".join(["%s=%s" % (vvs[x + 2], str(vvs[x])) for x in range(2) if vvs[x]]))
    
    def read(self, fldr, mps=None, hisoks=None, hiserrs=None, okfn="runOKs.dat", errfn="runErrs.dat"):
        """ read running/mps from target folder.
            If the folder already contains runnOKs.dat(inside it there should be jono/mps/runn columns
            data will be directly retrieved from it, or try to extract running/mps data from all the
            excel files and fetch the jo#, then finally generate 2 result files: runOKs.dat and runErrs.dat 
        @param mps: if no mps defined in the file, using this
        @param hisoks: the original OK result, should not be returned. A map or set with runn+","+mps as key 
        @param hiserrs: the original error result, should not be returned. A map or set with runn+","+mps as key
        @param okfn: the preferred file name for the OK runnings
        @param errfn: the preferred file name for the error runnings 
        @return: a map with runn+","+mps as key and a map with a dict with (runn/mps) as value
                where the mps is an MPS object, not string  
        """
        
        mp = {};fns = _getexcels(fldr)
        fldr = appathsep(fldr)
        fn = fldr + (okfn if okfn else "runOKs.dat")
        if _checkfile(fn, fns):
            with open(fn, "r") as f:
                rdr = csv.DictReader(f)
                for x in rdr:
                    key = x["runn"] + "," + x["mps"]
                    if not mp.has_key(key):
                        x["mps"] = MPS(x["mps"])
                        mp[key] = x
        else:
            killxls, app = xwu.app(False)
            try:
                rtomps = {}
                for x in fns:
                    wb = app.books.open(x)
                    for sht in wb.sheets:
                        vvs = xwu.usedrange(sht).value
                        if not vvs: continue
                        mps1 = self._readMPS(vvs)
                        if not mps1: mps1 = mps if mps else pc.PAJCHINAMPS
                        for row in vvs:
                            for x in [x for x in row if(x and isinstance(x, basestring) and x.lower().find("runn") >= 0)]:
                                mt = self._ptnrunn.search(x)
                                if mt:
                                    runn = mt.group()
                                    key = runn + "," + mps1.value
                                    if not mp.has_key(key): mp[key] = {"runn":runn, "mps":mps1}
                                    if not runn in rtomps: rtomps[runn] = mps1.value
                    wb.close()
                maps = self.dao.getjosbyrunns([x.split(",")[0] for x in mp.keys()])                
                if maps[1]:
                    logging.debug("some runnings(%s) do not have JO#" % mp.keys())
                    with open(fldr + (errfn if errfn else "runErrs.dat"), "w") as f:
                        wtr = csv.writer(f, dialect="excel")
                        wtr.writerow(["#failed to get JO# for below runnings"])
                        wtr.writerow(["Runn"])
                        for x in maps[1]:
                            wtr.writerow([x])
                if maps[0]:
                    with open(fn, "w") as f:
                        wtr = None
                        for x in maps[0].iteritems():
                            it = mp[x[0] + "," + rtomps[x[0]]];it["jono"] = x[1];
                            if not wtr: 
                                wtr = csv.DictWriter(f, it.keys())
                                wtr.writeheader()
                            wtr.writerow(it)
                if all(maps):
                    for x in maps[1]:
                        key = x + "," + rtomps[x]
                        if key in mp: del mp[key]
            except Exception as e:
                logging.debug(e)
                raise e
            finally:
                if(killxls): app.quit()
        rmvs = (hisoks, hiserrs)
        if mp and any(rmvs):
            ks = set()
            [ks.update(x.keys()) for x in rmvs if x]
            for k in ks:
                if k in mp: del mp[k]
        return mp
    
    def run(self, fldr, defaultmps=None, okfn="OKs.dat", errfn="Errs.dat",
            runokfn="runOKs.dat", runerrfn="runErrs.dat"):
        """do folder \\172.16.8.46\pb\dptfile\quotation\date\Date2018\0502\(2) quote to customer\ PAJ cost calculation
            find the JO of given running, then do the calc based on S=20;G=1350
            try to read running/mps from fldr and generate result files with PAJ related costs
            if the folder contains @runokfn, runnings will be from it, else from the excel files
            the files should contains MPS there, or the default mps will be applied
            @param fldr: the folder to generate data
            @param defaultmps: when no MPS provided in the source file(s), use this. should be an MPS() object
            @param okfn: file name of the ok result
            @param errfn: file name of the error result
            @param runokfn: filename of the ok runnings
            @param runerrfn: filename of the error runnings     
        """
        
        def _putmap(wnc, runn, orgjn, tarmps, themap):
            key = "%s,%6.1f" % (wnc["paj"]["pcode"], wnc["china"].china)
            if not themap.has_key(key):
                wnc["china"] = pc.PajCalc.calctarget(wnc["china"], tarmps)
                themap[key] = {"runn":runn, "jono":orgjn, "wnc":wnc}
                return True
        
        def _writeOks(wtroks, foks, fn , ttroks, oks, hisoks):
            if not wtroks:
                if not foks: foks = open(fn, "a+" if hisoks else "w")
                wtroks = csv.DictWriter(foks, ttroks)
                if not hisoks: wtroks.writeheader()
            
            for x in sorted(oks.values()):
                wnc = x["wnc"];jo = wnc["jo"];paj = wnc["paj"];cost = wnc["china"]
                jn0 = x["jono"];jn1 = jo["name"]
                rmk = "Actual" if jn0 == jn1 else "Candiate"
                rmk = dict(zip(ttroks, (x["runn"], jn0, jo["styno"], jo["cstname"], jn1, \
                    "N/A" if not jo["skuno"] else jo["skuno"], paj["pcode"], cost.china, \
                    cost.metalcost if cost.metalcost else 0, cost.mps.value, rmk, cost.discount * 1.25)))
                wtroks.writerow(rmk)
            foks.flush()                                            
            return wtroks, foks
        
        def _writeErrs(wtrerrs, ferrs, fnerrs, ttrerrs, errs, hiserrs):
            if not wtrerrs:
                if not ferrs: ferrs = open(fnerrs, "a+" if hiserrs else "w")
                wtrerrs = csv.DictWriter(ferrs, ttrerrs)
                if not hiserrs: wtrerrs.writeheader()               
            for x in sorted(errs):
                ar = [x["runn"], x["jono"]]
                for y in [(str(y.karat) + "=" + str(y.wgt) if y else "0") for y in x["wnc"]["jo"]["wgts"]]:
                    ar.append(y)
                ar.append(x["mps"])
                wtrerrs.writerow(dict(zip(ttrerrs, ar)))
            ferrs.flush()
            return wtrerrs, ferrs
            
        fldr = appathsep(fldr)
        ttroks = "Runn,OrgJO#,Sty#,Cstname,JO#,SKU#,PCode,ttcost,mtlcost,mps,rmks,discount".split(",")
        ttrerrs = "Runn,OrgJO#,mainWgt,auxWgt,partWgt,mps".split(",")
        errs = [];hiserrs = None;wtrerrs = None;ferrs = None
        hisoks = None;wtroks = None;foks = None
        commitcnt = 10        
        xlsx = _getexcels(fldr);
        
        fnoks = fldr + (okfn if okfn else "OKs.dat")
        fnerrs = fldr + (errfn if errfn else "Errs.dat")
        if _checkfile(fnoks, xlsx):
            with open(fnoks) as f:
                rdr = csv.DictReader(f)
                hisoks = {}
                for x in rdr:
                    hisoks[x["Runn"] + "," + x["mps"]] = x
        if _checkfile(fnerrs, xlsx):
            with open(fnerrs) as f:
                rdr = csv.DictReader(f)
                hiserrs = {}
                for x in rdr:
                    hiserrs[x["Runn"] + "," + x["mps"]] = x
        
        mp = self.read(fldr, defaultmps, hisoks, hiserrs, runokfn, runerrfn)
        if not mp: return             
        oks = {};dao = self.dao;stp = 0;cnt = len(mp)
        try:
            for x in mp.values():
                # if x["runn"] != "625254": continue
                # logging.debug("doing running(%s)" % x["runn"])
                if not "jono" in x:
                    logging.critical("No JO field in map of running(%s)" % x["runn"])
                    continue
                found = False
                wnc = dao.getjoandchina(x["jono"])
                if x["mps"].isvalid:
                    if wnc and all(wnc.values()):
                        found = True 
                        if not _putmap(wnc, x["runn"], x["jono"], x["mps"], oks):
                            logging.debug("JO(%s) is duplicated for same pcode/cost" % wnc["jo"]["name"])
                    else:                    
                        jo = wnc["jo"] if wnc["jo"] else dao.getjo(x["jono"])
                        jos = dao.extsearch(jo, 1)
                        if jos:
                            for jo in jos:
                                wnc1 = dao.getjoandchina(jo["name"])
                                if(all(wnc1.values())):
                                    found = True
                                    if not _putmap(wnc1, x["runn"], x["jono"], x["mps"], oks):
                                        logging.debug("JO(%s) is duplicated for same pcode/cost" % str(wnc1["jo"]["name"]))
                else:
                    found = False
                if not found:
                    errs.append({"runn":x["runn"], "jono":x["jono"], "wnc":wnc, "mps":x["mps"]})
                    if len(errs) > commitcnt:
                        wtrerrs, ferrs = _writeErrs(wtrerrs, ferrs, fnerrs, ttrerrs, errs, hiserrs)
                        errs = []
                if len(oks) > commitcnt:
                    wtroks, foks = _writeOks(wtroks, foks, fnoks, ttroks, oks, hisoks)
                    oks = {}         
                stp += 1
                if not (stp % 20): logging.debug("%d of %d done" % (stp, cnt))       
            if len(oks) > 0:
                wtroks, foks = _writeOks(wtroks, foks, fnoks, ttroks, oks, hisoks)
            if errs:
                wtrerrs, ferrs = _writeErrs(wtrerrs, ferrs, fnerrs, ttrerrs, errs, hiserrs)
        finally:
            if foks: foks.close()
            if ferrs: ferrs.close()
        
        return fnoks, fnerrs
    
    @classmethod
    def readquoprice(self,fldr, rstfn="costs.dat"):
        """read simple quo file which contains Running:xxx, Cost XX: excel
        @param fldr: the folder to read files from
        @return: the result file name or None if nothing is returned  
        """
        if not fldr: return
        fldr = appathsep(fldr)
        kxl, app = xwu.app(False)
        ptnRunn = re.compile("running\s?:\s?(\d*)", re.IGNORECASE)
        ptnCost = re.compile("cost\s?(\w*)\s?:", re.IGNORECASE)
        lst = []
        try:
            for fn in _getexcels(fldr):
                wb = app.books.open(fn)
                for sh in wb.sheets:
                    phase = 0;runns = {};costs = {};rowrunn = 0;lastii = 0
                    vvs = xwu.usedrange(sh).value
                    for hh in range(len(vvs)):
                        tr = vvs[hh]
                        for ii in range(len(tr)):
                            if not tr[ii]: continue
                            x = str(tr[ii])
                            if phase <= 1:
                                mt = ptnRunn.search(x)
                                if mt:
                                    if phase <> 1: 
                                        phase = 1
                                        rowrunn = hh
                                    runns[ii] = mt.group(1)
                                    lastii = ii
                                    continue
                            if phase >= 1:
                                mt = ptnCost.search(x)
                                if mt:
                                    cost = 0
                                    for jj in range(ii + 1, len(tr)):
                                        if isinstance(tr[jj], numbers.Number):
                                            cost = tr[jj]
                                            break
                                    if phase <> 2 and hh != rowrunn: phase = 2
                                    kk = ii if hh <> rowrunn else lastii                                       
                                    if kk in runns:
                                        costs[kk] = (mt.group(1) , cost, fn)
                                    else:
                                        print("error, no running found for cost %s" % cost)
                                        print("file(%s)" % fn)
                                        print("row(%d), data = %s " % (hh + 1, tr))                                        
                        if phase == 2:
                            for ii in runns.keys():
                                if ii in costs:
                                    cost = costs[ii]
                                    lst.append((runns[ii], cost[0], cost[1], cost[2]))
                            phase = 0
                            runns = {};costs = {}
                wb.close()
        except Exception as e:
            print(e)
            print(fn)
            print(tr)
        finally:
            if kxl: app.quit()
        fn = fldr + rstfn if rstfn else "costs.dat"
        if lst:
            with open(fn, "w") as f:
                wtr = csv.writer(f, dialect="excel")
                wtr.writerow("runn,karat,cost,file".split(","))
                for x in lst:
                    wtr.writerow(x)
        return fn
    
if __name__ == "__main__":
    for x in (r'd:\temp\1200&15.xls', r'd:\temp\1300&20.xls'):
        readagq(x)
        pass
