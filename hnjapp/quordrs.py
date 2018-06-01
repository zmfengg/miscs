# coding=utf-8
'''
Created on Apr 23, 2018
module try to read data from quo catalog
@author: zmFeng
'''

import csv
import logging
import math
import numbers
import os
import re
import sys
from collections import namedtuple
from os import path

import pajcc as pc
from hnjapp.pajcc import MPS,PrdWgt,WgtInfo
from hnjcore import JOElement, xwu, appathsep, utils as hnju


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
    ptncn = re.compile(r"CN\d{3,}")
    ptndec = re.compile(r"\d*\.\d+")
    ptnsty = re.compile(r"[\w/\.]+\s+")
    ptndscerr = re.compile(r"\(?[\w/\.]+\)?(\s{4,})")
    styPFX = "YOUR STYLE: ";lnPFX = len(styPFX)
    ttlPFX = "STYLE TOT"; lnTtlPFX = len(ttlPFX)
    lstall = []
    cnt = 0
    for fn in fns:
        cnt = cnt + 1
        with open(fldr + fn,"rb") as fh:
            cn = None;stage = 0
            lstfn = [];dct = {}
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
        pt = re.compile(r"cost\s*:", re.IGNORECASE)        
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
    except:
        pass
    finally:
        wb.close()
        if not wb1 and kxl:
            app.quit()
        else:
            if wb1: app.visible = True

class DAO(object):
    """a handy Hnjhk dao for data access in this tests
    now it's only bc services
    """
    _querysize = 20  # batch query's batch, don't be too large

    def __init__(self, bcdb=None):
        if bcdb: self._bcdb = bcdb
    
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
        except:
            pass
        finally:
            if cur: cur.close()
        return lst
    

class PajDataByRunn(object):
    r"""
    class to read such file as \\172.16.8.46\pb\dptfile\quotation\date
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
    _ptnrunn = re.compile(r"\d{6}")
    _duprunn = False

    def __init__(self, hksvc):
        self._hksvc = hksvc
    
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
                    except:
                        pass
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
        if not mp:
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
                                    if key not in mp: mp[key] = {"runn":runn, "mps":mps1}
                                    if runn not in rtomps: rtomps[runn] = mps1.value
                    wb.close()
                with self._hksvc.sessionctx() as sess:                
                    maps = self._hksvc.getjos(["r" + x.split(",")[0] for x in mp.keys()],psess = sess)
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
                            for x in dict([(str(x.running),x) for x in maps[0]]).iteritems():
                                runnstr = x[0]
                                it = mp[runnstr + "," + rtomps[runnstr]];it["jono"] = x[1].name.value
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
            key = "%s,%6.1f" % (wnc["PajShp"].pcode, wnc["china"].china)
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
                wnc = x["wnc"];jo = wnc["JO"];cost = wnc["china"]
                jn0 = x["jono"];jn1 = jo.name.value
                rmk = "Actual" if jn0 == jn1 else "Candiate"
                skuno = jo.po.skuno
                rmk = dict(zip(ttroks, (x["runn"], jn0, jo.style.name.value, jo.customer.name.strip()\
                    ,jn1,"N/A" if skuno else skuno, wnc["PajShp"].pcode, cost.china, \
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
                prd = x["wnc"]["wgts"]
                if not prd: prd = PrdWgt(WgtInfo(0,0))
                for y in [(str(y.karat) + "=" + str(y.wgt) if y else "0") for y in prd]:
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
        xlsx = _getexcels(fldr)
        
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
        oks = {};dao = self._hksvc;stp = 0;cnt = len(mp)
        try:
            with dao.sessionctx() as sess:
                for x in mp.values():
                    # if x["runn"] != "625254": continue
                    # logging.debug("doing running(%s)" % x["runn"])
                    if "jono" not in x:
                        logging.critical("No JO field in map of running(%s)" % x["runn"])
                        continue
                    found = False
                    if x["jono"] != "580191":
                        pass
                    print("doing " + x["jono"])
                    wnc = dao.calchina(x["jono"],psess = sess)
                    if x["mps"].isvalid:
                        if wnc and all(wnc.values()):
                            found = True 
                            if not _putmap(wnc, x["runn"], x["jono"], x["mps"], oks):
                                logging.debug("JO(%s) is duplicated for same pcode/cost" % wnc["JO"].name.value)
                        else:                    
                            jo = wnc["JO"]
                            if not jo:
                                jo = dao.getjos([x["jono"]],psess = sess)
                                jo = jo[0][0] if jo and jo[0] else None
                            if jo:                                
                                jos = dao.findsimilarjo(jo, 1,psess = sess)
                                if jos:
                                    for x1 in jos:
                                        wnc1 = dao.calchina(x1.name,psess = sess)
                                        if(all(wnc1.values())):
                                            found = True
                                            if not _putmap(wnc1, x["runn"], x["jono"], x["mps"], oks):
                                                logging.debug("JO(%s) is duplicated for same pcode/cost" % str(wnc1["JO"].name.value))
                    else:
                        found = False
                        jo = None
                    if not found:
                        if jo and not wnc["wgts"]: wnc["wgts"] = dao.getjowgts(jo, psess = sess)
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
        except:
            pass
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
