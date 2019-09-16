# coding=utf-8
'''
Created on Apr 23, 2018
module try to read data from quo catalog
@author: zmFeng
'''

import csv
from datetime import datetime
import math
import numbers
import os
import re
import sys
import time
from collections import namedtuple
from os import path
from .common import config

from sqlalchemy.orm import Query
from xlwings import App
from xlwings.constants import (
    FormatConditionOperator,
    FormatConditionType,
)

from hnjapp.c1rdrs import C1InvRdr
from hnjapp.svcs.db import jesin
from hnjapp.pajcc import MPS, PajCalc, PajChina, PrdWgt, WgtInfo
from hnjcore import JOElement
from hnjcore import utils as hnju
from hnjcore.models.hk import JO, PO, Customer, Orderma, PajAck, POItem, Style
from hnjcore.utils.consts import NA
from utilz import NamedList, NamedLists, appathsep, getfiles, splitarray, xwu
from utilz.miscs import list2dict, tofloat

from . import pajcc
from .common import _logger as logger


def _checkfile(fn, fns):
    """ check file exists and fn's date is newer than fns
    @param fn: the log file
    @param fns: the files that need to generate log from
    return True if file expired
    """
    flag = path.exists(fn)
    if flag:
        if fns:
            ld = path.getmtime(fn)
            flag = ld > max([path.getmtime(x) for x in fns])
    return flag


def _getexcels(fldr):
    """ return the excel(xls*) files from fldr, but exclude the shared lock files
    """
    if not fldr:
        return None
    fns = getfiles(fldr, "xls", True)
    p = appathsep(fldr)
    if fns:
        fns = [p + x for x in fns if not x.startswith("~")]
    return fns


def readsignetcn(fldr):
    """
    read file format like \\172.16.8.46\pb\dptfile\quotation\date\Date2018\0521\123
    for signet's "RETURN TO VENDOR" sheet
    """
    if not os.path.exists(fldr): return
    fns = getfiles(fldr, "txt")
    ptncn = re.compile(r"(CN|QC)\d{3,}")
    ptndec = re.compile(r"\d*\.\d+")
    ptnsty = re.compile(r"[\w/\.]+\s+")
    ptndscerr = re.compile(r"\(?[\w/\.]+\)?(\s{4,})")
    styPFX = "YOUR STYLE: "
    lnPFX = len(styPFX)
    ttlPFX = "STYLE TOT"
    lnTtlPFX = len(ttlPFX)
    lstall = []
    cnt = 0
    for fn in fns:
        cnt = cnt + 1
        with open(fn, "rb") as fh:
            cn = None
            stage = 0
            lstfn = []
            dct = {}
            for ln in fh:
                ln = str(ln)
                if not cn:
                    mt = ptncn.search(ln)
                    if mt: cn = ln[mt.start():mt.end()]
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
                            dct = {
                                "cn": cn,
                                "styno": ln[:idx1].strip(),
                                "fn": os.path.basename(fn)
                            }
                            dct["description"] = ln[idx1:].strip().replace(
                                "\\r\\n'", "")
                            lstfn.append(dct)
                            stage = 1
                    elif stage == 1:
                        if ptndec.search(ln): stage += 1
                    elif stage == 2:
                        idx = ln.find(ttlPFX)
                        mt = ptndec.search(ln)
                        if idx >= 0 and mt:
                            dct["ttl"] = float(mt.group())
                            dct["qty"] = float(
                                ln[idx + lnTtlPFX + 1:mt.start()].strip())
                            stage += 1
                    else:
                        pass
            if lstfn: lstall.extend(lstfn)
    fn = None
    if lstall:
        app = xwu.app(True)[1]
        wb = app.books.add()
        lstfn = []
        for x in lstall:
            if "qty" in x:
                lstfn.append((x["styno"], x["qty"], x["cn"], x["description"],
                              "", x["ttl"], x["fn"]))
            else:
                print("data error:(%s)" % x)
        sht = wb.sheets(1)
        rng = sht.range("A1")
        rng.value = lstfn
        wb.save(path.join(fldr, "data"))
        sht.autofit("c")
        fn = wb.fullname
    return fn


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
                    if isinstance(val, str) and pt.match(val):
                        if len(cidxs) < ccnt and not (cidx in cidxs):
                            cidxs.append(cidx)
            if len(cidxs) < ccnt: continue
            val = tr[cidxs[0]] if isinstance(tr[cidxs[0]], str) else None
            if not (val and pt.match(val)): continue
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
                if isinstance(val, str):
                    if val.lower().find("style#") == 0:
                        for x in val[len("style#"):].split(","):
                            je = JOElement(x.strip())
                            if len(je.alpha) == 1 and je.digit > 0:
                                stynos.append(str(je))
                        break
                    else:
                        if len(val) < 5: continue
                        if val.lower()[:2] == 'rg': rgridx = ridx
                        for x in val.split(","):
                            je = JOElement(x.strip())
                            if len(je.alpha) == 1 and je.digit > 0:
                                stynos.append(str(je))
            if not stynos:
                logger.debug("failed to get sty# for pt %s" % pt)
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
                        tp = "SS" if rmk.lower(
                        ) == "silver" else "RG+" if x == rgridx else "S+"
                        v0 = vals[pt.x][pt.y + jj]
                        if v0: rmk += ";" + v0
                        if x == rgridx:
                            mpsidx = 1
                        else:
                            mpsidx = (x - pt.x - 1) % 2
                        mps = "S=%3.2f;G=%3.2f" % (mpss[mpsidx + 2],
                                                   mpss[mpsidx])
                        for s0 in stynos:
                            items.append(Item(s0, mpsidx if x != rgridx else 2 , \
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


class QuoDatMkr(object):
    r"""
    provided a folder(\\172.16.8.46\pb\dptfile\quotation\date
    Date2018\0502\(2) quote to customer), read the quoted prices out into costs.dat.

    Source files in that folder should contains:
      .A field contains \d{6}, which will be treated as a running
      .Sth. like Silver@20.00/oz in the first 10 rows as MPS, if no, use the
         caller's MPS as default MPS

    costs.dat is a csv file contains "runn,mps,jono". Not sure all the runnings in costs.dat is valid, so a runOKs.dat csv file will be created, which holds only the valid runnings, runOKs.dat contains "jono,mps,runn" fields. The invalid runnings in costs.dat will be saved in runErrs.dat

    After that, OKs.dat csv file will be created, which holds the calculated cost info. OKs.dat contains fields "Runn,OrgJO#,Sty#,Cstname,JO#,SKU#,PCode,ttcost,mtlcost,mps,rmks,discount". Those failed to calculated will be saved to err.dat csv file, which contains "Runn,OrgJO#,mainWgt,auxWgt,partWgt,mps" fields

    one helper method createpodat() was appended to calculate a given runn's cost under given MPS using it's PO price

    Data will be fetch from both PAJ and C1's invoice history. When the extractly one is not found, getfamiliar() function will be called to find a suitable one. In the case of C1 data handling, usbtormb/usdtohkd/pkratio is needed, provide you own if the default value is not suitable

    An example to read data from folder "d:\temp"
        QuoDatMkr(hksvc,cnsvc).run(r"d:\temp", defaultmps=pajcc.MPS("S=20;G=1350"))
    """
    _ptnrunn = re.compile(r"\d{6}")
    _duprunn = False

    def __init__(self, hksvc, cnsvc=None, usdtormb=6.5, usdtohkd=7.8, pkratio=0.8,
        fnoks="OKs.dat", fnerrs="Errs.dat",fnrunoks="runOKs.dat", fnrunerrs="runErrs.dat"):
        self._hksvc = hksvc
        self._cnsvc = cnsvc
        self._usdtormb = usdtormb
        self._usdtohkd = usdtohkd
        self._pkratio = pkratio
        self._fnoks = path.basename(fnoks)
        self._fnerrs = path.basename(fnerrs)
        self._fnrunoks = path.basename(fnrunoks)
        self._fnrunerrs = path.basename(fnrunerrs)

    def _readMPS(self, rows):
        """try to parse out the MPS from the first 5 rows of an excel
        the rows should be obtained by xwu.usedranged(sht).value
        @return: A MPS() object or None
        """
        if not rows: return None
        vvs = [0, 0, "S", "G"]
        for row in [rows[x] for x in range(min(5, len(rows)))]:
            for s0 in row:
                if not isinstance(s0, str): continue
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
                        idx = 0 if kt.find("si") >= 0 or kt.find(
                            "925") >= 0 else 1
                        if not vvs[idx]: vvs[idx] = float(pr)
        if any(vvs[:2]):
            return MPS(";".join([
                "%s=%s" % (vvs[x + 2], str(vvs[x])) for x in range(2) if vvs[x]
            ]))

    def createpodat(self, fldr, fn="po.dat", stpfn=None, pfrmap=None):
        """ from the runOK.dat, read the JO/MPS, get the weights/poprice/pomps, translate them into the newpoprice/newmps
        @param stpfn: the file contains sty# that is marked stamping
        @param pfrmap: each customer's profit ratio, if not provided, each use default's 1.15
        """
        if not stpfn:
            stpfn = r"\\172.16.8.46\pb\dptfile\quotation\date\Date2018\0502\(06) PriceDrop\StpStynos.dat"
        with open(stpfn, "r") as fh:
            stpstys = set([x.replace("\n", "") for x in fh])
        mp = self.readreqs(fldr, checkdone=False)[0]
        if not mp: return
        fnpo = path.join(fldr, fn)
        exists = path.exists(fnpo)
        savecnt, allrc = 10, []
        if exists:
            with open(fnpo) as fh:
                lst = [x.split(",") for x in fh]
            ttl, oks, rc = lst[0], lst[1:], []
            allrc.extend(lst)
        else:
            ttl = tuple(
                "pono,jono,runn,pomps,poup,newmps,newup,lossrate,mainwgt,auxwgt,partswgt"
                .split(","))
            rc, allrc = [ttl], [ttl]
        if pfrmap is None:
            pfrmap = {}
        mp = dict([(x["runn"], x) for x in mp.values()])
        if exists:
            ompcnt, rvcnt = len(mp), 0
            for x in oks:
                if x[2] in mp:
                    del mp[x[2]]
                    rvcnt += 1
            logger.debug("%d original mps returned, %d existing item removed" %
                         (ompcnt, rvcnt))
        if not mp:
            logger.debug("No item need to be processed")
            return allrc
        ops, lst, mp = {}, [], list(mp.values())
        q = Query([
            JO.name.label("jono"), POItem.uprice, PO.mps,
            PO.name.label("pono"),
            Style.name.label("styno"),
            Customer.name.label("cstname")
        ]).join(POItem).join(PO).join(
            Orderma, JO.orderid == Orderma.id).join(Style).join(
                Customer, Orderma.cstid == Customer.id)
        fmt = ("%s," * len(ttl))[:-1]
        nl = NamedList("uprice,mps,pono,styno,cstname")
        with self._hksvc.sessionctx() as cur:
            for arr in splitarray(mp, savecnt):
                jns = [JOElement(x["jono"]) for x in arr]
                lst = q.filter(jesin(jns, JO)).with_session(cur).all()
                ops = dict([(x.jono.value, (float(x.uprice), MPS(x.mps),
                                            x.pono.strip(), x.styno.value,
                                            x.cstname.strip())) for x in lst])
                for x in arr:
                    jn = x["jono"]
                    op = ops.get(jn)
                    if not op:
                        logger.debug("no po returned for JO#(%s)" % jn)
                        continue
                    nl.setdata(op)
                    wgts = self._hksvc.getjowgts(jn)
                    cn = nl.cstname
                    pfr = pfrmap.get(cn, 1.15)
                    if pfr < 1:
                        pfr += 1.0
                    elif pfr > 3:
                        pfr = 1.0 + pfr * 1.0 / 100.0
                    mps0, np = nl.mps, nl.uprice
                    hasbg = sum([1 for x in wgts.wgts if x and x.karat == 9925])
                    #Natalie confirmed on 2018/07/20:bonded gold no metal cost, the target up the final PO#'s up
                    lr = 1.05 if nl.styno in stpstys else 1.10
                    if hasbg:
                        ww = sum([
                            1 if x and x.karat != 9925 else 0 for x in wgts.wgts
                        ])
                        #except bg, has other metal
                        if ww > 0:
                            wgts = PrdWgt(*[
                                x if x and x.karat != 9925 else None
                                for x in wgts.wgts
                            ])
                            logger.debug(
                                "JO(%s) is bonded gold + other, use the last po price (%s,%6.2f) and gold incr"
                                % (jn, op[2], np))
                        else:
                            logger.debug(
                                "JO(%s) is bonded gold only, use the last po price (%s,%6.2f)"
                                % (jn, nl.pono, np))
                            wgts = None
                    if wgts:
                        #get the loss rate
                        mc0 = PajCalc.calcmtlcost(
                            wgts,
                            mps0,
                            lossrate=lr,
                            vendor="HNJ",
                            oz2gm=31.1031)
                        mc1 = PajCalc.calcmtlcost(
                            wgts,
                            x["mps"],
                            lossrate=lr,
                            vendor="HNJ",
                            oz2gm=31.1031)
                        if mc0 > 0:
                            np = round((mc1 - mc0) * pfr + nl.uprice, 2)
                        else:
                            np = pajcc.MPSINVALID
                    y = [
                        nl.pono, jn, x["runn"], nl.mps.value, nl.uprice,
                        x["mps"].value, np, lr
                    ]
                    if wgts:
                        for kx in wgts:
                            y.append("%s=%s" % (kx.karat,
                                                kx.wgt) if kx else "0")
                    else:
                        y.extend("0,0,0".split(","))

                    rc.append(y)
                allrc.extend(rc)
                with open(fnpo, "at") as fh:
                    for yy in rc:
                        print(fmt % tuple(yy), file=fh)
                rc = []
        return allrc

    def readreqs(self, fldr, mps=None, checkdone=True):
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
        fns, mp = _getexcels(fldr), {}

        fldr = appathsep(fldr)
        fn = path.join(fldr, self._fnrunoks)
        if _checkfile(fn, fns):
            with open(fn, "r") as f:
                rdr = csv.DictReader(f)
                for x in rdr:
                    key = x["runn"] + "," + x["mps"]
                    if key not in mp:
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
                        if not mps1: mps1 = mps if mps else pajcc.PAJCHINAMPS
                        for row in vvs:
                            for x in [
                                    x for x in row
                                    if (x and isinstance(x, str) and
                                        x.lower().find("runn") >= 0)
                            ]:
                                mt = self._ptnrunn.search(x)
                                if mt:
                                    runn = mt.group()
                                    key = runn + "," + mps1.value
                                    if key not in mp:
                                        mp[key] = {"runn": runn, "mps": mps1}
                                    if runn not in rtomps:
                                        rtomps[runn] = mps1.value
                    wb.close()
                with self._hksvc.sessionctx():
                    maps = self._hksvc.getjos(
                        ["r" + x.split(",")[0] for x in mp.keys()])
                    if maps[1]:
                        logger.debug(
                            "some runnings(%s) do not have JO#" % mp.keys())
                        with open(path.join(fldr, self._fnerrs), "w") as f:
                            wtr = csv.writer(f, dialect="excel")
                            wtr.writerow(
                                ["#failed to get JO# for below runnings"])
                            wtr.writerow(["Runn"])
                            for x in maps[1]:
                                wtr.writerow([x])
                    if maps[0]:
                        with open(fn, "w") as f:
                            wtr = None
                            for x in dict(
                                [(str(x.running), x) for x in maps[0]]).items():
                                runnstr = x[0]
                                it = mp[runnstr + "," + rtomps[runnstr]]
                                it["jono"] = x[1].name.value
                                if not wtr:
                                    wtr = csv.DictWriter(f, it.keys())
                                    wtr.writeheader()
                                wtr.writerow(it)
                if all(maps):
                    for x in maps[1]:
                        key = x + "," + rtomps[x]
                        if key in mp: del mp[key]
            except Exception as e:
                logger.debug(e)
                raise e
            finally:
                if (killxls): app.quit()

        hisoks, hiserrs = {}, {}
        if checkdone:
            fnoks = path.join(fldr, self._fnoks)
            fnerrs = path.join(fldr, self._fnerrs)
            if _checkfile(fnoks, fns):
                with open(fnoks) as f:
                    rdr = csv.DictReader(f)
                    for x in rdr:
                        hisoks[x["Runn"] + "," + x["mps"]] = x
            if _checkfile(fnerrs, fns):
                with open(fnerrs) as f:
                    rdr = csv.DictReader(f)
                    for x in rdr:
                        hiserrs[x["Runn"] + "," + x["mps"]] = x
            rmvs = (hisoks, hiserrs)
            if mp and any(rmvs):
                ks = set()
                [ks.update(x.keys()) for x in rmvs if x]
                for k in ks:
                    if k in mp: del mp[k]
        return mp, hisoks, hiserrs

    def readcalcc1(self, root=None):
        if not root:
            root = r"\\172.16.8.46\pb\dptfile\quotation\2017外发工单工费明细"
        fldrs = [
            path.join(root, x)
            for x in os.listdir(root)
            if path.isdir(path.join(root, x))
        ]
        if not fldrs: return
        kxls, app = xwu.app(False)
        trmap = {"diff": "C1差额", "labor": "C1工费", "jono": "工单"}
        mp = {}
        for fldr in fldrs:
            if len([
                    1 for x in path.basename(fldr)
                    if ord(x) <= 31 or ord(x) >= 127
            ]) > 0:
                print("don't process non-ascii folder(%s)" % fldr)
                continue
            xlsx = getfiles(fldr, "xlsm")
            for fn in xlsx:
                fs = os.stat(fn)
                if fs.st_size > 1024 * 1024:
                    print("%s is too big" % fn)
                    continue
                wb = app.books.open(fn)
                for sht in wb.sheets:
                    rng = xwu.find(sht, "C1差额")
                    if rng: break
                if rng:
                    rng = rng.expand("table")
                    nls = NamedLists(rng.value, trmap)
                    for x in nls:
                        jn = x.jono
                        if not jn: continue
                        je = JOElement(jn)
                        if not je.isvalid: continue
                        try:
                            mp[je.value] = round(x.diff + x.labor, 2)
                        except:
                            print("file(%s), jo#(%s) error" % (fn, je.value))
                wb.close()
        if kxls: app.quit()
        return mp

    def _calcc1(self, c1s, c1invs):
        """ return a tuple, [0] as success, [1] for failed """
        #c1's metal is simple, 1.06 loss
        #c1's stone sometimes inside the labor, but when it's in labor, stones disappear
        #so won't be duplicated
        if not (c1s and c1invs): return
        #for demo only, I test 5 records only
        c1s = c1s[:5]
        rt0 = []
        jns = [x[0] for x in c1s]
        stcosts = self._cnsvc.getjostcosts(jns)
        jos = self._hksvc.getjos(jns)[0]
        if jos:
            jos = dict([(x.name.value, x) for x in jos])
        for jnmps in c1s:
            jn, mps = jnmps[0], jnmps[1]
            ci = c1invs[jn]
            lb = round((ci.labor + ci.setting) / self._usdtormb, 2)
            wgts = self._hksvc.getjowgts(jn)
            mtlcost = PajCalc.calcmtlcost(wgts, mps, vendor=None)
            jo = jos[jn]
            sts0 = ci.stones
            stc = stcosts.get(jn, 0)
            if not stc and sts0:
                stc = -1000
            elif stc:
                stc = round(
                    stc / float(jo.qty) / self._usdtohkd * self._pkratio, 2)
            mp = {
                "runn": jo.running,
                "jono": jn,
                "jono1": jn,
                "pcode": "C1",
                "skuno": jo.po.skuno,
                "styno": jo.style.name.value,
                "customer": jo.customer.name.strip(),
                "wgts": wgts,
                "labor": lb,
                "mtlcost": mtlcost,
                "stcost": stc
            }
            cn = PajChina(mtlcost + lb + stc, PajCalc.calcincrement(wgts), mps,
                          0, mtlcost)
            mp["china"] = cn
            rt0.append(mp)
        return rt0, None

    def run(self, fldr, defaultmps=None):
        """do folder \\172.16.8.46\pb\dptfile\quotation\date\Date2018\0502\(2) quote to customer\ PAJ cost calculation
            find the JO of given running, then do the calc based on S=20;G=1350
            try to read running/mps from fldr and generate result files with PAJ related costs
            if the folder contains @fnrunoks, runnings will be from it, else from the excel files
            the files should contains MPS there, or the default mps will be applied
            @param fldr: the folder to generate data
            @param defaultmps: when no MPS provided in the source file(s), use this. should be an MPS() object
            @param fnoks: file name of the ok result
            @param fnerrs: file name of the error result
            @param fnrunoks: filename of the ok runnings
            @param fnrunerrs: filename of the error runnings
        """

        def _putmap(wnc, runn, orgjn, tarmps, themap):
            key = "%s,%6.1f" % (wnc["PajShp"].pcode, wnc["china"].china)
            if key not in themap:
                mp = {
                    "runn": runn,
                    "jono": orgjn,
                    "china": pajcc.PajCalc.calctarget(wnc["china"], tarmps)
                }
                jo = wnc["JO"]
                mp["skuno"] = jo.po.skuno
                mp["jono1"] = jo.name.value
                mp["styno"] = jo.style.name.value
                mp["customer"] = jo.customer.name.strip()
                mp["pcode"] = wnc["PajShp"].pcode
                mp["wgts"] = wnc["wgts"]
                themap[key] = mp
                return True

        def _writeOks(wtroks, foks, fn, ttroks, oks, hisoks):
            if not wtroks:
                if not foks: foks = open(fn, "a+" if hisoks else "w")
                wtroks = csv.DictWriter(foks, ttroks)
                if not hisoks: wtroks.writeheader()

            for x in sorted(oks.values(), key=lambda x: x["jono"]):
                cost = x["china"]
                jn0 = x["jono"]
                jn1 = x["jono1"]
                rmk = "Actual" if jn0 == jn1 else "Candiate"
                skuno = x["skuno"]
                skuno = "N/A" if not skuno else skuno
                vals = [x["runn"], jn0, x["styno"], x["customer"],jn1,skuno, x["pcode"], cost.china, \
                    cost.metalcost if cost.metalcost else 0, cost.mps.value, rmk, cost.discount * 1.25]
                wgts = cost.increment.prdwgt
                [
                    vals.append("0" if not x else "%s=%f" % (x.karat, x.wgt))
                    for x in wgts
                ]
                rmk = dict(zip(ttroks, vals))
                wtroks.writerow(rmk)
            foks.flush()
            return wtroks, foks

        def _writeErrs(wtrerrs, ferrs, fnerrs, ttrerrs, errs, hiserrs):
            if not wtrerrs:
                if not ferrs: ferrs = open(fnerrs, "a+" if hiserrs else "w")
                wtrerrs = csv.DictWriter(ferrs, ttrerrs)
                if not hiserrs: wtrerrs.writeheader()
            for x in sorted(errs, key=lambda j: j["jono"]):
                ar = [x["runn"], x["jono"]]
                prd = x["wgts"]
                if not prd: prd = PrdWgt(WgtInfo(0, 0))
                for y in [
                    (str(y.karat) + "=" + str(y.wgt) if y else "0") for y in prd
                ]:
                    ar.append(y)
                ar.append(x["mps"])
                wtrerrs.writerow(dict(zip(ttrerrs, ar)))
            ferrs.flush()
            return wtrerrs, ferrs

        fldr = appathsep(fldr)
        fnoks = path.join(fldr, self._fnoks)
        fnerrs = path.join(fldr, self._fnerrs)
        ttroks = "Runn,OrgJO#,Sty#,Cstname,JO#,SKU#,PCode,ttcost,mtlcost,mps,rmks,discount,MainWgt,Auxwgt,PartWgt".split(
            ",")
        ttrerrs = "Runn,OrgJO#,mainWgt,auxWgt,partWgt,mps".split(",")
        errs = []
        hiserrs = None
        wtrerrs = None
        ferrs = None
        wtroks = None
        foks = None
        commitcnt = 5

        mp, hisoks, hiserrs = self.readreqs(fldr, defaultmps)
        if not mp:
            if len(hisoks) + len(hiserrs) > 0:
                logger.debug("everything is up to date")
            return
        oks, dao, stp, cnt, c1s = {}, self._hksvc, 0, len(mp), []
        c1invs = C1InvRdr().read()
        if c1invs:
            c1invs = {x.jono: x for y in c1invs for x in y[0]}
            #labor of c1 invoice is sometimes lower than actual(C1 calculation error)
            #use our monthly xlsm to fix it
            c1invx = self.readcalcc1()
            if c1invx:
                for yy in c1invx.items():
                    ci = c1invs.get(yy[0], None)
                    if not ci: continue
                    ov = ci.labor + ci.setting
                    if ov < yy[1]:
                        logger.debug(
                            "JO(%s)'s labor(%6.2f) in C1 invoice is lower, use ours(%6.2f)"
                            % (yy[0], ov, yy[1]))
                        c1invs[yy[0]] = ci._replace(labor=yy[1], setting=0)
        else:
            c1invs = {}
        try:
            with dao.sessionctx():
                for x in mp.values():
                    if "jono" not in x:
                        logger.critical(
                            "No JO field in map of running(%s)" % x["runn"])
                        continue
                    found = False
                    #if x["jono"] != "B103431": continue
                    jn = x["jono"]
                    if jn in c1invs:
                        c1s.append((jn, x["mps"]))
                        logger.debug("JO#%s is by C1" % jn)
                        continue
                    print("doing " + jn)
                    wnc = dao.calchina(jn)
                    if x["mps"].isvalid:
                        if wnc and all(wnc.values()):
                            found = True
                            if not _putmap(wnc, x["runn"], jn, x["mps"], oks):
                                logger.debug(
                                    "JO(%s) is duplicated for same pcode/cost" %
                                    wnc["JO"].name.value)
                        else:
                            jo = wnc["JO"]
                            if not jo:
                                jo = dao.getjos([jn])
                                jo = jo[0][0] if jo and jo[0] else None
                            if jo:
                                jos = dao.findsimilarjo(jo, 1)
                                if jos:
                                    for x1 in jos:
                                        wnc1 = dao.calchina(x1.name)
                                        if (all(wnc1.values())):
                                            found = True
                                            if not _putmap(
                                                    wnc1, x["runn"], jn,
                                                    x["mps"], oks):
                                                logger.debug(
                                                    "JO(%s) is duplicated for same pcode/cost"
                                                    % str(
                                                        wnc1["JO"].name.value))
                    else:
                        found = False
                        jo = None
                    if not found:
                        if jo and not wnc["wgts"]:
                            wnc["wgts"] = dao.getjowgts(jo)
                        errs.append({
                            "runn": x["runn"],
                            "jono": jn,
                            "wgts": wnc["wgts"],
                            "mps": x["mps"]
                        })
                        if len(errs) > commitcnt:
                            wtrerrs, ferrs = _writeErrs(wtrerrs, ferrs, fnerrs,
                                                        ttrerrs, errs, hiserrs)
                            errs = []
                    if len(oks) > commitcnt:
                        wtroks, foks = _writeOks(wtroks, foks, fnoks, ttroks,
                                                 oks, hisoks)
                        oks = {}
                    stp += 1
                    if not (stp % 20):
                        logger.debug("%d of %d done" % (stp, cnt))
                if c1s:
                    for arr in splitarray(c1s, commitcnt * 2):
                        xx = self._calcc1(arr, c1invs)
                        if xx[0]:
                            for x in xx[0]:
                                key = "%s,%6.1f" % (x["styno"],
                                                    x["china"].china)
                                oks[key] = x
                        if xx[1]: errs.extend(xx[1])
                        if len(oks) > commitcnt:
                            wtroks, foks = _writeOks(wtroks, foks, fnoks,
                                                     ttroks, oks, hisoks)
                            oks = {}
            if len(oks) > 0:
                wtroks, foks = _writeOks(wtroks, foks, fnoks, ttroks, oks,
                                         hisoks)
            if errs:
                wtrerrs, ferrs = _writeErrs(wtrerrs, ferrs, fnerrs, ttrerrs,
                                            errs, hiserrs)
        finally:
            if foks: foks.close()
            if ferrs: ferrs.close()

        return fnoks, fnerrs

    @classmethod
    def readquoprice(self, fldr, rstfn="costs.dat"):
        """read simple quo file which contains Running:xxx, Cost XX: excel
        @param fldr: the folder to read files from
        @return: the result file name or None if nothing is returned
        """
        if not fldr: return
        fldr = appathsep(fldr)
        app, kxl = xwu.appmgr.acq()
        ptnRunn = re.compile(r"running\s?:\s?(\d*)", re.IGNORECASE)
        ptnCost = re.compile(r"^(cost\s?(\w*)\s?:?)|(N\.cost\s?(\w*)\s?:?)",
                             re.IGNORECASE)
        def _get_nbrs(tr, idx_frm):
            flag, costs = False, []
            for x in tr[idx_frm:]:
                if isinstance(x, numbers.Number):
                    costs.append(x)
                    flag = True
                elif flag:
                    break
            return ";".join([str(x) for x in costs]) if costs else None

        lst, wb = [], None
        for fn in _getexcels(fldr):
            wb = app.books.open(fn)
            for sh in wb.sheets:
                phase = rowrunn = 0
                runns, costs = [], []
                vvs = xwu.usedrange(sh).value
                for hh in range(len(vvs)):
                    tr = [x for x in vvs[hh] if x]
                    if not tr:
                        continue
                    for ii, x in enumerate(tr):
                        if not isinstance(x, str):
                            continue
                        if phase <= 1:
                            mt = ptnRunn.search(x)
                            if mt:
                                if phase != 1:
                                    phase, rowrunn = 1, hh
                                runns.append(mt.group(1))
                                continue
                        if phase >= 1:
                            mt = ptnCost.search(x)
                            if not mt:
                                continue
                            x = _get_nbrs(tr, ii)
                            if x:
                                costs.append((mt.group(2) or mt.group(4), x, fn))
                            if phase != 2 and hh != rowrunn:
                                phase = 2
                    if phase == 2:
                        if len(costs) != len(runns):
                            logger.debug("full row data contains runnings(%s) contains invalid cost data in file(%s)" % (runns, fn))
                            costs = [(NA, NA, fn, ), ] * len(runns)
                        if runns:
                            for runn, cost in zip(runns, costs):
                                lst.append((runn, cost[0], cost[1], cost[2]))
                        phase, runns, costs = 0, [], []
            wb.close()
            wb = None
        if wb:
            wb.close()
        if kxl:
            xwu.appmgr.ret(kxl)
        if lst:
            fn = path.join(fldr, rstfn or "costs.dat")
            with open(fn, "w") as f:
                wtr = csv.writer(f, dialect="excel")
                wtr.writerow("runn,karat,cost,file".split(","))
                for x in lst:
                    wtr.writerow(x)
        return fn


class InvAnalysis(object):
    """ TODO:: do this after ack process the weekly PAJ Invoice Detail Analysis
    """

    def run(self, srcfldr, tarfile):
        xls, app = xwu.app(False)
        srcfldr = appathsep(srcfldr)
        fns = _getexcels(srcfldr)
        try:
            for fn in fns:
                wb = app.books.open(srcfldr + fn)
                for sht in wb.sheets():
                    rng = xwu.find(sht, "PAJ_REFNO")
                    if not rng: continue
                    lst = xwu.list2dict(xwu.usedrange(sht))
        except:
            pass
        finally:
            if wb: wb.close()
