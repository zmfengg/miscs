#! coding=utf-8
'''
* @Author: zmFeng
* @Date: 2018-11-08 08:40:11
* @Last Modified by:   zmFeng
* @Last Modified time: 2018-11-08 08:40:11
Shipment handler
'''
import random
import re
from datetime import date, datetime, timedelta
from os import (
    path,
    rename,
)
from time import clock
from copy import copy
from tkinter import filedialog, messagebox

from sqlalchemy import and_, func
from sqlalchemy.orm import Query, aliased
from xlwings.constants import (BorderWeight, Constants, FormatConditionOperator,
                               FormatConditionType, LineStyle)
from xlwings.utils import col_name

from hnjapp.c1rdrs import C1InvRdr, C3InvRdr
from hnjcore import JOElement, samekarat
from hnjcore.models.cn import MM, MMgd, MMMa
from hnjcore.models.hk import JO as JOhk
from hnjcore.models.hk import Orderma, PajAck, POItem
from hnjcore.utils.consts import NA
from utilz import (NamedList, NamedLists, ResourceCtx, easydialog, easymsgbox,
                   list2dict, splitarray, triml, trimu, xwu, deepget, karatsvc,
                   getfiles)

from .common import _logger as logger
from .dbsvcs import jesin
from .pajcc import cmpwgt
from .pajrdrs import PajBomHdlr, PajShpHdlr, PajBomDAO

_appmgr = xwu.appmgr


def _adjwgtneg(wgt):
    """
    sometimes PrdWgt.part contains negative value(for not sure), adjust it to pos
    """
    if wgt < 0:
        wgt = -wgt
        if wgt > 30:
            wgt /= 100.0
    return wgt


def _hl(rng, clidx=3):
    if not rng:
        return
    rng.api.interior.colorindex = clidx


class PajNSOFRdr(object):
    """
    class to read a NewSampleOrderForm's data out
    """
    _tplfn = r"\\172.16.8.46\pb\dptfile\pajForms\PAJSKUSpecTemplate.xlt"

    def readsettings(self, fn=None):
        """
        read the setttings from setting sheet of NSOF
        """
        usetpl, mp = False, None
        if not fn:
            fn, usetpl = self._tplfn, True
        app, kxl = _appmgr.acq()
        try:
            wb = app.books.open(fn) if not usetpl else xwu.fromtemplate(fn, app)
            shts = [x for x in wb.sheets if triml(x.name).find("setting") >= 0]
            if shts:
                rng = xwu.find(shts[0], "name")
                nls = NamedLists(rng.expand("table").value)
                mp = {triml(nl.name): nl for nl in nls}
        finally:
            if wb:
                wb.close()
            _appmgr.ret(kxl)
        return mp if mp else None


class _SMSns(object):
    """ util class for ShpMkr to organize the warning msgs and related operations """

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        msgs = "sn_rpt,Rpt,SHEET_NAME;sn_err,错误,SHEET_NAME;sn_warn,警告,SHEET_NAME;sn_bc,BCData,SHEET_NAME;ec_qty,数量错误,ERROR;ec_jn,工单号错误,ERROR;ec_wgt_not_sure,重量不确定,ERROR;ec_wgt_missing,无重量数据,ERROR;ec_wgt,重量数据严重不符,ERROR;ec_jmp,JMP号错误,ERROR;ec_date,落货日期错误,ERROR;ec_sh_error,格式错误,ERROR;ec_inv_none,无发票资料,ERROR;wc_wgt,重量不符警告,WARN;ec_karat,成色错误,ERROR;wc_ack,Ack警告,WARN;wc_date,日期警告,WARN;wc_qty,落货数量警告,WARN;wc_inv_qty,发票数量警告,WARN;wc_smp,样板相关,WARN"
        self._errs = {y[0]: y for y in (x.split(",") for x in msgs.split(";"))}

    def get(self, wb, sn, auto_add=True):
        """
        get or create sheet with given sheetname(sn)
        """
        sn = self.get_error(sn)[1] or sn
        try:
            sht = wb.sheets[sn]
        except:
            sht = wb.sheets.add(sn, after=wb.sheets[-1]) if auto_add else None
        return sht

    @classmethod
    def new_err(cls, *args):
        """
        new a dict holding the key error info
        the argument sequence are:
        jn, loc, etype, msg, objs
        """
        return {
            "jono": "'" + args[0],
            "location": "'" + args[1],
            "type": args[2],
            "msg": args[3],
            "objs": None if len(args) < 5 else args[4]
        }

    def eap(self, errlst, *args):
        """
        create an error item and append it to the first argument(it should be a list)
        """
        errlst.append(self.new_err(*args))
        type_name = args[2]
        #check if the weight is too critical. (append 20190312 but when it's a QC sample, don't do it. QCSample flag is inside args[5] if there is
        if type_name == "wc_wgt" and len(args) < 6:
            jwgt, shpwgt = args[4]
            if not cmpwgt(jwgt, shpwgt, 50):
                type_name = [x for x in args]
                type_name[2:] = "ec_wgt", "重量偏离指定值50%以上", None
                errlst.append(self.new_err(*type_name))

    def _gettype(self, name):
        """ return the type of given name, can be one of
        WARN/ERR
        """
        rc = self.get_error(name)
        return rc[2] if rc else None

    def get_error(self, name):
        """ given name, return name/type/msg error tuple
        if name is None, return tuple of error tuple
        """
        return self._errs.get(triml(name))

    def err_warn(self, errlst):
        return (tuple(x for x in errlst if self._gettype(x["type"]) == "ERROR"),
                tuple(x for x in errlst if self._gettype(x["type"]) != "ERROR"))


class ShpMkr(object):
    """ class to make the daily shipment, include below functions
    .build the report if there is not and maintain the runnings
    .build the bc data
    .make the import
    .do invoice comparision
    @param cnsvc/hksvc/bcsvc: the related db services
    @param(optional) nsofn: the NewSampleOrderForm file name,
        I used it to read Paj/C1 repositories
    @param(optional) cache: a sessionMgr pointing to a cache db where
        BOM hints can be saved/queried
    """
    _mergeshpjo = False
    _vdrname = _nsofsts = None
    _sns = _SMSns()
    _nsofn = None

    def __init__(self, cnsvc, hksvc, bcsvc, **kwds):
        self._cnsvc, self._hksvc, self._bcsvc = cnsvc, hksvc, bcsvc
        self._nsofn, self._cache = kwds.get("nsofn"), kwds.get("cache")
        # debug mode, qtyleft/running should be reset so that I can
        # generate proper report
        self._debug = kwds.get("debug", False)

    def _check_src_file(self, fldr):
        """
        check if the source file is already there. If no valid location
        provided, return None. if exists, return "tarfn" as key, else
        return a map with keys(tarfldr,fns,fn)
        """
        if not fldr:
            fldr = easydialog(
                filedialog.Directory(
                    title="Choose folder contains all raw files from PAJ"))
            if not path.exists(fldr):
                return None
        sts = self._nsofsettings()
        tarfldr, tarfn = path.dirname(sts.get("shp.template").value), None
        fns = getfiles(fldr, ".xls")
        var = re.compile(r"^HNJ \d+")
        for fn in fns:
            if var.search(path.basename(fn)):
                tdm = PajShpHdlr.get_shp_date(fn)
                if tdm:
                    tarfn = "HNJ %s 出货明细" % tdm.strftime("%Y-%m-%d")
                    break
        if not tarfn:
            return {"tarfn": None}
        sts = getfiles(tarfldr, tarfn)
        if sts:
            tarfn = sts[0]
            tdm = path.getmtime(tarfn)
            fds = copy(fns)
            fds.append(fldr)
            fds = max([path.getmtime(x) for x in fds])
            if fds > tdm:
                messagebox.showwarning("文件过期",
                                       "%s\n已过期,请手动删除或更新后再启动本程序" % tarfn)
                app = _appmgr.acq()[0]
                app.books.open(tarfn)
                app.visible = True
                return {"tarfn": None, "expired": True}
            logger.debug("result file(%s) already there" % tarfn)
            return {"tarfn": tarfn}
        if len(fns) == 1:
            return {"tarfn": fns[0]}
        return {"tarfldr": tarfldr, "fns": fns, "fn": tarfn}

    def _pajfldr2file(self, fldr):
        """ group the folder into one target file. If target file already exists,
        do date check
        @return : filename if succeeded
                  -1 if file expired
                  None if unexpected error occured
        """
        mp = self._check_src_file(fldr)
        if not mp:
            return None
        if "tarfn" in mp:
            return mp["tarfn"]
        tarfldr, fns = (mp[x] for x in "tarfldr fns".split())
        app, kxl = _appmgr.acq()
        wb = app.books.add()
        fds = [x for x in wb.sheets]
        for fn in (x for x in fns if x.find("对帐单") < 0):
            var = xwu.safeopen(app, fn)
            try:
                for sht in (
                        x for x in var.sheets
                        if x.api.visible == -1 and xwu.usedrange(x).size > 1):
                    sht.api.Copy(Before=fds[0].api)
            finally:
                var.close()
        for var in fds:
            var.delete()

        wb.save(path.join(tarfldr, mp["fn"]))
        tarfn = wb.fullname
        logger.debug("merged file saved to %s" % tarfn)
        wb.close()
        _appmgr.ret(kxl)
        return tarfn

    def _read_c1_3(self, sht, args, vdr):
        """
        read shipment data from c1 sheet, return result dict and
        an error list, inside the result map, the "shpdate" holds
        the shipment date for later reference
        """
        # determine the header row
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
            sht.range("1:%d" % ridx).api.entirerow.delete()
        mp = sht.name
        its = (C1InvRdr() if vdr == 1 else C3InvRdr()).read(sht)
        if not its:
            logger.debug("no valid data in sheet(%s)" % mp)
            return (None,) * 2
        its = its[0]
        mp, errs, args["shpdate"] = {}, [], its[1]
        for shp in its[0]:
            jn = shp.jono
            key = jn if self._mergeshpjo else jn + str(random.random())
            it = mp.setdefault(key, {
                "jono": jn,
                "qty": 0,
                "location": "%s" % jn
            })
            it["mtlwgt"] = shp.mtlwgt
            it["qty"] += shp.qty
            if shp.stones:
                it["_stone_data"] = self._stdto_c1_paj(shp.stones)
        if mp:
            mp["shpdate"] = args.get("shpdate")
        return mp, errs

    def _read_c1(self, sht, args):
        return self._read_c1_3(sht, args, 1)

    def _read_c3(self, sht, args):
        return self._read_c1_3(sht, args, 3)

    def _stdto_c1_paj(self, c1sts):
        '''
        dto the c1 stone to paj stone format
        '''
        #using the namedlist like pajrdrs._StMaker._nl_rc
        # NamedList('qty shape stone size wgt'.split())
        nl = NamedList('qty shape stone size wgt'.split())
        lsts = [nl, ]
        for c1 in c1sts:
            lsts.append(nl.newdata())
            nl.qty, nl.shape, nl.stone, nl.size, nl.wgt = c1.qty, c1.shape, c1.stone, c1.remark, c1.wgt
        return lsts

    def _read_c2(self, sht, args):
        pass

    def _read_paj(self, sht, args):
        """ return tuple(map,errlist)
        where errlist contain err locations
        """
        shps = PajShpHdlr.read_shp(args["fn"], args["shpdate"], args["fmd"],
                                   sht, args.get("bomwgts"))
        if not shps:
            return (None, None)
        if "_ERROR_" in shps:
            return (None, shps["_ERROR_"])
        mp, errs, shp_date = {}, [], args["shpdate"]
        for shp in shps.values():
            jn = shp.jono
            shp_date = max(shp_date, shp.invdate)
            key = jn if self._mergeshpjo else jn + str(random.random())
            it = mp.setdefault(key, {
                "jono": jn,
                "qty": 0,
                "location": "%s(%s)" % (jn, shp.pcode)
            })
            it["mtlwgt"] = shp.mtlwgt
            it["qty"] += shp.qty
            if shp.stwgt:
                sts = shp.stwgt
                # there might be a sn# there, if yes, put it into map for reference
                if isinstance(sts[0], str):
                    it["_snno"] = sts[0]
                    if len(sts) == 1:
                        continue
                    sts = sts[1:]
                it["_stone_data"] = sts
        mp["shpdate"] = shp_date
        return (mp, errs)

    def _check_db_error(self, shplst, invmp, errlst):
        """
        check the source data about weight/pajinv
        @param shpmp: the shipment map with JO# as key and map as value
        """
        jns = {x["jono"] for x in shplst}
        if self._vdrname == "paj":
            self._check_db_inv(shplst, invmp, errlst, jns)
        t0 = clock()
        logger.debug("Begin to verify shipment qty&wgt")
        with self._cnsvc.sessionctx():
            jos = self._cnsvc.getjos(jns)
            jos = {x.name.value: x for x in jos[0]}
            if self._debug:
                for jo in jos.values():
                    jo.qtyleft, jo.running = jo.qty, 0
            jwgtmp, jncmp, shpwgts = {}, {}, {}
            nmap = {
                "cstname": "customer.name",
                "styno": "style.name.value",
                "running": "running",
                "description": "description",
                "qtyleft": "qtyleft"
            }
            for mp in shplst:
                jn = mp["jono"]
                jncmp[jn] = jncmp.get(jn, 0) + 1
                if jn not in shpwgts:
                    shpwgts[jn] = mp["mtlwgt"]
                # place the ordertype into shplst for further usage
                mp["ordertype"] = jos.get(jn).ordertype
            for mp in shplst:
                jn = mp["jono"]
                jncmp[jn] -= 1
                jo = jos.get(jn)
                if not jo:
                    self._sns.eap(errlst, jn, mp["location"], "ec_jn",
                                  "工单号(%s)错误" % jn, None)
                    continue
                for y in nmap.items():
                    sx = deepget(jo, y[1])
                    if sx and isinstance(sx, str):
                        sx = sx.strip()
                    mp[y[0]] = sx
                jo.qtyleft = jo.qtyleft - mp["qty"]
                if jo.qtyleft < 0:
                    s0 = "数量不足"
                    self._sns.eap(errlst, mp["jono"], mp["location"], "ec_qty",
                                  s0, (jo.qtyleft + mp["qty"], mp["qty"]))
                    mp["errmsg"] = s0
                elif jo.qtyleft > 0 and not jncmp[jn]:
                    s0 = "数量有余"
                    self._sns.eap(errlst, mp["jono"], mp["location"], "wc_qty",
                                  s0, (jo.qtyleft + mp["qty"], mp["qty"]))
                    mp["errmsg"] = s0
                else:
                    mp["errmsg"] = ""
                jwgt = jwgtmp.get(jn)
                if not jwgt and jn not in jwgtmp:
                    jwgt = self._hksvc.getjowgts(jn)
                    if not jwgt:
                        jwgt = None
                    else:
                        # the weight from JO won't contains stone, this might lead to error prompting for weight error, so let them be the same
                        jwgt = jwgt._replace(netwgt=mp["mtlwgt"].netwgt)
                    jwgtmp[jn] = jwgt
                if not cmpwgt(jwgt, mp["mtlwgt"]):
                    haswgt = bool(
                        [x for x in mp["mtlwgt"].wgts if x and x.wgt > 0])
                    jn = [
                        errlst,
                        mp["jono"],
                        mp["location"],
                    ]
                    if haswgt:
                        if jo.ordertype not in ("O", 'Q'):
                            jn = None
                        else:
                            jn.extend(("wc_wgt", "重量不符"))
                    else:
                        jn.extend(("ec_wgt_missing", "欠重量资料"))
                    if jn:
                        jn.append((jwgt, mp["mtlwgt"], ))
                        if jo.ordertype == 'Q':
                            jn.append("_QCSAMPLE_")
                        self._sns.eap(*jn)
                jn = (
                    jo.karat,
                    mp["mtlwgt"].main.karat,
                )
                if jn[0] != jn[1]:
                    self._sns.eap(errlst, mp["jono"], mp["location"],
                                  "ec_karat", "主成色与工单成色不一致", jn)
            jncmp = {
                "PAJ,N": "新版,请向PAJ索要产品图",
                "PAJ,Q": "QC版，更新重量",
                "C1,N": "新版,请向C1索要(JCAD图),并编制《图文技术说明》(如无图烦请香港补照)",
                "C1,Q": "QC版,请查看是否需要《编制图文技术说明》(如无图烦请香港补照)"
            }
            for jo in jos.values():
                s0 = jncmp.get((self._vdrname + "," + jo.ordertype).upper())
                if not s0:
                    continue
                jn = jo.name.value
                self._sns.eap(errlst, jn, jn, "wc_smp", s0, shpwgts[jn])
        logger.debug("using %fs for above action" % (clock() - t0))

    def _ne(self, *args):
        return self._sns.new_err(*args)

    def _check_db_inv(self, shplst, invmp, errlst, jns):
        if not invmp:
            self._sns.eap(errlst, NA, "_all_", "ec_inv_none", "不应无发票资料")
            return
        logger.debug("begin to fetch ack/inv data")
        t0 = clock()
        with self._hksvc.sessionctx() as cur:
            q = Query([JOhk, PajAck]).join(PajAck).filter(
                jesin(set([JOElement(x) for x in jns]), JOhk))
            q = q.with_session(cur).all()
            logger.debug("using %fs to fetch %d JOs for above action" %\
                        (clock() - t0, len(jns)))
            tmp = "uprice,mps,ackdate,docno,mps,pcode".split(",")
            acks = {
                x[0].name.value: tuple(getattr(x[1], y) for y in tmp) for x in q
            } if q else {}
            if acks:
                nlack = NamedList(
                    list2dict(",".join(tmp), alias={"date": "ackdate"}))
        tmp = {}
        for x in invmp.values():
            tmp1 = tmp.setdefault(x.jono, {"jono": x.jono})
            if "inv" not in tmp1:
                tmp1["inv"], tmp1["invqty"] = x, 0
            else:
                x0 = tmp1["inv"]
                if abs(x0.uprice - x.uprice) > 0.001:
                    self._sns.eap(errlst, x.jono, x.jono, "wc_inv_diff",\
                            "工单(%s)对应的发票单价前后不一致" % x.jono,\
                            (x.uprice, x0.uprice))
            tmp1["invqty"] += x.qty
        for x in shplst:
            jn = x["jono"]
            tmp1 = tmp.get(jn)
            if tmp1:
                tmp1["qty"] = tmp1.get("qty", 0) + x["qty"]
        for x in tmp.values():
            if x.get("invqty") != x.get("qty"):
                if not x.get("qty"):
                    self._sns.eap(
                        errlst, x["jono"],
                        "Inv(%s),JO#(%s)" % (x["inv"].invno, x["jono"]),
                        "wc_inv_qty",
                        "工单(%s)有发票(%s)无落货" % (x["jono"], x["inv"].invno), None)
                else:
                    self._sns.eap(errlst, x["jono"], x["jono"],
                            "wc_inv_qty", "落货数量(%s)与发票数量(%s)不一致" %\
                            (str(x.get("qty", 0)), str(x.get("invqty", 0))),
                            (x.get("qty", 0), x.get("invqty")))
            if acks:
                ack = acks.get(x["jono"])
                if not ack:
                    continue
                inv = x["inv"]
                nlack.setdata(ack)
                if abs(inv.uprice - float(nlack.uprice)) > 0.01:
                    jn = x["jono"]
                    self._sns.eap(
                        errlst, jn, jn, "wc_ack", nlack.pcode, {
                            "inv": inv.uprice,
                            "ack": nlack.uprice,
                            "inv_mps": inv.mps,
                            "ack_mps": nlack.mps,
                            "file": nlack.docno,
                            "date": nlack.date.strftime("%Y-%m-%d")
                        })
        tmp = jns.difference(tmp.keys())
        if tmp:
            for x in tmp:
                self._sns.eap(errlst, x, x, "wc_qty", "工单(%s)有落货无发票" % x, None)

    def _nsofsettings(self, fn=None):
        if self._nsofsts is None:
            self._nsofsts = PajNSOFRdr().readsettings(fn or self._nsofn)
        return self._nsofsts

    def _write_rpts_pgsetup(self, sts, sht, shp_date, iorst):
        """
        setup sheet("rpt")'s margins/header/footer
        according to pylint, the sht arg is not used, it gives
        """
        s0 = sts.get("shipment.rptmgns.%s" % self._vdrname)
        if not s0:
            s0 = sts.get("shipment.rptmgns")
        pfx = "sht.api.pagesetup"
        shtcmds = [
            pfx + ".%smargin=%s" % tuple(y.split("="))
            for y in triml(s0.value).split(";")
        ]
        shtcmds.append(pfx + ".printtitlerows='$1:$1'")
        shtcmds.append(
            pfx + ".leftheader='%s'" %
            ("%s年%s月%s日落货资料" % tuple(shp_date.strftime("%Y-%m-%d").split("-"))))
        shtcmds.append(pfx + ".centerheader='%s'" % iorst["jmp#"])
        shtcmds.append(pfx + ".rightheader='%s'" % iorst["n#"])
        shtcmds.append(pfx + ".rightfooter='&P of &N'")
        shtcmds.append(pfx + ".fittopageswide=1")
        for x in shtcmds:
            exec(x)

    def _write_rpts(self, wb, shplst, newrunmp, shp_date):
        """
        send the shipment related sheets(Rpt/Err)
        """
        sts = self._nsofsettings()
        io_hlp = _SMIOHlpr(wb.app, sts, shp_date, self._vdrname)
        iorst = io_hlp.get()
        if not iorst:
            logger.debug("failed to get get valid IO item, generation failed")
            return None
        sht = self._sns.get(wb, "sn_rpt")
        self._write_rpts_pgsetup(sts, sht, shp_date, iorst)
        #now prepare and write the data
        s0 = sts.get("shipment.hdrs." + self._vdrname)
        if not s0:
            s0 = sts.get("shipment.hdrs")
        ttl, ns = [], {}
        for tl, s0 in (
                x.split("=") for x in s0.value.replace(r"\n", "\n").split(";")):
            ttl.append(tl)
            y1 = s0.split(",")
            if len(y1) > 1:
                ns[y1[0]] = tl
            sht.range(1, len(ttl)).column_width = float(y1[-1])
        ns["thisleft"] = "此次,"
        nl = NamedList(list2dict(ttl, alias=ns))
        ns = "jono running qty cstname styno description qtyleft errmsg".split()
        maxr, lenttl, lsts, hls = iorst["maxrun#"], len(ttl), [ttl], []
        shplst = sorted(
            shplst,
            key=lambda mpx: "A%06d%s" % (mpx.get("running") or 0, mpx["jono"]))
        for it in shplst:
            ttl = [""] * lenttl
            nl.setdata(ttl)
            if not it.get("running"):
                if it["jono"] not in newrunmp:
                    maxr += 1
                    it["running"] = nl["running"] = maxr
                    hls.append((len(lsts) + 1, nl.getcol("running")))
                    newrunmp[it["jono"]] = maxr
                else:
                    # sometimes it's a zero, just don't show it
                    if not it["running"]:
                        it["running"] = None
            for col in ns:
                nl[col] = it[col]
            nl.jono, karats = "'" + nl.jono, {}
            for wi in it["mtlwgt"].wgts:
                if not wi:
                    continue
                if wi.karat not in karats:
                    karats[wi.karat] = wi.wgt
                else:
                    karats[wi.karat] += wi.wgt
            karats = [(x[0], x[1]) for x in karats.items()]
            nl.karat1, nl.wgt1 = karats[0][0], karats[0][1]
            lsts.append(ttl)
            if nl.wgt1 < 0:
                hls.append((len(lsts), nl.getcol("wgt1")))
            if len(karats) > 1:
                jn = nl.jono
                for idx in range(1, len(karats)):
                    nl.setdata([""] * lenttl)
                    nl.jono = jn
                    nl.karat1, nl.wgt1 = karats[idx][0], karats[idx][1]
                    lsts.append(nl.data)
                    if nl.wgt1 < 0:
                        hls.append((len(lsts), nl.getcol("wgt1")))
        sht.range(1, 1).value = lsts
        if hls:
            rng = sht.range(1, 1)
            for x in hls:
                _hl(xwu.offset(rng, x[0] - 1, x[1]), 6)
        # the qtyleft formula
        s0 = {
            x: col_name(nl.getcol(x) + 1)
            for x in "qty qtyleft thisleft".split()
        }
        for idx in range(2, len(lsts) + 1):
            rng = sht.range("%s%d" % (s0["thisleft"], idx))
            rng.formula = "=%s%d-%s%d" % (s0["qtyleft"], idx, s0["qty"], idx)
            rng.api.numberformatlocal = "_ * #,##0_ ;_ * -#,##0_ ;_ * " "-" "_ ;_ @_ "
            rng.api.formatconditions.add(FormatConditionType.xlCellValue,
                                        FormatConditionOperator.xlLess, "0")
            rng.api.formatconditions(1).interior.colorindex = 3

        rng = sht.range(sht.range(1, 1), sht.range(len(lsts), len(nl.colnames)))
        rng.api.borders.linestyle = LineStyle.xlContinuous
        rng.api.borders.weight = BorderWeight.xlThin
        # the sum formula at the bottom
        s0 = int(nl.getcol("qty")) + 1
        rng = sht.range(len(lsts) + 1, s0)
        rng.formula = "=sum(%s1:%s%d)" % (col_name(s0), col_name(s0), len(lsts))
        rng.api.font.bold = True
        rng.api.borders.linestyle = LineStyle.xlContinuous
        rng.api.borders.weight = BorderWeight.xlThin
        sht.range("A2:A%d" % (len(lsts) + 1)).row_height = 18
        rng = xwu.usedrange(sht).api
        rng.VerticalAlignment = Constants.xlCenter
        rng.font.name, rng.font.size = "tahoma", 10

        iorst["maxrun#"] = maxr
        io_hlp.update(iorst)
        return 1

    def _get_file(self, fldr):
        """
        return the final file for report generation
        """
        sts = self._nsofsettings()
        if not fldr:
            fldr = easydialog(
                filedialog.Open(
                    "Choose a file to create shipment",
                    initialdir=path.dirname(
                        path.dirname(sts["shp.template"].value))))
        if not path.exists(fldr):
            return None
        if not path.isdir(fldr):
            if self._cache:
                var = self._commit_bom(fldr)
                if var:
                    return "_BOM_COMITTED_%d" % len(var)
            # if the file's parent folder not in ref_folder, treate it as pajraw files
            getrf = lambda x: triml(path.basename(path.dirname(x)))
            if getrf(fldr) not in [
                    getrf(sts[x].value)
                    for x in "shp.template shpc1.template".split()
            ]:
                fldr = path.dirname(fldr)
        return self._pajfldr2file(fldr) if path.isdir(fldr) else fldr

    def _get_reader(self, wb):
        """
        try to detect the reader/vender_type of given workbook
        """
        rdrmap = {
            "长兴珠宝": ("c2", self._read_c2),
            "诚艺,胤雅": ("c1", self._read_c1),
            "帝宝": ("c3", self._read_c3),
            "十七,物料编号,paj,diamondlite": ("paj", self._read_paj)
        }
        for sht in wb.sheets:
            if self._vdrname:
                break
            for x in rdrmap:
                if tuple(y for y in x.split(",") if xwu.find(sht, y)):
                    rdr, self._vdrname = rdrmap[x], rdrmap[x][0]
                    return rdr
        return None

    def _commit_bom(self, fn):
        tk = None
        if isinstance(fn, str):
            app, tk = _appmgr.acq()
            wb = app.books.open(fn)
        else:
            wb = fn
        fn = PajBomDAO(self._cache).cache(wb)
        _appmgr.ret(tk)
        return fn

    def build_rpts(self, fldr=None):
        """ create the rpt/err/bc sheets if they're not available
        @return: workbook if no error is found and None if err found during report generation
        """
        fn = self._get_file(fldr)
        if not fn:
            logger.debug("user does not specified any valid source file")
            return (None, ) * 2
        if fn.startswith("_BOM_COMITTED_"):
            # the user commit a bom file, return the process result
            return (fn, None)
        pajopts = {
            "fn": fn,
            "shpdate": PajShpHdlr.get_shp_date(fn),
            "fmd": datetime.fromtimestamp(path.getmtime(fn))
        }

        invmp, shplst, errlst, self._vdrname = {}, [], [], None
        app, kxl = _appmgr.acq()
        try:
            wb = app.books.open(fn)
            # check if Rpt sheet was already there
            var = tuple(self._sns.get(wb, x, False) for x in ("sn_rpt",))
            flag = var and any(var)
            if flag:
                rdr = xwu.usedrange(var[0])
                flag = bool(rdr) or [x for x in rdr.value if x]
            if flag:
                logger.debug("target file(%s) don't need regeneration", wb.name)
                # get reader method will detect the vendor name
                self._get_reader(wb)
                return wb, self._vdrname

            rdr = self._get_reader(wb)
            if not rdr:
                return (None, ) * 2
            if self._vdrname == "paj":
                pajopts["bomwgts"] = PajBomHdlr(part_chk_ver=1, cache=self._cache).readbom(wb)
            crt_err, shp_date = False, None  # critical error flag
            for sht in wb.sheets:
                flag, lst = False, rdr[1](sht, pajopts)
                if lst and any(lst):
                    flag, mp = True, lst[0]
                    if mp:
                        # shpdate overriding
                        if "shpdate" in mp:
                            shp_date = mp["shpdate"]
                            del mp["shpdate"]
                        shplst.extend(mp.values())
                    if lst[1]:
                        self._sns.eap(errlst, sht.name, sht.name,
                                    "ec_sh_error", lst[1])
                        crt_err = True
                if not flag and self._vdrname == "paj" and not lst[1]:
                    invno = PajShpHdlr.read_invno(sht)
                    if not invno:
                        continue
                    mp = PajShpHdlr.read_inv_raw(sht, invno)
                    if mp:
                        invmp.update(mp)
            if shp_date and isinstance(shp_date, str):
                shp_date = datetime.strptime(shp_date, "%Y-%m-%d")
            if shp_date:
                var = datetime.today().date() - shp_date
                if var.days > 2 or var.days < 0:
                    self._sns.eap(errlst, "_all_", "_日期_", "wc_date",
                                "落货日期%s可能错误" % shp_date.strftime("%Y-%m-%d"),
                                (var, shp_date))
                #rename c1's source file based on the shp_date
                if self._vdrname != "paj" and trimu(
                        path.basename(fn)).find(trimu(self._vdrname)) != 0:
                    wb.save()
                    wb.close()
                    var = path.splitext(fn)[1]
                    var = "%s %s 落货明细%s" % (trimu(self._vdrname), shp_date.strftime("%y%m%d"), var)
                    var = path.join(path.dirname(fn), var)
                    rename(fn, var)
                    fn, wb = var, app.books.open(var)
            if shplst:
                if self._debug or not crt_err:
                    self._check_db_error(shplst, invmp, errlst)
            if errlst:
                _SMLogWtr(self._cnsvc, self._sns).write(wb, errlst)
                errlst = self._sns.err_warn(errlst)[0]
            if not shplst:
                wb = None
            else:
                if not errlst or (errlst and self._debug):
                    mp = {} # map to hold the new created runnings
                    self._write_rpts(wb, shplst, mp, shp_date)
                    hls = _SMBCHdlr(self._bcsvc, self._hksvc, self._cnsvc, self._sns).write(wb, shplst, mp, shp_date)
                    self._write_wgts(wb, shplst)
                    if hls:
                        PajShpHdlr.build_bom_sheet(wb, min_rowcnt=10, main_offset=3, bom_check_level=1)
                else:
                    PajShpHdlr.build_bom_sheet(wb, min_rowcnt=10, main_offset=3, bom_check_level=1)
                    wb = None
        finally:
            _appmgr.ret(kxl)
        return wb, self._vdrname

    @staticmethod
    def _fmt_wgtinfo(wi, tn):
        return "%s%s:%4.2fgm" % (karatsvc.getkarat(wi.karat).name, tn, wi.wgt)

    @classmethod
    def _write_wgts(cls, wb, shplst, div='-' * 15):
        '''
        write the wgt_info for the new/qc/sample samples
        '''
        smps = [x for x in shplst if (x.get("ordertype") or 'O') != 'O']
        if not smps:
            return
        wgts = "Wgts"
        sht = xwu.findsheet(wb, wgts)
        if not sht:
            sht = wb.sheets.add(wgts, after=wb.sheets[-1])
        else:
            xwu.usedrange(sht).value = None
        nmp = {'jono': "工单", 'type': '类型', 'main': '主体重', 'chain': '链重', 'net': '连石重', 'aio': "全部备注"}
        # lst = tuple(nmp.values())
        nl, lst = NamedList(tuple(nmp.keys())), []
        for x in smps:
            lst.append(nl.newdata(True))
            nl.jono, nl.type = "'" + x["jono"], x["ordertype"]
            wgts, aio = x["mtlwgt"], []
            var = [cls._fmt_wgtinfo(x, '重') for x in wgts.metal]
            aio.extend(var)
            nl.main = "\n".join(var)
            # show netwgt only when there is stone
            if wgts.metal_stone:
                if div:
                    aio.append(div)
                aio.append("连石重:%4.2fgm" % wgts.metal_stone)
                nl.net = aio[-1]
            var = wgts.chain
            if var:
                if div:
                    aio.append(div)
                aio.append(cls._fmt_wgtinfo(var, "链重"))
                nl.chain = aio[-1]
            nl.aio = "\n".join(aio)
        lst = sorted(lst, key=lambda x: x[nl.getcol("jono")])
        lst.insert(0, tuple(nmp.values()))
        sht.cells(1, 1).value = lst
        sht.autofit()
        sht.cells(1, nl.getcol("aio") + 1).column_width = 20
        sht.autofit("r")
        xwu.freeze(sht.cells(2, 2))
        xwu.maketable(xwu.usedrange(sht), "Wgts")
        #sht.api.ListObjects.Add(1, xwu.usedrange(sht).api, None, 1).Name = "Wgts"

class _SMLogWtr(object):
    """
    the log writter for ShpMkr/ShpImptr
    """
    _wgt_threshold = 0.03

    def __init__(self, cnsvc, sns=None):
        self._cnsvc = cnsvc
        self._sns = sns or _SMSns()

    def write(self, wb, logs):
        """ write the logs to wb """
        for sn, log, wtr in zip(("sn_err", "sn_warn"),
                                 self._sns.err_warn(logs),
                                 (self._write_err, self._write_warn)):
            if not log:
                continue
            sht = wtr(self._sns.get(wb, sn), log)
            if not sht:
                continue
            xwu.freeze(sht.range("D2"))

    def _write_err(self, sht, logs):
        """
        write errs
        """
        '''
        nls = xwu.NamedRanges(sht.range(1, 1))
        if nls:
            nls = [x for x in nls]
            ttl = nls[0].colnames
            vvs = [nl.data for nl in nls]
        else:
        '''
        ttl, vvs = "location,type,msg".split(","), []
        for mp in logs:
            vvs.append(
                tuple("%s" % mp.get(x) if x != "type" else self._sns.
                      get_error(mp.get(x))[1] for x in ttl))
        # supress the duplicates
        vvs = list({"%s%s%s" % x: x for x in vvs}.values())
        vvs.insert(0, ttl)
        sht.range(1, 1).value = vvs
        sht.autofit("c")
        return sht

    def _write_warn(self, sht, logs):
        """
        write warnings with different encoder, different title
        """
        encs = {
            "wc_wgt": self._enc_wgt,
            "wc_ack": self._enc_ack,
            "wc_qty": self._enc_qty,
            "wc_smp": self._enc_smp
        }
        ridx, ttl = 0, "cstname,ordertype,jono,styno,location,type,msg".split(",")
        rmpfx = lambda x: (x[1:] if x[0] == "'" else x) if isinstance(x, str) else x

        jns = set(rmpfx(mp.get("jono")) for mp in logs)
        with self._cnsvc.sessionctx():
            jomp = self._cnsvc.getjos(jns)[0]
            jomp = {x.name.value: x for x in jomp}
            for mp in logs:
                jn = rmpfx(mp.get("jono"))
                if jn in jomp:
                    jn = jomp[jn]
                    mp["cstname"], mp["styno"], mp[
                        "ordertype"] = jn.customer.name.strip(
                        ), jn.style.name.value, jn.ordertype
                else:
                    mp["cstname"], mp["styno"], mp["ordertype"] = (NA,) * 3
        logs = sorted(
            logs, key=lambda x: [x.get(y) for y in "type cstname ordertype styno jono".split()])
        # hdrs holds the table ranges(if there is). Use table for better user experience
        jn, hdr, hdrs = None, None, []
        for mp in logs:
            jomp = encs.get(mp["type"])
            if not jomp:
                logger.debug("warning encoder for (%s) not found, default used",
                             mp["type"])
                jomp = self._enc_default
            vvs = jomp(mp)
            if vvs is None:
                #weight error of new sample won't be shown
                if mp["type"] == "wc_wgt":
                    continue
                vvs = []
            # write a title row for each warning category
            if mp["type"] != jn:
                jn = mp["type"]
                if hdr:
                    hdrs.append(hdr.expand("table"))
                ridx += 1
                sht.range(ridx, 1).value = ttl + jomp(None)
                hdr = sht.range(ridx, 1).expand("right")
                # when using table, high-lighting is not necessary
                # _hl(hdr, 37)
            ridx += 1
            sht.range(ridx, 1).value = [
                "%s" % mp.get(x) if x != "type" else self._sns.get_error(
                    mp.get(x))[1] for x in ttl
            ] + vvs
        sht.autofit("c")
        if True:
            if hdr:
                hdrs.append(hdr.expand("table"))
            for hdr in reversed(hdrs):
                xwu.maketable(hdr)
        return sht

    @classmethod
    def _enc_wgt(cls, opts):
        """ encoder for wgt error """
        # don't report weight error for new sample
        if opts is None:
            return "主成色,副成色,配件,备注".split(",")
        if opts["ordertype"] == 'N':
            return None
        wgts, flag, vvs = tuple(x.wgts for x in opts["objs"]), False, []
        for wgt in zip(*wgts):
            wgtexp, wgtact = tuple(_adjwgtneg(x.wgt) if x else 0 for x in wgt)
            if wgtact or wgtexp:
                wdf = (wgtact - wgtexp) / wgtexp if wgtexp else NA
                pfx = "%4.2f-%4.2f" % (wgtact, wgtexp)
                if wgtexp:
                    if abs(wdf) <= cls._wgt_threshold:
                        # vvs.append("OK")
                        vvs.append(pfx + "(OK)")
                    else:
                        flag = flag or wdf > cls._wgt_threshold
                        vvs.append(pfx + "(%s%%)" % ("%4.2f" % (wdf * 100.0)))
                else:
                    if not flag:
                        flag = True
                    vvs.append(pfx + "(%s)" % NA)
            else:
                vvs.append("'-")
        if flag:
            vvs.append("金控有误")
        return vvs

    @classmethod
    def _enc_ack(cls, opts):
        """ encoder for ack error """
        if opts is None:
            return "state inv ack date file".split()
        objs = opts["objs"]
        if isinstance(objs, dict):
            prs = tuple(objs[x] for x in "ack inv".split())
            prs = "New" if not prs[0] else "%4.2f" % (
                (float(prs[1]) - float(prs[0])) / float(prs[0]) * 100)
            return [
                prs,
            ] + [objs[x] for x in "inv ack date file".split()]
        return None

    @classmethod
    def _enc_qty(cls, opts):
        if opts is None:
            return [
                "剩余数量",
            ]
        objs = opts["objs"]
        if objs and len(objs) == 2:
            return [
                objs[0] - objs[1],
            ]
        return None

    @classmethod
    def _enc_smp(cls, opts):
        """
        sample, also provide the netwgt/metalwgt data
        """
        if opts is None:
            return ["连石重", "金重", "链重"]
        prdwgt = opts["objs"]
        rc = [round(prdwgt.metal_stone, 2)]
        lst = prdwgt.metal
        if lst:
            if len(lst) == 1:
                rc.append(round(lst[0].wgt, 2))
            else:
                rc.append(";".join([str(x) for x in lst if x.wgt > 0]))
        else:
            rc.append(0)
        lst = prdwgt.chain
        rc.append(0 if not lst else str(lst))
        return rc

    @classmethod
    def _enc_default(cls, opts):
        """ default encoder, just show the errmsg """
        return [] if opts is None else None


class ShpImptr():
    """
    import the shipment data(C1/PAJ/C2) into workflow system
    also generate BC/MMimport data for BCsystem and HK system
    @param cnsvc: A CNSvc instance
    @param hksvc: An HKSvc instance
    @param bcsvc: An BCSvc instance
    @param(optional) cache: a cache db for PajBom Cache save/query
    """

    def __init__(self, cnsvc, hksvc, bcsvc, **kwds):
        self._cnsvc, self._hksvc, self._bcsvc = cnsvc, hksvc, bcsvc
        self._groupsampjo = False
        self._sns = _SMSns()
        self._cache = kwds.get("cache")

    @classmethod
    def exacthdr(cls, sht):
        """ extract data/jmp#/n# from header """
        pts = (sht.api.pagesetup.leftheader, sht.api.pagesetup.centerheader,
               sht.api.pagesetup.rightheader)
        pts = [xwu.escapetitle(pt) for pt in pts]
        ptn = re.compile(r"\d+")
        pts[0] = date(*tuple(int(x) for x in ptn.findall(pts[0])))
        return pts

    def doimport(self, fn=None, **options):
        """
        options:
            verbose = True: show the errors or the complete state
        """
        if self._cache:
            options["cache"] = self._cache
        sm, verbose = ShpMkr(self._cnsvc, self._hksvc, self._bcsvc,
                             **options), options.get("verbose")
        wb, vdrname = sm.build_rpts(fn)
        if not wb:
            return None
        if not vdrname:
            # a _BOM_COMITTED_ item
            logger.info("totally %s bom check records were committed" % wb[len("_BOM_COMITTED_"):])
            return None
        # build_rpt checked, check again? because build_rpt do db check
        # _check_rpt_error only check data on rpt sheet
        mp = self._check_rpt_error(wb)
        ttl, errs = tuple(mp.get(x) for x in "ttl errs".split())
        if not ttl:
            if errs is None:
                mp["errs"] = errs = []
            sht, nlhdr = tuple(mp.get(x) for x in "sht nlhdr".split())
            sht.activate()
            sht.range(xwu.usedrange(sht).last_cell.row,
                      mp["cidxqty"] + 1).select()
            wb.app.visible = True
            msg = "文件=%s,\n日期=%s，落货纸号=%s\n总件数=%s，总重量=%s" % \
                (wb.name, nlhdr.date.strftime("%Y-%m-%d"), nlhdr.jmpno,
                mp["ttlqty"], str(round(mp["ttlwgt"], 2)))
            msg = messagebox.askyesno("确定要将以下资料导入落货系统?", msg)
            if not msg:
                return None
            msg = self._do_db(mp)
            if not msg:
                self._write_mm_in(wb, nlhdr, vdrname)
            else:
                errs.append("程序错误(%s)" % msg)
        if ttl:
            if errs:
                _SMLogWtr(self._cnsvc, self._sns).write(wb, errs)
                self._sns.get(wb, "sn_err").activate()
                ttl = ("检测到错误", "详情请参考Excel")
        xwu.appswitch(_appmgr.acq()[0], {"visible": True})
        if ttl and verbose:
            easymsgbox(messagebox.showinfo, ttl[0], ttl[1])
        return ttl or wb

    def _check_rpt_error(self, wb):
        """
        Check if there is still critical inside sheet("rpt")
        the returned dict must returned at least below columns:
            ttl errs
        then when the ttl is empty(no critical error), should also return
            cidxqty jns nlhdr nls sht ttlqty ttlwgt
        """
        ttl, var, mpx = (None,) * 3
        if not wb:
            xwu.appswitch(_appmgr.acq()[0], True)
            ttl = ("文件错误", "文件有误或不存在")
        else:
            sht, errs = self._sns.get(wb, "sn_rpt", False), []
            if not sht:
                ttl = ("文件错误", "关键页(%s)不存在" % self._sns.get_error("sn_rpt")[1])
        if not ttl:
            nlhdr = NamedList(list2dict("date,jmpno,iono"), self.exacthdr(sht))
            xwu.appswitch(wb.app, {"visible": True})
            if self._isimported(nlhdr.jmpno):
                var = "JMP#(%s)已导入" % nlhdr.jmpno
                self._sns.eap(errs, "_记录重复_", "_all_", "ec_jmp", var)
                ttl = ("错误", var)
        if not ttl:
            if nlhdr.jmpno[0] != "J":
                var = "落货纸#(%s)应该以J开头" % nlhdr.jmpno
                self._sns.eap(errs, "_落货纸错误_", "_all_", "ec_jmp", var)
            var = date.today() - nlhdr.date
            if var.days < 0 or var.days > 20:
                var = ("来至未来(%s)的资料" if var.days < 0 else
                       "太早以前(%s)的资料") % nlhdr.date.strftime("%Y-%m-%d")
                self._sns.eap(errs, "_日期错误_", "_all_", "ec_date", var)
            mpx = self._check_rpt_detail(sht, errs)
            # don't send below locals to the caller
        del var, wb
        mp = dict(x for x in locals().items() if x[0] not in (
            "mpx",
            "self",
        ) and x[1] is not None)
        if mpx:
            mp.update(mpx)
        return mp

    def _check_rpt_detail(self, sht, errs):
        """ check sheet(rpt)'s detail data row by row """
        nls = [
            x for x in xwu.NamedRanges(
                sht.range(1, 1),
                name_map={
                    "jono": "工单",
                    "qty": "件数",
                    "qtyleft": "此次,",
                    "running": "run#",
                    "karat": "成色"
                })
        ]
        ttlqty, ttlwgt, jns, var, cidxqty = 0, 0, set(), 0, 0
        for nl in nls:
            if not nl.jono:
                break
            if not cidxqty:
                cidxqty = nl.getcol("qty")
            if nl.qtyleft < 0:
                self._sns.eap(errs, nl.jono, nl.jono, "ec_qty", "数量不足")
            if not nl.wgt or nl.wgt < 0:
                ttl = ("重量错误", "存在重量不合规记录")
                if nl.wgt < 0:
                    self._sns.eap(errs, nl.jono, nl.jono, "ec_wgt_not_sure",
                                  "重量不确定，请人工复核")
                else:
                    self._sns.eap(errs, nl.jono, nl.jono, "ec_wgt_missing",
                                  "欠重量资料")
            if nl.qty:
                var = nl.qty
                ttlqty += var
            ttlwgt += nl.wgt * var
            jns.add(nl.jono)
        del var
        return dict(x for x in locals().items() if x[1] is not None)

    def _do_db_prep(self, mp):
        """ prepare data from db for inserting """
        maMap, mmMap, gdMap, updjos = {}, {}, {}, {}
        refid = refno = None
        nlhdr, errs = tuple(
            mp.get(x) for x in "nlhdr errs".split())
        jos = self._cnsvc.getjos(mp["jns"])[0]
        joqls = {x.name.value: x.qtyleft for x in jos}
        jos = {x.name.value: x for x in jos}
        for nl in mp["nls"]:
            jn = nl.jono
            if not jn:
                break
            jn = JOElement(nl.jono).value
            nl.jono = jn
            karat, jo, running = int(nl.karat), jos.get(jn), nl.running
            nl.karat = karat
            if running:
                if jn not in updjos and jo.running != running:
                    updjos[jn] = jo
                    jo.running = running
                    jo.lastupdate = datetime.today()
                if karat not in maMap:
                    if not refid:
                        refid, refno, mmid = self._lastrefid(), self._nextrefno(
                        ), self._lastmmid()
                    maMap[karat] = var = MMMa()
                    refid += 1
                    var.id, var.name, var.karat, var.refdate, var.tag = refid, refno, karat, nlhdr.date, 0
            var = jo.id if self._groupsampjo else random.randint(0, 9999999)
            if nl.qty:
                mm = mmMap.get(var)
                if not mm:
                    mmMap[var] = mm = MM()
                    mmid += 1
                    mm.id, mm.jsid, mm.name, mm.refid, mm.qty = mmid, jo.id, nlhdr.jmpno, refid, 0
                mm.qty += nl.qty
                # don't change the jo.qtyleft directly because this might cause a double-substract by both me and mm.insert trigger
                var = joqls[jn] - nl.qty
                if var < 0:
                    self._sns.eap(errs, nl.jono, nl.jono, "ec_qty", "数量不足")
                else:
                    mm.tag = 0 if var else 4
                joqls[jn] = var
            var = "%d/%d" % (mm.id, karat)
            gd = gdMap.get(var)
            if not gd:
                gdMap[var] = gd = MMgd()
                gd.id, gd.karat, gd.wgt = mm.id, nl.karat, 0
            gd.wgt += nl.wgt * (nl.qty if nl.qty else mm.qty)
        return {
            "errs": errs,
            "maMap": maMap,
            "mmMap": mmMap,
            "gdMap": gdMap,
            "updjos": updjos,
            "jos": jos
        }

    def _do_db(self, mp):
        app, tk = _appmgr.acq()
        try:
            xwu.appswitch(app, {"visible": False})
            with ResourceCtx((self._cnsvc.sessmgr(),
                              self._hksvc.sessmgr())) as curs:
                mpx = self._do_db_prep(mp)
                jos = mpx.get("errs")
                if jos:
                    return jos
                # update those tag = 0 to max + 1
                cncmds, jos = [], curs[0].query(MMMa).filter(
                    MMMa.tag == 0).all()
                if jos:
                    cncmds.append(jos)
                    hkjos = curs[0].query(func.max(MMMa.tag)).first()
                    hkjos = (hkjos[0] if hkjos else 0) + 1
                    for ma in jos:
                        ma.tag = hkjos
                # the insert sequence, ma -> mm -> gd -> jo
                cncmds.extend([
                    tuple(y.values())
                    for y in (
                        mpx.get(x) for x in "maMap mmMap gdMap updjos".split())
                    if y
                ])
                # HK jo running update
                hkjos = self._hksvc.getjos(mp.get("jns"),
                                        JOhk.running == 0)[0] or []
                hkcmds, jos = [], mpx.get("jos")
                for y, x in ((x, jos.get(x.name.value), ) for x in hkjos):
                    if not (x and x.running):
                        continue
                    y.running = x.running
                    hkcmds.append(y)
                for x in cncmds:
                    for y in x:
                        curs[0].add(y)
                    curs[0].flush()
                curs[0].commit()
                for x in hkcmds:
                    curs[1].add(x)
                curs[1].commit()
        except:
            pass
        finally:
            _appmgr.ret(tk)
        return None

    def _write_mm_in(self, wb, nlhdr, vdrname, new_wb=True):
        """
        write the data for mm import to a new workbook or a sheet inside existing wb
        """
        if new_wb:
            wb1 = wb.app.books.add()
            sht = wb1.sheets[0]
            sht.name = "mmimptr"
            sht.range(1, 1).value = ((
                nlhdr.iono,
                nlhdr.jmpno,
                nlhdr.date,
            ),)
            sht.autofit("c")
            fldr, pfx, ext, cntr = path.dirname(
                wb.fullname), nlhdr.date.strftime(
                    "%Y%m%d"
                ), ".xls" if wb.app.version.major < 12 else ".xlsx", 0
            pfx0 = pfx
            if vdrname and vdrname == "c1":
                pfx0 = "C1 " + pfx0
            while True:
                fn = path.join(fldr, pfx + ext)
                if path.exists(fn):
                    cntr += 1
                    pfx = "%s_%d" % (pfx0, cntr)
                else:
                    break
            wb1.save(fn)
            wb = wb1
        else:
            sht = self._sns.get(wb, "mmimptr")
            sht.range(1, 1).value = ((nlhdr.iono, nlhdr.jmpno, nlhdr.date))
            sht.autofit("c")
        return wb

    def _isimported(self, jn):
        """
        check if given JO# has been imported
        """
        with self._cnsvc.sessionctx() as cur:
            cnt = cur.query(func.count(MM.name)).filter(MM.name == jn).first()
            return cnt[0] > 0

    def _nextrefno(self):
        """
        get next available ref# for mmma
        """
        pf, pl = "J", 7
        with self._cnsvc.sessionctx() as cur:
            name = cur.query(func.max(MMMa.name)).filter(MMMa.tag == 0).first()
            if not name[0]:
                mtag = cur.query(func.max(MMMa.tag)).first()[0]
                name = cur.query(func.max(
                    MMMa.name)).filter(MMMa.tag == mtag).first()
        name = name[0] if name else pf & "0"
        je = JOElement(name)
        je.digit += 1
        return je.alpha + ("%%0%dd" % (pl - len(je.alpha))) % je.digit

    def _lastmmid(self):
        """
        get the last (max) mmid from database
        """
        with self._cnsvc.sessionctx() as cur:
            mmid = cur.query(func.max(MM.id)).first()
        return mmid[0] if mmid else 0

    def _lastrefid(self):
        """
        return the last (max) refid of mmma from database
        """
        with self._cnsvc.sessionctx() as cur:
            x = cur.query(func.max(MMMa.id)).first()[0]
        return x if x else 0


class _SMIOHlpr(object):
    """ helper class for ShpMkr to get/set the IO sheet """

    def __init__(self, *args):
        self._app, self._sts, self._shp_date, self._vdrname = args
        self._shtio, self._itio, self._ridx = (None,) * 3

    def get(self):
        """
        fetch IO data based on the vendor name to a dict
        The returned dict contains below keys:
            .n#,date,jmp#,"maxrun#"
        """
        fn = self._sts.get(triml("Shipment.IO")).value
        wbio, mp = self._app.books.open(fn), {}
        self._shtio = wbio.sheets["master"]
        nls = [x for x in xwu.NamedRanges(self._shtio.range(1, 1))]
        self._itio, self._ridx = nls[-1], len(nls) + 2
        je = JOElement(self._itio["n#"])
        mp["n#"], mp["date"] = "%s%d" % (je.alpha, je.digit + 1), self._shp_date
        pfx = self._shp_date.strftime("%y%m%d")
        pfx = 'J' + (pfx if self._vdrname == "paj" else pfx[1:])
        existing = [
            x["jmp#"]
            for x in nls[-20:]
            if x["jmp#"] and x["jmp#"].find(pfx) == 0
        ]
        if existing:
            if self._vdrname != "paj":
                logger.debug(
                    "%s should not have more than one shipment in one date" %
                    self._vdrname)
                return None
            sfx = "%d" % (int(max(existing)[-1]) + 1)
        else:
            sfx = "1" if self._vdrname == "paj" else trimu(self._vdrname)
        mp["jmp#"], idx = pfx + sfx, -1
        for idx in range(len(nls) - 1, 0, -1):
            jn = nls[idx]["jmp#"]
            if not jn:
                continue
            if (jn.find("C") >= 0) ^ (self._vdrname == "paj"):
                break
        mp["maxrun#"] = int(nls[idx]["maxrun#"])
        return mp

    def update(self, iorst):
        """ send the updated result to the last row of IO """
        for knv in iorst.items():
            self._shtio.range(self._ridx,
                              self._itio.getcol(knv[0]) + 1).value = knv[1]


class _SMBCHdlr(object):
    '''
    class help to handle the bc issue for shipment
    '''

    def __init__(self, bcsvc, hksvc, cnsvc, sns):
        self._bcsvc, self._hksvc, self._cnsvc, self._sns = bcsvc, hksvc, cnsvc, sns


    def write(self, wb, shplst, newrunmp, shp_date):
        """
        create a bc template
        """
        lsts, rcols = [], "lymd,lcod,styn,mmon,mmo2,runn,detl,quan,gwgt,gmas,jobn,ston,descn,desc,rem1,rem2,rem3,rem4,rem5,rem6,rem7,rem8".split(
            ",")
        dups = len("rem")
        dups = [int(x[dups:]) for x in rcols if x.find("rem") == 0]
        rems = (
            min(dups),
            max(dups) + 1,
        )
        bc_by_sty, bc_by_jn = self._get_data(shplst, shp_date, rems)
        dups, hls, lymd = {}, [], shp_date.strftime("%Y%m%d %H:%M%S")

        lsts.append(rcols)
        nl = NamedList(list2dict(rcols))
        shplst = sorted(
            shplst,
            key=
            lambda mpx: "A%06d%s" % (mpx["running"], mpx["jono"]) if mpx["running"] else "B%06d%s" % (0, mpx["jono"])
        )
        for it in shplst:
            jn = it["jono"]
            if jn in dups:
                continue
            pfx = "XX" if jn not in newrunmp else ""
            dups[jn], styno = 1, it["styno"]
            flag, bc, rmks = self._select(bc_by_jn, bc_by_sty, jn,
                                                   styno, rems)
            nl.setdata([None] * len(rcols))
            nl.lymd, nl.lcod, nl.styn, nl.mmon = lymd, styno, styno, "'" + lymd[
                2:4]
            nl.mmo2, nl.runn, nl.detl = lymd[4:6], "'%d" % it["running"] if it[
                "running"] else NA, it["cstname"]
            nl.quan, nl.jobn = it["qty"], "'" + jn
            nl.descn = pfx + ("SKU一致:" if flag else "") + it["description"]
            prdwgt = it["mtlwgt"]
            nl.gmas, nl.gwgt = prdwgt.main.karat, "'" + str(prdwgt.main.wgt)
            if not bc:
                it["_raw_data"] = nl.data
                bc = self._bcsvc.build_from_jo(jn, self._hksvc, self._cnsvc, it)
                if not bc:
                    nl.ston, nl.desc = "--", "TODO"
                else:
                    #append a mark in description for REF
                    bc.descn = 'NEW:' + bc.descn
                    nl.setdata(bc.data)
                    for idx in range(*rems):
                        rmk = getattr(bc, "rem%d" % idx)
                        if not rmk:
                            continue
                        bc["rem%d" % idx] = "'" + rmk
            else:
                nl.ston, nl.desc = bc.ston, bc.desc
                rmks = (getattr(bc, "rem%d" % y) for y in range(*rems))
                rmks = [y.strip() for y in rmks if y]
            nrmks = []
            for x in ((prdwgt.aux, "*%s %4.2f"), (prdwgt.part, "*%sPTS %4.2f")):
                if x[0] and x[0].karat:
                    nrmks.append(x[1] % (karatsvc.getkarat(x[0].karat).name,
                                         _adjwgtneg(x[0].wgt)))
            if prdwgt.part and prdwgt.part.karat:
                wgt = prdwgt.part.wgt
                if wgt < 0:
                    hls.append((len(lsts), nl.getcol("rem%d" % len(nrmks))))
                else:
                    if prdwgt.part.karat == 925:
                        if wgt < 1.0 or wgt > 2.0:
                            hls.append((len(lsts),
                                        nl.getcol("rem%d" % len(nrmks))))
                    else:
                        if wgt < 0.3 or wgt > 1.0:
                            hls.append((len(lsts),
                                        nl.getcol("rem%d" % len(nrmks))))
            cn = len(nrmks) + len(rmks) - rems[1] + 1
            if cn > 0:
                rmks[-cn - 1] = ";".join(rmks[-cn - 1:])
                nrmks.extend(rmks[:-cn])
            else:
                nrmks.extend(rmks)
            for idx, rmk in enumerate(nrmks):
                nl["rem%d" % (idx + 1)] = "'" + rmk
            lsts.append(nl.data)
        sht = self._sns.get(wb, "sn_bc")
        sht.range(1, 1).value = lsts
        self._fmt_sht(sht, hls)
        return hls

    @staticmethod
    def _fmt_sht(sht, hls):
        sht.autofit()
        if hls:
            rng = sht.range(1, 1)
            for x in hls:
                _hl(xwu.offset(rng, x[0], x[1]), 6)
        xwu.maketable(xwu.usedrange(sht), "BCData")
        # hide cols A:H, free to M column
        sht.range("A:H").api.EntireColumn.Hidden = True
        xwu.freeze(sht.range("M2"), False)


    def _get_data(self, shplst, shp_date, rems):
        """
        get candidates for writing bc report
        """
        # sometimes different customer has same SKU#, but, their orderid will be different, so orderma table don't need to be invoked
        refjo, refpo = aliased(JOhk), aliased(POItem)
        with self._hksvc.sessionctx() as cur:
            dt = shp_date - timedelta(days=720)
            jes = set(JOElement(x["jono"]) for x in shplst)
            logger.debug("begin to select same sku items for BC")
            t0 = clock()
            # below query runs quite fast(0.01s) under common client, but very slow(11+s) in sqlalchemy, don't know the reason, maybe check fry pyodbc/sqlchemy
            q = Query([JOhk.name, refjo.running]).join(
                (refjo, and_(JOhk.orderid == refjo.orderid, JOhk.id != refjo.id)), (POItem, JOhk.poid == POItem.id),
                (refpo,
                 and_(POItem.skuno != '', refpo.id == refjo.poid,
                      refpo.skuno == POItem.skuno))).filter(
                          and_(POItem.id > 0, refjo.createdate > dt, refjo.createdate < shp_date))
            lst = []
            for arr in splitarray(jes, 20):
                qx = q.filter(jesin(arr, JOhk))
                lst0 = qx.with_session(cur).all()
                if lst0:
                    lst.extend(lst0)
            logger.debug("using %fs to fetch %d records for above action" %
                         (clock() - t0, len(lst)))
            josku = {x[1]: x[0].value for x in lst if x[1] > 0} if lst else {}

        bc_by_jn = self._bcsvc.getbcs([x for x in josku]) or {}
        # get the longest rmks for each sku# because manually input might sometimes
        # missing some data
        bc_by_styn, lens = {}, {}
        for x in bc_by_jn:
            q = self._extr_rmks(x, rems)
            q = (len(q), sum([len(x) for x in q]))
            qx = josku.get(int(x.runn))
            if q > lens.get(qx, (0, 0)):
                bc_by_styn[qx], lens[qx] = x, q
        bc_by_jn = bc_by_styn
        #no good candidates, fetch by sty#
        bc_by_styn = {
            x.get("styno") for x in shplst if x["jono"] not in bc_by_jn
        }
        bc_by_styn, bcs = {}, self._bcsvc.getbcs(bc_by_styn, True)
        if bcs:
            for it in bcs:
                bc_by_styn.setdefault(it.styn, []).append(it)
        for x in bc_by_styn:
            bc_by_styn[x] = sorted(
                bc_by_styn[x], key=lambda x: x.runn, reverse=True)
        return bc_by_styn, bc_by_jn

    @classmethod
    def _select(cls, bc_by_jn, bc_by_sty, jn, styno, rems):
        """
        select a bc record based on provided arguments. First return the extract,
        if not, select the one with same karat and more remarks
        @return flag: True if the extract JO# is found
        @return bc: an BCSystem instance
        @return rmks: the remarks fields as an list
        """
        bc, rmks = bc_by_jn.get(jn), []
        flag = bool(bc)
        if not flag:
            bcs = bc_by_sty.get(styno)
            if bcs:
                for bcx in bcs[:10]:
                    if not samekarat(jn, bcx.jobn):
                        continue
                    mc0 = cls._extr_rmks(bcx, rems)
                    if len(mc0) > len(rmks):
                        rmks, bc = mc0, bcx
                if not bc:
                    bc, rmks = bcs[0], cls._extr_rmks(bcs[0], rems)
        return flag, bc, rmks

    @staticmethod
    def _extr_rmks(bc, rems):
        return [x for x in (getattr(bc, "rem%d" % y).strip() for y in range(*rems)) if x]
