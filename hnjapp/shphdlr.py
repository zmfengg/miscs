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
from os import path
from time import clock
from tkinter import filedialog, messagebox

from sqlalchemy import and_, func
from sqlalchemy.orm import Query, aliased
from xlwings.constants import (BorderWeight, Constants, FormatConditionOperator,
                               FormatConditionType, LineStyle)
from xlwings.utils import col_name

from hnjapp.c1rdrs import C1InvRdr
from hnjcore import JOElement, samekarat
from hnjcore.models.cn import MM, MMgd, MMMa
from hnjcore.models.hk import JO as JOhk
from hnjcore.models.hk import Orderma, PajAck, POItem
from hnjcore.utils.consts import NA
from utilz import (NamedList, NamedLists, ResourceCtx, easydialog, easymsgbox,
                   list2dict, splitarray, triml, trimu, xwu, deepget, karatsvc, getfiles)

from .common import _logger as logger
from .dbsvcs import jesin
from .pajcc import cmpwgt
from .pajrdrs import PajBomHhdlr, PajShpHdlr

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


class ShpSns(object):

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

    def new_err(self, jn, loc, etype, msg, objs=None):
        """ new a dict holding the key error info """
        return {
            "jono": "'" + jn,
            "location": "'" + loc,
            "type": etype,
            "msg": msg,
            "objs": objs
        }

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

    def _errandwarn(self, errlst):
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
    """
    _mergeshpjo = False
    _vdrname = _nsofsts = None
    _shpsns = ShpSns()
    _nsofn = None
    
    def __init__(self, cnsvc, hksvc, bcsvc, **kwds):
        self._cnsvc, self._hksvc, self._bcsvc = cnsvc, hksvc, bcsvc
        self._nsofn = kwds.get("nsofn")
        # debug mode, qtyleft/running should be reset so that I can
        # generate proper report
        self._debug = kwds.get("debug", False)

    def _pajfldr2file(self, fldr):
        """ group the folder into one target file. If target file already exists,
        do date check
        @return : filename if succeeded
                  -1 if file expired
                  None if unexpected error occured
        """
        if not fldr:
            fldr = easydialog(
                filedialog.Directory(
                    title="Choose folder contains all raw files from PAJ"))
            if not path.exists(fldr):
                return
        sts = self._nsofsettings()
        tarfldr, tarfn = path.dirname(sts.get("shp.template").value), None
        fns = getfiles(fldr, ".xls")
        ptn = re.compile(r"^HNJ \d+")
        for fn in fns:
            if ptn.search(path.basename(fn)):
                sd = PajShpHdlr.get_shp_date(fn)
                if sd:
                    tarfn = "HNJ %s 出货明细" % sd.strftime("%Y-%m-%d")
                    break
        if not tarfn:
            return None
        sts = getfiles(tarfldr, tarfn)
        if sts:
            tarfn = sts[0]
            tdm = path.getmtime(tarfn)
            fds = [path.getmtime(x) for x in getfiles(fldr)]
            fds.append(path.getmtime(fldr))
            if max(fds) > tdm:
                messagebox.showwarning("文件过期",
                                       "%s\n已过期,请手动删除或更新后再启动本程序" % tarfn)
                app, kxl = _appmgr.acq()
                wb = app.books.open(tarfn)
                return None
            else:
                logger.debug("result file(%s) already there" % tarfn)
                return tarfn
        if len(fns) == 1:
            return fns[0]

        app, kxl = _appmgr.acq()
        wb = app.books.add()
        new_shts = [x for x in wb.sheets]
        before_sht = wb.sheets[0]
        for fn in fns:
            if fn.find("对帐单") >= 0:
                continue
            wbx = xwu.safeopen(app, fn)
            try:
                for sht in wbx.sheets:
                    if sht.api.visible == -1 and xwu.usedrange(sht).size > 1:
                        sht.api.Copy(Before=before_sht.api)
            finally:
                wbx.close()
        for x in new_shts:
            x.delete()
        if tarfn:
            wb.save(path.join(tarfldr, tarfn))
            tarfn = wb.fullname
            logger.debug("merged file saved to %s" % tarfn)
            wb.close()
        _appmgr.ret(kxl)
        return tarfn

    def read_c1(self, sht, args):
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
            sht.range("1:%d" % ridx).api.entirerow.delete()
        mp = sht.name
        its = C1InvRdr.read_c1(sht)
        if not its:
            logger.debug("no valid data in sheet(%s)" % mp)
            return (None,) * 2
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
        if mp:
            mp["shpdate"] = args.get("shpdate")
        return mp, errs

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
            for mp in shplst:
                jn = mp["jono"]
                jncmp[jn] -= 1
                jo = jos.get(jn)
                if not jo:
                    self._eap(errlst, jn, mp["location"], "ec_jn",\
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
                    self._eap(errlst, mp["jono"], mp["location"], "ec_qty",\
                        s0, (jo.qtyleft + mp["qty"], mp["qty"]))
                    mp["errmsg"] = s0
                elif jo.qtyleft > 0 and not jncmp[jn]:
                    s0 = "数量有余"
                    self._eap(errlst, mp["jono"], mp["location"], "wc_qty",\
                        s0, (jo.qtyleft + mp["qty"], mp["qty"]))
                    mp["errmsg"] = s0
                else:
                    mp["errmsg"] = ""
                jwgt = jwgtmp.get(jn)
                if not jwgt and jn not in jwgtmp:
                    jwgt = self._hksvc.getjowgts(jn)
                    if not jwgt:
                        jwgt = None
                    jwgtmp[jn] = jwgt
                if not cmpwgt(jwgt, mp["mtlwgt"]):
                    haswgt = bool(
                        [x for x in mp["mtlwgt"].wgts if x and x.wgt > 0])
                    jn = [errlst, mp["jono"], mp["location"], ]
                    if haswgt:
                        if jo.ordertype != "O":
                            jn = None
                        else:
                            jn.extend(("wc_wgt", "重量不符"))
                    else:
                        jn.extend(("ec_wgt_missing", "欠重量资料"))
                    if jn:
                        jn.append((jwgt, mp["mtlwgt"]))
                        self._eap(*jn)
                jn = (jo.karat, mp["mtlwgt"].main.karat, )
                if jn[0] != jn[1]:
                    self._eap(errlst, mp["jono"], mp["location"], "ec_karat", "主成色与工单成色不一致", jn)
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
                errlst.append(self._ne(jn, jn, "wc_smp", s0, shpwgts[jn]))
        logger.debug("using %fs for above action" % (clock() - t0))

    def _eap(self, errlst, *args):
        errlst.append(self._ne(*args))
        type_name = args[2]
        if type_name == "wc_wgt":
            #check if the weight is too critical
            jwgt, shpwgt = args[4]
            if not cmpwgt(jwgt, shpwgt, 50):
                type_name = [x for x in args]
                type_name[2:] = "ec_wgt", "重量偏离指定值50%以上", None
                errlst.append(self._ne(*type_name))

    def _ne(self, *args):
        return self._shpsns.new_err(*args)

    def _check_db_inv(self, shplst, invmp, errlst, jns):
        if not invmp:
            self._eap(errlst, NA, "_all_", "ec_inv_none", "不应无发票资料")
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
                x[0].name.value: tuple(getattr(x[1], y) for y in tmp)
                for x in q
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
                    self._eap(errlst, x.jono, x.jono, "wc_inv_diff",\
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
                    self._eap(errlst, x["jono"],
                            "Inv(%s),JO#(%s)" % (x["inv"].invno, x["jono"]),
                            "wc_inv_qty", "工单(%s)有发票(%s)无落货" %
                            (x["jono"], x["inv"].invno), None)
                else:
                    self._eap(errlst, x["jono"], x["jono"],
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
                    self._eap(errlst, jn, jn, "wc_ack", nlack.pcode, {"inv": inv.uprice,"ack": nlack.uprice, "inv_mps": inv.mps, "ack_mps": nlack.mps,"file": nlack.docno, "date": nlack.date.strftime("%Y-%m-%d")})
        tmp = jns.difference(tmp.keys())
        if tmp:
            errlst.extend([self._ne(x, x, "wc_qty",\
                "工单(%s)有落货无发票" % x, None) for x in tmp])

    def _write_bc(self, wb, shplst, newrunmp, shp_date):
        """
        create a bc template
        """
        dmp, lsts, rcols = {}, [], "lymd,lcod,styn,mmon,mmo2,runn,detl,quan,gwgt,gmas,jobn,ston,descn,desc,rem1,rem2,rem3,rem4,rem5,rem6,rem7,rem8".split(
            ",")
        refjo, refpo, refodma = aliased(JOhk), aliased(POItem), aliased(Orderma)
        rems, nl, hls = [999, 0], NamedList(list2dict(rcols)), []
        for x in nl.colnames:
            if not x.find("rem"):
                idx = int(x[len("rem"):])
                if idx < rems[0]:
                    rems[0] = idx
                if idx > rems[1]:
                    rems[1] = idx
        rems[1] += 1
        with self._hksvc.sessionctx() as cur:
            dt = datetime.today() - timedelta(days=365)
            jes = set(JOElement(x["jono"]) for x in shplst)
            logger.debug("begin to select same sku items for BC")
            t0 = clock()
            q = Query([JOhk.name, func.max(refjo.running)]).join(
                (refjo, JOhk.id != refjo.id), (POItem, JOhk.poid == POItem.id),
                (Orderma, JOhk.orderid == Orderma.id),
                (refodma, refjo.orderid == refodma.id),
                (refpo,
                 and_(POItem.skuno != '', refpo.id == refjo.poid,
                      refpo.skuno == POItem.skuno))).filter(
                          and_(POItem.id > 0, refjo.createdate > dt,
                               Orderma.cstid == refodma.cstid)).group_by(
                                   JOhk.name)
            lst = []
            for arr in splitarray(jes, 20):
                qx = q.filter(jesin(arr, JOhk))
                lst0 = qx.with_session(cur).all()
                if lst0:
                    lst.extend(lst0)
            logger.debug("using %fs to fetch %d records for above action" %
                         (clock() - t0, len(lst)))
            josku = {x[1]: x[0].value for x in lst if x[1] > 0} if lst else {}

        joskubcs = self._bcsvc.getbcs([x for x in josku])
        joskubcs = {josku[int(x.runn)]: x for x in joskubcs} if joskubcs else {}

        stynos = set(
            [x.get("styno") for x in shplst if x["jono"] not in joskubcs])
        bcs = self._bcsvc.getbcs(stynos, True)
        if bcs:
            for it in bcs:
                dmp.setdefault(it.styn, []).append(it)
        for x in dmp:
            dmp[x] = sorted(dmp[x], key=lambda x: x.runn, reverse=True)
        bcmp, dmp, lymd = dmp, {}, shp_date.strftime("%Y%m%d %H:%M%S")

        lsts.append(rcols)
        shplst = sorted(
            shplst,
            key=lambda mpx: "A%06d%s" % (mpx["running"], mpx["jono"]) if mpx["running"] else "B%06d%s" % (0, mpx["jono"])
        )
        for it in shplst:
            jn = it["jono"]
            if jn in dmp:
                continue
            pfx = "XX" if jn not in newrunmp else ""
            dmp[jn], styno = 1, it["styno"]
            bc, rmks = joskubcs.get(jn), []
            if not bc:
                bcs = bcmp.get(styno)
                if bcs:
                    # find the same karat and longest remarks as template
                    for bcx in bcs[:10]:
                        if not samekarat(jn, bcx.jobn):
                            continue
                        mc0 = [
                            x for x in [
                                getattr(bcx, "rem%d" % y).strip()
                                for y in range(*rems)
                            ] if x
                        ]
                        if len(mc0) > len(rmks):
                            rmks, bc = mc0, bcx
                    if not bc:
                        bc = bcs[0]
                        rmks = [
                            x for x in [
                                getattr(bc, "rem%d" % y).strip()
                                for y in range(*rems)
                            ] if x
                        ]
                flag = False
            else:
                flag = True
            nl.setdata([None] * len(rcols))
            nl.lymd, nl.lcod, nl.styn, nl.mmon = lymd, styno, styno, "'" + lymd[
                2:4]
            nl.mmo2, nl.runn, nl.detl = lymd[4:6], "'%d" % it["running"] if it[
                "running"] else NA, it["cstname"]
            nl.quan, nl.jobn = it["qty"], "'" + jn
            nl.descn = pfx + ("---" if flag else "") + it["description"]
            prdwgt = it["mtlwgt"]
            nl.gmas, nl.gwgt = prdwgt.main.karat, "'" + str(prdwgt.main.wgt)
            if not bc:
                nl.ston, nl.desc = "--", "TODO"
            else:
                nl.ston, nl.desc = bc.ston, bc.desc
                rmks = [
                    x for x in
                    [getattr(bc, "rem%d" % y).strip() for y in range(*rems)]
                    if x
                ]
            nrmks = []
            for x in ((prdwgt.aux, "*%s %4.2f"), (prdwgt.part, "*%sPTS %4.2f")):
                if x[0]:
                    nrmks.append(x[1] % (karatsvc.getkarat(x[0].karat).name,
                                         _adjwgtneg(x[0].wgt)))
            if prdwgt.part:
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
                nl["rem%d" % (idx + 1)] = rmk
            lsts.append(nl.data)
        sht = self._shpsns.get(wb, "sn_bc")
        sht.range(1, 1).value = lsts
        if hls:
            rng = sht.range(1, 1)
            for x in hls:
                _hl(rng.offset(x[0], x[1]), 6)
        sht.autofit()

    def _nsofsettings(self, fn=None):
        if self._nsofsts is None:
            self._nsofsts = PajNSOFRdr().readsettings(fn or self._nsofn)
        return self._nsofsts

    def _write_rpts(self, wb, shplst, newrunmp, shp_date):
        """
        send the shipment related sheets(Rpt/Err)
        """
        app = wb.app
        sts = self._nsofsettings()

        fn = sts.get(triml("Shipment.IO")).value
        wbio, iorst = app.books.open(fn), {}
        shtio = wbio.sheets["master"]
        nls = [x for x in xwu.NamedRanges(shtio.range(1, 1))]
        itio, ridx = nls[-1], len(nls) + 2
        je = JOElement(itio["n#"])
        iorst["n#"], iorst["date"] = "%s%d" % (je.alpha, je.digit + 1), shp_date
        pfx = shp_date.strftime("%y%m%d")
        if self._vdrname != "paj":
            pfx = pfx[1:]
        pfx = 'J' + pfx
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
                return
            sfx = "%d" % (int(max(existing)[-1]) + 1)
        else:
            sfx = "1" if self._vdrname == "paj" else trimu(self._vdrname)
        iorst["jmp#"], idx = pfx + sfx, -1
        for idx in range(len(nls) - 1, 0, -1):
            jn = nls[idx]["jmp#"]
            if not jn:
                continue
            if (jn.find("C") >= 0) ^ (self._vdrname == "paj"):
                break
        iorst["maxrun#"] = int(nls[idx]["maxrun#"])

        s0 = sts.get("shipment.rptmgns.%s" % self._vdrname)
        if not s0:
            s0 = sts.get("shipment.rptmgns")
        sht = self._shpsns.get(wb, "sn_rpt")
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

        s0 = sts.get("shipment.hdrs." + self._vdrname)
        if not s0:
            s0 = sts.get("shipment.hdrs")
        ttl, ns, hls = [], {}, []
        for x in s0.value.replace(r"\n", "\n").split(";"):
            y = x.split("=")
            y1 = y[1].split(",")
            ttl.append(y[0])
            if len(y1) > 1:
                ns[y1[0]] = y[0]
            sht.range(1, len(ttl)).column_width = float(y1[len(y1) - 1])
        ns["thisleft"] = "此次,"
        nl, maxr, lenttl = NamedList(list2dict(
            ttl, alias=ns)), iorst["maxrun#"], len(ttl)
        lsts, ns, hls = [
            ttl
        ], "jono,running,qty,cstname,styno,description,qtyleft,errmsg".split(
            ","), []
        shplst = sorted(
            shplst,
            key=
            lambda mpx: "A%06d%s" % (mpx.get("running") or 0, mpx["jono"])
        )
        for it in shplst:
            ttl = [""] * lenttl
            nl.setdata(ttl)
            if not it.get("running"):
                if it["jono"] not in newrunmp:
                    maxr += 1
                    it["running"], nl["running"] = maxr, maxr
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
                _hl(rng.offset(x[0] - 1, x[1]), 6)
        # the qtyleft formula
        s0, s1, s2 = col_name(nl.getcol("qty") + 1), col_name(
            nl.getcol("qtyleft") + 1), col_name(nl.getcol("thisleft") + 1)
        for idx in range(2, len(lsts) + 1):
            rng = sht.range("%s%d" % (s2, idx))
            rng.formula = "=%s%d-%s%d" % (s1, idx, s0, idx)
            rng.api.numberformatlocal = "_ * #,##0_ ;_ * -#,##0_ ;_ * " "-" "_ ;_ @_ "
            rng.api.formatconditions.add(FormatConditionType.xlCellValue,
                                         FormatConditionOperator.xlLess, "0")
            rng.api.formatconditions(1).interior.colorindex = 3

        rng = sht.range(sht.range(1, 1), sht.range(len(lsts), len(nl.colnames)))
        rng.api.borders.linestyle = LineStyle.xlContinuous
        rng.api.borders.weight = BorderWeight.xlThin

        # write sum formula at the bottom
        s0 = int(nl.getcol("qty")) + 1
        rng = sht.range(len(lsts) + 1, s0)
        rng.formula = "=sum(%s1:%s%d)" % (col_name(s0), col_name(s0), len(lsts))
        rng.api.font.bold = True
        rng.api.borders.linestyle = LineStyle.xlContinuous
        rng.api.borders.weight = BorderWeight.xlThin
        sht.range("A2:A%d" % (len(lsts) + 1)).row_height = 18
        rng = xwu.usedrange(sht).api
        rng.VerticalAlignment = Constants.xlCenter
        rng.font.name = "tahoma"
        rng.font.size = 10

        # write IOs back
        iorst["maxrun#"] = maxr
        for knv in iorst.items():
            shtio.range(ridx, itio.getcol(knv[0]) + 1).value = knv[1]

        return fn

    def _err_enc_wgt(self, opts):
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
                    if abs(wdf) <= 0.05:
                        vvs.append("OK")
                        # vvs.append(pfx + "(-)")
                    else:
                        flag = flag or wdf > 0.05
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

    def _err_enc_ack(self, opts):
        """ encoder for ack error """
        if opts is None:
            return "state inv ack date file".split()
        objs = opts["objs"]
        if isinstance(objs, dict):
            prs = tuple(objs[x] for x in "ack inv".split())
            prs = "New" if not prs[0] else "%4.2f" % ((float(prs[1]) - float(prs[0])) / float(prs[0]) * 100)
            return [prs, ] + [objs[x] for x in "inv ack date file".split()]
        return None

    def _err_enc_qty(self, opts):
        if opts is None:
            return ["剩余数量", ]
        objs = opts["objs"]
        if objs and len(objs) == 2:
            return [objs[0] - objs[1], ]
        return None

    def _err_enc_smp(self, opts):
        """
        sample, also provide the netwgt/metalwgt data
        """
        if opts is None:
            return ["净重", "金重", ]
        prdwgt = opts["objs"]
        wgts, cats = [x for x in prdwgt.wgts if x and x.wgt > 0], {}
        for x in wgts:
            cats[x.karat] = cats.get(x.karat, 0) + x.wgt
        if len(cats) == 1:
            x = "%4.2f" % iter(cats.values()).__next__()
        else:
            x = ";".join("%s=%4.2f" % (x[0], x[1]) for x in cats.items())
        return [prdwgt.netwgt, x]

    def _err_enc_default(self, opts):
        """ default encoder, just show the errmsg """
        return [] if opts is None else None

    def write_logs(self, wb, errlst):

        def _write_err(sht, logs):
            """
            write errs
            """
            nls = xwu.NamedRanges(sht.range(1, 1))
            if nls:
                ttl = nls[0].colnames
                vvs = [nl.data for nl in nls]
            else:
                ttl, vvs = "location,type,msg".split(","), []
            for mp in logs:
                vvs.append(tuple("%s" % mp.get(x) if x != "type" else self._shpsns.get_error(mp.get(x))[1] for x in ttl))
            # supress the duplicates
            vvs = list({"%s%s%s" % x: x for x in vvs}.values())
            vvs.insert(0, ttl)
            sht.range(1, 1).value = vvs
            return sht

        def _write_warn(sht, logs):
            """
            write warnings with different encoder, different title
            """
            encs = {
                "wc_wgt": self._err_enc_wgt,
                "wc_ack": self._err_enc_ack,
                "wc_qty": self._err_enc_qty,
                "wc_smp": self._err_enc_smp
            }
            ridx, ttl = 0, "cstname,jono,styno,location,type,msg".split(",")
            rmpfx = lambda x: (x[1:] if x[0] == "'" else x) if isinstance(x, str) else x

            jns = set(rmpfx(mp.get("jono")) for mp in logs)
            with self._cnsvc.sessionctx():
                jomp = self._cnsvc.getjos(jns)[0]
                jomp = {x.name.value: x for x in jomp}
                for mp in logs:
                    jn = rmpfx(mp.get("jono"))
                    if jn in jomp:
                        jn = jomp[jn]
                        mp["cstname"], mp[
                            "styno"], mp["ordertype"] = jn.customer.name.strip(
                            ), jn.style.name.value, jn.ordertype
                    else:
                        mp["cstname"], mp["styno"], mp["ordertype"] = (NA,) * 3
            logs = sorted(logs, key=lambda x: x["type"] + "," + x["styno"] + "," + x["jono"])
            jn = None
            for mp in logs:
                jomp = encs.get(mp["type"])
                if not jomp:
                    logger.debug("warning encoder for (%s) not found, default used", mp["type"])
                    jomp = self._err_enc_default
                vvs = jomp(mp)
                if vvs is None:
                    #weight error of new sample won't be shown
                    if mp["type"] == "wc_wgt":
                        continue
                    vvs = []
                # write a title row for each warning category
                if mp["type"] != jn:
                    jn = mp["type"]
                    ridx += 1
                    sht.range(ridx, 1).value = ttl + jomp(None)
                    _hl(sht.range(ridx, 1).expand("right"), 37)
                ridx += 1
                sht.range(ridx,
                          1).value = ["%s" % mp.get(x) if x != "type" else self._shpsns.get_error(mp.get(x))[1] for x in ttl] + vvs
            return sht

        for sn, logs, wtr in zip(("sn_err", "sn_warn"),
                                 self._shpsns._errandwarn(errlst),
                                 (_write_err, _write_warn)):
            if not logs:
                continue
            sht = wtr(self._shpsns.get(wb, sn), logs)
            if not sht:
                continue
            xwu.freeze(sht.range("D2"))
            sht.autofit("c")

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
            # if the file's parent folder not in ref_folder, treate it as pajraw files
            getrf = lambda x: triml(path.basename(path.dirname(x)))
            if getrf(fldr) not in [
                    getrf(sts[x].value)
                    for x in "shp.template,shpc1.template".split(",")
            ]:
                fldr = path.dirname(fldr)
        return self._pajfldr2file(fldr) if path.isdir(fldr) else fldr

    def _get_reader(self, wb):
        """
        try to detect the reader/vender_type of given workbook
        """
        rdrmap = {
            "长兴珠宝": ("c2", self._read_c2),
            "诚艺,胤雅": ("c1", self.read_c1),
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

    def build_rpts(self, fldr=None):
        """ create the rpt/err/bc sheets if they're not available
        @return: workbook if no error is found and None if err found during report generation
        """
        fn = self._get_file(fldr)
        if not fn:
            logger.debug("user does not specified any valid source file")
            return None
        pajopts = {
            "fn": fn,
            "shpdate": PajShpHdlr.get_shp_date(fn),
            "fmd": datetime.fromtimestamp(path.getmtime(fn))
        }

        invmp, shplst, errlst, self._vdrname = {}, [], [], None
        app, kxl = _appmgr.acq()
        try:
            wb, flag = app.books.open(fn), False
            # check if Rpt sheet was already there
            var = tuple(self._shpsns.get(wb, x, False) for x in ("sn_rpt",))
            flag = var and any(var)
            if flag:
                rng = xwu.usedrange(var[0])
                flag = bool(rng) or [x for x in rng.value if x]
            if flag:
                logger.debug("target file(%s) don't need regeneration", wb.name)
                return wb

            rdr = self._get_reader(wb)
            if not rdr:
                return None
            if self._vdrname == "paj":
                pajopts["bomwgts"] = PajBomHhdlr().readbom(wb)
            crt_err, shp_date = False, None  # critical error flag
            for sht in wb.sheets:
                flag, lst = False, rdr[1](sht, pajopts)
                if lst and any(lst):
                    flag, mp = True, lst[0]
                    if mp:
                        if "shpdate" in mp:
                            shp_date = mp["shpdate"]
                            del mp["shpdate"]
                        shplst.extend(mp.values())
                    if lst[1]:
                        self._eap(errlst, sht.name, sht.name,\
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
                td = datetime.today().date() - shp_date
                if td.days > 2 or td.days < 0:
                    self._eap(errlst, "_all_", "_日期_", "wc_date",
                            "落货日期%s可能错误" % shp_date.strftime("%Y-%m-%d"),
                            (td, shp_date))
            if shplst:
                if self._debug or not crt_err:
                    self._check_db_error(shplst, invmp, errlst)
            if errlst:
                self.write_logs(wb, errlst)
                errlst = self._shpsns._errandwarn(errlst)[0]
            if not shplst:
                wb = None
            else:
                if not errlst or (errlst and self._debug):
                    mp = {}
                    self._write_rpts(wb, shplst, mp, shp_date)
                    self._write_bc(wb, shplst, mp, shp_date)
                else:
                    wb = None
        finally:
            if kxl:
                _appmgr.ret(kxl)
        return wb


class ShpImptr():
    """
    import the shipment data(C1/PAJ/C2) into workflow system
    also generate BC/MMimport data for BCsystem and HK system
    """

    def __init__(self, cnsvc, hksvc, bcsvc):
        self._cnsvc, self._hksvc, self._bcsvc = cnsvc, hksvc, bcsvc
        self._groupsampjo = False
        self._shpsns = ShpSns()

    def exacthdr(self, sht):
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
        sm, verbose = ShpMkr(self._cnsvc, self._hksvc,
                             self._bcsvc, **options), options.get("verbose")
        wb = sm.build_rpts(fn)
        if not wb:
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
            msg = "文件=%s,\n日期=%s，落货纸号=%s\n总件数=%s，总重量=%s" % \
                (wb.name, nlhdr.date.strftime("%Y-%m-%d"), nlhdr.jmpno,
                mp["ttlqty"], str(round(mp["ttlwgt"], 2)))
            msg = messagebox.askyesno("确定要将以下资料导入落货系统?", msg)
            if not msg:
                return
            msg = self._do_db(mp)
            if not msg:
                self._write_mm_in(wb, nlhdr)
            else:
                errs.append("程序错误(%s)" % msg)
        if ttl:
            if errs:
                sm.write_logs(wb, errs)
                self._shpsns.get(wb, "sn_err").activate()
                ttl = ("检测到错误", "详情请参考Excel")
        xwu.appswitch(_appmgr.acq()[0], {"visible": True})
        if ttl and verbose:
            easymsgbox(messagebox.showinfo, ttl[0], ttl[1])
        return ttl or wb

    def _check_rpt_error(self, wb):
        if not wb:
            xwu.appswitch(_appmgr.acq()[0], True)
            ttl = ("文件错误", "文件有误或不存在")
        else:
            sht, errs = self._shpsns.get(wb, "sn_rpt", False), []
            if not sht:
                ttl = ("文件错误", "关键页(%s)不存在" % self._shpsns.get_error("sn_rpt")[1])
            else:
                nlhdr = NamedList(list2dict("date,jmpno,iono"), self.exacthdr(sht))
                xwu.appswitch(wb.app, {"visible": True})
                if self._isimported(nlhdr.jmpno):
                    var = "JMP#(%s)已导入" % nlhdr.jmpno
                    errs.append(
                        self._shpsns.new_err("_记录重复_", "_all_", "ec_jmp", var))
                    ttl = ("错误", var)
                else:
                    if nlhdr.jmpno[0] != "J":
                        var = "落货纸#(%s)应该以J开头" % nlhdr.jmpno
                        errs.append(
                            self._shpsns.new_err("_落货纸错误_", "_all_", "ec_jmp", var))
                    var = date.today() - nlhdr.date
                    if var.days < 0 or var.days > 20:
                        var = ("来至未来(%s)的资料" if var.days < 0 else
                            "太早以前(%s)的资料") % nlhdr.date.strftime("%Y-%m-%d")
                        errs.append(
                            self._shpsns.new_err("_日期错误_", "_all_", "ec_date", var))
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
                            errs.append(
                                self._shpsns.new_err(nl.jono, nl.jono, "ec_qty",
                                                    "数量不足"))
                        if not nl.wgt or nl.wgt < 0:
                            ttl = ("重量错误", "存在重量不合规记录")
                            if nl.wgt < 0:
                                errs.append(
                                    self._shpsns.new_err(nl.jono, nl.jono, "ec_wgt_not_sure", "重量不确定，请人工复核"))
                            else:
                                errs.append(
                                    self._shpsns.new_err(nl.jono, nl.jono, "ec_wgt_missing", "欠重量资料"))
                        if nl.qty:
                            ttlqty += nl.qty
                            var = nl.qty
                        ttlwgt += nl.wgt * (nl.qty if nl.qty else var)
                        jns.add(nl.jono)
                    # don't send below locals to the caller
                    del var, wb
        return dict(x for x in locals().items() if x[1] is not None)

    def _do_db(self, mp):
        app, tk = _appmgr.acq()
        try:
            xwu.appswitch(app, {"visible": False})
            refid = refno = None
            maMap, mmMap, gdMap, updjos = {}, {}, {}, {}
            nls, nlhdr, errs, jns = tuple(
                mp.get(x) for x in "nls nlhdr errs jns".split())
            with ResourceCtx((self._cnsvc.sessmgr(),
                              self._hksvc.sessmgr())) as curs:
                jos = self._cnsvc.getjos(jns)[0]
                joqls = {x.name.value: x.qtyleft for x in jos}
                jos = {x.name.value: x for x in jos}
                for nl in nls:
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
                                refid, refno, mmid = self._lastrefid(
                                ), self._nextrefno(), self._lastmmid()
                            maMap[karat] = ma = MMMa()
                            refid += 1
                            ma.id, ma.name, ma.karat, ma.refdate, ma.tag = refid, refno, karat, nlhdr.date, 0
                    ma = maMap.get(karat)
                    mmmapid = jo.id if self._groupsampjo else random.randint(
                        0, 9999999)
                    if nl.qty:
                        if mmmapid not in mmMap:
                            mmMap[mmmapid] = mm = MM()
                            mmid += 1
                            mm.id, mm.jsid, mm.name, mm.refid, mm.qty = mmid, jo.id, nlhdr.jmpno, refid, 0
                        else:
                            mm = mmMap[mmmapid]
                        mm.qty += nl.qty
                        # don't change the jo.qtyleft directly because this might cause a double-substract by both me and mm.insert trigger
                        ql = joqls[jn] - nl.qty
                        if ql < 0:
                            errs.append(
                                self._shpsns.new_err(nl.jono, nl.jono, "ec_qty",
                                                     "数量不足"))
                        else:
                            mm.tag = 0 if ql else 4
                        joqls[jn] = ql
                    key = "%d/%d" % (mm.id, karat)
                    if key not in gdMap:
                        gdMap[key] = gd = MMgd()
                        gd.id, gd.karat, gd.wgt = mm.id, nl.karat, 0
                    else:
                        gd = gdMap[key]
                    gd.wgt += nl.wgt * (nl.qty if nl.qty else mm.qty)
                if not errs:
                    cncmds, xx = [], curs[0].query(MMMa).filter(
                        MMMa.tag == 0).all()
                    if xx:
                        cncmds.append(xx)
                        mmid = curs[0].query(func.max(MMMa.tag)).first()
                        mmid = (mmid[0] if mmid else 0) + 1
                        for ma in xx:
                            ma.tag = mmid
                    cncmds.extend([
                        tuple(y) for y in (maMap.values(), mmMap.values(),
                                           gdMap.values(), updjos.values()) if y
                    ])
                    hkjos, hkcmds = self._hksvc.getjos(
                        jns, JOhk.running == 0)[0], []
                    for jo in hkjos:
                        if jo.running:
                            continue
                        x = jos.get(jo.name.value)
                        if not (x and x.running):
                            continue
                        jo.running = x.running
                        hkcmds.append(jo)
                    if cncmds:
                        for x in cncmds:
                            for y in x:
                                curs[0].add(y)
                            curs[0].flush()
                        curs[0].commit()
                    if hkcmds:
                        for x in hkcmds:
                            curs[1].add(x)
                        curs[1].commit()
        except Exception as exp:
            return exp
        finally:
            if tk and app:
                _appmgr.ret(tk)
        return None

    def _write_mm_in(self, wb, nlhdr, new_wb=True):
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
            # TODO: if vendor is C1, append C1 to before prefix
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
            sht = self._shpsns.get(wb, "mmimptr")
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
