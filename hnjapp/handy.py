#! coding=utf-8
'''
* @Author: zmFeng
* @Date: 2018-10-19 13:41:06
* @Last Modified by:   zmFeng
* @Last Modified time: 2018-10-19 13:41:06
handy utils for daily life
'''

from datetime import datetime
from itertools import chain
from numbers import Number
from os import listdir, makedirs, path, remove, rename, sep, utime, walk
from re import compile as compile_r
from shutil import copy

from PIL import Image
from sqlalchemy import desc
from sqlalchemy.orm import Query
from xlwings.constants import LookAt

from hnjapp.c1rdrs import C1InvRdr, _fmtbtno
from hnjapp.pajcc import (MPS, PajCalc, PrdWgt, WgtInfo, addwgt, cmpwgt,
                          karatsvc)
from hnjapp.svcs.misc import StylePhotoSvc
from hnjcore import JOElement
from hnjcore.models.cn import JO as JOcn
from hnjcore.models.cn import StoneMaster
from hnjcore.models.cn import Style as Stycn
from hnjcore.models.hk import JO, Orderma, PajShp, Style
from utilz import (
    NamedLists, ResourceCtx, getfiles, splitarray, triml, trimu, xwu)
from utilz.exp import AbsResolver, Exp
from utilz.odbctpl import getXBase
from utilz.xwu import (NamedRanges, appmgr, find, freeze, maketable, offset,
                       usedrange)

from .common import _getdefkarat
from .common import _logger as logger
from .common import karatsvc
from .svcs.db import HKSvc

try:
    import pandas as pd
except:
    pandas = None



def _df2sht(df, sht):
    '''
    dataframe to tuple of tuple, for excel value assignment
    '''
    lsts = [df.columns.to_list()]
    for *_, row in df.iterrows():
        lsts.append(row.tolist())
    sht.cells[0, 0].value = lsts
    rng = sht.range(sht.cells[0, 0], sht.cells(len(lsts), len(lsts[0])))
    rng.row_height = 18
    freeze(sht.cells[1, 1])
    maketable(rng)
    sht.autofit('c')
    return lsts

class CadDeployer(object):
    """
    when (maybe C1)'s JCAD file comes, send them to related folder
    also check if they're already there, if already exists, prompt the user
    also decrease the pending list
    """

    def __init__(self, tar_fldr=None):
        self._tarfldr = tar_fldr

    def deploy(self, src_fldr, tar_fldr=None):
        """
        deploy the jcad files in src_fldr to tar_fldr, if tar_fldr is ommitted,
        deploy to self._tarfldr
        return a tuple as (list(stynos deployed.), list(dup. stynos))
        """
        fns, stynos, dups = [path.join(src_fldr, x) for x in listdir(src_fldr)], [], []
        if not fns:
            return None
        if tar_fldr is None:
            tar_fldr = self._tarfldr
        for fn in fns:
            styno = path.splitext(path.basename(fn))
            styno = (trimu(styno[0]), styno[1])
            var0 = self._exists(styno[0])
            if var0:
                dups.append((fn, var0))
            else:
                stynos.append("".join(styno))
                dt = datetime.fromtimestamp(path.getmtime(fn))
                dt = "%s%s%s%s%s" % (dt.strftime("%Y"), sep, dt.strftime("%m%d"), sep, styno[0])
                dt = path.join(tar_fldr, dt)
                if not path.exists(dt):
                    makedirs(dt)
                rename(fn, path.join(dt, stynos[-1]))
        self._modlist(stynos, "delete")
        return (stynos, dups)

    def _exists(self, styno):
        #FIXME: check if styno exists current folder or child folders
        return []

    def addlist(self, stynos):
        """
        add stynos to the pending list(waiting for CAD file)
        """
        self._modlist(stynos, "add")

    def _modlist(self, stynos, action="add"):
        """
        modify(add/delete) the pending list
        """
        #FIXME
        pass


def ren_paj_imgs(src_fldr, sm_hk, keep_org=True, shortsz=1500):
    """
    rename the Paj image(for example, 23ARVXA062-0V07400.JPG)
        to our style_jono.jpg(for example, B13996_575459.jpg)
    @param src_fldr: the folder contains the jpgs
    @param sm_hk: HK server's session manager
    @param keep_org: keep the original file
    @param shortsz: the short side's min length
    """
    ptn = compile_r(r"\w{10}-\w{7}")
    fns = getfiles(src_fldr)
    if not fns:
        return None
    pcs = {path.splitext(path.basename(fn))[0].replace("-", ""): fn for fn in fns if ptn.search(fn)}
    if not pcs:
        return None
    with ResourceCtx(sm_hk) as cur:
        lst, jns = Query([PajShp.pcode, Style.name.label("styno"), JO.name.label("jono")]).join(JO).join(Orderma).join(Style).filter(PajShp.pcode.in_(pcs)).with_session(cur).all(), set()
        for x in lst:
            fn = x.jono.name
            if fn in jns:
                continue
            jns.add(fn)
            fn = pcs.get(x.pcode)
            if not (fn and path.exists(fn)):
                continue
            root, bn = path.split(fn)
            bn = path.splitext(bn)
            img = Image.open(fn)
            ptn, cp = img.size, keep_org
            if max(ptn) > shortsz * 1.1:
                ptn = (shortsz, shortsz / ptn[0] * ptn[1]) if ptn[0] < ptn[1] else (shortsz / ptn[1] * ptn[0], shortsz)
                ptn = tuple(int(x) for x in ptn)
                if keep_org:
                    fn = path.join(root, bn[0] + "_" + bn[1].lower())
                    cp = False
                img.thumbnail(ptn)
                img.save(fn)
            img.close()
            ptn = (fn, path.join(root, "%s_%s%s" % (x.styno, x.jono, bn[1])))
            if cp:
                copy(*ptn)
            else:
                rename(*ptn)
            del pcs[x.pcode]
        if pcs:
            for fn in pcs.values():
                print("failed to find JO record about (%s)" % path.basename(fn))
            for fn in pcs.values():
                root, ptn = path.split(fn)
                rename(fn, path.join(root, "_" + ptn))


def makecrab(act="MAKE"):
    """ prepare file for fish or crab handling """

    srcstyroot = r"\\172.16.8.91\Jpegs\Style"
    srcjoroot = r"\\172.16.8.91\Jpegs\JPEG"
    #tarroot = r"d:\temp\styphoto"
    tarroot = r"\\172.16.8.46\pb\dptfile\quotation\date\Date2018\0703"
    fn = r"\\172.16.8.46\pb\dptfile\quotation\date\Date2018\0703\龙虾扣&鱼勾扣.xls"
    # also get the operations defined, the user defined by back color
    # 5296274 is green, 65535 is yellow, 16777215 is no color

    def _chkcpy(src, tar):
        if not path.exists(src):
            return False
        if path.exists(tar):
            mts = (path.getmtime(tar), path.getmtime(src))
            if mts[0] >= mts[1]:
                if mts[0] > mts[1]:
                    logger.debug("target file(%s) is newer than source(%s)" % (tar, src))
                return False
        copy(src, tar)
        utime(tar, (path.getatime(src), path.getmtime(src)))
        return True

    if act == "MAKE":
        app, tk = appmgr.acq()

        _jn_str = JOElement.tostr
        opmap = {52377: "工单图", 65535: "款图", 16777215: "SN及配件"}
        try:
            wb = app.books.open(fn)
            sht = wb.sheets[0]
            vvs = usedrange(sht).value[1:]
            stynos = []
            ridx = 1
            for x in vvs:
                ridx += 1
                it = dict({"styno": x[0], "type": x[1]})
                stynos.append(it)
                lst = dict([(_jn_str(y), "") for y in x[2:] if y])
                if lst:
                    it["jonos"] = lst
                    for idx in range(2, len(x)):
                        rg = sht.range(ridx, idx+1)
                        jn, clr = _jn_str(rg.value), int(rg.api.interior.color)
                        if jn:
                            it["jonos"][jn] += ("-" + opmap[clr] if clr in opmap else "未知")
        finally:
            appmgr.ret(tk)
        cnt = 0

        def mksrcfldr(x, y):
            return path.join(x, path.sep.join([y[:i] for i in range(2, 4)]))

        def _ensure(fldr):
            if not path.exists(fldr):
                makedirs(fldr)
            return fldr

        for x in stynos:
            cnt += 1
            # if cnt > 20: break
            pfx, sfx = x["styno"], x["type"]
            if not sfx:
                continue
            srcfldr = mksrcfldr(srcstyroot, pfx)
            fns = getfiles(srcfldr, pfx)
            if fns:
                tarfldr = _ensure(path.join(tarroot, pfx + ("_" + sfx if sfx else pfx), "style"))
                for y in fns:
                    tfn = path.join(tarfldr, path.basename(y))
                    _chkcpy(y, tfn)
                with open(path.join(tarfldr, "files.dat"), "w") as fh:
                    print("#the origingal file names:", file=fh)
                    for y in fns:
                        print(path.basename(y), file=fh)
                if "jonos" in x:
                    tarfldr = _ensure(path.join(tarroot, pfx + ("_" + sfx if sfx else pfx), "jo" + sfx))
                    for y in x["jonos"].items():
                        jn, ops = y[0], y[1]
                        srcfldr = mksrcfldr(srcjoroot, jn)
                        fns = getfiles(srcfldr, jn)
                        for z in fns:
                            jn = path.basename(z).split(".")
                            jn = "%s%s.%s" % (jn[0], ops[0] + sfx * 2 + ops[1:], jn[1])
                            _chkcpy(z, path.join(tarfldr, jn))
    elif act == "GENLIST":
        # check get a list about the files removed
        rvs, nochgs, jfs = [], [], []
        for x in walk(tarroot):
            rt, fns = x[0], x[2]
            rt = trimu(rt)
            if rt.find("STYLE") >= 0:
                with open(path.join(rt, "files.dat")) as fh:
                    orgfns = set(x[:-1] for x in fh.readlines() if not x.startswith("#"))
                    dfs = orgfns.difference(fns)
                    if dfs:
                        rvs.extend(dfs)
                    else:
                        nochgs.append(path.basename(path.split(rt)[0]))
            elif rt.find("JO") >= 0:
                jfs.extend([path.join(rt, x) for x in fns if trimu(x).endswith("JPG")])
        with open(path.join(tarroot, "prosslog.txt"), "wt") as fh:
            for x in zip((rvs, nochgs), ("#style files removed", "#Style files without changes")):
                if not x[0]:
                    continue
                print(x[1], file=fh)
                for y in x[0]:
                    print(y, file=fh)
        for fn in jfs:
            _chkcpy(fn, path.join(r"d:\temp\xx", path.basename(fn).split("-")[0]+".jpg"))


def _format_btchno():
    """ target folder's batch# is malform, format them """
    tk, app = appmgr.acq()
    root = r"\\172.16.8.46\pb\dptfile\pajForms\miscs\现存宝石\宝石部分已寄"
    try:
        for fn in getfiles(root, "xls"):
            wb, upd = app.books.open(fn), 0
            for sht in wb.sheets:
                rng = find(sht, "水号")
                if not rng:
                    continue
                rng.api.entirecolumn.insert()
                lsts, idx = [], 0
                rnga = usedrange(sht)
                rnga = sht.range(rng, rnga.last_cell)
                nls = NamedLists(rnga.value)
                for nl in nls:
                    try:
                        btno = _fmtbtno(nl["水号"])
                        if btno:
                            lsts.append(("'" + btno,))
                        else:
                            lsts.append((" ",))
                    except:
                        print("error(%s) in file(%s)" % (nl["水号"], fn))
                        lsts.append(("'-",))
                    idx += 1
                rng.offset(1, -1).value = lsts
                upd += len(lsts)
            if upd:
                wb.save()
            wb.close()
    finally:
        appmgr.ret(tk)

def mtl_cost_forc1(c1calc_fn):
    """
    don't use pajcc's calculator method because I can not adjust the loss rate
    """
    app, tk = appmgr.acq()
    try:
        wb = app.books.open(c1calc_fn)
        version = offset(find(wb.sheets["背景资料"], "Version", lookat=LookAt.xlWhole), 1, 0).value
        sht = wb.sheets["计价资料"]
        rng = find(sht, "镶石费$")
        lossrates = {"GOLD": 1.08, "SILVER": 1.09} if version else {"GOLD": 1.07, "SILVER": 1.08}
        oz2gm = 31.1035
        org = [rng.row, 0]
        nls = NamedRanges(rng, {"jono": "工单,", "styno": "款号,", "karat0": "成色1", "wgt0": "金重1", "karat1": "成色2", "wgt1": "金重2", "karat2": "成色3", "wgt2": "金重3", "mtlcost": "金费", "mps": "金价"})
        cc, idx = PajCalc(), 0
        for nl in nls:
            idx += 1
            if not nl.jono:
                continue
            if not org[1]:
                org[1] = rng.column + nl.getcol("mtlcost")
            wgt = tuple(WgtInfo(getattr(nl, "karat%d" % idx), getattr(nl, "wgt%d" % idx)) for idx in range(3))
            if True:
                rc, mps = 0, MPS(nl.mps)
                for x in wgt:
                    if not x.karat:
                        continue
                    kt = karatsvc.getkarat(x.karat)
                    mp = mps.silver if x.karat == 925 else 0 if x.karat == 200 else mps.gold
                    if not mp and x.karat != 200:
                        rc = -1
                        break
                    rc += (x.wgt * kt.fineness * lossrates[kt.category] * mp / oz2gm)
                wgt = rc
            else:
                wgt = cc.calcmtlcost(PrdWgt(*wgt), nl.mps, lossrate=1.08, vendor="C1")
            sht.range(org[0] + idx, org[1]).formula = "= round(%f * if($C$4>1,1,6.5),2)" % wgt
        #wb.close()
    finally:
        if app:
            app.visible = True
            #appmgr.ret(tk)

def check_c1_wgts(cand=None, src=None):
    '''
    c1's weight data from kang, need validation
    '''

    kt_mp = {"Silver": 925, "9K": 9, "18K": 18, "14K": 14}
    def _wi(nl):
        kt = nl.karat
        if isinstance(kt, Number):
            kt = int(kt)
        return WgtInfo(kt_mp.get(kt, kt), nl.main)

    def _addwgt(wgts, wi, ispart=False):
        nw = wgts.netwgt or 0
        return addwgt(wgts, wi, ispart)._replace(netwgt=nw+wi.wgt)

    cand = cand or r"p:\aa\bc\明哥落货资料_汇总.xlsx"
    # src = src or r"\\172.16.8.46\pb\dptfile\quotation\2017外发工单工费明细\CostForPatrick\AIO_F.xlsx"

    app, tk = appmgr.acq()

    wb = app.books.open(cand)
    cand = NamedRanges(usedrange(wb.sheets[0]), {"jono": "工单号", "styno": "款号", "karat": "成色", "main": "金重", "stone": "石重", "metal_stone": "连石重", "chain": "配件重"})
    mp, idx, chns = {}, 0, {}
    for nl in cand:
        idx += 1
        jn = JOElement.tostr(nl.jono)
        if not jn:
            break
        if jn not in mp:
            wgts = _addwgt(PrdWgt(), _wi(nl))
        else:
            wgts = _addwgt(mp[jn], _wi(nl))
        if nl.stone:
            wgts = wgts._replace(netwgt=round(wgts.metal_jc + nl.stone, 2))
        if nl.chain:
            chns[jn] = (nl.styno, nl.chain)
        mp[jn] = wgts
    wb.close()
    for jn, x in chns.items():
        wgts = mp[jn]
        mp[jn] = _addwgt(wgts, WgtInfo(wgts.main.karat, x[1]), x[0] and x[0].find("P") >= 0)
    if not src:
        appmgr.ret(tk)
        return mp

    cand = mp
    # the source
    wb = app.books.open(src)
    mp = {x.jono: x for y in C1InvRdr().read(wb) for x in y[0]}
    wb.close()

    lsts = [('JO#', 'Expected', "Actual"), ]
    for jn, wgts in cand.items():
        # below 3 items were modified by me, for test only, so just skip it
        # if jn in ('463347', '463468', '463490'):
        #    continue
        wgts_exp = mp.get(jn)
        if not wgts_exp:
            print("Error, no source weight found for JO#(%s)" % jn)
            continue
        wgts_exp = wgts_exp.mtlwgt
        # because C1InvRdr return a 4-digit result, need to convert it to 2-digit
        nl = [None if not x else WgtInfo(x.karat, round(x.wgt, 2)) for x in (wgts_exp.main, wgts_exp.aux, wgts_exp.part)]
        wgts_exp = PrdWgt(nl[0], aux=nl[1], part=nl[2], netwgt=wgts_exp.netwgt)
        if not cmpwgt(wgts_exp, wgts):
            lsts.append((jn, wgts_exp.terms(), wgts.terms()))
    wb = app.books.add()
    wb.sheets[0].cells(1, 1).value = lsts
    app.visible = True
    return cand

def style_photos(dat_fn, tar_fldr):
    '''
    request from Murphy according to website:
    https://enterprise.atelier.technology/home/login
    provided a text file of sty# only, extract the photo that contains
    pure white background
    @return (tar_fldr, missing_fn) where
        @tar_fldr is the folder contains the result files
        @missing_fn: None or name of text file containing the missing styles
    '''
    var, stynos = trimu(path.splitext(dat_fn)[1][1:]), None
    if var.find("XL") == 0:
        app, tk = appmgr.acq()
        styn, fns = app.books.open(dat_fn), []
        for var in [x for x in styn.sheets]:
            x = usedrange(var).value
            if x:
                fns.extend(x)
        styn.close()
        appmgr.ret(tk)
        if not fns:
            return None
        tk = compile_r(r"^[A-Z]{1,2}\d{3,6}$")
        stynos = [x for y in fns for x in y if x and isinstance(x, str) and tk.findall(x)]
    else:
        with open(dat_fn) as fh:
            stynos = [x for y in fh.readlines() for x in y.split()]
    if not stynos:
        return None
    svc, rmp = StylePhotoSvc(), {}
    for styn in stynos:
        if styn in rmp:
            logger.debug("sty#(%s) is duplicated")
            continue
        fns = svc.getPhotos(styn)
        if not fns:
            rmp[styn] = None
            continue
        flag, fns = True, sorted(fns, key=path.getmtime, reverse=True)
        for var in fns:
            flag = StylePhotoSvc.isGood(var)
            if flag:
                rmp[styn] = (var, True)
                break
        if not flag and fns:
            # use the largest instead of the lastest
            fns = sorted(fns, key=path.getsize, reverse=True)
            rmp[styn] = (fns[0], False)
    roots = (tar_fldr, path.join(tar_fldr, "ref"))
    err_fn = path.join(tar_fldr, "_missing.txt")
    if path.exists(err_fn):
        remove(err_fn)
    for var in roots:
        if not path.exists(var):
            makedirs(var)
    fns = []
    for styn, var in rmp.items():
        if var:
            copy(var[0], path.join(roots[0 if var[1] else 1], path.basename(var[0])))
        else:
            fns.append(styn)
    if fns:
        with open(err_fn, "w+t") as fh:
            fh.writelines(("\n".join(fns), ))
        fns = err_fn
    return tar_fldr, fns

def stocktake_data(sessMgr, fn=None):
    fn = fn or r"\\172.16.8.46\pb\DptFile\pajForms\miscs\现存宝石\总数统计.xlsx"
    app = appmgr.acq()[0]
    mp = {}
    try:
        wb = app.books.open(fn)
        for sn in ('1', '2', '4.1', '4.2'):
            nls = [x for x in NamedRanges(usedrange(wb.sheets[sn]))]
            print("%d records from sheet(%s)" % (len(nls), sn))
            for nl in (nl for nl in nls if nl['已寄'] == 'N'):
                pfx, qty, wgt = [nl[x] for x in ('包头', '数量', '重量')]
                pfx = pfx[:2]
                lst = mp.setdefault(pfx, [0, 0])
                lst[0] += qty or 0
                lst[1] += round(wgt or 0, 3)
        with ResourceCtx(sessMgr) as cur:
            pkmp = cur.query(StoneMaster).filter(StoneMaster.name.in_([x for x in mp])).all()
            pkmp = {pk.name: pk for pk in pkmp}
            lst = ['石料 中文描述 数量 重量(卡)'.split()]
            for pk, qnw in mp.items():
                lst.append([pk, pkmp[pk].cdesc, qnw[0], qnw[1]])
        wb.close()
        wb = app.books.add()
        wb.sheets[0].cells(1, 1).value = lst
        wb.sheets[0].autofit()
    finally:
        app.visible = True
        #appmgr.ret(tk)

class JOResolver(AbsResolver):
    '''
    extract JO's property by name
    '''

    def __init__(self):
        self._jo = None

    def setjo(self, jo):
        self._jo = jo

    def resolve(self, arg):
        rc = arg
        if isinstance(arg, str):
            sig = '${jo}.'
            idx = arg.find(sig)
            if idx >= 0:
                fld, jo = triml(arg[idx + len(sig):]), self._jo
                if fld == 'cstname':
                    rc = jo.customer.name.strip()
                elif fld in ('qty', 'quantity'):
                    rc = jo.qty
                elif fld == 'karat':
                    rc = jo.karat
                elif fld == 'description':
                    rc = jo.description
        return rc

class ExtExp(Exp):
    def _eval_ext(self, op, l, r):
        if op == 'find':
            return l.find(r) >= 0
        return super()._eval_ext(op, l, r)

class CoreTraySvc(object):
    r''' help to find running of specified sty# for core tray,
    an example is in \\172.16.8.46\pb\dptfile\quotation\date\Date2019\0611\Candy Stamping Creoles_Rst.xlsx, where sheet("Stamping") is the source
    '''

    def __init__(self):
        self._qmp = {}
        self._rsv = JOResolver()
        self._KEY_RUNN = "Runn"

    def find_running(self, fn, cnsvc, hints):
        ''' write result progressive instead of block writing because
        program is not so stable
        Args:
            fn:     the source file to read sty# from
            cnsvc:  china db service
            hints:  dict(styno(string), running(string or int))
        '''
        app = appmgr.acq()[0]
        app.visible = True
        wb = app.books.open(fn)
        sns = self._select_req_sheets(wb)
        if not hints:
            hints = {}
        for sn in sns:
            # it's strange that usedranged might return A2:xx case, so
            sht = wb.sheets[sn]
            vvs = usedrange(sht).last_cell.address.split(":")[-1]
            vvs = sht.range("A1:" + vvs).value
            # vvs = usedrange(sht).value
            sht = wb.sheets.add(before=sht)
            sht.name = sn + '-R'
            sht.activate()
            self._find_one_sheet(sht, vvs, cnsvc, hints)
            sht.autofit('c')

    def _find_one_sheet(self, sht, vvs, cnsvc, hints):
        def _write(sht, ridx, cidx, runn, fromHints=False):
            if fromHints:
                val = runn
            else:
                val = (se.name + "@%s" % runn) if runn else "%s@Error" % se.name
            sht.cells[ridx, cidx].value = val
        locs = {}
        for ridx, row in enumerate(vvs):
            for cidx, val in enumerate(row):
                if not val:
                    continue
                se = JOElement(val)
                if not se.isvalid():
                    continue
                styno = se.name
                if styno in locs and styno in hints:
                    del hints[styno]
                runn = (None if styno in locs else hints.get(styno)) or self._find(cnsvc, se)
                _write(sht, ridx, cidx, runn, styno in hints)
                if styno in locs:
                    cidx = locs[styno]
                    if cidx:
                        locs[styno] = None
                        # this style is duplicated, override the prior result
                        _write(sht, cidx[0], cidx[1], self._find(cnsvc, se))
                else:
                    locs[styno] = (ridx, cidx)

    @staticmethod
    def _select_req_sheets(wb):
        lst, sns = [], {triml(x): x for x in (x.name for x in wb.sheets)}
        for sn in sns:
            idx = sn.find('-')
            if idx <= 0:
                continue
            lst.append(sn)
            lst.append(sn.split('-')[0])
        if lst:
            for sn in lst:
                del sns[sn]
        return sns.values()

    def _find(self, cnsvc, styn):
        with cnsvc.sessionctx() as cur:
            q = self._find_q.filter(Stycn.name == styn)
            jos = q.with_session(cur).all()
            worst = None
            if not jos:
                return None
            worst = jos[0]
            for dsc, exps in self._find_lvl.items():
                for jo in jos:
                    self._rsv.setjo(jo)
                    if exps.eval(self._rsv):
                        return dsc + ":%d" % jo.running
            return 'Bad:%d' % worst.running if worst else None

    @property
    def _find_q(self):
        key = '_find_q'
        if key not in self._qmp:
            self._qmp[key] = Query(JOcn).join(Stycn).filter(JOcn.id > 0).filter(JOcn.tag > -10).filter(JOcn.running > 200000).order_by(desc(JOcn.createdate))
        return self._qmp[key]

    @property
    def _find_lvl(self):
        key = '_find_level'
        if key not in self._qmp:
            _pfx = lambda x: "${jo}." + x
            _k = lambda k: Exp(_pfx("karat"), '==', k)

            dd = ExtExp(_pfx("description"), 'find', '钻').or_(
                ExtExp(_pfx("description"), 'find', '占'))
            naa = Exp(_pfx("cstname"), '==', 'NAA')
            self._qmp[key] = {'naa+9+dd': naa.chain('and', _k(9), dd),
                        '9+dd': _k(9).and_(dd),
                        'naa+9': naa.and_(_k(9)),
                        '9': _k(9),
                        '925+dd': _k(925).and_(dd),
                        '925': _k(925)}
        return self._qmp[key]

    def read_hints(self, wb):
        ''' read the sty -> runn hints from a getRunn file
        '''
        rst = [('shtname styno running'.split())]
        dx = dy = fc = None
        for sht in wb.sheets:
            if not dy:
                dy = self._detect_y_dist(sht)
                if not dy:
                    continue
                dy, fc = dy
            sn = sht.name
            vvs = usedrange(sht).value
            for ridx, row in enumerate(vvs):
                if not dx or dx == -1:
                    dx = self._detect_x_dist(fc, row)
                    if not dx:
                        continue
                rns = []
                if dx > 0:
                    ln = len(row)
                    for idx in range(fc, fc + 10000, dx):
                        if idx >= ln:
                            break
                        runn = row[idx]
                        if self._is_runn(runn):
                            rns.append((idx, runn))
                else:
                    rns = [(fc, row[fc])]
                for ln, runn in rns:
                    rst.append((sn, vvs[ridx + dy][ln], runn.split(':')[-1].strip()))
        return rst

    def _is_runn(self, val):
        return val and isinstance(val, str) and val.find(self._KEY_RUNN) == 0 and val.find(':') > 0

    def _detect_y_dist(self, sht):
        rng = find(sht, self._KEY_RUNN)
        if not rng:
            return None
        if self._is_runn(rng.value):
            return rng.offset(-5, 0).end('up').row - rng.row, rng.column - 1
        return None

    def _detect_x_dist(self, fc, row):
        ''' detect the x distinct between each block
        Returns:
            None if nothing found at all.
            Positive integer if found.
            -1 if only one block is there
        '''
        if not self._is_runn(row[fc]):
            return None
        rns = [(idx, runn) for idx, runn in enumerate(row) if self._is_runn(runn)]
        if len(rns) > 1:
            return rns[1][0] - rns[0][0]
        return -1

class MMStTk(object):
    r''' create stocktake report based on the period provided
    Args:
        cnnHK(func):     a function call that can return a connection to HNJHK db
        period(string):  period data to extract, if not provided, use current year
        root(string):   root folder to place the data to, default is d:\joblog\MMStTk
    '''
    def __init__(self, cnnHK, **args):
        self._func_hk = cnnHK
        self._perid = args.get('period')
        if not self._perid:
            td = datetime.today()
            y, m = td.year, td.month
            if m < 4:
                y -= 1
            self._perid = 'per%d' % y
        self._root = args.get('root', r'd:\joblog\MMStTk')
        self._year = self._perid[-4:]

    @property
    def period(self):
        ''' the period for report generation
        '''
        return self._perid

    @property
    def root(self):
        ''' root file to place files to
        '''
        return self._root

    def _fileName(self, tn):
        fn = None
        if tn == 'rpt':
            fn = path.join(self._root, 'STK%s_Diff.xlsx' % self._year)
        elif tn == 'cache':
            fn = path.join(self._root, '_cache')
            if not path.exists(fn):
                makedirs(fn)
        elif tn == 'mm':
            fn = path.join(self._fileName('cache'), '_' + self._perid + '_mm.csv')
        elif tn == 'mm_calc':
            fn = path.join(self._fileName('cache'), '_' + self._perid + '_mm_calc.csv')
        elif tn == 'mm_calc_sql':
            fn = path.join(self._root, 'STK%s_upd.sql' % self._year)
        elif tn == 'tk':
            fn = path.join(self._fileName('cache'), '_' + self._perid + '_tk.csv')
        return fn

    @property
    def _cnnHK(self):
        cnn = self._func_hk()
        print('connection established')
        return cnn

    def _getStock(self):
        fn = self._fileName('mm')
        if path.exists(fn):
            df = pd.read_csv(fn)
        else:
            theSql = 'select stockcode, description, qtyleft as qty from stockobjectma where qtyleft <> 0'
            with self._cnnHK as cnn:
                df = pd.read_sql_query(theSql, cnn)
            df.to_csv(fn, index=False)
            print('%d stock records from database' % len(df))
        df_c = self._getStockCalc()
        mp_df = dict(zip(df.stockcode, df.qty))
        mp_dfc = dict(zip(df_c.stockcode, df_c.qty))
        cmds = []
        keys = (set(mp_df.keys()), set(mp_dfc.keys()))
        diff = keys[1].difference(keys[0])
        theSql = "update stockobjectma set qtyleft = %f where stockcode = '%s'"
        if diff:
            lsts = []
            for x in diff:
                lsts.append((x, 'N/A', mp_dfc[x]))
                cmds.append(theSql % (mp_dfc[x], x))
            df = pd.concat((df, pd.DataFrame(lsts, columns=df.columns)))
        for idx, row in df.iterrows():
            cqty = float(mp_dfc.get(row.stockcode, 0))
            if abs(float(row.qty) - cqty) > 0.1:
                df.loc[idx, 'qty'] = cqty
                cmds.append(theSql % (cqty, row.stockcode))
        if cmds:
            with open(self._fileName('mm_calc_sql'), 'w+t') as fh:
                cmds.insert(0, 'use cstbld')
                cmds.insert(1, 'go')
                cmds.insert(2, '')
                cmds.append('go')
                for x in cmds:
                    print(x, file=fh)
            print("pls. execute file(%s) to sync stockobjectma's qtyleft" % self._fileName('mm_calc_sql'))
        return df

    def _getStockCalc(self):
        fn = self._fileName('mm_calc')
        if path.exists(fn):
            df = pd.read_csv(fn)
        else:
            theSql = 'select * from (select so.stockcode, sum(case when inv.locationidfrm = 2 then -d.qty else d.qty end) as qty from invoicema inv join invoicedtl d on inv.invid = d.invid join stockobjectma so on d.srid = so.srid where (inv.locationidto <> 9 and (inv.locationidfrm = 2 or inv.locationidto = 2)) group by so.stockcode) v where qty <> 0'
            with self._cnnHK as cnn:
                df = pd.read_sql_query(theSql, cnn)
            df.to_csv(fn, index=False)
            print('%d recalculated stock records from database' % len(df))
        return df

    def _getTake(self):
        fn = self._fileName('tk')
        if path.exists(fn):
            df = pd.read_csv(fn)
        else:
            # because there might be many records, one query can be very slow, into 10 logs
            theSql = "select min(id), max(id) from mmstocktake where perno = '%s'" % self._perid
            with self._cnnHK as cnn:
                frm, to = cnn.execute(theSql).fetchone()
                theSql = "select sm.stockcode, sm.description, st.docno, st.qty as qty from mmstocktake st join stockobjectma sm on st.srid = sm.srid where perno  = '%s' and id >= %d and id < %d"
                stp = int((to - frm) / 10)
                ids = [frm + x * stp for x in range(11)]
                ids[-1] = to + 1
                df = None
                for to, frm in enumerate(ids[:-1]):
                    to = ids[to+1]
                    print('fetching id(%d)' % frm)
                    var = pd.read_sql_query(theSql % (self._perid, frm, to), cnn)
                    if var is None:
                        df = var
                    else:
                        df = pd.concat((df, var))
            df.to_csv(fn, index=False)
            print('%d take records from database' % len(df))
        return df

    def run(self):
        ''' create the report file
        '''
        tk = self._getTake()
        mm = self._getStock()
        tk = tk.groupby(['stockcode', 'description'])
        docs = tk.docno.apply(lambda x: ','.join(sorted(set(x))))
        tk = tk.sum().reset_index()
        tk['docno'] = docs.tolist()

        tars = [[mm, set(mm.stockcode)], [tk, set(tk.stockcode)], ]
        for fIdx, df in enumerate(tars):
            tIdx = len(tars) - 1 - fIdx
            ni = df[1].difference(tars[tIdx][1])
            if ni:
                ni = df[0].loc[df[0].stockcode.isin(ni)]
                for var in ('docno', 'qty'):
                    if var in ni.columns:
                        del ni[var]
                ni['qty'] = [0] * len(ni)
                if fIdx == 0:
                    ni['docno'] = ['NoDoc',] * len(ni)
                ni.columns = tars[tIdx][0].columns
                tars[tIdx][0] = pd.concat((tars[tIdx][0], ni))
        mm, tk = (tars[i][0] for i in range(len(tars)))
        ni = dict(zip(mm.stockcode, mm.qty))
        tk['mo'] = tk['stockcode'].apply(lambda x: ni[x])
        tk['diff'] = tk.apply(lambda row: row.mo - row.qty, axis=1)
        df = tk['stockcode description docno mo qty diff'.split()]
        df.columns = 'stockcode description docno mo st diff'.split()
        df = df.sort_values(['diff', 'stockcode'])
        fn = self._fileName('rpt')
        # use xlwings to beautify the result
        app, tk = appmgr.acq()
        try:
            wb = app.books.add()
            sht = wb.sheets[0]
            sht.name = 'StkData'
            _df2sht(df, sht)
            wb.save(fn)
            wb.close()
        finally:
            appmgr.ret(tk)
        return fn

class HKMeltdown(object):
    '''
    handle the HK meltdown issue
    Args:
        fn(string): the source excel with sheet('00_List') as source
        sm_cn(func): a function that will create sessionMgr to China db
        sm_hk(func): a function that will create sessionMgr to HK db
    '''

    def __init__(self, fn, sm_cn, sm_hk):
        self._fn_stk = fn
        self._x_cn = sm_cn
        self._x_hk = sm_hk
        self._wgt_flds = 2 # only write wgt0/wgt1 to result

    @property
    def _cnn_cn(self):
        return self._x_cn().engine.connect()

    @property
    def _cnn_hk(self):
        return self._smHK.engine.connect()

    @property
    def _smHK(self):
        return self._x_hk()

    def _fileName(self, tn):
        fn = None
        def _bn(pth, sfx, pfx):
            if pfx is None:
                pfx = '_'
            return pfx + path.splitext(path.basename(pth))[0] + sfx
        if tn == 'srouce':
            fn = self._fn_stk
        elif tn == 'root':
            fn = path.dirname(self._fn_stk)
        elif tn == 'cache':
            fn = path.join(self._fileName('root'), '_cache')
            if not path.exists(fn):
                makedirs(fn)
        elif tn == 'lst2fj':
            fn = path.join(self._fileName('root'), _bn(self._fn_stk, '_lst2fj.csv', ''))
        elif tn == 'jowgts':
            fn = path.join(self._fileName('cache'), 'jowgts.csv')
        elif tn == 'mmwgts':
            fn = path.join(self._fileName('cache'), 'mmwgts.csv')
        elif tn == 'mitwgts':
            fn = path.join(self._fileName('cache'), 'mitwgts.csv')
        elif tn == 'fjfolder':
            fn = path.join(path.dirname(self._fn_stk), '_FJ_Data')
        else:
            if tn == 'jn2run':
                fn = path.join(self._fileName('cache'), _bn(self._fn_stk, '_jn2run.csv', '_'))
            elif tn == 'lst2jn':
                fn = path.join(self._fileName('cache'), _bn(self._fn_stk, '_lst2jn.csv', '_'))
            elif tn == 'fjdat':
                fn = path.join(self._fileName('cache'), _bn(self._fn_stk, '_fjdat.csv', '_'))
        return fn

    def run(self):
        '''
        create the FJ sheet based on a raw list, should be run under PY environment because most data source is PY based.
        list should be placed in sheet("00_List"), result file will be saved as a csv file in the same folder
        If not all aucos is found, update _fj_data\fj_stock files from HK
        '''
        self._list2FJ()

    def _valid(self, mmWgt, joWgt, mitInJC):
        _sum = lambda prdwgt: sum(x.wgt if x else 0 for x in prdwgt.wgts)
        wgts = [_sum(mmWgt), _sum(joWgt)]
        if mitInJC and joWgt.part and joWgt.part:
            wgts[1] -= joWgt.part.wgt
        return wgts[1] == 0 or abs(wgts[1] - wgts[0]) / wgts[1] < 0.1

    def _getMMWgts(self, allJns):
        if path.exists(self._fileName('mmwgts')):
            df_cache = pd.read_csv(self._fileName('mmwgts'))
        else:
            df_cache = pd.DataFrame(None, columns='jono styno running qty karat wgt'.split())
        if path.exists(self._fileName('mitwgts')):
            df_mcache = pd.read_csv(self._fileName('mitwgts'))
        else:
            df_mcache = pd.DataFrame(None, columns='jono karat wgt'.split())

        jonos = allJns.difference(set(df_cache.jono.unique()))
        if jonos:
            cnn = self._cnn_cn
            lsts = []
            mits = []
            def _persist(df_cache, lsts, df_mcache, mits):
                lsts = [(JOElement(x[0]).value, x[1], x[2], x[3], x[4], x[5]) for x in lsts]
                df_cache = pd.concat((df_cache, pd.DataFrame(lsts, columns='jono styno running qty karat wgt'.split())))
                df_cache.to_csv(self._fileName('mmwgts'), index=False)
                if mits:
                    mits = [(JOElement(x[0]).value, x[1], x[2]) for x in mits]
                    df_mcache = pd.concat((df_mcache, pd.DataFrame(mits, columns='jono karat wgt'.split())))
                    df_mcache.to_csv(self._fileName('mitwgts'), index=False)
                return df_cache, df_mcache

            theSql = 'select jo.cstbldid_alpha + convert(varchar(10), jo.cstbldid_digit) jono, sty.alpha + convert(varchar(10), sty.digit) styno, jo.running, mm.qty, gd.karat, gd.wgt from mm join mmgd gd on mm.mmid = gd.mmid join b_cust_bill jo on mm.jsid = jo.jsid join styma sty on jo.styid = sty.styid where (%s)'

            theSql_mit = 'SELECT jo.cstbldid_alpha + convert(varchar(10), jo.cstbldid_digit) jono, ma.karat, mit.wgt/jo.quantity FROM jomit mit join mitma ma on ma.mitid = mit.mitid join jocost c on c.jocostid = mit.jocostid join b_cust_bill jo on jo.jsid = c.jsid where (%s)'

            for jns in splitarray(tuple(jonos), 20):
                jns = [JOElement(x) for x in jns]
                s0 = ' or '.join(["cstbldid_alpha='%s' and cstbldid_digit = %d" % (x.alpha, x.digit) for x in jns])
                lst = cnn.execute(theSql % s0).fetchall()
                lsts.extend(lst)
                lst = cnn.execute(theSql_mit % s0).fetchall()
                if lst:
                    mits.extend(lst)
                if len(lsts) > 100:
                    df_cache, df_mcache = _persist(df_cache, lsts, df_mcache, mits)
                    lsts = []
                    mits = []
            if lsts or mits:
                df_cache, df_mcache = _persist(df_cache, lsts, df_mcache, mits)
        # finally, I need a map of map with JO# as key, submap with karat as key and wgt as value, by unit
        # because the cache file is global, so make it local
        df_cache = df_cache.loc[df_cache.jono.isin(allJns)]
        df_cache = df_cache.groupby(['jono', 'styno', 'running', 'karat']).sum().reset_index()
        df_cache['uwgt'] = df_cache.wgt / df_cache.qty
        df_cache['scode'] = df_cache.apply(lambda row: (row.styno, row.running), axis=1)
        jo_sty = dict(zip(df_cache.jono, df_cache.scode))

        df_mcache = df_mcache.loc[df_mcache.jono.isin(allJns)]
        if not df_mcache.empty:
            df_mcache['wi'] = df_mcache.apply(lambda row: WgtInfo(row.karat, -row.wgt), axis=1)
            df_mcache = dict(zip(df_mcache.jono, df_mcache.wi))
        else:
            df_mcache = {}
        mp = {}
        for *_, row in df_cache.iterrows():
            wgts = mp.get(row.jono) or PrdWgt()
            mp[row.jono] = addwgt(wgts, WgtInfo(row.karat, row.uwgt), False, True)
        for jono in mp:
            if jono not in df_mcache:
                continue
            # many MITs, include ear pin, but it's unitwgt won't be too high, so don't check by sty#, just by unitwgt
            wi = df_mcache[jono]
            if jo_sty[jono][0].find('P') >= 0:
                mp[jono] = addwgt(mp[jono], wi, False, True)
        return {x[0]: (x[1], jo_sty[x[0]], x[0] in df_mcache) for x in mp.items()}

    def _getJOWgts(self, allJNs):
        ''' return the joweight based on a jo list
        '''
        if path.exists(self._fileName('jowgts')):
            df_cache = pd.read_csv(self._fileName('jowgts'), converters={'jono': self._jonoCvt})
            df_cache.jono = df_cache.jono.apply(lambda x: JOElement(x).name)
        else:
            df_cache = pd.DataFrame(None, columns=('jono k0 w0 k1 w1 k2 w2'.split()))
        joDone = set(df_cache.jono)
        df = allJNs.difference(joDone)
        if df:
            def _persist(df_jowgts, lsts):
                df_jowgts = pd.concat((df_jowgts, pd.DataFrame(lsts, columns=df_jowgts.columns)))
                df_jowgts.to_csv(self._fileName('jowgts'), index=False)
                print('persisted')
                return df_jowgts

            hksvc = HKSvc(self._smHK)
            lsts = []
            for jn in df:
                try:
                    wgts = hksvc.getjowgts(jn)
                    lst = [jn, ]
                    lst.extend(chain(*[(wi.karat, wi.wgt) if wi else (0, 0) for wi in wgts.wgts]))
                except:
                    continue
                lsts.append(lst)
                if len(lsts) > 20:
                    df_cache = _persist(df_cache, lsts)
                    lsts = []
            if lsts:
                df_cache = _persist(df_cache, lsts)
        # make the global cache local
        df = df_cache.loc[df_cache.jono.isin(allJNs)]
        # normalize the karats like 91 to 9
        mp = {
            81: 8,
            88: 8,
            91: 9,
            98: 9,
            101: 10,
            108: 10,
            141: 14,
            148: 14,
            181: 18,
            188: 18
        }
        for kt in 'k0 k1 k2'.split():
            df[kt] = df[kt].apply(lambda x: mp.get(x, x))
        mp = {row.jono: PrdWgt(WgtInfo(row.k0, row.w0), WgtInfo(row.k1, row.w1), WgtInfo(row.k2, row.w2)) for *_, row in df.iterrows()}
        return mp

    @staticmethod
    def _runnoCvt(runno):
        if isinstance(runno, str):
            return runno
        return '%d' % runno

    @staticmethod
    def _jonoCvt(jono):
        return JOElement(jono).value

    def _list2FJ(self):
        ''' build-up the JO/runn/qty list based on the raw list provides RP#
        '''
        if not path.exists(self._fileName('lst2fj')):
            df_src = pd.read_excel(self._fn_stk, '00_List', header=5).dropna(subset=['RP#'])
            df_src['rpno'] = df_src['RP#'].apply(lambda x: 'RP%d' % x)
            df_src['snnum'] = df_src['No'].apply(lambda x: '%d' % x)
            df_src['fbymd'] = df_src['FB Date'].apply(lambda x: x.strftime('%Y%m%d'))
            df_src['fbnum'] = df_src['FB#'].apply(lambda x: 'FB%d' % x)
            cnn = None
            if path.exists(self._fileName('lst2jn')):
                df = pd.read_csv(self._fileName('lst2jn'), converters={'runno': self._runnoCvt})
            else:
                s0 = "','".join(set(df_src.rpno))
                theSql = "select ma.docno, sm.styno, sm.running as runno, sm.description, d.qty as quant, ma.inoutno as refno, convert(varchar(20), docdate, 112) as doymd from invoicema ma join invoicedtl d on ma.invid = d.invid join stockobjectma sm on d.srid = sm.srid where docno in ('%s') and locationidfrm = 2"
                cnn = self._cnn_hk
                df = pd.read_sql_query(theSql % s0, cnn)
                df.to_csv(self._fileName('lst2jn'), index=False)
            df['scode'] = df['ncode'] = df.apply(lambda row: '%s/%s' % (row.styno, row.runno), axis=1)
            s0 = len(df)
            df['fmloc'] = ['MM', ] * s0
            df['toloc'] = ['MT', ] * s0
            cns = 'rpno snnum fbymd fbnum'.split()
            mp = tuple(df_src.apply(lambda row: [row[x] for x in cns], axis=1))
            mp = {x[0]: x[1:] for x in mp}
            # now lookup the fields from df_src
            for idx, cn in enumerate(cns[1:]):
                df[cn] = df['docno'].apply(lambda x: mp[x][idx])
            if path.exists(self._fileName('jn2run')):
                mp = pd.read_csv(self._fileName('jn2run'), converters={'runno': self._runnoCvt})
            else:
                mp = pd.DataFrame(None, columns='runno jono'.split())
                theSql = 'select convert(varchar(10), running) as runno, convert(varchar(10), jo.cstbldid_alpha) + convert(varchar(10), jo.cstbldid_digit) jono from b_cust_bill jo where running in (%s)'
                if cnn is None:
                    cnn = self._cnn_cn
                for s0 in splitarray(tuple(df.runno.apply(lambda x: '%s' % x)), 100):
                    mp = pd.concat((mp, pd.read_sql_query(theSql % ','.join(s0), cnn)))
                mp.to_csv(self._fileName('jn2run'), index=False)
                cnn.close()
            mp = mp.apply(lambda row: (row.runno, row.jono), axis=1)
            mp = {x[0]: JOElement(x[1]).value for x in mp}
            df['jobno'] = df['runno'].apply(mp.get)
            df = df.sort_values(['snnum', 'docno', 'runno'])
            cns = [x for x in 'snnum scode ncode fbymd fbnum docno refno doymd fmloc toloc jobno styno runno descn quant aucos catag'.split() if x in df.columns]
            # re-order the columns
            df = df[cns]
            df.to_csv(self._fileName('lst2fj'), index=False)
        else:
            df = pd.read_csv(self._fileName('lst2fj'), converters={'runno': self._runnoCvt})
        updCnt = 0
        if 'aucos' not in df.columns:
            self._fillFJData(df)
            updCnt += 1
        if 'k0' not in df.columns:
            df = self._fillWgts(df)
            updCnt += 1
        if updCnt > 0:
            df.to_csv(self._fileName('lst2fj'), index=None)
        if sum(df.aucos.isna()) > 0:
            print('folder(%s) expired, get it from HK again' % self._fileName('fjdat'))
        else:
            self._fj2xls(df)

    def _fj2xls(self, df):
        df = df.sort_values(['runno', 'docno'])
        app, tk = appmgr.acq()
        wb = xwu.safeopen(app, self._fn_stk, readonly=False)
        fn = 'FJ'
        sht = xwu.findsheet(wb, fn)
        updCnt = 0
        if not sht:
            _df2sht(df, wb.sheets.add(fn))
            updCnt += 1
        wgts = {}
        fn = 'MtlWgt'
        sht = xwu.findsheet(wb, fn)
        if not sht:
            for *_, row in df.iterrows():
                for idx in range(self._wgt_flds):
                    wgt = row['w%d' % idx]
                    if not wgt:
                        continue
                    kt = karatsvc.getkarat(row['k%d' % idx])
                    cat = kt.category
                    wgts[cat] = round(wgts.get(cat, 0) + kt.fineness * row.quant * wgt, 2)
            wgts = sorted(list(wgts.items()), key=lambda x: x[0])
            sht = wb.sheets.add(fn)
            _df2sht(pd.DataFrame(wgts, columns='Metal Wgt'.split()), sht)
            updCnt += 1
        if updCnt > 0:
            wb.save()
        wb.close()
        appmgr.ret(tk)

    def _fillWgts(self, df):
        allJns = set(df.jobno)
        if pd.np.nan in allJns:
            allJns.remove(pd.np.nan)
        joWgts = self._getJOWgts(allJns)
        lsts = []
        for key, val in self._getMMWgts(allJns).items():
            s_r, mitInJC = val[1], val[2]
            val = val[0]
            wgts = joWgts.get(key)
            if val.empty and wgts:
                # in chaos stage(around 2017/05), JOs like B103235 without weight
                val = addwgt(PrdWgt(wgts.main, wgts.aux, None), wgts.part, False, True)
            rmk = 'OK' if self._valid(val, wgts, mitInJC) else 'suspicious'
            # because mmWgt has already substract mit, so let it be the final result
            lst = [key, s_r[0], s_r[1]]
            if not mitInJC and wgts:
                pts = wgts.part
                if not (pts is None or not pts.wgt):
                    val0 = addwgt(val, WgtInfo(pts.karat, -pts.wgt), False, True)
                    if sum([1 if x and x.wgt < 0 else 0 for x in val0]):
                        rmk = 'CHAIN ATTACHED LATER'
                    else:
                        val = val0
            val = val.follows(_getdefkarat(key))
            lst.extend(chain(*[(x.karat if x.wgt else 0, x.wgt) if x else (0, 0) for x in val.wgts]))
            lst.append(rmk)
            lsts.append(lst)
        val = {x[0]: x[3:] for x in lsts}
        for rmk in range(self._wgt_flds):
            df['k%d' % rmk] = df['jobno'].apply(lambda x: val[x][rmk * 2] if x in val else 0)
            df['w%d' % rmk] = df.apply(lambda row: val[row.jobno][rmk * 2 + 1] * row.quant if row.jobno in val else 0, axis=1)
        df['remark'] = df['jobno'].apply(lambda x: val[x][-1] if x in val else None)
        return df

    def _fillFJData(self, df):
        ''' from the df, fetch fj's aucos and catag
        '''
        if path.exists(self._fileName('fjdat')):
            df_fj = pd.read_csv(self._fileName('fjdat'))
        else:
            fn = self._fileName('fjfolder')
            if not path.exists(fn):
                return None
            # the in operation is quite slow even if the given field is indexed, so use pandas to cache again.
            # strange that the slowness happens only inside VFP, the odbc works quite well
            scs = set(df.scode.unique())
            df_fj = None
            theSql = "SELECT trim(lcode) as lcode, catag, aucos FROM fj_stock WHERE lcode in ('%s')"
            cnn = getXBase(fn)
            for sc in splitarray(scs, 20):
                var = pd.read_sql_query(theSql % "','".join(sc), cnn)
                if df_fj is None:
                    df_fj = var
                else:
                    df_fj = pd.concat((df_fj, var))
            nfs = scs.difference(set(df_fj.lcode))
            if nfs:
                # in Raymond's MORPSUM.PRG, there is still more ways to find
                nfs = self._getFJAdd(nfs, cnn)
            cnn.close()
            if nfs is not None:
                df_fj = pd.concat((df_fj, nfs))
            df_fj.to_csv(self._fileName('fjdat'), index=False)
        cns = 'lcode catag aucos'.split()
        var = df_fj.apply(lambda row: [row[x] for x in cns], axis=1)
        mp = {x[0].strip(): x[1:] for x in var}
        for idx, cn in enumerate(cns[1:]):
            df[cn] = df.scode.apply(lambda x: mp.get(x)[idx] if x in mp else None)
        return df

    def _getFJAdd(self, scs, cnn):
        theSql = "SELECT catag, aucos FROM fj_stock WHERE lcode = '%s'"
        lsts = []
        for sc in scs:
            lst = None
            for s0 in (sc.replace('/', ''), sc[:-1] if sc[-1].isalpha() else ''):
                if not s0:
                    continue
                lst = cnn.execute(theSql % s0).fetchone()
                if lst:
                    lsts.append([sc, lst[1], lst[2]])
                    break
        return pd.DataFrame(lsts, columns='lcode catag aucos'.split()) if lsts else None
