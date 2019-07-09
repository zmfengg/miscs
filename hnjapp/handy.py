#! coding=utf-8
'''
* @Author: zmFeng
* @Date: 2018-10-19 13:41:06
* @Last Modified by:   zmFeng
* @Last Modified time: 2018-10-19 13:41:06
handy utils for daily life
'''

from datetime import datetime
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
from utilz import ResourceCtx, getfiles, triml, trimu
from utilz.exp import AbsResolver, Exp
from utilz.xwu import NamedLists, NamedRanges, appmgr, find, offset, usedrange

from .common import _logger as logger


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
