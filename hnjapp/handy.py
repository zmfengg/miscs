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
from os import listdir, makedirs, path, rename, sep, utime, walk
from re import compile as compile_r
from shutil import copy

from PIL import Image
from sqlalchemy.orm import Query

from hnjapp.c1rdrs import _fmtbtno
from hnjcore.models.hk import JO, Orderma, PajShp, Style
from utilz import ResourceCtx, getfiles, trimu
from utilz.xwu import NamedLists, appmgr, find, usedrange

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
        return None

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
        jns, lst = Query([PajShp.pcode, Style.name.label("styno"), JO.name.label("jono")]).join(JO).join(Orderma).join(Style).filter(PajShp.pcode.in_(pcs)).with_session(cur).all(), set()
        for x in lst:
            fn = x.jono.name
            if fn in jns:
                continue
            jns.add(fn)
            fn = pcs[x.pcode]
            if not path.exists(fn):
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

        _jn_str = lambda x: "%d" % int(x) if isinstance(x, Number) else x
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
