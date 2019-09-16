'''
#! coding=utf-8
@Author:        zmFeng
@Created at:    2019-08-09
@Last Modified: 2019-08-09 2:10:55 pm
@Modified by:   zmFeng

ack related
'''
from csv import DictReader, DictWriter
from datetime import date, datetime
from numbers import Real
from os import listdir, mkdir, path
from random import randint
from tempfile import TemporaryFile, gettempdir
from collections import defaultdict

from sqlalchemy import desc
from sqlalchemy.orm import Query
from xlwings.constants import (FormatConditionOperator, FormatConditionType,
                               LineStyle, HAlign)

from hnjapp.pajcc import MPS, PajCalc, PrdWgt, WgtInfo
from hnjcore import JOElement
from hnjapp.svcs.db import jesin
from hnjcore.models.hk import JO, PajAck, PajInv, PajShp, Orderma, Style
from utilz import xwu
from utilz.miscs import NamedList, NamedLists, getfiles, na, newline, tofloat, trimu

from .common import _logger as logger
from .common import config

try:
    import matplotlib as mpl
    from matplotlib import pyplot as plt
    from matplotlib.backends.backend_pdf import PdfPages
    import pandas as pd
except:
    pass


def _fetch_invs(hksvc, fldr, pcodes):
    pcodes = [x for x in pcodes if x[0] != '#']
    if not pcodes:
        return None
    fn_cns = path.join(fldr, '_cns.csv')
    cns = 'pcode styno jono qty description invdate mps uprice china mtlcost ocost'.split()
    df = pd.read_csv(fn_cns, parse_dates=['invdate']) if path.exists(fn_cns) else pd.DataFrame(None, columns=cns)
    lns = {x for x in pcodes if x[0] != '#'} - set(df.pcode)
    if lns:
        wgtsvc, lsts = LocalJOWgts(path.join(fldr, '_wgts.csv'), hksvc), []
        def _flush(df, lsts):
            if lsts:
                df = df.append(pd.DataFrame(lsts, columns=df.columns))
                df.to_csv(fn_cns, index=None)
            return df, []
        with hksvc.sessionctx() as cur:
            for pcode in lns:
                try:
                    lst = cur.query(JO, PajShp, PajInv).join(PajShp).join(PajInv).filter(PajShp.pcode == pcode).order_by(desc(PajShp.invdate)).all()
                except UnicodeDecodeError:
                    print('failed to get invoice for pcode(%s) because of encoding error, to next' % pcode)
                    continue
                if not lst:
                    lsts.append([pcode, na, na, 0, na, date.today(), na, 0, 0, 0, 0])
                    continue
                for jo, shp, inv in lst:
                    wgts = wgtsvc.wgts(pcode)
                    cn = PajCalc.calchina(wgts, inv.uprice, inv.mps, shp.invdate)
                    if not cn:
                        print('failed to calc cn for JO(%s)' % jo.name.value)
                        continue
                    lsts.append([pcode, jo.orderma.style.name.value, jo.name.value, jo.qty, jo.description.replace(' ', ''), shp.invdate, inv.mps, inv.uprice, cn.china, cn.metalcost, cn.china - cn.metalcost])
                if len(lsts) % 10 == 0:
                    df, lsts = _flush(df, lsts)
                    logger.debug('%d invoice fetched' % len(lsts))
        if lsts:
            _flush(df, lsts)
        # an unknown but appear in the liner() method if df is not loaded
        df = pd.read_csv(fn_cns, parse_dates=['invdate'])
    return df.loc[df.pcode.isin(pcodes)]

class AckPriceCheck(object):
    """
    check given folder(not the sub-folders)'s acks. I'll firstly check if
    the folder has been analysed. if yes, no thing will be done
    """
    _fnsrc, _fndts = "_src.dat", "_fdates.dat"

    @property
    def fnsrc(self):
        return path.join(self._fldr, self._fnsrc)

    @property
    def fndts(self):
        return path.join(self._fldr, self._fndts)
    
    def __init__(self, fldr, hksvc):
        self._fldr = fldr
        self._hksvc = hksvc
        self._fmtr = _AckFmt()
        self._nl_src = NamedList("jono,date,file,mps,pcode,pajprice,styno,qty")
        self._nl_rst = self._fmtr.nl_rst
        self._wgtsvc = LocalJOWgts(path.join(self._fldr, '_wgts.csv'), hksvc)

    def _uptodate(self, fns):
        ''' return True if data in fnsrc is up-to-date
        '''
        if not fns:
            return True
        fn = self.fndts
        flag = path.exists(fn)
        if flag:
            with open(fn, "r") as fh:
                rdr = DictReader(fh)
                mp = {x["file"]: x for x in rdr}
            flag = len(fns) == len(mp)
            for x in fns:
                fn = path.basename(x)
                flag = fn in mp
                if not flag:
                    break
                flag = float(mp[fn]["date"]) >= path.getmtime(x)
                if not flag:
                    break
        return flag

    def _get_src_fns(self):
        fns = getfiles(self._fldr, "xls", True)
        if fns:
            fns = [path.join(self._fldr, x) for x in fns if x.find("_") != 0]
        return fns

    def persist(self, srcs=None, jos=None):
        ''' save the ack source data to db

        Args:

            srcs=None:   the data from self._read_srcs()

            jos=None:    a tuple of JO instance that present JOs inside the source files

        '''
        if not srcs:
            var = self._get_src_fns()
            srcs = self._read_srcs(var, self._uptodate(var))
        if not srcs or isinstance(srcs, str):
            return None
        logger.debug("begin persisting")
        fds = {x: datetime.fromtimestamp(
                path.getmtime(path.join(self._fldr, x))).replace(
                second=0, microsecond=0)
                    for x in listdir(self._fldr)}
        def _newinst(td):
            ins = PajAck()
            ins.tag, ins.filldate = 0, td
            return ins

        with self._hksvc.sessionctx() as cur:
            try:
                var = set(x.file for x in srcs.values())
                dds = Query([PajAck.docno, PajAck.lastmodified]).filter(
                    PajAck.docno.in_(list(var))).distinct()
                dds = dds.with_session(cur).all()
                dds = {x[0]: x[1] for x in dds}
            except:
                dds = {}
            exps, jes = set(), set()
            lst, td = [], datetime.today()
            for x in srcs.values():
                if not (x.pajprice and x.date and x.mps):
                    continue
                var = x.file
                if var in dds:
                    if dds[var] >= fds[var]:
                        continue
                    exps.add(var)
                ins = _newinst(td)
                ins.lastmodified = fds[var]
                ins.ackdate, ins.mps = (x[y] for y in ("date", "mps"))
                ins.joid = JOElement(x.jono)
                if not jos:
                    jes.add(ins.joid)
                ins.uprice, ins.docno, ins.pcode = x.pajprice, var, x.pcode[0]
                if ins.uprice < 0:
                    ins.uprice = -1
                lst.append(ins)
            if exps:
                cur.query(PajAck).filter(PajAck.docno.in_(
                    list(exps))).delete(synchronize_session=False)
            if lst:
                if not jos:
                    jos = self._hksvc.getjos(jes)
                    if jos[1]:
                        for x in jos[1]:
                            logger.debug("invalid JO#(%s)" % x.value)
                        return None
                    jos = jos[0]
                flag, jos = True, {x.name: x for x in jos}
                try:
                    for x in lst:
                        x.joid = jos[x.joid].id
                        cur.add(x)
                    cur.flush()
                except Exception as err:
                    flag = False
                    logger.debug(
                        "error occur while persisting: %s" % err)
                finally:
                    if flag:
                        cur.commit()
                        logger.debug("persisted")
                    else:
                        cur.rollback()
                        lst = None
        return lst

    def analyse(self, tar_fn='_rst', persist=True):
        """ execute the check process against the given folder, return the result's full filename

        Args:

            tar_fn='_rst': target file name(without path) to save in self._fldr

            persist=True: True to persist the result

        """
        err = "process of folder(%s) completed" % path.basename(self._fldr)
        fns = self._get_src_fns()
        if not fns:
            logger.debug(err + " but no source file found")
            return None
        rc, utd = getfiles(self._fldr, tar_fn), self._uptodate(fns)
        if rc:
            rc = rc[0]
        if rc and utd and \
            path.getmtime(rc) >= path.getmtime(self.fndts):
            logger.info(err)
            logger.info("data up to date, don't need further process")
            return rc
        rc = err = jos = None
        logger.debug("begin to analyse acknowledgements (%s)" % self._fldr)
        tk = xwu.appmgr.acq()[1]
        with self._hksvc.sessionctx():
            # keep session to make sure self.persist() can still access the JOs
            try:
                srcs = self._read_srcs(fns, utd)
                if isinstance(srcs, str):
                    err = srcs[1:]
                elif srcs:
                    logger.debug("%d JOs returned" % len(srcs))
                    rc = self._process_all(srcs)
                    if rc:
                        jos = rc[1]
                        rc = self._write_wb(rc[0], path.join(self._fldr, tar_fn))
            finally:
                xwu.appmgr.ret(tk)
            if err:
                logger.info("exception(%s) occured" % err[1:])
            elif rc and persist:
                self.persist(srcs, jos)
        return rc

    def _read_dat(self, datfn):
        mp = {}
        if not path.exists(datfn):
            return None
        with open(datfn) as fh:
            lns = fh.readlines()
            nl = newline(lns[0])
            nl = [x[:nl].split(',') for x in lns]
            nls = NamedLists([x for x in nl if len(x[0])])
            for nl in nls:
                jono = nl.jono
                nl0 = mp.get(jono)
                if not nl0:
                    mp[jono] = nl0 = nl
                    nl0.pajprice = float(nl0.pajprice)
                    nl0.qty = float(nl0.qty)
                    nl0.pcode = [nl0.pcode]
                    nl0.date = self._fmtr.d2s(nl0.date)
                else:
                    nl0.qty += float(nl.qty)
                    nl0.pcode.append(nl.pcode)
        return mp

    def _write_dat(self, mp, fn):
        '''
        Args:
            mp: a {name(string): NL} or a collection of NL
        '''
        with open(fn, "w") as fh:
            cns = wtr = None
            for nl in mp.values() if isinstance(mp, dict) else mp:
                if not cns:
                    cns = nl.colnames
                    wtr = DictWriter(fh, cns, dialect="excel")
                    wtr.writeheader()
                dct = {cn: nl[cn] for cn in cns if cn != 'date'}
                dct['date'] = self._fmtr.d2s(nl.date)
                idx = 0
                for pc in tuple(nl.pcode):
                    if idx > 0:
                        dct["qty"] = 0
                    dct["pcode"] = pc
                    wtr.writerow(dct)
                    idx += 1

    def _read_src(self, fn, app, fds):
        bfn = path.basename(fn)
        logger.debug("Reading file(%s)" % bfn)
        fds[bfn] = str(path.getmtime(fn))
        wb, shtcnt, data, err = xwu.safeopen(app, fn), 0, {}, None
        for sht in wb.sheets:
            adate, sp, gp = None, 0, 0
            adate = self._getvalue(sht, "Order Date:")
            sp, gp = (tofloat(self._getvalue(sht, x)) for x in ("Silver*:", "gold*:"))
            if not (adate and any((sp, gp))):
                if any((adate, sp, gp)):
                    err = "Key argument missing in (%s)" % bfn
                    logger.debug("sheet(%s) has not enough arguments" % bfn)
                break
            shtcnt += 1
            mps = MPS("S=%f;G=%f" % (sp, gp)).value
            #don't use the NO field, sometimes it's blank, use JO# instead
            rng = xwu.find(sht, "Job*")
            nls = NamedLists(xwu.usedrange(sht).value[rng.row - 1:],\
                self._fmtr.reader_def)
            exc = set(('jono', 'qty', 'pcode', 'pajprice', ))
            for nl in nls:
                jono = nl.jono
                if not jono:
                    break
                it = self._nl_src.clone()
                jono = JOElement(jono).value
                if jono in data:
                    it = data[jono]
                else:
                    data[jono] = it
                    it.pajprice = max(nl.pajprice or 0, 0)
                    it.pcode, it.jono, it.qty = [], jono, 0
                    it.file, it.mps, it.date = bfn, mps, adate
                for cn in it.colnames:
                    if cn not in exc and cn in nl.colnames:
                        it[cn] = nl[cn]
                it.qty += tofloat(nl.qty)
                if nl.pcode not in it.pcode:
                    it.pcode.append(nl.pcode)
        if shtcnt <= 0:
            logger.critical("file(%s) doesn't contains any valid sheet" % bfn)
        wb.close()
        return data, err

    def _read_srcs(self, fns, utd):
        """read necessary data from the ack excel file.

        Args:

            fns:    a collection of file name

            udt:    use dat file instead of the source files

        Returns:

            When no error occur, returns {jono, nl} where namedlist with below columns:

                "jono,pajprice,file,mps,styno,date,qty,pcode"

            When there is error, return err(string)
        """

        if not fns:
            return None, None
        mp = None
        if utd and path.exists(self.fnsrc):
            mp = self._read_dat(self.fnsrc)
        if mp:
            return mp
        mp, fds = {}, {}
        app, kxl = xwu.appmgr.acq()
        try:
            err = None
            for fn in fns:
                dx, err = self._read_src(fn, app, fds)
                if err:
                    break
                mp.update(dx)
        except Exception as e:
            err = "file(%s),err(%s)" % (path.basename(fn), e)
        if kxl:
            xwu.appmgr.ret(kxl)
        if err:
            return "_" + err
        if mp:
            self._write_dat(mp, self.fnsrc)
            lst = [','.join(x) for x in fds.items()]
            lst.insert(0, "file,date")
            with open(self.fndts, "w") as fh:
                for x in lst:
                    print(x, file=fh)
        return mp

    def _process_one(self, nl_src, jo_mp, smlookup=False):
        """
        Ack based on revised/history/similar

        Args:

        nl_src: NamedList of the source data

        jo_mp: {jono(string), {"jo,mps,wgts".split(), ...}

        Returns:

            the result NamedList
        """
        pajup = nl_src.pajprice
        nl = self._nl_rst.clone()
        pdx = jo_mp.get(nl_src.jono)
        for cn in 'date file jono mps pajprice qty styno'.split():
            nl[cn] = nl_src[cn]
        if pdx and pdx['poprice'] > 0:
            nl.wgts = pdx["wgts"]
            x = pdx["poprice"]
            if x:
                nl.poprice = x
                nl.profit = x - pajup
                nl.profitt = nl.profit * nl_src.qty
                if pajup:
                    nl.ratio = x / pajup * 100.0
            nl.result = self._fmtr.classify_pft(pajup, x)
        else:
            nl.wgts = pdx["wgts"] or self._wgtsvc.wgts(nl_src.pcode)
            nl.result = self._fmtr.classify_pft(pajup, 0)

        pfx = self._p_frm_rev(nl_src, nl)
        if not pfx:
            pfx = self._p_frm_his(nl_src, nl)
        if not pfx:
            if smlookup:
                pfx = self._p_frm_sml(nl_src, nl, jo_mp)
        if not pfx:
            # new item, create mock place for it
            nl.ref = self._fmtr.label('labels', 'rf.noref')
            nl.pcode = nl_src.pcode[0]
        else:
            pfx, tar = pfx
            nl.expected = tar.china
            nl.diff = nl_src.pajprice - tar.china
            nl.ref = pfx  + self._fmtr.classify_ref(pajup, tar.china)
            for cn in self._nl_rst.colnames:
                val = nl[cn]
                if isinstance(val, Real):
                    nl[cn] = self._fmtr.rd(val)
        return nl

    def _p_frm_rev(self, nl_src, nl):
        cn = None
        for pcode in nl_src.pcode:
            cn = self._hksvc.getrevcns(pcode)
            if cn:
                nl.pcode = pcode
                break
        if not cn:
            return None
        revs = cn
        cn = PajCalc.newchina(cn[0].uprice, nl.wgts)
        tar = PajCalc.calctarget(cn, nl_src.mps, affdate=nl_src.date)
        nl.rev = self._mk_rev(revs[0].uprice, revs[0].revdate)
        if "revs" in nl.colnames and len(revs) > 1:
            nl.revs = ",".join([self._mk_rev(x.uprice, x.revdate)
                for x in revs
            ])
        return self._fmtr.label('labels', 'rf.rev'), tar

    def _p_frm_his(self, nl_src, nl):
        for pcode in nl_src.pcode:
            shp = self._hksvc.getpajinvbypcode(pcode, \
                maxinvdate=nl_src.date, limit=2)
            if shp:
                shp = shp[0]
                nl.pcode = pcode
                break
        if not shp:
            return None
        # nl.wgts don't need changes because same pcode should have same wgts
        tar = self._calc_his(shp, nl_src, nl)
        return self._fmtr.label('labels', 'rf.his'), tar

    def _mk_rev(self, cn, invdate, jo=None):
        if not jo:
            return '%4.2f@%s' % (cn, self._fmtr.d2s(invdate))
        return '%s:%4.2f@%s' % (jo.name.value, cn, self._fmtr.d2s(invdate))

    def _calc_his(self, shp, nl_src, nl):
        '''
        Args:
            nl: the result NamedList. To make correct answer, put wgt of history's pcode in nl.wgts before calling me
        '''
        refup, refmps = shp.PajInv.uprice, MPS(shp.PajInv.mps)
        cn = PajCalc.calchina(nl.wgts, refup, refmps, shp.PajShp.invdate)
        nl.rev = self._mk_rev(cn.china, shp.PajShp.invdate, shp.JO)
        return PajCalc.calctarget(cn, nl_src.mps, nl_src.date)


    def _p_frm_sml(self, nl_src, nl, jo_mp):
        jos = None
        try:
            jos = self._hksvc.findsimilarjo(jo_mp[nl_src.jono]["jo"], level=1)
            # the result is sorted by fill_date desc
        except:
            pass
        if not jos:
            return None
        for jo in jos:
            shp = self._hksvc.getpajinvbyjes((jo.name, ))
            if not shp:
                continue
            shp = shp[0]
            jono0, wgt0 = nl.jono, nl.wgts
            # change wgts to history
            nl.wgts = self._fetchjos([jo, ])[jo.name.value]['wgts']
            nl_srcx = nl_src.clone(False)
            nl_srcx.jono = jo.name.value
            nl.pcode = shp.PajShp.pcode
            tar = self._calc_his(shp, nl_srcx, nl)
            nl.wgts, nl.jono = wgt0, jono0
            return self._fmtr.label('labels', 'rf.sml'), tar
        nl.ref = self._fmtr.label('labels', 'rf.noref')
        return None

    def _fetchjos(self, jos):
        """
        return a dict with jono as key and dict with columns: "poprice,mps,wgts"
        """
        mp = {}
        for jo in jos:
            try:
                mp[jo.name.value] = {"jo": jo,
                    "poprice":float(jo.po.uprice),
                    "mps": jo.poid,
                    "wgts": self._hksvc.getjowgts(jo)
                }
            except:
                logger.debug('failed to get po/wgts for JO(%s), maybe BIG5 encoding error' % jo.name.value)
        return mp

    def _process_all(self, src_mp):
        """
        @param all: a dict with jono as key and a dict with these keys: "jono,pajprice,file,mps,styno,date,qty,pcode". ref @_read_srcs() FMI.
        """
        if not src_mp:
            return None
        hksvc, rsts = self._hksvc, {}
        with hksvc.sessionctx():
            nls = src_mp.values()
            jo_mp = [JOElement(x.jono) for x in nls]
            jo_mp = self._fetchjos(hksvc.getjos(jo_mp)[0])
            jos = [x['jo'] for x in jo_mp.values()]
            idx = 0
            for nl_src in nls:
                try:
                    nl = self._process_one(nl_src, jo_mp, True)
                    rsts.setdefault(nl.result, []).append(nl)
                    idx += 1
                    if idx % 10 == 0:
                        logger.info("%d of %d done" % (idx, len(nls)))
                except:
                    nl_src.pcode = nl_src.pcode[0]
                    rsts.setdefault(self._fmtr.label('labels', 'prg.error'), []).append(nl_src)
        return rsts, jos

    def _write_wb(self, rsts, fn):
        app, kxl = xwu.appmgr.acq()

        wb = app.books.add()
        self._write_readme(wb)
        for k, v in rsts.items():
            self._write_sht(k, v, wb)

        _ex = self._fmtr.label('labels', 'prg.error')
        # prg.error contains source result set only, so void it
        lst = [tmp for k, v in rsts.items() for tmp in v if k != _ex]
        _ex = lambda val, kwds: [0 for k in kwds if val.find(k) >= 0]
        _exs = lambda kwds: [k for k in lst if _ex(k.ref, kwds)]
        mp = {
            "_except": _exs(self._fmtr.setting('pft.ref.classify')['labels'][1:]),
            "_new_sml": _exs((self._fmtr.label('labels', 'rf.sml'), )),
            "_new": _exs((self._fmtr.label('labels', 'rf.noref'), )),
        }
        if mp['_except']:
            # when there is exception, prepare one for enqurying zhengyuting
            tmp = _exs([self._fmtr.label('labels', x) for x in ("rf.rev", "rf.his", "rf.sml")])
            tmp = [k.clone(False) for k in tmp if self._fmtr.classify_ref(k.pajprice, k.expected, False) > 0]
            for k in tmp:
                k.ref = k.ref.split('_')[0] # remove the profit result from REF for PAJ
            mp['_paj_enq'] = tmp
        tmp = set()
        # when item in higher level, don't show them in lower one
        for k, v in mp.items():
            if k == '_paj_enq':
                lst = v
            else:
                lst = []
                for nl in v:
                    var = nl.jono
                    if var not in tmp:
                        tmp.add(var)
                        lst.append(nl)
            if not lst:
                continue
            self._write_sht(k, lst, wb)
        for var in wb.sheets:
            xwu.usedrange(var).row_height = 18
        if fn:
            wb.save(fn)
            fn = wb.fullname
        wb.close()
        xwu.appmgr.ret(kxl)
        return fn

    def _write_sht(self, name, nls, wb):
        if not nls:
            return
        lst = []
        excl = self._fmtr.excludes(name) or ()
        hdr = next(iter(nls)).colnames
        if excl:
            hdr = [x for x in hdr if x not in excl]
        if len(nls) > 1:
            nls = self._fmtr.sort(name, nls)
        hdr_ttl = [self._fmtr.label('cat.cns', x) for x in hdr]
        hdr_ttl = [x[0] if isinstance(x, list) else x for x in hdr_ttl]
        lst.append(hdr_ttl)
        hdr_mp = {x: idx for idx, x in enumerate(hdr)}
        for nl in nls:
            nl = nl.clone(False)
            nl.jono = "'" + nl.jono
            if 'ratio' in nl.colnames and nl.ratio is not None: # ERROR without ratio or wgts
                nl.ratio = "%4.2f%%" % (nl.ratio or 0)
            nl.wgts = ';'.join('%s=%4.2f' % (x.karat, x.wgt) for x in nl.wgts.wgts if x and x.wgt)
            lst.append([nl[x] for x in hdr])
        sht = wb.sheets.add(self._fmtr.label('cats', name) or name)
        sht.range(1, 1).value = lst
        fidx = hdr_mp.get('file', None)
        if fidx is not None:
            for idx in range(1, len(lst)):
                rng = sht.cells[idx, fidx]
                x = rng.value
                rng.add_hyperlink(x)
                rng.value = x.split()[1]
        if name != self._fmtr.label('labels', 'prg.error'):
            self._fmtr.adj_colwidth(sht, hdr_mp)
        else:
            sht.autofit('c')
        xwu.freeze(sht.range(2, 4), False)
        xwu.maketable(xwu.usedrange(sht))
        if name == '_except':
            sht.activate()
            rng = xwu.usedrange(sht)
            rng.select()
            fidx = self._fmtr.setting('pft.ref.classify')['labels'][-1]
            fidx = '=SEARCH("%s",$%s1)>0' % (fidx, xwu.col(hdr_mp['ref'] + 1))
            rng = rng.api
            rng.formatconditions.add(FormatConditionType.xlExpression, FormatConditionOperator.xlEqual, fidx)
            rng.formatconditions(rng.formatconditions.count).interior.colorindex = 40
            sht.cells[1, 0].select()
            self._write_his(nls, wb.sheets.add())

    def _write_his(self, nls, sht):
        ''' write the history of exception pcodes to a sheet and plot
        '''
        fn = path.join(self._fldr, '_invhis.csv')
        df = pd.read_csv(fn, parse_dates=['invdate', ]) if path.exists(fn) else  pd.DataFrame(None, columns='pcode jono invdate uprice mps china mtlcost ocost'.split())
        reqs = [nl.pcode for nl in nls if df.loc[df.pcode == nl.pcode].empty]
        if reqs:
            var, nls = [], {nl.pcode: nl for nl in nls}
            df = _fetch_invs(self._hksvc, self._fldr, reqs)
            # append myself to show this trend
            for pcode in reqs:
                wgts = self._wgtsvc.wgts(pcode)
                nl = nls[pcode]
                cn = PajCalc.calchina(wgts, nl.pajprice, nl.mps, nl.date)
                var.append([pcode, '*' + nl.jono, nl.qty, nl.date, nl.pajprice, nl.mps, cn.china, cn.metalcost, cn.china - cn.metalcost])
            df.append(pd.DataFrame(var, columns=df.columns))
        ttl = 'Other Cost (China - Metal) Trend of "%s"' % path.basename(self._fldr)
        _HisPltr(df, title=ttl).plot(path.join(self._fldr, '_invHis.pdf'), sht, method='changes')
        try:
            sht.name = '_InvHis'
        except:
            pass

    def _write_readme(self, wb):
        """
        create a readme sheet
        """
        mp = self._fmtr.setting('pft.ref.classify')
        ab, rel, lbl = (mp[x] for x in ('absolute', 'relative', 'labels'))
        cnt = len(ab)
        lst = [("Ref. Classifying:", "", "")]
        lst.append("Ref.Suffix,Diff$,DiffRatio".split(","))
        for ii in range(cnt):
            lst.append((lbl[ii], "'%s" % ab[ii], "'%s%%" % (rel[ii]*100)))
        lst.append((lbl[cnt], "'-", "'-"))

        def _mtb(rng, data):
            rng.value = data
            xwu.maketable(rng.offset(1, 0).expand('table'))

        sht = wb.sheets.add("Readme")
        _mtb(sht.cells[0, 0], lst)

        rowidx = len(lst) + 2
        lst = ["Ref.Prefix,Meaning".split(",")]
        _lbl = self._fmtr.label
        lst.append((_lbl('labels', 'rf.rev'), "Found in PAJ's revised files"))
        lst.append((_lbl('labels', 'rf.his'),
                    "Not in PAJ's revised files, but has invoice history"))
        lst.append((_lbl('labels', 'rf.noref'), "No any PAJ price reference data"))
        _mtb(sht.range(rowidx, 1), lst)

        rowidx += len(lst) + 1
        pfr = self._fmtr.setting("misc")["pft.min.ratio"]
        pfr = "%s%%" % (pfr * 100)
        lst = [("Profit Margin(POPrice/PAJPrice) Classifying", "")]
        lst.append(("Spc. Sheet", "Meaning"))
        lst.append((_lbl('labels', 'pft.normal'),
                    "Profit margin greater or equal than %s" % pfr))
        lst.append((_lbl('labels', 'pft.low'),
                    "Profit margin less than %s" % pfr))
        lst.append((_lbl('labels', 'pft.error'),\
                    "Not enough data for profit calculation"))
        _mtb(sht.range(rowidx, 1), lst)

        rowidx += len(lst) + 1
        lst = [("Spc. Sheet records are already inside other sheet", "")]
        lst.append(("Spc. Sheet", "Meaning"))
        lst.append((_lbl('cats', '_new'),\
                "Item does not have any prior PAJ price data"))
        lst.append((_lbl('cats', '_except'),
                "PAJ price exception with rev./previous data"))
        _mtb(sht.range(rowidx, 1), lst)

        for c, w in ((0, 15), (1, 24), (2, 10), ):
            sht.cells[1, c].column_width = w

        for sht in wb.sheets:
            if sht.name.lower().find("sheet") >= 0:
                sht.delete()

    def _getvalue(self, sht, kw, direct="right"):
        rng = xwu.find(sht, kw)
        if not rng:
            return
        return rng.end(direct).value

    @classmethod
    def plot_history(cls, pcodes, prefix, hksvc, fldr=None, ttl=None, method=None):
        ''' plot given pcodes to folder/_rst as excel + pdf(chart), a folder _rst will be created under given folder
        Args:

            pcodes: a collection of pcodes. use:
                JO:JO# to indicate JO#
                STY:Sty# to indicate Sty#

            preifx: prefix of the result file name

            hksvc: HKSvc instance help to retrieve JO data in necessary

            fldr=None: the folder to hold the result file, omitting will be set to temp folder

            ttl:    the title appear in the chart

            method=None: method for filtering the history data, can be one of:
                None:   the same as 'all'
                'all':  all histories
                'changes': only the history that contains changes
                'up':   the history whose last value is less than prior
                'down': the history whose last value is greater than prior
        '''
        if not fldr:
            fldr = gettempdir()
        # check if there is title or method
        if not (ttl and method):
            tr = [x for x in pcodes if x[0] == '#']
            if tr:
                lns = {x.split('=')[0]: x.split('=')[1] for x in tr[0][1:].split(';')}
                if not ttl:
                    ttl = lns.get('title')
                if not method:
                    method = lns.get('method', 'all')
            else:
                if not ttl:
                    ttl = prefix
                if not method:
                    method = 'all'
        app = tk = pltr = None
        s_fldr = path.join(fldr, '_rst')
        if not path.exists(s_fldr):
            mkdir(s_fldr)
        for mt in method.split(','):
            fn = prefix + '_%s' % mt
            fn_xls = path.join(s_fldr, fn + '.xlsx')
            if path.exists(fn_xls):
                continue
            # translate those JO# or Sty# to pcodes
            var = [x for x in pcodes if x.find('title') < 0 and x.find('method') < 0 and (len(x) != 17 or x.find(':') > 0)]
            if var:
                pcodes = cls._extend(pcodes, var, hksvc)
            if not pltr:
                pltr = _HisPltr(_fetch_invs(hksvc, fldr, pcodes), title=ttl)
                app, tk = xwu.appmgr.acq()
            wb = app.books.add()
            pltr.plot(path.join(s_fldr, fn + '.pdf'), wb.sheets[0], method=mt)
            wb.save(fn_xls)
            wb.close()
        if tk:
            xwu.appmgr.ret(tk)

    @classmethod
    def _extend(cls, pcodes, other, hksvc):
        apcs = [x for x in pcodes if x not in set(other)]
        mp = defaultdict(list)
        for x in (x.split(':') for x in other):
            t, n = ('JO', x[0]) if len(x) < 2 else (trimu(x[0]), x[1])
            mp[t].append(n)
        exted = []
        with hksvc.sessionctx() as cur:
            for t, ns in mp.items():
                if t == 'JO':
                    lst = cur.query(PajShp.pcode).join(JO).filter(jesin([JOElement(x) for x in ns], JO)).distinct().all()
                elif t == 'STY':
                    lst = cur.query(PajShp.pcode).join(JO).join(Orderma).join(Style).filter(jesin([JOElement(x) for x in ns], Style)).distinct().all()
                if lst:
                    exted.extend([x[0] for x in lst])
                    lst = None
        return list(set(apcs + exted))


class _AckFmt(object):
    ''' class for Ack's format settings
    '''
    def __init__(self):
        self._cfg_root = mp = config.get('pajcc.ack.chk')
        cns = mp['cat.cns']
        self._cns = {key: (val if isinstance(val, list) else (val, 0))\
            for key, val in cns.items()}
        tmp = mp['cat.excl']
        self._cat_excl = {x[0]: set(self._parse(x[1], tmp).split(',')) for x in\
            tmp.items()}
        self.nl_rst = NamedList([x for x in self._cns])

    @property
    def reader_def(self):
        return self._cfg_root.get('reader.def')

    def setting(self, name):
        ''' return the sub-setting under cfg_root

        Args:

            name:   name of the key
        '''
        return self._cfg_root.get(name)


    def d2s(self, d0):
        ''' date to string or verse
        '''
        if not isinstance(d0, (str, datetime, date)):
            return d0
        df = self._cfg_root['misc']["date.fmt"]
        if isinstance(d0, str):
            return datetime.strptime(d0, df)
        return d0.strftime(df)

    def excludes(self, cat):
        ''' the exclude columns of given cat
        '''
        return self._cat_excl.get(cat.lower())

    def adj_colwidth(self, sht, hdr_idx):
        ''' setup the column widths of given sheet

        Args:

            hdr_mp: a {colname(string), colIdx(int)} map

        '''
        sht.autofit('c')
        cands = [(n, spec) for n, spec in self._cns.items() if n in hdr_idx and spec[1]]
        if not cands:
            return
        for n, spec in cands:
            sht.cells[0, hdr_idx[n]].column_width = spec[1]

    def sort(self, cat, nls):
        ''' sort nls based on the cat and definations in conf.json

        Args:

            cat:    the cat name
            nls:    NamedLists to be sorted
        '''
        if not nls or len(nls) < 2:
            return nls
        sdef = self.setting("cat.sorting")
        cns = self._parse(sdef.get(cat) or sdef.get('_default'), sdef)
        cnsx = next(iter(nls)).colnames
        cns = [x for x in cns if x in cnsx]
        return sorted(nls, key=lambda nl: ",".join([str(nl[x]) for x in cns]))

    def rd(self, d0):
        ''' round the argument by 2
        '''
        return round(d0, 2)

    @classmethod
    def _parse(cls, x, mp):
        if not x:
            return None
        if isinstance(x, list):
            rst = []
            for y in x:
                rst.append(cls._parse(y, mp))
            return rst
        if isinstance(x, str):
            if x[:2] == '${':
                return cls._parse(mp[x[2:-1]], mp)
        return x

    def label(self, cat, name=None, args=None, ret_all=False):
        ''' return the label by given name
        Args:

        cat:    category iame, sth. like cat.cns

        args=None:   argument to format the return label

        ret_all=False:  return all the labels

        '''
        mp = self.setting(cat)
        if not mp:
            return None
        if ret_all:
            lst = []
            for x in mp.values():
                if not isinstance(x, (str, list)):
                    continue
                lst.append(self._parse(x, mp))
            return lst
        lbl = self._parse(mp.get(name), mp)
        if not lbl:
            return lbl
        if args:
            lbl = lbl % args
        return lbl

    def classify_pft(self, pajup, poup):
        ''' profit classify

        Args:

            pajup:  the paj unit price

            poup:   the PO unit price
        '''
        if not (poup and pajup):
            return self.label('labels', 'pft.error')
        lvl = self.setting('misc')['pft.min.ratio']
        cat = 'pft.normal' if poup / pajup >= lvl else "pft.low"
        return self.setting('labels')[cat]

    def classify_ref(self, pajup, expup, rtTtl=True):
        """return a classified string based on pajuprice/expected

        Args:

            rtTtl=True: True to return title, else index
        """
        diff = pajup - expup
        rdiff = diff / expup
        flag = False
        cfg = self.setting('pft.ref.classify')
        anr = [cfg[x] for x in ('absolute', 'relative')]
        for ii, ab in enumerate(anr[0]):
            if diff <= ab and rdiff <= anr[1][ii]:
                flag = True
                break
        if not flag:
            ii = len(anr[0])
        return cfg['labels'][ii] if rtTtl else ii

class _HisPltr(object):
    ''' save given dataframe to an existing worksheet and product a
    chart file as pdf
    Args:

        df: a DataFrame that should contains at least these columns:
            pcode, invdate, ocost
    '''

    def __init__(self, df, **kwds):
        self._df = df.drop_duplicates(['pcode', 'jono', 'invdate'])
        self._args = args = config.get('pajcc.ack.chk')['plot.his.args'].copy()
        args.update(kwds)
        args['page_size'] = [x / 25.4 for x in args['page_size']]

    def _group(self, pvs):
        grps = []
        def _ngrp():
            grps.append([])
            return grps[-1], 0, 0
        lst, cnt, lb = _ngrp()
        for pv in pvs:
            if cnt >= self._args['lines_per_ax']:
                lst, cnt, lb = _ngrp()
            if lb and pv[1] - lb > self._args['max_diff_per_ax']:
                lst, cnt, lb = _ngrp()
            if not lb:
                lb = pv[1]
            lst.append(pv[0])
            cnt += 1
        return grps

    def _plt_fmt(self):
        mpl.rcParams.update({'font.size': self._args['font.size']})
        # mpl.rcParams['figure.constrained_layout.use'] = True
        # After using plt.subplots_adjust, the figure's size won't work
        # plt.subplots_adjust(top=0.92, bottom=0.08, left=0.08, right=0.92, wspace=0.1) make the figure un-scalable

    def _liner(self, df, method='all'):
        _to_discard, pcode, modcnt = [], None, 0
        th = config.get('pajcc.ack.chk')['pft.ref.classify']
        th = [th[x][0] for x in ("relative", "absolute")]
        def _discard():
            if not pcode:
                return
            flag = False
            if method != 'all':
                flag = modcnt < 2
            if not flag:
                if method == 'up':
                    flag = trend != 1
                elif method == 'down':
                    flag = trend != -1
            if flag:
                _to_discard.append(pcode)

        for idx, row in df.iterrows():
            if row.pcode != pcode:
                _discard()
                trend = loc = modcnt = 0
                pcode = row.pcode
            if loc and abs(loc / row.ocost - 1) < th[0] and abs(loc - row.ocost) < th[1]:
                df.loc[idx, 'ocost'] = loc
                # below 2 methods might has warning or does not work at all
                # df.ocost[idx] = loc
                # row.ocost = loc won't work because it's a view
            else:
                if loc and (method == 'up' or method == 'down'):
                    trend = -1 if loc > row.ocost else 1
                loc = row.ocost
                modcnt += 1
        _discard()
        if _to_discard:
            # in/notin query can be df.pcode.isin(mx)
            # notin: ~df.pcode.isin(mx)
            # df.query('pcode not in @mx')
            df = df.loc[~df.pcode.isin(_to_discard)]
        return df

    def plot(self, fn=None, sht=None, method='changes'):
        ''' plot to a file
        Args:
            fn=None: the pdf file to plot to, omitting will create a temp file
            sht=None: the sheet to write data to, omitting won't produce any excel thing
            method='changes': what to show, can be one of:
                'all'     -> any series
                'changes' -> the series that contains changes
                'up'      -> the series that belongs to up-trend
                'down'    -> the series that belongs to down-trend
        '''
        df = self._df.sort_values(by=['styno', 'pcode', 'invdate'])
        df = self._liner(df, method)
        if df.empty:
            return None
        self._plt_fmt()
        grps = [(n, d.ocost.mean(), len(d)) for n, d in df.groupby('pcode')]
        var = [y for x in grps for y in range(1, int(x[2]) + 1)]
        df = df.assign(idx=var)
        # https://stackoverflow.com/questions/11067027/re-ordering-columns-in-pandas-dataframe-based-on-column-name
        # var = [df.columns[0], 'idx'] + list(df.columns[1:-1])
        var = 'pcode idx styno jono qty description invdate mps uprice china mtlcost ocost'.split()
        df = df.reindex(var, axis=1)
        grps = sorted(grps, key=lambda x: x[1])
        grps = self._group(grps)
        ccnt = self._args['col_cnt']
        rcnt = (len(grps) + ccnt - 1) // ccnt
        if rcnt == 1 and len(grps) % ccnt:
            ccnt = 1
        if not fn:
            fn = TemporaryFile().name + '.pdf'
        grps, plt_mp, mks, done = iter(grps), defaultdict(list), self._args['marks'], False
        for var in df.itertuples(False):
            plt_mp[var.pcode].append(var.ocost)
        with PdfPages(fn) as pdf:
            while rcnt > 0 and not done:
                rcnt = rcnt - 2 # max 2 rows per page
                row = 2 if rcnt >= 0 else rcnt + 2
                fig, _axs = plt.subplots(row, ccnt, figsize=self._args['page_size'], constrained_layout=True)
                _axs = [x for x in _axs.flat] if ccnt > 1 else [_axs]
                for ax in _axs:
                    var = self._args['ax_labels']
                    ax.set_xlabel(var[0])
                    ax.set_ylabel(var[1])
                    var = self._args['title']
                    if var:
                        fig.suptitle(var + '(filter=%s)' % method)
                    var = self._args['grids']
                    if var:
                        ax.grid(b=True, which='major', color='k', linestyle='-.', alpha=0.6)
                        if var > 1:
                            ax.grid(b=True, which='minor', linestyle='-.')
                        ax.minorticks_on()
                    try:
                        for pcode in next(grps):
                            lst = plt_mp[pcode]
                            ax.plot([x + 1 for x in range(len(lst))], lst, '-' + mks[randint(0, len(mks) - 1)], label=pcode)
                    except StopIteration:
                        done = True
                        break
                    finally:
                        ax.legend()
                pdf.savefig(fig)
                plt.close(fig)
        plt.close()
        if sht:
            self._write_sht(df, sht, [fn, ])
        return fn

    def _write_sht(self, df, sht, atts):
        # df.jono = "'" + df.jono
        cns = [x for x in df.columns]
        sht.cells[0, 0].value = cns
        sht.cells[1, 0].value = df.values
        # uprice field might be formatted as currency, turn it to standard
        rng = df.columns.get_loc('uprice')
        sht.range(sht.cells[0, rng], sht.cells[sht.used_range.rows.count - 1, rng]).NumberFormatLocal = 'G/General'
        sht.autofit('c')
        rng = xwu.usedrange(sht)
        rng.row_height = 18
        rng = sht.range(sht.cells[1, 0], sht.cells[len(df), len(cns) - 1])
        rng.select()
        nl, rng = NamedList(cns), rng.api
        nms = {x: xwu.col(nl.getcol(x) + 1) for x in ('pcode', 'ocost')}
        nms['th'] = 0.05
        # Green
        fml = '=AND($%(pcode)s2=$%(pcode)s1,$%(ocost)s2-$%(ocost)s1<-%(th)4.2f)' % nms
        c = rng.formatconditions.add(FormatConditionType.xlExpression, FormatConditionOperator.xlEqual, fml)
        c.interior.colorindex = 35
        # RED
        fml = '=AND($%(pcode)s2=$%(pcode)s1,$%(ocost)s2-$%(ocost)s1>%(th)4.2f)' % nms
        c = rng.formatconditions.add(FormatConditionType.xlExpression, FormatConditionOperator.xlEqual, fml)
        c.interior.colorindex = 40
        # title
        fml = '=$%(pcode)s2<>$%(pcode)s1' % nms
        c = rng.formatconditions.add(FormatConditionType.xlExpression, FormatConditionOperator.xlEqual, fml)
        c.interior.colorindex = 37
        c.borders.linestyle = LineStyle.xlContinuous
        rng = sht.cells[0, 0]
        rng.select()
        rng.api.AutoFilter()
        xwu.freeze(sht.cells[1, 2])
        for x in (('mps', 7), ('description', 10)):
            sht.cells[0, df.columns.get_loc(x[0])].column_width = x[1]
        c = df.columns.get_loc('description')
        sht.range(sht.cells[0, c], sht.cells[(len(df), c)]).api.HorizontalAlignment = HAlign.xlHAlignRight
        if atts:
            for fn in atts:
                if path.splitext(fn)[-1].lower() in ('png', 'jpg', 'jpeg'):
                    pass
                else:
                    # can not use OLEObjects of api, use hyper-link instead
                    # sht.api.OLEObjects.Add(Filename=fn, Link=False, DisplayAsIcon=False)
                    rng = sht.cells[0, len(cns)]
                    rng.add_hyperlink(fn)
                    rng.value = 'Inv.Trend'

class LocalJOWgts(object):
    ''' joweight cache from csv or database
    '''

    def __init__(self, csv_fn, hksvc):
        self._modcnt, self._fn = 0, csv_fn
        self._df = pd.read_csv(csv_fn, keep_default_na=False) if path.exists(csv_fn) else pd.DataFrame(None, columns='pcode k0 w0 k1 w1 k2 w2'.split())
        self._hksvc = hksvc

    def wgts(self, pcode):
        ''' return weights of JO#
        '''
        sr = self._df.loc[self._df.pcode == pcode]
        if sr.empty:
            with self._hksvc.sessionctx():
                jo = self._jo_of(pcode)
                wgts = self._hksvc.getjowgts(jo)
                if not wgts:
                    return None
            wgts = [pcode, ] + [y for x in [(x.karat, x.wgt) if x else (None, ) * 2 for x in wgts.wgts] for y in x]
            sr = pd.DataFrame((wgts, ), columns=self._df.columns)
            self._df = self._df.append(sr)
            self._modcnt += 1
            if self._modcnt % 10 == 0:
                self._flush()
        return None if sr.empty else self._mk_wgt(sr.iloc[0])

    @staticmethod
    def _mk_wgt(row):
        return PrdWgt(*[WgtInfo(float(row['k%d' % i]), float(row['w%d' % i])) if row['k%d' % i] else None for i in range(3)])

    def _jo_of(self, pcode):
        ''' the latest JO of given pcode, for better weight value
        '''
        with self._hksvc.sessionctx() as cur:
            lst = cur.query(JO).join(PajShp).filter(PajShp.pcode == pcode).order_by(desc(PajShp.invdate)).limit(1).one()
            return lst or None

    def _flush(self):
        if self._modcnt > 0:
            self._df.to_csv(self._fn, index=None)
            self._modcnt = 0

    def __del__(self):
        self._flush()
