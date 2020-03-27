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
from utilz.miscs import NamedList, NamedLists, getfiles, na, newline, tofloat, trimu, splitarray

from .common import _logger as logger
from .common import config


try:
    import matplotlib as mpl
    from matplotlib import pyplot as plt
    from matplotlib.backends.backend_pdf import PdfPages
    import pandas as pd
except:
    pass

_inv_cols = 'pcode styno jono qty description invdate mps uprice china mtlcost ocost'.split()

def _fetch_invs(hksvc, fldr, pcodes, fn_cns=None, fn_wgts=None):
    pcodes = [x for x in pcodes if x[0] != '#']
    if not pcodes:
        return None
    if not fn_cns:
        fn_cns = path.join(fldr, '_cns.csv')
    cns = _inv_cols
    df = pd.read_csv(fn_cns, parse_dates=['invdate']) if path.exists(fn_cns) else pd.DataFrame(None, columns=cns)
    lns = {x for x in pcodes if x[0] != '#'} - set(df.pcode)
    if lns:
        wgtsvc, lsts = LocalJOWgts(fn_wgts or path.join(fldr, '_cache', '_wgts.csv'), hksvc), []
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

    @property
    def _fn_src(self):
        return self._cache('_src.dat')

    @property
    def _fn_dts(self):
        return self._cache('_fdates.dat')
    
    def __init__(self, fldr, hksvc):
        self._fldr, self._hksvc = fldr, hksvc
        self._fmtr = _AckFmt()
        self._nl_src = NamedList("jono,date,file,mps,pcode,pajprice,styno,qty")
        self._nl_rst = self._fmtr.nl_rst
        self._wgtsvc = LocalJOWgts(self._cache('_wgts.csv'), hksvc)
        self._src_mp = None
        self._jc = lambda x: None if not x else JOElement(x).name

    def _uptodate(self, fns):
        ''' return True if data in _fn_src is up-to-date
        '''
        if not fns:
            return True
        fn = self._fn_dts
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
        if srcs is None or srcs.empty:
            var = self._get_src_fns()
            srcs = self._read_srcs(var, self._uptodate(var))
        if srcs is None or isinstance(srcs, str):
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

        var = set(srcs.file)
        with self._hksvc.sessionctx() as cur:
            try:
                dds = Query([PajAck.docno, PajAck.lastmodified]).filter(
                    PajAck.docno.in_(list(var))).distinct()
                dds = dds.with_session(cur).all()
                dds = {x[0]: x[1] for x in dds}
            except:
                dds = {}
            exps, jes = set(), set()
            lst, td = [], datetime.today()
            for idx, x in srcs.iterrows():
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
            path.getmtime(rc) >= path.getmtime(self._fn_dts):
            logger.info(err)
            logger.info("data up to date, don't need further process")
            return rc
        rc = err = jos = None
        logger.debug("begin to analyse acknowledgements (%s)" % self._fldr)
        srcs = self._read_srcs(fns, utd)
        if isinstance(srcs, str):
            err = srcs[1:]
        elif not srcs.empty:
            logger.debug("%d JOs returned" % len(srcs))
            self._src_mp = srcs
            with self._hksvc.sessionctx():
                # keep session to make sure self.persist() can still access the JOs
                rc = self._process_all()
                if rc:
                    jos = rc[1]
                    rc = self._write_wb(rc[0], path.join(self._fldr, tar_fn))
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

    def _read_src(self, fn, fds):
        bfn, err = path.basename(fn), None
        logger.debug("Reading file(%s)" % bfn)
        fds[bfn] = str(path.getmtime(fn))
        df = pd.read_excel(fn, 'To Factory -Other_1')
        sp = [df.loc[df["Purchase Order"] == x].iloc[0]['Unnamed: 3'] for x in ('Silver:', 'Gold :', 'Order Date:')]
        if not all(sp):
            err = "Key argument(sp/gp/date) missing in (%s)" % bfn
            logger.debug("sheet(%s) has not enough arguments" % bfn)
        # TODO:: PAJ's form keep changing, so try to detect the row header, find Job#
        df = pd.read_excel(fn, 'To Factory -Other_1', header=9).dropna(subset=['Job#'])
        df = df.assign(mps=MPS("S=%f;G=%f" % tuple(sp[:2])).value, adate=sp[-1])
        return df, err

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
        df = None
        if utd and path.exists(self._fn_src):
            df = pd.read_csv(self._fn_src, parse_dates=['date'])
        if df is None or df.empty:
            df, fds = None, {}
            try:
                err = None
                for fn in fns:
                    dx, err = self._read_src(fn, fds)
                    if err:
                        break
                    dx['file'] = [path.basename(fn), ] * len(dx)
                    df = dx if (df is None or df.empty) else pd.concat([df, dx], ignore_index=True)
                if not err:
                    # name convert and just fetch some columns
                    ttls = ('file,Job#,Style#,Quantity,Price,mps,adate,Item No'.split(','), 
                    'file jono styno qty pajprice mps date pcode'.split())
                    df = pd.concat([df[x] for x in ttls[0]], keys=ttls[1], axis=1)
                    df.to_csv(self._fn_src, index=None)
                    lst = [','.join(x) for x in fds.items()]
                    lst.insert(0, "file,date")
                    with open(self._fn_dts, "w") as fh:
                        for x in lst:
                            print(x, file=fh)
            except Exception as e:
                err = "file(%s),err(%s)" % (path.basename(fn), e)
            if err:
                return "_" + err
        # merge the duplicateds and make array, 2 cold to be merged: qty and pcode
        df['pajprice'] = df.pajprice.apply(lambda n: 0 if pd.isnull(n) else n)
        x = df.groupby('file jono styno pajprice mps date'.split())
        df = x['qty'].sum().reset_index()
        # direct assign because the physical order is the same
        df['pcode'] = x['pcode'].apply(','.join).values
        df['pcode'] = df.apply(lambda row: row.pcode.split(','), axis=1)
        df['jono'] = df.jono.apply(self._jc) # force the JO# to string
        return df

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
        pdx = jo_mp.loc[jo_mp.jono == nl_src.jono].iloc[0]
        for cn in 'date file jono mps pajprice qty styno'.split():
            nl[cn] = nl_src[cn]
        if not pdx.empty and pdx.poprice > 0:
            nl.poprice = x = float(pdx.poprice) #pdx.poprice sometimes might be Decimal
            nl.profit = x - pajup
            nl.profitt = nl.profit * nl_src.qty
            if pajup:
                nl.ratio = x / pajup * 100.0
            nl.result = self._fmtr.classify_pft(pajup, x)
            pcodes = pdx.pcode
        else:
            nl.result = self._fmtr.classify_pft(pajup, 0)
            pcodes = nl_src.pcode

        for pcode in pcodes:
            wgts = self._wgtsvc.wgts(pcode)
            if wgts:                    
                nl.wgts = wgts
                break

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
        pcs = nl_src.pcode
        if not isinstance(pcs, (list, tuple)):
            pcs = (pcs, )
        for pcode in pcs:
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
            nl.wgts = self._wgtsvc.wgts(self._src_mp[jo.name.value].pcode[0])
            nl_srcx = nl_src.clone(False)
            nl_srcx.jono = jo.name.value
            nl.pcode = shp.PajShp.pcode
            tar = self._calc_his(shp, nl_srcx, nl)
            nl.wgts, nl.jono = wgt0, jono0
            return self._fmtr.label('labels', 'rf.sml'), tar
        nl.ref = self._fmtr.label('labels', 'rf.noref')
        return None

    def _fetch_jos(self, jos):
        """
        return a dict with jono as key and dict with columns: "poprice,mps,wgts"
        """
        mp = []
        for jo in jos:
            try:
                jn = jo.name.value
                x = self._src_mp.loc[self._src_mp.jono == jn]
                if x.empty:
                    pcode = None
                else:
                    x = x.iloc[0]
                    pcode = x.pcode
                po = jo.po
                mp.append((jn, po.uprice, pcode, ))
            except:
                logger.debug('failed to get po/wgts for JO(%s), maybe BIG5 encoding error' % jo.name.value)
        return pd.DataFrame(mp, columns='jono poprice pcode'.split())
    
    def _cache(self, fn):
        fldr = path.join(self._fldr, '_cache')
        if not path.exists(fldr):
            mkdir(fldr)
        return path.join(fldr, fn)

    def _process_all(self):
        """
        @param all: a dict with jono as key and a dict with these keys: "jono,pajprice,file,mps,styno,date,qty,pcode". ref @_read_srcs() FMI.
        """
        src_mp = self._src_mp
        if src_mp is None or src_mp.empty:
            return None
        fn = self._cache('_jos.csv')
        df = pd.read_csv(fn) if path.exists(fn) else None
        if df is not None:
            df['pcode'] = df.pcode.apply(lambda x: x.replace("'", '')[1:-1].split(','))
            df.jono = df.jono.apply(self._jc)
        reqs = set(src_mp.jono)
        if not (df is None or df.empty):
            reqs = reqs - set(df.jono)
        if reqs:
            with self._hksvc.sessionctx():
                for var in splitarray([JOElement(x) for x in reqs], 10):
                    logger.debug('fetching POs(%s)' % var)
                    var = self._fetch_jos(self._hksvc.getjos(var)[0])
                    df = pd.concat([df, var], sort=False)
                    df.to_csv(fn, index=None)
        rsts = {}
        for idx, nl_src in self._src_mp.iterrows():
            try:
                nl = self._process_one(nl_src, df, True)
                rsts.setdefault(nl.result, []).append(nl)
                idx += 1
                if idx % 10 == 0:
                    logger.debug('analysing %d of %d' % (idx, len(self._src_mp)))
            except:
                nl_src.pcode = nl_src.pcode[0]
                rsts.setdefault(self._fmtr.label('labels', 'prg.error'), []).append(nl_src)
        # map value time translation
        var = {}
        for k, v in rsts.items():
            nl = v[0]
            if isinstance(nl, NamedList):
                var[k] = pd.DataFrame([x.data for x in v], columns=nl.colnames)
            elif isinstance(nl, pd.Series):
                var[k] = pd.DataFrame(v)
        return var, None # TODO:: return jos to avoid persist-double-retrieve

    def _write_wb(self, rsts, fn):
        app, kxl = xwu.appmgr.acq()

        wb = app.books.add()
        self._write_readme(wb)
        for k, v in rsts.items():
            self._write_sht(k, v, wb)

        df = self._fmtr.label('labels', 'prg.error')
        # prg.error contains source result set only, so void it
        df = [v for k, v in rsts.items() if k != df]
        df = pd.concat(df) if df else pd.DataFrame()
        def _exs(kwds):
            if df.empty:
                return df
            df['flag'] = df.ref.apply(lambda ref: sum([1 for y in kwds if ref.find(y) >= 0]))
            kwds = df.loc[df.flag > 0]
            del df['flag']
            del kwds['flag']
            return kwds
        mp = {
            "_except": _exs(self._fmtr.setting('pft.ref.classify')['labels'][1:]),
            "_new_sml": _exs((self._fmtr.label('labels', 'rf.sml'), )),
            "_new": _exs((self._fmtr.label('labels', 'rf.noref'), )),
        }
        if not mp['_except'].empty:
            # when there is exception, prepare one for enqurying zhengyuting
            tmp = _exs([self._fmtr.label('labels', x) for x in ("rf.rev", "rf.his", "rf.sml")])
            tmp['flag'] = tmp.apply(lambda k: self._fmtr.classify_ref(k.pajprice, k.expected, False), axis=1)
            tmp = tmp.loc[tmp.flag > 0]
            del tmp['flag']
            tmp['ref'] = tmp.ref.apply(lambda x: x.split('_')[0]) # remove the profit result from REF for PAJ            
            mp['_paj_enq'] = tmp
        tmp = set()
        # when item in higher level, don't show them in lower one
        for k, v in mp.items():
            if v.empty:
                continue
            if k == '_paj_enq':
                df = v
            else:
                df = v.loc[~v.jono.isin(tmp)]
                tmp = tmp.union(set(df.jono))
            if df.empty:
                continue
            self._write_sht(k, df, wb)
        for var in wb.sheets:
            xwu.usedrange(var).row_height = 18
        if fn:
            wb.save(fn)
            fn = wb.fullname
        wb.close()
        xwu.appmgr.ret(kxl)
        return fn

    def _write_sht(self, name, nls, wb):
        if nls is None:
            return
        lst = []
        excl = self._fmtr.excludes(name) or ()
        hdr = [x for x in nls.columns]
        if excl:
            hdr = [x for x in hdr if x not in excl]
        if len(nls) > 1:
            nls = self._fmtr.sort(name, nls)
        hdr_ttl = [self._fmtr.label('cat.cns', x) for x in hdr]
        hdr_ttl = [x[0] if isinstance(x, list) else x for x in hdr_ttl]
        lst.append(hdr_ttl)
        hdr_mp = {x: idx for idx, x in enumerate(hdr)}
        _flag = lambda cn: cn in nls.columns
        if _flag('jono') and nls.iloc[0].jono[0] != "'":
            nls.jono = nls.jono.apply(lambda x: "'" + x)
        _flag = lambda cn: cn in nls.columns and not isinstance(nls.iloc[0][cn], str)
        if  _flag('wgts'):
            nls.wgts = nls.wgts.apply(lambda x: ';'.join('%d=%4.2f' % (x.karat, x.wgt) for x in x.wgts if x and x.wgt))
        if _flag('ratio'):
            nls.ratio = nls.ratio.apply(lambda x: "%4.2f%%" % (x or 0))
        for idx, nl in nls.iterrows():
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
        fn = self._cache('_invhis.csv')
        df = pd.read_csv(fn, parse_dates=['invdate', ]) if path.exists(fn) else  pd.DataFrame(None, columns=_inv_cols)
        reqs = set([x for x in nls.pcode])
        if not df.empty:
            reqs = reqs.difference(set([x for x in df.pcode]))
        # reqs = [nl.pcode for nl in nls if df.loc[df.pcode == nl.pcode].empty] #df.pcode is an array
        if reqs:
            df = _fetch_invs(self._hksvc, self._fldr, reqs, fn, self._cache('_wgts.csv'))
            df.to_csv(fn, index=None)
        # append myself to show this trend
        var = []
        # df['jono'] = df.jono.apply(self._jc)
        df['jono'] = df.jono.apply(lambda x: "'%s" % x)
        for pcode in df.pcode:
            wgts = self._wgtsvc.wgts(pcode)
            nl = nls.loc[nls.pcode == pcode].iloc[0]
            cn = PajCalc.calchina(wgts, nl.pajprice, nl.mps, nl.date)
            var.append([pcode, nl.styno, '*' + nl.jono[1:], nl.qty, 'n/a', nl.date, nl.mps, nl.pajprice, cn.china, cn.metalcost, cn.china - cn.metalcost])
        df = df.append(pd.DataFrame(var, columns=df.columns))
        ttl = 'Other Cost (China - Metal) Trend of "%s"' % path.basename(self._fldr)
        _HisPltr(df, title=ttl).plot(path.join(self._fldr, '_invhis.pdf'), sht, methods=['changes'])
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
        pcodes = [x for x in pcodes if x[0] != '#']

        flag, methods, fn = False, method.split(','), path.join(s_fldr, prefix)
        for mt in methods:
            fn_xls = _HisPltr.make_fn(fn, mt, 'xlsx')
            if path.exists(fn_xls):
                continue
            flag = True
        if flag:
            # translate those JO# or Sty# to pcodes
            var = [x for x in pcodes if len(x) != 17 or x.find(':') > 0]
            if var:
                pcodes = cls._extend(pcodes, var, hksvc)
            pltr = _HisPltr(_fetch_invs(hksvc, fldr, pcodes), title=ttl)
            pltr.plot(path.join(s_fldr, fn), None, methods=methods)

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
            nls:    NamedLists to be sorted, since 20191022, it becomes dataframe
        '''
        if nls is None or len(nls) < 2:
            return nls
        sdef = self.setting("cat.sorting")
        cns = self._parse(sdef.get(cat) or sdef.get('_default'), sdef)
        if isinstance(nls, pd.DataFrame):
            cns = [x for x in cns if x in nls.columns]
            if cns:
                return nls.sort_values(cns)
            return nls
        else:
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
        self._df = df.drop_duplicates(['pcode', 'jono', 'invdate']) if df is not None else None
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

    def _liner(self, df, methods='all'):
        accepts, pcode, trcnt = defaultdict(list,), None, 0
        th = config.get('pajcc.ack.chk')['pft.ref.classify']
        th = [th[x][0] for x in ("relative", "absolute")]
        def _accept():
            if not pcode:
                return
            for method in methods:
                flag, modcnt = False, len(trend)
                flag = method == 'all'
                if not flag:
                    if method == 'changes':
                        flag = modcnt > 0
                if not flag and modcnt:
                    if method in ('up', 'down'):
                        flag = trend[-1][0] == ('U' if method == 'up' else 'D')
                    elif method in ('high', 'low'):
                        flag = (his[-1] > his[0]) == (method == 'high')
                    elif method == 'sea-food':
                        flag = modcnt >= 3 and modcnt / trcnt >= 0.3
                    elif method != 'all':
                        shape = ''.join([x[0] for x in trend])
                        if method == 'shape-v':
                            flag = shape.find('DU') >= 0
                        elif method == 'shape-n':
                            flag = shape.find('UD') >= 0
                        elif method == 'shape-2u':
                            flag = shape.find('UU') >= 0
                        elif method == 'shape-2d':
                            flag = shape.find('DD') >= 0
                        elif method == 'suspicious':
                            flag = max(his) / min(his) >= 1.2
                if flag:
                    accepts[method].append(pcode)

        for idx, row in df.iterrows():
            # bonded gold, ignore
            if row.pcode[0] == 'K':
                continue
            if row.pcode != pcode:
                _accept()
                loc = trcnt = 0
                pcode = row.pcode
                trend, his = [], []
            trcnt += 1
            if loc and abs(loc / row.ocost - 1) < th[0] and abs(loc - row.ocost) < th[1]:
                df.loc[idx, 'ocost'] = loc
                # below 2 methods might has warning or does not work at all
                # df.ocost[idx] = loc
                # row.ocost = loc won't work because it's a view
            else:
                if loc:
                    trend.append(('D' if loc > row.ocost else 'U', trcnt))
                loc = row.ocost
                his.append(loc)
        _accept()
        if not accepts:
            return None
        return {k: df.loc[df.pcode.isin(v)] for k, v in accepts.items()}

    def plot(self, fn=None, sht=None, methods=['changes']):
        ''' plot to a file or a set of files
        Args:
            fn=None: the pdf file to plot to, omitting will create a temp file
            sht=None: the sheet to write data to, omitting won't produce any excel thing
            method=['changes']: what to show, a list of one of:
                'all'     -> any series
                'changes' -> the series that contains changes
                'up'      -> the series that belongs to up-trend
                'down'    -> the series that belongs to down-trend
        '''
        if self._df is None:
            return None
        df = self._df.sort_values(by=['styno', 'pcode', 'invdate'])
        accepts = self._liner(df, methods)
        if not accepts:
            return None
        self._plt_fmt()
        app = tk = None
        if sht is None:
            fldr, fn = (gettempdir(), '_') if not fn else (path.dirname(fn), path.basename(fn))
            app, tk = xwu.appmgr.acq()
        for method, df in accepts.items():
            if app:
                fn_pdf = self.make_fn(path.join(fldr, fn), method, 'pdf')
                if path.exists(fn_pdf):
                    continue
                df = self._plot_df(df, method, fn_pdf)
                wb = app.books.add()
                self._write_sht(df, wb.sheets[0], [fn_pdf, ])
                wb.save(self.make_fn(path.join(fldr, fn), method, 'xlsx'))
                wb.close()
            else:
                df = self._plot_df(df, method, fn)
                if sht:
                    self._write_sht(df, sht, [fn, ])
        if app:
            xwu.appmgr.ret(tk)

    @classmethod
    def make_fn(cls, pfx, method, ext):
        return pfx + '_' + method + '.' + ext

    def _plot_df(self, df, method, fn):
        grps = [(n, d.ocost.mean(), len(d)) for n, d in df.groupby('pcode')]
        # grps is sorted by pcode, so need to keep the physical order for idx
        ccnt, rcnt = df.pcode.unique(), {x[0]: x[-1] for x in grps}
        var = [y + 1 for x in ccnt for y in range(rcnt[x])]
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
        return df

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
            if self._modcnt % 1 == 0: #network is very bad, save one by one
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
