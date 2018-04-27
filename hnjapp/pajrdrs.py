# coding=utf-8
'''
Created on Apr 17, 2018

the replacement of the Paj Shipment Invoice Reader, which was implmented
in PAJQuickCost.xls#InvMatcher

@author: zmFeng
'''

from collections import namedtuple
from datetime import datetime, date
import numbers
import os, sys
import re

from xlwings.constants import LookAt

import logging as logging
from models.utils import JOElement
from utils import p17u
from utils import xwu
import xlwings.constants as const


def _accdstr(dt):
    """ make a date into an access date """ 
    return dt.strftime('%Y-%m-%d %H:%M:%S') if(dt and isinstance(dt, date)) else dt


def _fmtjono(jn):
    return ("%d" % jn if(isinstance(jn, numbers.Number)) else jn.strip()) if jn else None


def _removenonascii(s0):
    """remove thos non ascii characters from given string"""
    if(isinstance(s0, basestring)): return "".join([x for x in s0 if ord(x) > 31 and ord(x) < 127])
    return s0


def _getjoids(jonos, hnjhkdb):        
        """get the joIds by the provided jonos
        @param jonos: a list of JOElement  
        """
        
        rc = None
        s0 = "or".join([" (alpha = '%s' and digit = %d) " % (x.alpha, x.digit) for x in jonos])
        s0 = "select alpha,digit,joid from jo where (%s)" % s0
        cur = hnjhkdb.cursor()
        try:
            cur.execute(s0)
            rows = cur.fetchall()
            if(rows):
                rc = dict((JOElement(x.alpha, x.digit).value, x.joid) for x in rows)
        finally:
            if(cur): cur.close()
        return rc


class ShpReader:

    def __init__(self, accdb, hnjhkdb):
        self._accessdb = accdb
        self._hnjhkdb = hnjhkdb
        self._ptnfd = re.compile(r"(\d{2,4})-(\d{1,2})-(\d{1,2})")
        self._ptngwt = re.compile(r"(\d*\.\d*)\s*[gG]+")
        self._ptnswt = re.compile(r"\d*\.\d*")
    
    def _hasread(self, fn):
        """
            check if given file has been read, by default, from AccessDb
            @param fn: the full-path filename 
        """
        rc = 0
        cur = self._accessdb.cursor()
        try:
            rfn = os.path.basename(fn).replace("_", "")
            cur.execute(r"SELECT max(lastModified) from jotop17 where fn='%s'" % rfn)
            row = cur.fetchone()
            if(row and row[0]):
                rc = 2 if(row[0] < datetime.fromtimestamp(os.path.getmtime(fn)).replace(microsecond=0))\
                    else 1
        except BaseException as e:
            rc = -1
            print(e)
        finally:
            if(cur): cur.close()
        return rc
    
    def _getshpdate(self, fn):
        """extract the shipdate from file name"""
        import datetime as dt
        mt = self._ptnfd.search(os.path.basename(fn))
        if(mt and len(mt.groups()) == 3):
            dg = mt.groups()
            return dt.date(int(dg[0]), int(dg[1]), int(dg[2]))

    def _readquodata(self, sht, qmap):
        """extract gold/stone weight data from the QUOXX sheet
        @param sht:  the DL_QUOTATION sheet that need to read data from
        @param qmap: the dict with p17 as key and (goldwgt,stwgt) as value
        """
        rng = xwu.find(sht, "Item*No", lookat=LookAt.xlPart)
        if(not rng): return
        # because there is merged cells rng.expand('table').value 
        # or sht.range(rng.end('right'),rng.end('down')).value failed
        vvs = sht.range(rng, rng.current_region.last_cell).value 
        nmap = {"p17":0}        
        tr = vvs[0]
        for it in {"stone":re.compile(r"stone\s*weight", re.I), \
            "metal":re.compile(r"metal\s*weight", re.I)}.iteritems():            
            cns = [x for x in range(len(tr)) if it[1].search(tr[x])]
            if(len(cns) > 0): nmap[it[0]] = cns[0]            
        for x in range(1, len(vvs)):
            tr = vvs[x]
            p17 = tr[nmap["p17"]]
            if(not p17): continue
            if(p17u.isvalidp17(p17) and not qmap.has_key(p17)):
                sw = 0 if(not tr[nmap["stone"]]) else \
                    sum([float(x) for x in self._ptnswt.findall(tr[nmap['stone']])])
                mw = sum([float(x) for x in self._ptngwt.findall(tr[nmap['metal']])])
                qmap[p17] = (mw, sw)
                
    def _persist(self, dups, items):
        """save the data to db
        @param dups: a list contains file names that need to be removed
        @param items: all the ShipItems that need to be persisted
        """
        
        if(len(dups) + len(items) <= 0): return 0, None
        dbs = (self._accessdb, self._hnjhkdb)
        curs = [x.cursor() for x in dbs]
        e = None
        try:
            if(len(dups) > 0):
                curs[0].execute("delete from jotoP17 where fn in ('%s')" % "','".join(dups))
                curs[1].execute("delete from pajshp where fn in ('%s')" % "','".join([_removenonascii(x) for x in dups])) 
            
            if(len(items) > 0):
                dcts = list([x._asdict() for x in items.values()])
                jns = [JOElement(x.jono) for x in items.values()]
                jns = _getjoids(jns, self._hnjhkdb)                
                for dct in dcts:
                    dct["joid"] = jns[dct["jono"]]
                    dct['fnnc'] = _removenonascii(dct["fn"])
                    ss = (("insert into jotop17 (fn,jono,p17,fillDate,lastModified,invno,qty" + \
                        ",InvDate,ShpDate,OrdNo,MtlWgt,StWgt) values('%(fn)s','%(jono)s','%(p17)s'" + \
                        ",#%(fillDate)s#,#%(lastModified)s#,'%(invno)s',%(qty)f,#%(invDate)s#" + \
                        ",#%(shpDate)s#,'%(ordno)s',%(mtlWgt)f,%(stWgt)f)") % dct, \
                        ("insert into pajshp (fn,joid,pcode,filldate,lastmodified,invno,qty" + \
                        ",invdate,shpdate,orderno,mtlwgt,stwgt) values('%(fnnc)s',%(joid)d,'%(p17)s'" + \
                        ",'%(fillDate)s','%(lastModified)s','%(invno)s',%(qty)f,'%(invDate)s'" + \
                        ",'%(shpDate)s','%(ordno)s',%(mtlWgt)6.2f,%(stWgt)6.2f)") % dct)
                    for ii in range(len(curs)):
                        curs[ii].execute(ss[ii])
        except Exception as e:
            pass
        finally:
            for cur in curs:
                if(cur): cur.close()
            if(e):
                for db in dbs:
                    db.rollback()                
            else:
                for db in dbs:
                    db.commit()
        return -1 if(e) else 1 , e
    
    def read(self, fldr):
        """
        read the shipment file and send to 2dbs
        @param fldr: the folder contains the files. sub-folders will be ignored 
        """
        
        ptn = re.compile("HNJ\s*\d*-", re.IGNORECASE)
        root = fldr + os.path.sep if fldr[len(fldr) - 1] <> os.path.sep else ""
        fns = [root + unicode(x, sys.getfilesystemencoding()) for x in os.listdir(fldr) if ptn.match(x)]
        if(len(fns) == 0): return
        errors = list()
        killxls, app = xwu.app(False)
        PajShpItem = namedtuple("PajShpItem", "fn,ordno,jono,qty,p17,invno,invDate" + \
            ",mtlWgt,stWgt,shpDate,lastModified,fillDate")
        try:
            for fn in fns:
                idx = self._hasread(fn)
                toRv = list();items = {}
                if(idx == 1):
                    # logging.debug("%s has been read" % fn)
                    continue
                elif(idx == 2):
                    logging.debug("%s is expired" % fn)
                    toRv.append(fn)
                logging.debug("processing file(%s)" % fn)
                lmd = _accdstr(datetime.fromtimestamp(os.path.getmtime(fn)))
                wb = app.books.open(fn)
                try:
                    # in new sample case, use DL_QUOXXX sheet's weight, so prepare it if there is
                    qmap = {}
                    for sht in wb.sheets:
                        rng = xwu.find(sht, u"十七*", lookat=LookAt.xlPart)
                        if(not rng): rng = xwu.find(sht, u"物料", lookat=LookAt.xlPart)
                        if(not rng): continue
                        # don't use this, sometimes the stupid user skip some table header
                        # vvs = rng.end('left').expand("table").value
                        vvs = xwu.usedrange(sht).value
                        th = vvs[0]
                        tm = xwu.arrtodict(th, {u"订单号":"odx", u"发票日期":"invdate", u"订单序号":"odseq", \
                            u"平均单件石头,XXX":"stwgt", u"发票号":"invno", u"订单号序号":"ordno", u"十七位,十七":"p17", \
                            u"平均单件金,XX":"mtlwgt", u"工单,job":"jono", u"数量":"qty"})
                        x = [x for x in "invno,p17,jono,qty,invdate".split(",") if not tm.has_key(x)]
                        if(len(x) > 0):
                            logging.debug("failed to find key columns(%s) in sheet(%s)" % (x, sht.name))
                            continue                        
                        bfn = os.path.basename(fn).replace("_", "")
                        if(tm.has_key("mtlwgt")):                            
                            for ridx in range(1, len(vvs)):
                                tr = vvs[ridx]
                                if(not tr[tm['p17']]): break
                                mwgt = tr[tm["mtlwgt"]]
                                if(not (isinstance(mwgt, numbers.Number) and mwgt > 0)): continue 
                                invno = tr[tm["invno"]] if [tm["invno"]] else "N/A"
                                if(tm.has_key('ordno')):
                                    odno = tr[tm['ordno']]
                                elif(tm.has_key('odx') and tm.has_key('odseq')):
                                    odno = tr[tm['odx']] + "-" + tr[tm["odseq"]]                       
                                else:
                                    odno = "N/A"
                                # 'use fn+JO#+P17+invno as map key
                                thekey = bfn + "," + tr[tm['p17']] + "," + invno
                                if(items.has_key(thekey)):
                                    si = items[thekey]
                                    items[thekey] = si._replace(qty=si.qty + tr[tm["qty"]]) 
                                else:
                                    si = PajShpItem(bfn, odno, _fmtjono(tr[tm["jono"]]) , tr[tm["qty"]], tr[tm['p17']], \
                                        invno, _accdstr(tr[tm['invdate']]), mwgt, tr[tm['stwgt']], \
                                        _accdstr(self._getshpdate(bfn)), lmd, _accdstr(datetime.today()))
                                    items[thekey] = si
                        else:
                            # new sample case, extract weight data from
                            if(len(qmap) == 0):
                                for x in [xx for xx in wb.sheets if xx.name.lower().find('dl_quotation') >= 0]:
                                    self._readquodata(x, qmap)
                            if(len(qmap) > 0):
                                import random
                                for x in range(1, len(vvs)):
                                    tr = vvs[x]
                                    odno = tr[tm['ordno']] if tm.has_key('ordno') else "N/A"
                                    p17 = tr[tm['p17']]
                                    if(not p17): break                                    
                                    si = PajShpItem(bfn, odno, _fmtjono(tr[tm["jono"]]), tr[tm["qty"]], p17, tr[tm["invno"]], \
                                        _accdstr(tr[tm['invdate']]), qmap[p17][0], qmap[p17][1], \
                                        _accdstr(self._getshpdate(bfn)), lmd, _accdstr(datetime.today()))
                                    # new sample won't have duplicated items                                    
                                    items[random.random()] = si
                            else:
                                qmap["_SIGN_"] = 0
                finally:
                    if(wb): wb.close()
                x = self._persist(toRv, items)
                if(x[0] <> 1): errors.append(x[1])
        finally:
            if(killxls): app.quit()
        return -1 if len(errors) > 0 else 1, errors


class InvReader(object):
    """
    read invoices from folder
    """
    
    def __init__(self, accdb, hnjhkdb):
        self._accessdb = accdb
        self._hnjhkdb = hnjhkdb

    def _getinv(self, fn):
        """
            extract the inv# from the file name
            @param fn: the full path just the file name 
        """
        s0 = os.path.basename(fn)
        return s0[:s0.find(".")].strip()
    
    def _hasread(self, fn):
        """
            check if given file has been read, by default, from AccessDb
            @param fn: the full-path filename 
        """
        rc = 0
        cur = self._accessdb.cursor()
        try:
            cur.execute(r"select lastModified from PajInv where invno = '%s'" % self._getinv(fn))
            row = cur.fetchone()
            if(row and row[0]):
                rc = 2 if(row[0] < \
                    datetime.fromtimestamp(os.path.getmtime(fn)).replace(microsecond=0)) else 1
        finally:
            if(cur): cur.close()
        return rc
    
    def _getjonos(self, p17, invno):
        """
        return the JO#s of the given p17 and inv#
        @param p17:     a valid p17 code
        @param invno:   the invoice# 
        """
        cur = self._accessdb.cursor()
        try:
            cur.execute("select jono from jotop17 where p17 = '%s' and invno = '%s'" % (p17, invno))
            rows = cur.fetchall()
            rc = set(x.jono.trim() for x in rows) if(rows) else None
        finally:
            if(cur): cur.close()
        return rc
    
    def _persist(self, dups, items):
        """persist the data
        @param dups:  a list of invnos
        @param items: the InvItems that need to be inserted
        """ 
        x = (len(dups), len(items))
        e = None
        if(sum(x) > 0):
            try:
                dbs = (self._accessdb, self._hnjhkdb) 
                curs = [db.cursor() for db in dbs]
                
                if(x[0]):
                    for x0 in ["delete from PajInv where invno = '%s'" % x0 for x0 in dups]:
                        curs[0].execute(x0)
                if(x[1]):                
                    dcts = list([x0._asdict() for x0 in items.values()])
                    jns = [JOElement(x0.jono) for x0 in items.values()]
                    jns = _getjoids(jns, self._hnjhkdb)
                    for dct in dcts:
                        dct["joid"] = jns[dct["jono"]]
                        dct["china"] = 0  # todo::make the china value for the user
                        ss = (("insert into PajInv(Invno,Jono,StSpec,Qty,UPrice,Mps,LastModified) values" \
                            + "('%(invno)s','%(jono)s','%(stone)s',%(qty)f,%(price)f,'%(mps)s',#%(lastmodified)s#)") % dct, \
                            ("insert into pajinv(invno,joid,stspec,qty,uprice,mps,lastmodified,china) values" \
                            + "('%(invno)s',%(joid)d,'%(stone)s',%(qty)f,%(price)f,'%(mps)s','%(lastmodified)s',%(china)f)") % dct)
                        for ii in range(len(curs)):
                            curs[ii].execute(ss[ii])
            except BaseException as e:  # there might be pyodbc.IntegrityError if dup found
                logging.debug("Error occur in InvRdr:%s" % e)                
            finally:
                for cur in curs: cur.close()
                if(not e):                    
                    for db in dbs: db.commit()                        
                else:
                    for db in dbs: db.rollback()
                    
        return (0 if sum(x) == 0 else -1 if e else 1) , e
    
    def read(self, invfldr, writeJOBack=True):
        """
        read files back, instead of using os.walk(root), use os.listdir()
        @param invfldr: the folder contains the invoices
        @param writeJOBack: write the JO# back to the source sheet 
        """
        
        if(not os.path.exists(invfldr)): return
        invs = {};errs = list()
        PajInvItem = namedtuple("PajInvItem", "invno,p17,jono,qty,price,mps,stone,lastmodified")
        if(invfldr[len(invfldr) - 1:] <> os.path.sep): invfldr += os.path.sep
        fns = [invfldr + unicode(x, sys.getfilesystemencoding()) for x in os.listdir(invfldr) \
            if x.lower()[2:4] == "pm" and x.lower().find(".xls")]
        if(not len(fns)): return
        killexcel, app = xwu.app(False)
        
        for fn in fns:  
            invno = self._getinv(fn).upper()
            if(invs.has_key(invno)): continue            
            dups = ()
            idx = self._hasread(fn)
            if(idx == 1):
                # logging.debug("%s has been read" % fn)
                continue
            elif(idx == 2):
                logging.debug("%s is expired" % fn)
                dups.append(invno)
            lmd = _accdstr(datetime.fromtimestamp(os.path.getmtime(fn)))
            wb = app.books.open(fn)
            updcnt = 0;items = {}
            try:
                for sh in wb.sheets:
                    rng = xwu.find(sh, "Invo No:")
                    if(not rng): continue
                    rng = xwu.find(sh, "Item*No", lookat=const.LookAt.xlWhole)
                    if(not rng): continue
                    rng = rng.expand("table")
                    vals = rng.value
                    # table header map
                    tm = {}
                    tm["p17"] = 0
                    tr = [xx.lower() for xx in vals[0]]
                    tm = xwu.arrtodict(tr, {"gold,":"gold", "silver,":"silver", u"job#,工单":"jono", \
                        "price,":"price", "unit,":"qty", "stone,":"stone"})
                    x = [x for x in "price,qty,stone".split(",") if not tm.has_key(x)]
                    if(len(x)):
                        logging.info("key columns(%s) missing in sheet('%s') of file (%s)" % (x, sh.name, fn))
                        continue
                    for jj in range(1, len(vals)):
                        tr = vals[jj]
                        if(not tr[tm["price"]]): continue
                        p17 = tr[tm["p17"]]
                        if(not p17u.isvalidp17(p17)):
                            logging.debug("invalid p17 code(%s) in %s" % (p17, fn))
                            continue
                        jn = _fmtjono(tr[tm["jono"]]) if tm.has_key("jono") else None
                        if(not jn):
                            jns = self._getjonos(p17, invno)
                            if(jns): jn = jns[0]
                            if(jn and writeJOBack):
                                sh.range(rng.row + jj, rng.column + tm["jono"]).value = jn
                                updcnt += 1
                        else:
                            jn = "%d" % jn if(isinstance(jn, numbers.Number)) else jn.strip()
                        if(not jn):
                            logging.debug("No JO# found for p17(%s) in file %s" % (tr[tm["p17"]], fn))
                            continue
                        key = invno + "," + jn
                        if(items.has_key(key)):
                            it = items[key]
                            items[key] = it._replace(qty=it.qty + tr[tm["qty"]])
                        else:
                            mps = "S=%3.2f;G=%3.2f" % (tr[tm["silver"]], tr[tm["gold"]]) \
                                if tm.has_key("gold") and tm.has_key("silver") else "S=0;G=0" 
                            it = PajInvItem(invno, p17, jn, tr[tm["qty"]], tr[tm["price"]], mps, tr[tm["stone"]], lmd)                           
                            items[key] = it
            finally:
                if updcnt > 0 : wb.save()
                wb.close()
            x = (0, None) if len(dups) + len(items) == 0 else self._persist(dups, items)
            if(x[0] == -1):
                errs.append(x[1])
            else:
                logging.debug("invoice (%s) processed" % fn + ("" if x[0] else " but all are repairing"))
        if(killexcel): app.quit()
        return x[0], items if(x[0] == 1) else () if x[0] == 0  else x[1]    
    
    '''
    def __del__(self):
        """release the database"""
        if(self._accessdb): 
            self._accessdb.close()
            logging.debug("database close()")
    '''
