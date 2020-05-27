'''
#! coding=utf-8
@Author:        zmFeng
@Created at:    2019-07-09
@Last Modified: 2019-07-09 3:32:33 pm
@Modified by:   zmFeng
Based on Workbook Controller, write/read data between workbook and data source
'''

from datetime import date
from numbers import Number

from hnjapp.c1calc import Utilz, Writer
from hnjapp.svcs.misc import LKSizeFinder, StyleFinder, StylePhotoSvc
from utilz import NA, karatsvc
from utilz.miscs import NamedList, NamedLists, triml, trimu
from utilz.xwu import (FormHandler, apirange, appmgr, appswitch, esctext,
                       fromtemplate, usedrange, insertphoto)

from ._nrlib import (BaseNrl, _NRInvoker)
from ._utilz import _Tbl_Optr, _Utilz, config
from ._misc import _SNS2Wgt


class JOFormHandler(FormHandler):
    ''' A ProductSpec FormHandler with Nrls binded
    '''

    def __init__(self, wb, hdlr=None, **kwds):
        super().__init__(wb, hdlr=hdlr, **kwds)

    def read(self, names=None, nrls=None):
        if not nrls:
            nrls = _NRInvoker(self, True, BaseNrl.LEVEL_ADVICE)
        return super().read(names=names, nrls=nrls)

    @staticmethod
    def read_field_map(wb):
        ''' read the column mapping fields out from the product form's meta sheet
        '''
        sht = wb.sheets('metadata')
        addr = apirange(sht.api.listObjects('m_fmp').Range)
        lst = addr.value
        if lst[0][2] == 'firstaddress':
            lst[0][2] = 'address'
        return NamedLists(lst)

    def _read_field_map(self):
        return JOFormHandler.read_field_map(self._wb)

    def insert_imgs(self, styno, jono=None):
        ''' insert image to the form
        '''
        self.sheet.api.Unprotect()
        stysvc = StylePhotoSvc.getInst()
        fns = stysvc.getPhotos(styno, hints=jono)
        if not fns:
            return
        pns = ['photo%d' % x for x in range(2)]
        for idx, pn in enumerate(pns):
            if idx >= len(fns):
                break
            rng = self.get(pn).get(merged=True)
            insertphoto(fns[idx], rng).name = pn
        self.sheet.api.Protect(DrawingObjects=True, Contents=True, Scenarios=True)


class FromJO(object):
    '''
    read data from JO and return it as a ready-to-fill data set
    for :class: FormHandler/FromJO.write()
    '''
    def __init__(self, hdlr, hksvc, cnsvc, cache_sm):
        '''
        need hdlr to read the field maps
        '''
        self._hdlr, self._hksvc, self._cnsvc = hdlr, hksvc, cnsvc
        self._mpr = _Tbl_Optr(hdlr)
        self._styfndr = StyleFinder(cache_sm) if cache_sm else None
        self._sns2wgt = _SNS2Wgt()

    def nrm(self, name):
        return self._hdlr.normalize(name)

    @staticmethod
    def from_jns(jns, tpl_file=None, **kwds):
        '''
        fetch data of given jns from HK system and write them back to
        a form

        Args:
            jns:    tuple(String)
            hksvc=None:  An :class: HKSvc instance help to get data from HK system

            cnsvc=None:  An :class: CNSvc instance help to get data from CN system

            cache_sm=None:   A sessionManager help to get data(lksize/familiar style) from cache

            noimg=None: don't insert image
        '''
        cnt = 0
        app, tk = appmgr.acq()
        sws = appswitch(app, {'EnableEvents': False})
        if not tpl_file:
            tpl_file = config.get("prdspec.template")
        hdlr = JOFormHandler(fromtemplate(tpl_file, app))
        rmp = FromJO(hdlr, kwds.get('hksvc'), kwds.get('cnsvc'), kwds.get('cache_sm')).retrieve(jns)
        noimg = kwds.get('noimg')
        for jn, rst_hints in rmp.items():
            cnt += 1
            if cnt > 1:
                hdlr = JOFormHandler(fromtemplate(tpl_file, app), hdlr)
            nlzr = _NRInvoker(hdlr, True, BaseNrl.LEVEL_ADVICE)
            mp = nlzr.normalize(rst_hints)[0]
            styn = mp["_styno"]
            jn = esctext(jn)
            if not noimg:
                hdlr.insert_imgs(styn, jn)
            yield (jn, styn), hdlr.book, mp
        appmgr.ret(tk)
        appswitch(app, sws)

    def retrieve(self, jns):
        '''
        retrieve the data of given jonos from HK system or excel workbook

        Args:

            jns(tuple of String): the jO#s. Can be one JO# or file name of a workbook.
            When it refers to a workbook, the hksvc paramater should be None
        Returns:
            A {jono(String): map} where map has the same format as Handler.read(), that means it can be write back to excel directly by Handler.write().
            One more thing, the map is not normalized by _NRInvoker, do it yourself.
        '''
        if isinstance(jns, str) or not self._hksvc:
            raw, jodtls = self._from_wb(jns), None
        else:
            wtr = Writer(self._hksvc, self._cnsvc, keep_dtl=True)
            raw = wtr.for_x(jns)
            if not raw:
                return None
            raw, jodtls = raw
        nl_src, rmp = raw[0], {}
        # because _fetch_stone might generate finishing data that should override
        # the one _fetch_finishing. so two loop needed
        _nrm = self._hdlr.normalize
        for row in raw[1:]:
            nl_src.setdata(row)
            jn = nl_src.jono # this jo# has a ' as prefix, for excel inserting
            if jn:
                self._mpr.mp = rmp[jn] = mp = {_nrm('docno'): jn}
                self._fetch_metal(nl_src)
                mp['_styno'] = styno = nl_src['styno'] # _styno for convenience only
                mp['_wgtinfo'] = jodtls[esctext(jn)]['wgt']
                if self._styfndr:
                    sf = self._styfndr
                    lst = LKSizeFinder(sf, stynos=(styno, )).find()
                    if lst:
                        mp[_nrm('size')] = ','.join([x[0] for x in lst])
                    lst = sf.find_by_feature(styno)
                    if lst:
                        mp[_nrm('_family')] = lst
                    lst = sf.getstyx(styno, 'snno')
                    if lst:
                        # next processor will get it from below
                        jodtls[esctext(jn)]['sn'] = ";".join(lst)
                    lst = sf.getstyx(styno, 'keyword')
                    if lst:
                        mp[_nrm('_keyword')] = ";".join(sorted(lst))
                    lst = sf.getstyx(styno, 'text')
                    if lst:
                        mp[_nrm('_text')] = ";".join(lst)
                mp[_nrm('type')] = Utilz.getStyleCategory(styno)
            self._fetch_stone(nl_src)
        for row in raw[1:]:
            nl_src.setdata(row)
            jn = nl_src.jono # this jo# has a ' as prefix, for excel inserting
            if jn:
                self._mpr.mp = mp = rmp[jn]
            self._fetch_finishing(nl_src)
        return self._normalize(rmp, jodtls)

    def _normalize(self, rmp, jodtls):
        for mp in rmp.values():
            mp['author'] = '_BY_PRG'
            mp['createdate'] = mp['lastmodified'] = date.today()
            self._mpr.mp = mp
            self._nrl_metal()
            self._nrl_parts()
            self._nrl_stone()
            _FromJO_BL(self, self._hdlr, mp, jodtls.get(esctext(mp[self.nrm('docno')]))).bl()
        return rmp

    def _nrl_parts(self):
        tn, mp = self.nrm('parts'), self._mpr.mp
        nls = mp.get(tn)
        if not nls:
            return
        kt = mp.get('_m_karat')
        for nl in nls:
            wgt = _Utilz.get_parts_wgt(nl['matid'], kt, True)
            if wgt:
                nl.matid, nl.remarks, wgt = wgt[1], nl.matid, wgt[0]
            nl.karat, nl.qty, nl.wgt = kt, 1, wgt or 0.001
            pt = nl.matid
            if pt.find('"') > 0 or pt.find("'") > 0:
                nl.type = 'XP'
        if len(nls) > 1:
            tmp = {'XP': 0, 'MP': 1}
            nls = sorted([nl for nl in nls], key=lambda nl: tmp.get(nl.type, 99))
        mp[tn] = nls

    def _nrl_metal(self):
        tn, mp = self.nrm('metal'), self._mpr.mp
        nls = mp.get(tn)
        if not nls:
            return
        if len(nls) > 1:
            nls = sorted([x for x in nls], key=lambda x: x.wgt, reverse=True)
            mp[tn] = nls
        for nl in nls:
            kt = nl.karat
            if isinstance(kt, Number):
                nl.karat = str(int(kt))
        if len(nls) > 1:
            mp[tn] = sorted(nls, key=lambda nl: nl.wgt, reverse=True)
        mp['_m_karat'] = nls[0].karat

    def _nrl_stone(self):
        tn, mp = self.nrm('stone'), self._mpr.mp
        nls = mp.get(tn)
        if not nls:
            return
        for nl in nls:
            nl.matid = st = nl.name
            if not nl.remarks:
                nl.remarks = st
            nl.name = st[:2]
            # when it's diamond, the size is sieze
            if self._is_dd(st):
                nl.size = nl.size or 'XX' + "SZ"
            var = nl.shape or st[2]
            if var:
                var = config.get("prdspec.fromjo.stshape_map").get(var)
                if var:
                    nl.shape = var
            if st == 'CZ' and nl.wgt and nl.wgt >= 1 or nl.wgt and nl.qty and abs(nl.wgt - nl.qty) < 0.001: # CZ or fake weight from HK
                nl.wgt = 0
            else:
                nl.wgt = round(nl.wgt or 0, 3)
            nl.wgtunit = 'CT'
            if not nl.wgt:
                lvl, uwgt = config.get("prdspec.stsnscvt.level"), None
                if len(st) > 5 and self._cnsvc and lvl > 0:
                    uwgt = self._cnsvc.getavgpkwgt(st, nl.size)
                elif lvl > 1:
                    uwgt = self._sns2wgt.get(st, nl.shape, nl.size)
                    if not uwgt and self._cnsvc and lvl > 2:
                        uwgt = self._cnsvc.getavgstwgt(st, nl.shape, nl.size)
                if uwgt:
                    nl.unitwgt = uwgt
            var = nl.setting
            nl.main = 'Y' if var.find('主') >= 0 else 'N'
            stmp = config.get("prdspec.fromjo.setting_map")
            nl.setting = stmp.get(var[:2]) or stmp.get(var) or var
        if len(nls) > 1:
            mp[tn] = sorted(nls, key=lambda nl: (nl.main, nl.size or '0', -ord(nl.name[0])), reverse=True)

    def _fetch_metal(self, nl_src):
        for idx in range(2):
            kt = nl_src['f_karat%d' % idx]
            if not kt:
                break
            self._mpr.set_metal(kt, nl_src['f_wgt%d' % idx])

    def _fetch_finishing(self, nl_src):
        exts = self._mpr.mp.get(self.nrm('finishing'))
        exts = set(x.remarks for x in exts) if exts else {}
        vmp = None
        for fld, rmk, vc in (('f_pen', 'HIGHLIGHT', 'VW'), ('f_chain', 'CHAIN', 'V?'), ('f_tt', 'VT/T', 'VX')):
            if rmk in exts:
                continue
            var = nl_src[fld]
            if not var:
                continue
            if vc == 'V?':
                mtls = self._mpr.mp.get(self.nrm('metal'))
                if mtls:
                    mtls = mtls[0]
                    mtls = karatsvc.getkarat(mtls['karat']).color
                    if not vmp:
                        vmp = {x['color']: x['name'] for x in config.get('vermail.defs')}
                    vc = vmp.get(mtls, 'VW')
            self._mpr.add_finishing(vc, '0.125 MIC', rmk)
        var = nl_src['f_micron']
        if var:
            for vc in var.split(";"):
                vc = vc.split("=")
                self._mpr.add_finishing(vc[0], vc[1], None)

    def _fetch_stone(self, nl_src):
        st = nl_src['stone']
        if not st:
            return
        st = st.split(";")
        spec = None
        if st and len(st) > 1:
            spec = ";".join(st[1:])
        st = st[0]
        if not st or st == 'MIT':
            if st:
                rmk = nl_src['setting']
                idx = rmk.find('相')
                if idx < 0:
                    x = Utilz.extract_micron(rmk)
                    if x:
                        self._mpr.add_finishing('VK', '%d MIC' % x, 'CHAIN')
                    idx, eidx = rmk.find('('), -1
                    if idx < 0 and x:
                        idx = rmk.find('電')
                        if idx >= 0:
                            idx, eidx = idx - 1, None
                    if idx > 0:
                        x, rmk = rmk[idx + 1: eidx], rmk[:idx]
                    else:
                        x = None
                    self._mpr.add_parts(x, 'MP', 0, rmk)
                elif not self._mpr.mp.get(self.nrm('size')):
                    self._mpr.mp[self.nrm('size')] = _Utilz.get_lksz(rmk)
            return
        nl = self._mpr.append('stone')
        nl.name = st
        nl.shape, nl.qty, nl.wgt, nl.setting = [nl_src[x] for x in 'shape stqty stwgt setting'.split()]
        nl.size, nl.remarks = esctext(nl_src['stsize']), spec

    def _is_dd(self, stname):
        return stname[0:2] in ('DD', 'DF', 'DY')

    def _from_wb(self, fn):
        app, tk = appmgr.acq()
        if not fn:
            fn = r"d:\pyhome\python\tests\c1calc\_expects.xlsx"
            fn = r"d:\pyhome\python\tests\prdspec\C1CC_MStone.xlsx"
            # fn = r"d:\pyhome\python\tests\prdspec\C1CC_CZ_TT_MP_MICInMIT.xlsx"
        wb = app.books.open(fn)
        sht = wb.sheets('计价资料')
        # max row detection
        m_row = usedrange(sht).rows.count
        er = lambda c: sht.range("%s%d" % (c, m_row)).end('up').row
        row = max(er('Y'), er('E'))
        raw = sht.range("E10:Y%d" % row).value
        wb.close()

        wb = fromtemplate(config.get('default.file.c1cc.tpl'), app)
        nl = Writer(None, None).find_sheet(wb)[-1]
        wb.close()
        appmgr.ret(tk)

        raw.insert(0, nl)
        return raw


class _FromJO_BL(object):

    def __init__(self, frmjo, hdlr, mp, jodtl):
        self._frmjo, self._mp, self._jodtl = frmjo, mp, jodtl
        self._hdlr = hdlr
        self._mpr = _Tbl_Optr(self._hdlr, mp)
        self._meta_mp = None

    def _nrm(self, name):
        return self._hdlr.normalize(name)

    def bl(self):
        self._mpr.set_feature('XSTYN', self._mp['_styno'])
        self._mpr.set_feature('SN', self._mp.get('_snno') or self._jodtl.get('sn') or NA)
        kwds = None
        bits = self._build_bits()
        #sometimes there is casting locket, but just let it be
        if bits["locket"] or bits['stp']:
            self._mpr.set_feature('REMARKS', '片厚_C')
        self._mp['hallmark'] = ";".join(bits['hm'])
        if bits['stp']:
            self._mp[self._nrm('craft')] = 'MIXTURE' if bits['cst'] else 'STAMPING'
        if bits['enamel']:
            self._mpr.set_item('finishing', 'method', 'ENAMEL', 'spec', '_CLR')
        if bits['cat'] == 'earring':
            x = ('夾針', '夾針@xxmm') if bits['stp'] else ('耳迫', '大中耳迫')
            self._mpr.add_parts(x[0], 'MP', 2, x[1])
        if bits["locket"]:
            rmk = self._jodtl['rmk']
            idx = rmk.find('節')
            x = "一_三節鉸" if idx < 0 else rmk[idx - 1] + "節鉸"
            self._mpr.set_feature('REMARKS', x)
            self._mpr.add_parts(self._textile(rmk), 'PKM', 2)
            self._mpr.add_parts('膠片', 'PKM', 2)
            x = self._nrm('size')
            if not self._mp.get(x):
                shp, self._mp[x] = '_HEART', '_HL:_MM'
            else:
                shp = config.get('locket.shapes.eng').get(self._mp[x].split(':')[0], '_HEART')
            kwds = '%s;LOCKET' % shp
        for key, val in {'vw': 'VW', 'vy': 'VY'}.items():
            if bits[key]:
                self._mpr.add_finishing(val, '0.125 MIC', 'PROTECTION')
        if bits["cat"] == 'ring':
            self._hdlr.get('size').value = '_N'
        self._check_glue()
        self._check_cards(bits)
        self._check_fin_by_mtl()
        # this make text has higher position than keywords. maybe not good
        txt = self._check_text()
        if txt:
            kwds = (kwds or '') + ";TEXT"
        x = self._mp.get('_keyword')
        if x:
            kwds = (kwds or '') + ";" + x
        if kwds:
            self._mpr.set_feature('KEYWORDS', kwds)
        if txt:
            self._mpr.set_feature('TEXT', txt)

    @staticmethod
    def _textile(rmk):
        idx = rmk.find('色')
        if idx > 0:
            x = -1
            while x > -5:
                if ord(rmk[idx + x]) < 250:
                    break
                x -= 1
            if x > -5:
                x += 1
        else:
            x = -1
        return ('_' if idx < 0 else rmk[idx + x: idx]) + "色绒布"

    def _check_glue(self):
        sts = self._mp.get(self._nrm('stone'))
        if not sts:
            return
        xtr = lambda idx: ss[idx][idxs[idx] - 1]
        for nl in sts:
            if nl.name.find('GC') < 0: # stone name here not yet extended, still abbr
                continue
            nl.setting = 'GLUE'
            ss = [self._jodtl[x] for x in ('description', 'rmk')]
            idxs = [x.find('色') for x in ss]
            if any((x > 0 for x in idxs)):
                clr = xtr(0) if idxs[0] > 0 else xtr(1)
            else:
                clr = '白'
            clr += '色泥膠'
            self._mpr.add_finishing('OP', NA, '力架')
            self._mpr.add_parts(clr, 'GEL', matid='GLUE')
            return nl
        return None

    def _check_cards(self, bits):
        cn, cnt = self._jodtl['cstname'], 0
        def apt(pn, qty=1):
            self._mpr.add_parts(pn, 'PKM', qty)
            nonlocal cnt
            cnt = cnt + 1
        if cn == 'ESO':
            for nm, cd in (('925', '銀卡'), ('kw', '白金卡')):
                if bits[nm]:
                    apt(cd)
            cat = bits['cat']
            for nm, cd in (('earring', '耳環卡'), ('pendant', '吊墜卡'),):
                if cat == nm:
                    apt(cd)
                    break
        elif cn in ('ELH', 'EJE'):
            if bits['locket']:
                apt('相盒卡')
            if bits['tc'] > 1 and bits['925']:
                apt('金夾銀卡')
            if self._jodtl['description'].find('童') >= 0:
                apt('特別/童裝卡')
        elif cn == 'EJW' and self._mp['_styno'] in ('P27220', 'P27221'):
            apt('EJW專用卡')
        elif self._mp['_styno'] == 'P39731':
            apt('四頁相盒教學卡')
        if cnt > 0:
            self._mpr.set_feature('REMARKS', '卡要逐一隨貨包裝')
        if bits['kw']:
            self._mpr.set_feature('REMARKS', '不可含NICK')

    def _check_fin_by_mtl(self):
        ''' the finishing for protection
        '''
        mtls = self._mp.get(self._nrm('metal'))
        if not mtls:
            return
        fins = self._mp.get(self._nrm('finishing'))
        fins = set((nl['method'] for nl in fins)) if fins else set()
        for mtl in mtls:
            kt = mtl.karat
            if kt.find('BG') >= 0:
                if 'VY' not in fins:
                    self._mpr.add_finishing('V_Y', '0.125 MIC', 'PROTECTION')
                    fins.add('VY')
            elif str(kt) == '925':
                if not any((1 for x in ('VW', 'VK') if x in fins)):
                    self._mpr.add_finishing('V_W', '0.125 MIC', 'PROTECTION')
                    fins.add('VW')
            elif kt == '9K' and self._jodtl.get('cstname') == 'GAM':
                if 'VY' not in fins:
                    self._mpr.add_finishing('VY', '0.125 MIC', 'PROTECTION')
                    fins.add('VY')

    def _check_text(self):
        ''' check if there is text inside JO's remark, if yes, extract it to keywords
        '''
        flag = False
        txt = self._mp.get('_text')
        if not txt:
            flag, txt = True, _Utilz.get_text(self._jodtl['rmk'])
        if txt:
            txt = ('_' if flag else '') + txt
        return txt

    def _build_bits(self):
        # tc for tune_count, sz for size, hm for hall-mark, cst for casting, stp for stamping, cat for style category
        bits = NamedList('locket cat ky kw kr 925 engraving tc vy vw vr enamel sz hm cst stp'.split())
        bits['hm'] = []

        desc, rmk = (self._jodtl.get(x) for x in ('description', 'rmk'))
        cns = [desc.find(x) > 0 for x in ('倒', '啤')]
        if any(cns):
            bits['cst'], bits['stp'] = cns[0], True
        if not bits['stp']:
            bits['stp'] = self._check_stp_by_sn()
        if not bits['stp']:
            bits['cst'] = True
        bits['locket'] = flag = desc.find('相') >= 0
        if flag:
            bits['sz'] = "_HL:_MM"
        styno = self._mp['_styno']
        bits['cat'] = triml(StylePhotoSvc.getCategory(styno, desc))
        if styno[0] == 'B' and desc.find('鏈') > 0:
            # TODO:: detect chain length
            bits['sz'] = '_"'
        for var in (('engraving', '批花'), ('enamel', '燒青')):
            for x in (desc, rmk):
                if x.find(var[1]) >= 0:
                    bits[var[0]] = True
        for var in (('vy', '電黃'), ('vy', '電王'), ('vw', '電白'), ('vr', '電玫瑰')):
            bits[var[0]] = rmk.find(var[1]) >= 0
        self._build_bits_hm(bits)
        return bits

    def _check_stp_by_sn(self):
        sn = self._jodtl.get('sn')
        if not sn:
            return False
        baleset = set(config.get('snno.translation').values())
        for var in sn.split(';'):
            if var[:2] == 'HB' and var not in baleset:
                return True
        return False

    def _build_bits_hm(self, bits):
        ''' hall-marks
        '''
        kts = [karatsvc.getkarat(x.karat) for x in self._mp[self._nrm('metal')]]
        bits['tc'], cn = len(kts), self._jodtl['cstname']
        hms = bits['hm']
        for var in kts:
            hm = config.get('prdspec.metal.mark').get(var.name)
            if hm:
                if var.karat == 925:
                    bits['925'] = True
                if var.category != 'GOLD' and any((bits[x] for x in ('vy', 'vr'))):
                    # non gold, vgold, mark
                    x = config.get("prdspec.customer.vgmark").get(cn)
                    if x:
                        hms.append(x)
                hms.append(hm)
            if var.category != 'GOLD':
                continue
            var = trimu(var.color)
            for x in (('y', "YELLOW"), ('w', "WHITE"), ('r', "ROSE")):
                if var.find(x[1]) >= 0:
                    bits['k' + x[0]] = True
                    break
        hm = config.get('prdspec.customer.mark').get(cn)
        if hm:
            hms.append(hm)
        var = self._mp.get(self._nrm('stone'))
        if var and any((cn.find(x)) >= 0 for x in ('ESO', 'ELH', 'EJE')):
            cntmp = set()
            for x in var:
                sn = x.name
                if sn == 'CZ' and sn not in cntmp:
                    hms.append('CZ')
                    cntmp.add(sn)
                elif 'DIA' not in cntmp and self._frmjo._is_dd(sn):
                    hms.append('DIA')
                    cntmp.add('DIA')
        return bits
