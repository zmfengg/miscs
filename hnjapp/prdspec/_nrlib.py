#! coding=utf-8
'''
* @Author: zmFeng
* @Date: 2019-05-29 09:53:51
* @Last Modified by:   zmFeng
* @Last Modified time: 2019-05-29 09:53:51
local normalizers for a product specification form
'''

import inspect
from abc import abstractmethod
from difflib import SequenceMatcher
from os import path

from hnjapp.c1rdrs import _fmtpkno
from hnjapp.svcs.db import _JO2BC, formatsn
from hnjcore import JOElement
from utilz import NA, karatsvc, stsizefmt, trimu
from utilz.miscs import NamedLists
from utilz.xwu import BaseNrl, NrlsInvoker, SmpFmtNrl, TBaseNrl, apirange, esctext

from ..common import config
from ._utilz import _Utilz

thispath = path.abspath(path.dirname(inspect.getfile(inspect.currentframe())))

# not sure
_nsr = lambda s: s and s.find('_') >= 0
# level for not sure
_nsl = lambda s, lv: BaseNrl.LEVEL_ADVICE if _nsr(s) else lv


class JENrl(SmpFmtNrl):
    """ the JOElement like normalizer
    """
    @staticmethod
    def _fmt_je(jn):
        je = JOElement(jn)
        if not je.isvalid():
            return jn
        return je.value

    def __init__(self, *args, **kwds):
        super().__init__(JENrl._fmt_je, *args, **kwds)


class KaratNrl(BaseNrl):
    '''
    normalizer for karat
    '''
    def __init__(self, strict=True, **kwds):
        """
        Args:
            strict: False to allow karat that's not defined
        """
        super().__init__(**kwds)
        self._strict = strict

    def normalize(self, pmap, name):
        k0 = esctext(pmap[name])
        nv = rmk = None
        kt = karatsvc.getkarat(k0)
        if kt:
            nv, ext = kt.name, False
            # excel might thread sth. like 925 as numeric
            if nv.isnumeric():
                nv, ext = "'" + nv, True
            if nv != k0:
                lvl = self.LEVEL_RPLONLY if ext else self.LEVEL_ADVICE
                rmk = 'Karat(%s) => (%s)' % (k0, nv)
        elif self._strict:
            rmk, lvl = 'Invalid Karat(%s)' % k0, self.LEVEL_CRITICAL
        return  (self._new_log(k0, nv, name=name, remarks=rmk, level=lvl), ) if rmk else None

class SizeNrl(BaseNrl):
    '''
    formatting stone size or object dimension
    '''
    def normalize(self, pmap, name):
        if name in pmap:
            ov = trimu(str(pmap[name]))
            if not ov or ov.find(':') > 0: # Locket size case
                return None
            sfx = ov.find('SZ')
            if sfx > 0:
                ov, sfx = ov[:sfx], ov[sfx:]
            else:
                sfx = ''
            nv = stsizefmt(str(ov), True) + sfx
            if nv == ov:
                if nv and _nsr(nv):
                    return  (self._new_log(ov, nv, level=self.LEVEL_ADVICE),)
                return super().normalize(pmap, name)
            return  (self._new_log(ov, nv, name=name, remarks=None, level=_nsl(nv, self.LEVEL_RPLONLY)),)
        return super().normalize(pmap, name)

class DescNrl(BaseNrl):
    ''' make the description based on the data map
    '''

    def __init__(self, **kwds):
        self._nrm = kwds.get('nrm') or _Utilz.nrm
        super().__init__(**kwds)

    def _knc(self, pmap):
        lsts, wgts = [], pmap.get('_wgtinfo')
        if wgts:
            nls = pmap.get(self._nrm('finishing'))
            if nls:
                vcmp = {x['name']: x['color'] for x in config.get('vermail.defs')}
                nls = {(vcmp.get(nl['method']), ) for nl in nls if nl['method'][0] == 'V'}
            lsts.append(_JO2BC().knc_mix(wgts, nls),)
        else:
            nls = pmap.get(self._nrm('metal'))
            if nls:
                lsts.append("&".join([nl.karat for nl in nls]))
            nls = pmap.get(self._nrm('finishing'))
            if nls:
                # merge the duplicated
                nls = set(nl['method'] for nl in nls if nl['method'][0] == 'V')
                if nls:
                    lsts.append("&".join(nls))
        return lsts

    def normalize(self, pmap, name):
        lsts, ov = self._knc(pmap), pmap.get(name)
        nls = self._make_desc_st(pmap)
        if nls:
            lsts.append(nls)
        ptype = pmap[self._nrm('type')]
        if ptype == 'PENDANT':
            nls = pmap.get(self._nrm('feature'))
            if nls:
                nls = [nl for nl in nls if nl['catid'] == 'KEYWORDS' and nl['value'].find('LOCKET') >= 0]
                if nls:
                    ptype = 'LOCKET PENDANT'
        lsts.append(ptype)
        nls = pmap.get(self._nrm('parts'))
        if nls:
            if [nl for nl in nls if nl['type'] == 'XP']:
                lsts.append('ROPE CHAIN')
        nv = " ".join(lsts)
        pmap[name] = nv
        return  (self._new_log(ov, nv, name=name, remarks=ov, level=self.LEVEL_ADVICE),)

    def _make_desc_st(self, pmap):
        ''' sort by stone size desc, then stone name. Dia will always be at the end
        '''
        sts = pmap.get(self._nrm('stone'))
        if not sts:
            return None
        sts = sorted([st for st in sts], key=lambda st: str(st['size']), reverse=True)
        pfx, sfx = set(), None
        for st in sts:
            sn = st['name']
            if sn.find('DIAMOND') >= 0:
                sfx = sn
            else:
                pfx.add(sn)
        pfx = " ".join(pfx) if len(pfx) < 3 else "RAINBOW"
        if sfx:
            pfx = pfx + " " + sfx
        return pfx

class NWgtNrl(BaseNrl):
    ''' NetWgt calculator
    '''

    def __init__(self, **kwds):
        self._nrm = kwds.get('nrm') or _Utilz.nrm
        super().__init__(**kwds)

    def normalize(self, pmap, name):
        nwgt = 0
        nls = pmap.get(self._nrm('metal'))
        if nls:
            nwgt += sum([nl.wgt for nl in nls if nl.wgt])
        nls = pmap.get(self._nrm('stone'))
        if nls:
            uv = {'CT': 0.2, 'GM': 1, 'TL': 37.429}
            nwgt += sum([nl.wgt * uv.get(nl['wgtunit'], 0.2) for nl in nls if nl.wgt])
        nls = pmap.get(self._nrm('parts'))
        if nls:
            nwgt += sum([nl.wgt for nl in nls if nl['type'] == 'XP' and nl['karat'] and nl.wgt])
        ov, nwgt = pmap.get(name), round(nwgt, 3)
        pmap[name] = nwgt
        return  (self._new_log(ov, nwgt, name=name, remarks=None, level=self.LEVEL_ADVICE if ov else self.LEVEL_RPLONLY),)


class _TBaseNrl(TBaseNrl):

    def __init__(self, **kwds):
        super().__init__(**kwds)

    @abstractmethod    
    def _nrl_row(self, name, row, nl, logs):
        ...

    @abstractmethod
    def _get_nrls(self):
        ...

    def _append_log(self, logs, ov, nv, rmk, **kwds):
        if nv == ov:
            return
        if rmk and 'remarks' not in kwds:
            kwds['remarks'] = rmk
        nl = tuple(kwds.get(x) for x in ('nl', 'colname'))
        if all(nl):
            nl, cn = nl
            if cn in nl.colnames:
                nl[cn] = nv
        logs.append(self._new_log(ov, nv, **kwds))


class MetalNrl(_TBaseNrl):
    '''
    Metal table normalizer
    '''

    def __init__(self, **kwds):
        super().__init__(**kwds)
        self._nrl_mp = None


    def _get_nrls(self):
        if not self._nrl_mp:
            self._nrl_mp = {
                'remarks': self.get_base_nrm('tu')
            }
        return self._nrl_mp

    def _nrl_row(self, name, row, nl, logs):
        colname = 'wgt'
        def _a_log(ov, nv, lvl, rmks):
            self._append_log(logs, ov, nv, rmks, name=name, row=row,
                colname=colname, nl=nl, level=lvl)
        wgt = nl[colname] or 0
        if wgt <= 0:
            _a_log(1, None, self.LEVEL_CRITICAL, 'wgt <= zero')
        else:
            kts, flag = str(nl['karat']).split("-"), True
            if len(kts) > 1:
                kts[1] = karatsvc.getkarat(kts[1])
                wgt1 = karatsvc.convert(kts[0], wgt, kts[1])
                if wgt != wgt1:
                    _a_log(wgt, wgt1, self.LEVEL_ADVICE if wgt1 else self.LEVEL_CRITICAL,
                        '(%s=%4.2f) => (%s=%4.2f)' % (kts[0], wgt, kts[1].name, wgt1))
                    wgt = wgt1
                flag, nl['karat'], colname = False, kts[1].name, 'karat'
                _a_log(None, kts[1].name, self.LEVEL_RPLONLY, None)
            if flag:
                self._normalize_item(KaratNrl(), nl, 'karat', logs=logs, row=row, name=name)
        # basic BL snippet
        kt = karatsvc.getkarat(nl['karat'])
        if kt and kt.category == 'GOLD' and wgt > 3:
            colname = 'wgt'
            _a_log(None, wgt, self.LEVEL_ADVICE, 'gold weight(%4.2f) greater than %4.2f, maybe error' % (wgt, 3))


class FinishingNrl(_TBaseNrl):
    '''
    normalizer for finishing
    '''

    def __init__(self, **kwds):
        super().__init__(**kwds)
        self._nrl_mp = None
        self._nrm = kwds.get('nrm') or _Utilz.nrm
        self._hints = kwds.get("hints")

    def _get_nrls(self):
        if not self._nrl_mp:
            tu = self.get_base_nrm('tu')
            self._nrl_mp = {
                'remarks': self.get_base_nrm('tr'),
                'method': tu,
                'spec': tu
            }
        return self._nrl_mp

    def _nrl_row(self, name, row, nl, logs):
        if not self._hints:
            return
        for colname in 'spec method'.split():
            mt, cand = self._hints(self._nrm('finishing.%s' % colname), nl[colname])
            if mt == 2 and cand:
                cand = trimu(cand)
                self._append_log(logs, nl[colname], cand, '(%s) => (%s)' % (nl[colname], cand), name=name, row=row, colname=colname, nl=nl, level=self.LEVEL_ADVICE)
        self._fix_vk(nl)

    @staticmethod
    def _fix_vk(nl):
        spec = nl['spec']
        if not spec:
            return
        idx = spec.find('MIC')
        if idx <= 0:
            return
        h = float(spec[:idx].strip())
        if h >= 1:
            spec = nl['method']
            if spec[1] != 'K':
                nl['method'] = 'VK' if spec[1] == 'Y' else spec[0] + 'K' + spec[1:]

class PartsNrl(_TBaseNrl):
    '''
    normalizer for Parts
    '''

    def __init__(self, **kwds):
        super().__init__(**kwds)
        self._nrl_mp = None
    
    def _nrl_row(self, name, row, nl, logs):
        pass

    def _get_nrls(self):
        if not self._nrl_mp:
            tu = self.get_base_nrm('tu')
            self._nrl_mp = {'remarks': self.get_base_nrm('tr'), 'type': tu, 'matid': tu, 'karat': (tu, KaratNrl(False))}
        return self._nrl_mp

class StoneNrl(_TBaseNrl):
    '''
    normalizer for stone
    .minor BL as:
        .Too many stone can not be main stone(ER=2, other=1)
        .There can be only one main stone item
        .Stone Name correction
        .Size to weight
    '''

    def __init__(self, **kwds):
        super().__init__(**kwds)
        self._nrl_mp = {'size': SizeNrl()}
        self._hints = kwds.get('hints')
        self._nrm = kwds.get('nrm') or _Utilz.nrm

    def _get_nrls(self):
        return self._nrl_mp

    def _nrl_row(self, name, row, nl, logs):
        colname = 'matid'
        ov, nv = nl[colname], None
        def _log(ov, nv, rmk, level=self.LEVEL_ADVICE):
            self._append_log(logs, ov, nv, rmk, name=name, nl=nl, row=row, colname=colname, level=level)
        if ov:
            nv = _fmtpkno(ov)
            if nv:
                nv = nv[0]
            _log(ov, nv, 'PK# formatted')
        qty, uw, wgt = [nl.get(x) or 0 for x in ('qty', 'unitwgt', 'wgt')]
        if not qty or qty <= 0:
            colname = 'qty'
            nv = qty = 1
            _log(ov, nv, 'Invalid qty(%r), set to %s' % (ov, nv))
        if uw * wgt == 0:
            if uw + wgt == 0:
                for colname in ('unitwgt', 'wgt'):
                    _log(ov, None, '(%s) not nullable' % colname, self.LEVEL_CRITICAL)
            else:
                if wgt:
                    colname, uw = 'unitwgt', round(wgt / qty, 3)
                else:
                    colname, uw = 'wgt', round(uw * qty, 3)
                _log(None, uw, '(%s) calculated' % colname)
        for colname, df in (('main', 'N'), ('wgtunit', 'CT')):
            ov = nl[colname]
            nv = trimu(ov or df)
            _log(ov, nv, None, self.LEVEL_RPLONLY)

        if self._hints:
            for colname in 'shape name'.split():
                ov = nl[colname]
                mt, nv = self._hints(self._nrm('stone.%s' % colname), ov)
                if mt > 0 and nv and nv != ov:
                    nv = trimu(nv)
                    self._append_log(logs, ov, nv, '%s -> %s' % (nl[colname], nv), name=name, row=row, colname=colname, nl=nl, level=self.LEVEL_ADVICE)

    def normalize(self, pmap, name):
        logs = super().normalize(pmap, name)
        if not logs:
            logs = []
        stones, ms = pmap.get(name), []
        if not stones or len(stones) == 1:
            return None
        for row, nl in enumerate(stones):
            valid = nl.qty == (1 if pmap.get(self._nrm('type')) != "EARRING" else 2) and stsizefmt(str(nl.size), False) >= "0300"
            ov = nv = nl[self._nrm('main')]
            if valid ^ (ov == 'Y'):
                nv = 'Y' if valid else 'N'
                self._append_log(logs, ov, nv, 'should be %s' % ('MAIN STONE' if valid else 'SIDE STONE'), nl=nl, name=name, row=row, colname='main', level=self.LEVEL_ADVICE)
            if nv == 'Y':
                ms.append(row)
        if len(ms) > 1:
            for row in ms:
                self._append_log(logs, None, 'Y', 'MULTI-MAIN is not allowed', name=name, nl=stones[row], row=row, colname='main', level=self.LEVEL_CRITICAL)
        return logs if logs else None


class FeatureNrl(_TBaseNrl):
    """ feature based normalizer
    """

    def __init__(self, **kwds):
        self._nrm = kwds.get('nrm') or _Utilz.nrm
        super().__init__(**kwds)

    def _get_nrls(self):
        return None

    def _nrl_row(self, name, row, nl, logs):
        colname = 'value'
        def _a_log(logs, ov, nv, rmk, lvl=self.LEVEL_RPLONLY):
            self._append_log(logs, ov, nv, rmk, nl=nl, name=name, row=row, colname=colname, level=lvl)
        catid = self._nrm(nl['catid'])
        if catid == self._nrm('xstyn'):
            ov = nl[colname]
            nv = JOElement(ov)
            if not nv.isvalid():
                _a_log(logs, ov, None, 'Sty# is invalid', self.LEVEL_CRITICAL)
            else:
                nv = nv.value
                if nv != ov:
                    _a_log(logs, ov, nv, None)
        elif catid == self._nrm('sn'):
            ov = nl[colname]
            nv = formatsn(nl[colname], 0, True) or (NA, )
            if nv:
                nv = ";".join(sorted(nv))
            if nv != ov:
                _a_log(logs, ov, nv, "SN# formatted")
        elif catid == self._nrm('keywords'):
            ov = nl[colname]
            nv = (x for x in trimu(ov).split(";") if x)
            nv = ";".join(sorted(tuple(set(nv))))
            if nv != ov or _nsr(nv):
                _a_log(logs, ov, nv, None, self.LEVEL_ADVICE)
        elif catid == self._nrm('text'):
            ov = nl[colname]
            if _nsr(ov):
                _a_log(logs, None, ov, None, self.LEVEL_ADVICE)

class _NRInvoker(NrlsInvoker):
    '''
    class help to validate the form's field one by one
    mainly make use of the sub classes of BaseNrl to complete the task
    '''

    def __init__(self, hdlr=None, upd_src=True, hl_level=BaseNrl.LEVEL_MINOR):
        """
        Args:
            hdlr: an Handler instance help to access the excel if there is
        """
        self._hdlr, self._upd_src, self._hl_level = hdlr, upd_src, hl_level
        self._hints = _HintsHdlr(hdlr.book)
        super().__init__(self._init_nrls())

    def _hl(self, rng, level, rmks):
        if not rng:
            return
        ci, api = self._get_hl_color(level), None
        if ci > 0:
            api = rng.api
            api.interior.colorindex = ci
        if rmks:
            if not api:
                api = rng.api
            api.ClearComments()
            api.AddComment()
            api = api.Comment
            api.Text(Text=rmks)

    @staticmethod
    def _get_hl_color(level):
        if level < BaseNrl.LEVEL_MINOR:
            return 19   # light yellow
        if level < BaseNrl.LEVEL_CRITICAL:
            return 6    # yellow
        return 3    # red

    def _init_nrls(self):
        nrls = {}
        # maybe from the config file
        tu = BaseNrl.get_base_nrm('tu')
        _nrm = self._hdlr.normalize if self._hdlr else _Utilz.nrm
        kwds = {'nrm': _nrm, 'hints': self._hints.get_hints}
        for n in 'styno qclevel parent craft'.split():
            nrls[_nrm(n)] = tu
        mp = {
            'docno': JENrl,
            'size': SizeNrl,
            'metal': MetalNrl,
            'finishing': FinishingNrl,
            'parts': PartsNrl,
            'stone': StoneNrl
        }
        for n, cn in mp.items():
            nrls[_nrm(n)] = cn(name=n, **kwds)
        tu = FeatureNrl(name='feature', **kwds)
        for n in 'feature feature1'.split():
            nrls[_nrm(n)] = tu
        n = _nrm('netwgt')
        nrls[n] = NWgtNrl(name=n, **kwds)
        n = _nrm('description')
        nrls[n] = (BaseNrl.get_base_nrm('tu'), DescNrl(name=n, **kwds), )
        return nrls

    def normalize(self, mp):
        '''
        normalize the result map
        Args:
            mp({string, BaseNrl}): a map generated by prdspec.Handler.read() method
            upd_src: update the source excel for those normalized
            hl_level: hight light those changes has high value than this
        '''
        mp, logs = super().normalize(mp)
        if self._hdlr and self._upd_src:
            self._hdlr.write(mp)
            if logs:
                self._hl_xls(logs)
        return mp, logs

    def _hl_xls(self, logs):
        '''
        high-light the given invalid items(a collection of self._nl_vld)
        '''
        for nl in logs:
            if not nl.name:
                continue
            if nl.level >= self._hl_level:
                nd = self._hdlr.get(nl.name)
                rng = nd.get(nl.row, nl.colname) if nl.row else nd.get()
                self._hl(rng, nl.level, nl.remarks)

class _HintsHdlr(object):
    ''' read and return hints
    '''

    def __init__(self, wb):
        self._wb, self._meta_mp, self._nrm = wb, None, None

    @property
    def nrm(self):
        return self._nrm or _Utilz.nrm

    def _read_hint_defs(self):
        '''
        read meta data except the field mappings from the meta sheet
        '''
        if self._meta_mp:
            return
        self._meta_mp = {}
        nrm = self.nrm
        for sn, mp in config.get('prodspec.meta_tables').items():
            sht = self._wb.sheets(sn)
            for meta_type, tblname in mp.items():
                addr = apirange(sht.api.listObjects(tblname).Range)
                self._meta_mp[nrm(meta_type)] = {nrm(nl['name']): nl for nl in NamedLists(addr.value)}
        # in the case of stone name, build an abbr -> name hints map
        mp = {nrm(nl['hp.abbr']): nrm(nl['name']) for nl in self._meta_mp[nrm('stone.name')].values()}
        self._meta_mp[nrm('stone.abbr')] = mp

    def get_hints(self, meta_type, cand):
        '''
        find a best match meta data item based on the cand string.
        Args:
            meta_type(string): sth. like 'producttype'/'finishingmethod', which was defined in conf.json under key 'prodspec.meta_tables'
            cand(string): the candidate string that need matching
        Returns:
            A tuple as:
                [0]: 1 if meta_type is defined and match is perfect
                     2 if meta_type is defined and match is not perfect
                     0 if meta_type is not defined
                [1]: the best-match item or None when [0] is True or None
        '''
        if not cand:
            return 0, None
        if not self._meta_mp:
            self._read_hint_defs()
        meta_type = self.nrm(meta_type)
        mp = self._meta_mp[meta_type]
        if not mp:
            return 0, None
        ncand = self.nrm(cand)
        if meta_type == self.nrm('stone.name'):
            ncand = self._meta_mp[self.nrm('stone.abbr')].get(ncand, ncand)
        nl = mp.get(ncand)
        if nl:
            return 1, self.nrm(nl['name']) # TODO::return based on meta_type
        sm = SequenceMatcher(None, ncand, None)
        ln = len(cand)
        rts, min_mt = [], 2 * max(2, ln * 0.6) # at least 2 or 1/2 match
        for s in mp:
            sm.set_seq2(s)
            rt = sm.ratio()
            if rt >= min_mt / (ln + len(s)):
                rts.append((rt, s,))
        if not rts:
            return 2, None
        rts = sorted(rts, key=lambda x: x[0])
        return 2, rts[-1][1]
