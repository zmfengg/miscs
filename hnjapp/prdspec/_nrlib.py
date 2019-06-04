#! coding=utf-8
'''
* @Author: zmFeng
* @Date: 2019-05-29 09:53:51
* @Last Modified by:   zmFeng
* @Last Modified time: 2019-05-29 09:53:51
local normalizers for a product specification form
'''

from abc import ABC, abstractmethod

from hnjapp.c1rdrs import _fmtpkno
from hnjapp.common import config
from hnjapp.dbsvcs import formatsn
from hnjcore import JOElement
from utilz import NamedList, karatsvc, stsizefmt, triml, trimu, na


class _Utilz(object):
    _nrm = _cn_seqid = None
    _base_nrls = {}

    @classmethod
    def nrm(cls, s):
        """ normalize a string
        Args:
            s: the string to normalize
        """
        if not cls._nrm:
            c = config.get("prdspec.case") or 'UPPER'
            cls._nrm = triml if  c == 'LOWER' else trimu
        return cls._nrm(s)

    @classmethod
    def cn_seqid(cls):
        """ colname of the sequence id
        """
        if not cls._cn_seqid:
            cls._cn_seqid = cls.nrm('seqid')
        return cls._cn_seqid

    @classmethod
    def get_base_nrm(cls, name):
        '''
        return the often used normalizer
        Args:
            name: only support tu/tr
        '''
        n = cls.nrm(name)
        nrl = cls._base_nrls.get(n)
        if nrl:
            return nrl
        if n == cls.nrm('tu'):
            nrl = TUNrl(name="base trim and upper handler")
        elif n == cls.nrm('tr'):
            nrl = TrimNrl(name="base trim and upper handler")
        if nrl:
            cls._base_nrls[n] = nrl
        return nrl

_nrm = _Utilz.nrm

class BaseNrl(ABC):
    '''
    base normalizer
    '''
    # level:
    # 0 - 99: very small case, can be ignored
    # 100 - 199: advices only, won't affect original value
    # 200 - 255: critical error, fail the program
    nl = NamedList('name row colname oldvalue newvalue level remarks'.split())
    LEVEL_RPLONLY = 0   #only need to do replacement
    LEVEL_MINOR = 50
    LEVEL_ADVICE = 100
    LEVEL_CRITICAL = 200

    def __init__(self, **kwds):
        self._default_level = kwds.get('level', self.LEVEL_MINOR)
        self._hdlr = None
        self._name = kwds.get('name', na)

    @property
    def hdlr(self):
        '''
        handler for getting/setting data with excel
        '''
        return self._hdlr

    @hdlr.setter
    def hdlr(self, hdlr):
        self._hdlr = hdlr

    @abstractmethod
    def normalize(self, pmap, name):
        '''
        return None if value is valid, or a tuple of _nl items
        @param pmap: the parent map that contains 'name'
        @param name: the item that need to be validate
        '''

    def _new_log(self, oldval, newval, **kwds):
        '''
        just return the array, returning the NamedList object will lose the array
        '''
        nl = self.nl
        nl.newdata()
        nl.oldvalue, nl.newvalue, nl.level = oldval, newval, self._default_level
        if kwds:
            cm = set(nl.colnames)
        for n, v in kwds.items():
            if n not in cm:
                continue
            if n == 'row':
                nl[n] = v + 1  # the caller is zero based, so + 1
            else:
                nl[n] = v
        return nl.data

class TrimNrl(BaseNrl):
    '''
    perform trim only action
    '''

    def __init__(self, *args, **kwds):
        if 'level' not in kwds:
            kwds['level'] = self.LEVEL_RPLONLY
        super().__init__(*args, **kwds)

    def normalize(self, pmap, name):
        old = pmap[name]
        nv = old.strip()
        if nv != old:
            pmap[name] = nv
            return (self._new_log(old, nv, name=name,
                remarks='(%s) trimmed to (%s)' % (old, nv)),)
        return super().normalize(pmap, name)


class TUNrl(BaseNrl):
    '''
    perform trim and upper actions
    '''

    def __init__(self, *args, **kwds):
        if 'level' not in kwds:
            kwds['level'] = self.LEVEL_RPLONLY
        super().__init__(*args, **kwds)

    def normalize(self, pmap, name):
        old = pmap[name]
        nv = trimu(old)
        if nv != old:
            pmap[name] = nv
            return (self._new_log(old, nv, name=name),)
        return super().normalize(pmap, name)


class JENrl(BaseNrl):
    """ the JOElement like normalizer
    """
    def __init__(self, *args, **kwds):
        if 'level' not in kwds:
            kwds['level'] = self.LEVEL_RPLONLY
        super().__init__(*args, **kwds)

    def normalize(self, pmap, name):
        ov = pmap[name]
        nv = JOElement(ov)
        if nv.isvalid():
            nv = nv.name
            if nv != ov:
                pmap[name] = nv
                return (self._new_log(ov, nv, name=name),)
        return super().normalize(pmap, name)


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
        k0 = pmap[name]
        nv = rmk = None
        kt = karatsvc.getkarat(k0)
        if kt and kt.name != k0:
            nv, lvl = kt.name, self.LEVEL_ADVICE
            rmk = 'Karat(%s) => (%s)' % (k0, nv)
        elif self._strict and not kt:
            rmk, lvl = 'Invalid Karat(%s)' % k0, self.LEVEL_CRITICAL
        return  (self._new_log(k0, nv, name=name, remarks=rmk, level=lvl), ) if rmk else None

class SizeNrl(BaseNrl):
    '''
    formatting stone size or object dimension
    '''
    def normalize(self, pmap, name):
        ov = trimu(str(pmap[name]))
        if not ov:
            return None
        sfx = ov.find('SZ')
        if sfx > 0:
            ov, sfx = ov[:sfx], ov[sfx:]
        else:
            sfx = ''
        nv = stsizefmt(str(ov), True) + sfx
        return  (self._new_log(ov, nv, name=name, remarks=None, level=self.LEVEL_RPLONLY),)

class TBaseNrl(BaseNrl):
    '''
    base class for table-based normalizer
    '''

    def __init__(self, *args, **kwds):
        super().__init__(*args, **kwds)
        self._nrl_mp = None
        self._reg_nrls()

    def _reg_nrls(self):
        pass

    def _nrl_one(self, name, row, nl, logs):
        pass

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

    def _normalize_item(self, nrl, nl, cn, **kwds):
        '''
        normal just one item inside a row, used for those table property that can
        be normalized by other Normalizer
        Args:
            kwds: must contain: name,row,logs
        '''
        var = nrl.normalize(nl, cn)
        if not var:
            return
        var = self.nl.setdata(var[0])
        nl[cn] = var.newvalue
        var.name, var.row, var.colname = kwds.get('name'), kwds.get('row') + 1, cn
        if 'logs' in kwds:
            kwds['logs'].append(var.data)

    def normalize(self, pmap, name):
        its = pmap.get(name)
        if not its:
            return super().normalize(pmap, name)
        logs = []
        for row, nl in enumerate(its):
            if self._nrl_mp:
                for cn, nrls in self._nrl_mp.items():
                    if isinstance(nrls, BaseNrl):
                        nrls = (nrls, )
                    for nrl in nrls:
                        self._normalize_item(nrl, nl, cn, logs=logs, name=name, row=row)
            self._nrl_one(name, row, nl, logs)
        return logs or None


class MetalNrl(TBaseNrl):
    '''
    Metal table normalizer
    '''

    def _reg_nrls(self):
        self._nrl_mp = {
            'remarks': _Utilz.get_base_nrm('tu')
        }

    def _nrl_one(self, name, row, nl, logs):
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


class FinishingNrl(TBaseNrl):
    '''
    normalizer for finishing
    '''

    def _reg_nrls(self):
        tu = _Utilz.get_base_nrm('tu')
        self._nrl_mp = {
            'remarks': _Utilz.get_base_nrm('tr'),
            'method': tu,
            'spec': tu
        }

    def _nrl_one(self, name, row, nl, logs):
        if not self.hdlr:
            return
        for colname in 'spec method'.split():
            mt, cand = self.hdlr.get_hints(_nrm('finishing.%s' % colname), nl[colname])
            if mt == 2 and cand:
                cand = trimu(cand)
                self._append_log(logs, nl[colname], cand, '(%s) => (%s)' % (nl[colname], cand), name=name, row=row, colname=colname, nl=nl, level=self.LEVEL_ADVICE)

class PartsNrl(TBaseNrl):
    '''
    normalizer for Parts
    '''

    def _reg_nrls(self):
        tu = _Utilz.get_base_nrm('tu')
        self._nrl_mp = {'remarks': _Utilz.get_base_nrm('tr'), 'type': tu, 'matid': tu, 'karat': (tu, KaratNrl(False))}

class StoneNrl(TBaseNrl):
    '''
    normalizer for stone
    .minor BL as:
        .Too many stone can not be main stone(ER=2, other=1)
        .There can be only one main stone item
        .Stone Name correction
        .Size to weight
    '''

    def _reg_nrls(self):
        self._nrl_mp = {'size': SizeNrl()}

    def _nrl_one(self, name, row, nl, logs):
        colname = 'matid'
        ov = nl[colname]
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
            _log(ov, 1, 'Invalid qty(%r), set to %d' % (ov, nv))
        if uw * wgt == 0:
            if uw + wgt == 0:
                for colname in ('unitwgt', 'wgt'):
                    _log(ov, None, self.LEVEL_CRITICAL, '(%s) not nullable' % colname)
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

        if self.hdlr:
            for colname in 'shape name'.split():
                ov = nl[colname]
                mt, nv = self.hdlr.get_hints(_nrm('stone.%s' % colname), ov)
                if mt > 0 and nv and nv != ov:
                    nv = trimu(nv)
                    self._append_log(logs, ov, nv, '%s -> %s' % (nl[colname], nv), name=name, row=row, colname=colname, nl=nl, level=self.LEVEL_ADVICE)

    def normalize(self, pmap, name):
        logs = super().normalize(pmap, name)
        if not logs:
            logs = []
        stones, ms = pmap[name], []
        for row, nl in enumerate(stones):
            valid = nl.qty == (1 if pmap[_nrm('type')] != "EARRING" else 2) and stsizefmt(str(nl.size), False) >= "0300"
            ov = nv = nl[_nrm('main')]
            if valid ^ (ov == 'Y'):
                nv = 'Y' if valid else 'N'
                self._append_log(logs, ov, nv, 'should be %s' % ('MAIN STONE' if valid else 'SIDE STONE'), nl=nl, name=name, row=row, colname='main', level=self.LEVEL_ADVICE)
            if nv == 'Y':
                ms.append(row)
        if len(ms) > 1:
            for row in ms:
                self._append_log(logs, None, 'Y', 'MULTI-MAIN is not allowed', name=name, nl=stones[row], row=row, colname='main', level=self.LEVEL_CRITICAL)
        return logs if logs else None

class FeatureNrl(TBaseNrl):
    """ feature based normalizer
    """
    def _nrl_one(self, name, row, nl, logs):
        colname = 'value'
        def _a_log(logs, ov, nv, rmk, lvl=self.LEVEL_RPLONLY):
            self._append_log(logs, ov, nv, rmk, nl=nl, name=name, row=row, colname=colname, level=lvl)
        catid = _nrm(nl['catid'])
        if catid == _nrm('xstyn'):
            ov = nl[colname]
            nv = JOElement(ov)
            if not nv.isvalid():
                _a_log(logs, ov, None, 'Sty# is invalid', self.LEVEL_CRITICAL)
            else:
                nv = nv.value
                if nv != ov:
                    _a_log(logs, ov, nv, None)
        elif catid == _nrm('sn'):
            ov = nl[colname]
            nv = formatsn(nl[colname], 0)
            if nv != ov:
                _a_log(logs, ov, nv, "SN# formatted")
        elif catid == _nrm('keyword'):
            ov = nl[colname]
            nv = ";".join(sorted(trimu(ov).split(";")))
            if nv != ov:
                _a_log(logs, ov, nv, None, self.LEVEL_ADVICE)
