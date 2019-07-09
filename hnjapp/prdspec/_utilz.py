'''
#! coding=utf-8
@Author:        zmFeng
@Created at:    2019-06-17
@Last Modified: 2019-06-17 11:00:54 am
@Modified by:   zmFeng
utilz for this package
'''
import re
from difflib import SequenceMatcher
from numbers import Number

from hnjapp.common import config
from utilz import NamedList, karatsvc, triml, trimu


class _Utilz(object):
    _nrm = _cn_seqid = None
    _base_nrls = {}
    _pw_mp = {}

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
    def get_parts_wgt(cls, pn, karat="925", alias=False):
        ''' return the weight of given parts name in karat
        Args:
            pn: name of the parts
            karat: target karat
        '''
        if not cls._pw_mp:
            cls._pw_mp = {x[0]: config.get(x[1]) for x in (('_sn_cvt', 'snno.translation'), ('_self', 'parts.wgt.self'), ('_sc', 'parts.wgt.sc'))}
        if isinstance(karat, Number):
            karat = '%d' % karat
        pn = trimu(pn)
        mp = cls._pw_mp['_self']
        kt0, w = mp.get('_df_karat'), mp.get(pn)
        if not w:
            w = mp.get(cls._pw_mp['_sn_cvt'].get(pn))
            if alias and w:
                pn = cls._pw_mp['_sn_cvt'].get(pn)
        if not w:
            # do a stupid diff cmp, get the matches one
            mp = cls._pw_mp['_sc']
            w, r0, cand = mp.get(pn), 0, None
            if not w:
                sm = SequenceMatcher(a=pn)
                for k, v in mp.items():
                    sm.set_seq2(k)
                    r = sm.ratio()
                    # print('%f = %s -> %s' % (r, pn, k))
                    if r > r0:
                        r0 = r
                        cand = (r, k, v)
                if not (cand and cand[0] > 0.6):
                    return None
                pn, w = cand[1:]
        if not isinstance(w, Number):
            w, kt0 = w.split("@")
            w = float(w)
        if kt0 != karat:
            w = karatsvc.convert(kt0, w, karat)
        return w if not alias else (w, pn)

    @staticmethod
    def get_lksz(rmk):
        ''' extract the locket size from a description
        '''
        idx = rmk.find('相')
        if not idx:
            return None
        sz = re.findall(config.get('pattern.locket.size'), rmk)
        if not sz:
            return None
        sz = "X".join(sz)
        shp, t = rmk[:idx].strip()[0], None
        for s, d in config.get('locket.shapes').items():
            if d[0] == shp:
                t = s.strip()
                break
        return '%s:%sMM' % (t or '_HL', sz)

    @staticmethod
    def get_text(rmk):
        ''' extract the text stated inside the remark
        '''
        if not rmk:
            return None
        m, idx = rmk.find('字'), -1
        while m >= 0:
            if rmk[m + 1] == '印':
                m = rmk.find('字', m + 2)
            else:
                idx = m
                break
        if idx <= 0:
            return None
        # because the user might even lazy to type a space, so find the first ascii
        ln = len(rmk)
        for m in range(1, 10):
            if idx + m >= ln:
                return None
            if ord(rmk[idx + m]) < 127:
                idx += m
                break
        ptn = re.compile(r'[\w\(\),.]*')
        wds = [s for s in ptn.findall(rmk[idx:]) if s]
        pts = [[], []]
        for idx, s in enumerate(wds):
            if s.find('(') >= 0:
                pts[0].append(idx)
            elif s.find(')') >= 0:
                pts[1].append(idx)
        if any(pts):
            cand = []
            def _pc(x):
                if not x:
                    return
                if x.find('(') >= 0 or x.find(')') >= 0 or x.find('17D') >= 0:
                    return
                cand.append(x)
            if pts[0]:
                _pc(_Utilz._raw_get_text(wds[:pts[0][-1]]))
            if pts[1]:
                _pc(_Utilz._raw_get_text(wds[pts[1][-1]+1:]))
            if not cand:
                return None
            if len(cand) == 1:
                return cand[0]
            return sorted(cand, key=len)[-1]
        return _Utilz._raw_get_text(wds)

    @staticmethod
    def _raw_get_text(wds):
        lst = []
        for s in wds:
            idx = s.find(',')
            if idx > 0:
                lst.extend(x + ',' for x in s.split(',') if x)
            else:
                lst.append(s)
        s = -1
        for idx, s0 in enumerate(lst):
            if not s0:
                continue
            if ord(s0[0]) > 127:
                if s >= 0:
                    break
            else:
                if s < 0:
                    s = idx
        if s < 0:
            return None
        if idx == len(lst) - 1:
            idx = None
        s = ' '.join((x for x in lst[s:idx] if x))
        return s[:-1] if s[-1] == ',' else s

_nrm = _Utilz.nrm

class _Tbl_Optr(object):
    '''  convenient dict's table operations
    '''
    def __init__(self, hdlr, mp=None):
        self._hdlr, self._mp = hdlr, mp
        self._nlmp = {}

    @property
    def mp(self):
        ''' get the map to operator on
        '''
        return self._mp

    @mp.setter
    def mp(self, mp):
        ''' set the map to operator on
        '''
        self._mp = mp

    def append(self, tblname, append=True):
        ''' append one record to given tblname, when the table is blank
        a record will be append no matter append=True or not
        Args:
            tblname: the table name to append
            apppend: when table found, append record or not
        Returns:
            A NamedList object pointing to the record
        '''
        tn = _nrm(tblname)
        nl = self._nlmp.get(tn)
        if not nl:
            self._nlmp[tn] = nl = NamedList(self._hdlr._get_raw_colnames(tn))
        lst = self._mp.get(tn)
        if not lst:
            nl = nl.clone()
            self._mp[tn] = lst = [nl]
            append = False
        if append:
            nl = nl.clone()
            lst.append(nl)
        return nl

    def set_item(self, tblname, kcolname, kid, colname, val, unique=True):
        """ get or new item from feature by given name
        Args:
            tblname: the tblname
            kcolname: the key column name
            kid: the key id
            colname: the colname you need to set value to
            val: the value for the column to set
            unique: True while the key column is unique
        """
        tblname, kid = _nrm(tblname), trimu(kid)
         # when there is no tn item, below statement will create it
        self.append(tblname, append=False)
        flag, lst = False, self._mp[tblname]
        for nl in lst:
            x = nl[kcolname]
            flag = x is None or unique and x == kid
            if flag:
                if x is None:
                    nl[kcolname] = kid
                break
        if not flag:
            nl = lst[0].clone()
            lst.append(nl)
            nl[kcolname] = kid
        s0 = nl[colname] or ""
        if s0:
            s0 += ";"
        nl[colname] = s0 + val
        return nl

    def add_parts(self, pname, ptype, qty=1, matid=None):
        ''' add parts
        '''
        if not matid:
            matid = config.get("snno.translation").get(pname, pname)
        nl = self.set_item('parts', 'type', ptype, 'matid', matid, False)
        nl['remarks'], nl['qty'], nl['wgt'] = pname, qty, 0.001
        return nl

    def set_metal(self, kt, wgt):
        ''' add or set metal
        Args:
            kt: the karat name, not object
            wgt: karat weight
        '''
        nl = self.append('metal')
        wgt0 = nl.wgt or 0
        nl.karat, nl.wgt = kt, wgt0 + wgt
        return nl

    def set_feature(self, catid, val):
        ''' add or set value to feature of given catid
        '''
        return self.set_item('feature', 'catid', catid, 'value', val)

    def add_finishing(self, method, spec, rmk):
        ''' add finishing
        '''
        nl = self.set_item('finishing', 'method', method, 'spec', spec, False)
        if rmk:
            nl['remarks'] = rmk
        return nl
