'''
#! coding=utf-8
@Author:        zmFeng
@Created at:    2019-07-12
@Last Modified: 2019-07-12 2:41:19 pm
@Modified by:   zmFeng

Data normalization interface and some common normalizers like trim/trimu/triml
'''

from abc import ABC, abstractmethod

from ..miscs import NA, NamedList, triml, trimu, NamedLists


class BaseNrl(ABC):
    '''
    base normalizer. Sub-class must override

    :func: normalize to perform actual normalizing
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
    _base_nrls = {}

    def __init__(self, **kwds):
        '''
        Args:

        level=LEVEL_MINOR:  the default log level of this normalizer
        name=NA:        name of this normalizer

        '''
        self._default_level = kwds.get('level', self.LEVEL_MINOR)
        self._name = kwds.get('name', NA)

    @abstractmethod
    def normalize(self, pmap, name):
        '''
        fetch pmap[name] and normalize it if necessary. When nrl is needed, the result will be directly write th pmap[name] and change log will be created.

        Args:

        pmap: the property map that contains 'name'

        name: the property name need to be normalized

        Returns:
            None if value is valid, or a tuple of list which created by BaseNrl.nl.data. Using tuple of list for saving memory.

            BaseNrl.nl contains below columns:

                name, row, colname, oldvalue, newvalue, level, remarks
            where level should be one of `BaseNrl.LEVEL_XXX`
        '''

    def _new_log(self, oldval, newval, **kwds):
        '''
        just return the array, returning the NamedList object will lose the array
        '''
        nl = self.nl
        nl.newdata()
        lvl = kwds.get('level')
        if lvl is None:
            lvl = self._default_level
        nl.oldvalue, nl.newvalue, nl.level = oldval, newval, lvl
        if kwds:
            cm = set(nl.colnames)
        if 'name' not in kwds:
            kwds['name'] = self._name
        for n, v in kwds.items():
            if n not in cm:
                continue
            if n == 'row':
                nl[n] = v + 1  # the caller is zero based, so + 1
            else:
                nl[n] = v
        return nl.data

    @classmethod
    def get_base_nrm(cls, name):
        '''
        return the often used normalizer
        Args:
            name: only support tu/tr/tl
        '''
        # use lower-case internally
        n = triml(name)
        nrl = cls._base_nrls.get(n)
        if nrl:
            return nrl
        if n == 'tu':
            nrl = TUNrl(name="trim and to_upper")
        elif n == 'tr':
            nrl = TrimNrl(name="trim")
        elif n == 'tl':
            nrl = TLNrl(name="trim and to_lower")
        if nrl:
            cls._base_nrls[n] = nrl
        return nrl


class SmpFmtNrl(BaseNrl):
    '''
    perform normalizer that can be executed by just calling one method with one argument
    '''

    def __init__(self, fmt_func, **kwds):
        if 'level' not in kwds:
            kwds['level'] = self.LEVEL_RPLONLY
        self._fmt_func = fmt_func
        super().__init__(**kwds)

    def normalize(self, pmap, name):
        if name not in pmap:
            return None
        old = pmap[name]
        nv = self._fmt_func(old) if old else None
        if nv != old:
            pmap[name] = nv
            return (self._new_log(old, nv, name=name, level=self._default_level, remarks=self._name),)
        return super().normalize(pmap, name)


class TrimNrl(SmpFmtNrl):
    '''
    trim only
    '''
    _tm = lambda x: x.strip()

    def __init__(self, **kwds):
        super().__init__(TrimNrl._tm, **kwds)


class TUNrl(SmpFmtNrl):
    '''
    trim and upper
    '''

    def __init__(self, **kwds):
        super().__init__(trimu, **kwds)


class TLNrl(SmpFmtNrl):
    ''' trim and lower
    '''
    def __init__(self, **kwds):
        super().__init__(triml, **kwds)

class TBaseNrl(BaseNrl):
    '''
    base class for table-based normalizer
    Descendant classes must override

        :func: _get_nrls and
        :func: _nrl_row

    '''

    def __init__(self, **kwds):
        super().__init__(**kwds)
        self._nrl_mp = None

    @abstractmethod
    def _nrl_row(self, name, row, nl, logs):
        ''' normal one table item
        '''

    @abstractmethod
    def _get_nrls(self):
        ''' return the nrls used to normalize column items
        '''

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
        nrlmp = self._get_nrls()
        for row, nl in enumerate(its):
            if nrlmp:
                for cn, nrls in nrlmp.items():
                    if isinstance(nrls, BaseNrl):
                        nrls = (nrls, )
                    for nrl in nrls:
                        self._normalize_item(nrl, nl, cn, logs=logs, name=name, row=row)
            self._nrl_row(name, row, nl, logs)
        return logs or None

class NrlsInvoker(object):
    ''' invoke a map of BaseNrls and return the logs
    '''

    def __init__(self, nrls):
        self._nrls = nrls

    def normalize(self, mp):
        ''' execute all the normalizes and merge the logs if necessary
        Args:

        mp: the data to be normalized, Map(name(String), value or List(NamedList))
        '''
        logs = None
        if not self._nrls:
            return None
        logs = []
        for pp_name, vdrs in self._nrls.items():
            if not vdrs:
                continue
            if isinstance(vdrs, BaseNrl):
                vdrs = (vdrs, )
            for vdr in vdrs:
                logx = vdr.normalize(mp, pp_name)
                if logx:
                    logs.extend(logx)
        if logs:
            logs.insert(0, tuple(BaseNrl.nl.colnames))
            logs = self._merge_logs(NamedLists(logs))
        return mp, logs

    def _merge_logs(self, logs):
        ''' merge the logs if necessary. logs with the same name/row/colname will be merged to one iten
        Args:

            logs: tuple(NamedList) generaeted by sub-classes of BaseNRL

        Returns:
            tuple(NamedList), see @BaseNrl for colname definitions
        '''
        if not logs:
            return None
        _mk = lambda x: (x.name, x.row, x.colname)
        mp = {}
        for log in logs:
            key = _mk(log)
            if key not in mp:
                mp[key] = log.clone(False)
            else:
                log0 = mp[key]
                log0.newvalue = log.newvalue
                lvls = log.level, log0.level
                if lvls[0] > lvls[1]:
                    log0.level = lvls[0]
                rmks = log0.remarks, log.remarks
                rmks = [x for x in rmks if x]
                if rmks:
                    log0.remarks = ";".join(rmks)
        return tuple(mp.values())
