#! coding=utf-8
'''
* @Author: zmFeng
* @Date: 2019-05-21 14:00:29
* @Last Modified by:   zmFeng
* @Last Modified time: 2019-05-21 14:00:29
services help to handle the product specification sheet
'''

from xlwings.constants import BordersIndex, LineStyle

from utilz import NamedLists, triml
from utilz.xwu import addr2rc, apirange, nextcell, nextrc, rc2addr
from ._nrls import Normalizer

_nrm = triml
_cn_seqid = _nrm('seqid')


class Controller(object):
    '''
    read data out from requested workbook
    '''

    def __init__(self, wb):
        self._wb = wb
        # hold the often used constans
        self._nmps = self._cnstmp = self._sht_tar = self._full_inited = None
        self._init_consts()
        self._read_mapping()

    def _init_consts(self):
        self._cnstmp = {
            '_dir_ud': ('down', 'up'),
            '_dir_ul': ('up', 'left'),
            '_dir_2_bdr': {
                'left': BordersIndex.xlEdgeLeft,
                'right': BordersIndex.xlEdgeRight,
                'down': BordersIndex.xlEdgeBottom,
                'up': BordersIndex.xlEdgeTop
            },
            '_d2d': {
                _nrm('d'): 'down',
                _nrm('u'): 'up',
                _nrm('l'): 'left',
                _nrm('r'): 'right'
            }
        }

    def _read_mapping(self):
        '''
        read the field mapping from the requested workbook
        '''
        sht = self._wb.sheets('metadata')
        addr = apirange(sht.api.listObjects('m_fmp').Range)
        n2addr, addr2n, mp = {}, {}, None
        # don't use NamedRanges(addr, newinst=False) because the shape is obvisous
        # and NamedRanges is slow
        # for nl in NamedRanges(addr, newinst=False):
        for nl in NamedLists(addr.value, newinst=False):
            name, addr = _nrm(nl.name), nl.firstaddress
            if not self._sht_tar:
                self._sht_tar = self._wb.sheets(
                    addr[addr.find(']') + 1:addr.find('!')])
            idx, addr = name.find("."), addr[addr.find('!') + 1:]
            mp = n2addr.setdefault(name[:idx], {}) if idx >= 0 else n2addr
            if idx >= 0:
                name = name[idx + 1:]
            else:
                addr2n[_nrm(addr)] = name
            di = self._d2d(nl.direction)
            mp[name] = (addr, di)
            #also put the index to name reserve lookup table
            if idx > 0:
                # get the adjusted directional from mp
                idx = addr2rc(addr)[0][1 if di in self._cnstmp['_dir_ud'] else 0]
                mp.setdefault('_i2n', {})[idx] = name
        self._nmps = {x[0]: x[1] for x in
            zip('n2addr addr2n reg2n'.split(), (n2addr, addr2n, {}))}
        self._nmps['_table_key_cols'] = {nl.tblname: nl.cols for nl in NamedLists(apirange(sht.api.listObjects('m_fkc').Range).value)}

    def _calc_n2addr(self, name):
        '''
        do some range calculation based on the mappings
        '''
        mp = self._nmps['n2addr'][name]
        if not isinstance(mp, dict) or '_max_cnt' in mp:
            return
        _rg = self._sht_tar.range
        ud = self._cnstmp['_dir_ud']
        cn = _cn_seqid if _cn_seqid in mp else next(self._get_cns(mp))
        var, di = mp[cn]
        rng = self._sht_tar.range(var)
        if cn == _cn_seqid:
            # don't use expand, quite slow
            # rng = _rg(rng, rng.end(di))# if di in self._cnstmp('_dir_ul') else rng.expand(di)
            rng = _rg(rng.address + ':' + rng.end(di).address)
        else:
            # save time, 3 steps forward because in this case, won't be
            # less than 4 rows
            cn = nextcell(rng, di, detect_merge=False, steps=4)
            bi, his = self._cnstmp['_dir_2_bdr'][di], []
            while cn.api.Borders(bi).LineStyle != LineStyle.xlLineStyleNone:
                cn = nextcell(cn, di, detect_merge=False)
                his.append(cn)
            rng = _rg(rng.address + ':' + his[-2].address)
        mc = rng.rows.count if di in ud else rng.columns.count
        mp['_max_cnt'], mp['_org'], mp['_dir'] = mc, addr2rc(var)[0], di
        mp['_region'] = var = self._calc_n2addr_table(mp)
        self._nmps['reg2n'][name] = addr2rc(var)

    @staticmethod
    def _end(rng, di):
        '''
        excel's end will stop when it reach a merged cell, but in my case, I need to continue if the merged cell contains sth.
        '''
        rngs, hc = [rng, ], 0
        while True:
            rngs.append(rngs[-1].end(di))
            if not rngs[-1].api.mergecells:
                hc += 1
                if hc > 1:
                    break
            else:
                hc = 0
        return rngs[-1 if nextcell(rngs[-1], di, -1, False).value else -2]

    def _calc_n2addr_table(self, mp):
        ud = self._cnstmp['_dir_ud']
        di, mc = [mp[x] for x in ('_dir', '_max_cnt')]
        cols = [mp[x][0] for x in self._get_cns(mp)]
        idx = 1 if di in ud else 0
        cols = sorted([addr2rc(x)[0][idx] for x in cols])
        # for speed reason, don't use range(cell, cell), it's terribly slow(time spent: 0.338 -> 0.018)
        if idx:
            di, addr = 1 if di == 'down' else -1, mp['_org'][0]
            addr = rc2addr((addr, cols[0]), (addr + di * (mc - 1), cols[-1]))
        else:
            di, addr = 1 if di == 'right' else -1, mp['_org'][1]
            addr = rc2addr((cols[0], addr), (cols[-1], addr + di * (mc - 1)))
        return addr

    def _full_init(self):
        # only get the first cell to avoid multi-cell case
        if not self._full_inited:
            for x in [x[0] for x in self._nmps['n2addr'].items() if isinstance(x[1], dict) and '_max_cnt' not in  x[1]]:
                self._calc_n2addr(x)
            self._full_inited = True

    def _d2d(self, d0):
        '''
        translate the direction in m_fmp to the one for xlwings
        '''
        return self._cnstmp['_d2d'][_nrm('R' if not d0 else d0[0])]

    def get(self, name, get_hotpoint=False):
        '''
        return a node related to given name(or address), for example, get('author')
        should return a Node who's get('') will be Range("$L$2")
        '''
        if get_hotpoint:
            return self._get_hot_point(name)
        mp = self._nmps['n2addr']
        nl = _nrm(name).split(".")[0]
        if nl in mp:
            self._calc_n2addr(name)
            return Node(self._sht_tar, mp[nl])
        return None

    def get_hotpoint(self, addr):
        '''
        convenience way for self.get(name, True)
        '''
        return self._get_hot_point(addr)

    def _get_hot_point(self, addr):
        '''
        return hotpoint of given address
        '''
        mp, ud = self._nmps['addr2n'], self._cnstmp['_dir_ud']
        di = mp.get(_nrm(addr))
        if di:
            return di
        self._full_init()
        pt = addr2rc(addr)[0]
        for name, pts in self._nmps['reg2n'].items():
            pt0, pt1 = pts
            # TODO:: find the exit point
            if all([pt0[x] <= pt[x] <= pt1[x] for x in range(2)]):
                mp = self._nmps['n2addr'][name]
                di, pt0 = 0 if mp['_dir'] in ud else 1, mp['_org']
                row = abs(pt0[di] - pt[di]) + 1
                return name, row, mp['_i2n'].get(pt[0 if di else 1])
        return None

    @staticmethod
    def _get_cns(mp):
        return (x for x in mp if x.find('_') < 0)

    def read(self):
        '''
        read the data inside given workbook as a map
        For single values, it's key-value form. For table values, it's key-(namedlist, list) form
        '''
        rmp = {}
        for name, var in self._nmps['n2addr'].items():
            nd = self.get(name)
            if isinstance(var, dict):
                off, di = var['_org'], var['_dir']
                n2imp = {x[1]: x[0] - off[0 if di not in self._cnstmp['_dir_ud'] else 1] for x in var['_i2n'].items() if x[1] != _cn_seqid}
                vvs = nd.get().value # all value in the table without seqid
                rmp[name] = self._transform(name, n2imp, vvs, di)
            else:
                rmp[name] = nd.get().value
        vdr = Normalizer(self)
        rmp, chgs = vdr.normalize(rmp)
        if chgs:
            vdr.updateXl(chgs)
        return rmp, chgs

    def _transform(self, tblname, n2imp, vvs, di):
        cns = [x for x in n2imp]
        lst = [cns, ]
        if di == 'up':
            vvs = reversed(vvs)
        elif di == 'right':
            vvs = zip(*vvs)
        elif di == 'left':
            vvs = reversed(tuple(zip(*vvs)))
        for i in vvs:
            lst.append([i[n2imp[x]] for x in cns])
        nls = NamedLists(lst)
        cols = self._nmps['_table_key_cols'].get(tblname)
        if not cols:
            return [x for x in nls]
        lst, cols = [], cols.split(',')
        for nl in nls:
            if not all((nl[x] for x in cols)):
                break
            lst.append(nl)
        return lst


class Node(object):
    '''
    a class represent a hot point or hot area
    '''

    def __init__(self, sht_tar, arg):
        self._sht_tar, self._arg = sht_tar, arg

    @property
    def isTable(self):
        '''
        is this node a table
        '''
        return isinstance(self._arg, dict)

    @property
    def maxCount(self):
        '''
        the maximum record count that this node can hold
        '''
        return self._arg['_max_cnt'] if self.isTable else 1

    def get(self, idx=0, name=None, getmerged=False):
        '''
        return the excel range of idx_th row and name
        when this is not a table, any argument passed into will be ignored
        '''
        if self.isTable:
            if idx < 0 or idx > self.maxCount:
                return None
            if idx == 0:
                return self._sht_tar.range(self._arg['_region'])
            di = self._arg['_dir']
            try:
                addr = self._arg[_nrm(name)][0]
                if idx > 1:
                    addr = nextrc(addr, di, idx - 1)
                return self._sht_tar.range(addr)
            except:
                return None
        rng = self._sht_tar.range(self._arg[0])
        if getmerged and rng and rng.api.mergecells:
            rng = apirange(rng.api.mergearea)
        return rng
