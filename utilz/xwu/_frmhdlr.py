'''
#! coding=utf-8
@Author:        zmFeng
@Created at:    2019-07-09
@Last Modified: 2019-07-09 4:20:09 pm
@Modified by:   zmFeng

 class help to read from/write to a sheet based on a set of field mapping
'''

from xlwings.constants import BordersIndex, LineStyle

from ..miscs import NamedList, triml
from ._xwu import NamedLists, addr2rc, apirange, nextcell, nextrc, rc2addr

# constances, keys are all triml
_consts = {
    '_dir_ud': ('down', 'up'),
    '_dir_ul': ('up', 'left'),
    '_dir_2_bdr': {
        'left': BordersIndex.xlEdgeLeft,
        'right': BordersIndex.xlEdgeRight,
        'down': BordersIndex.xlEdgeBottom,
        'up': BordersIndex.xlEdgeTop
    },
    '_d2d': {
        'd': 'down',
        'u': 'up',
        'l': 'left',
        'r': 'right'
    }
}

def _get_const(name):
    return _consts.get(triml(name))

def _d2d(d0):
    '''
    translate the direction in m_fmp to the one for xlwings
    '''
    return _get_const('_d2d')['r' if not d0 else triml(d0[0])]


class FormHandler(object):
    '''
    read data out from requested workbook
    Args:
        wb: the workbook to process on

        hdlr(FormHandler): if exists, the settings will be inited from hdlr instead of loading from wb

        `sheet_name=None: name of the target form(sheet)

        `field_nls`=None: a NamedLists with 3 columns: name,address,direction, direction should be one of u/d/l/r. if this is omitted, the wb must contains a sheet('MetaData') and inside the sheet, a table named 'm_fmp' must be there and the table at least contains 3 columns as above

        cn_seqid='seqid': id name of the table-like data. This is must-have, used for testing table size because the field_nls specify only the column name, no table size

        _nrm=triml: a method for normalize the name for dict access(deprecated)

        nrls=None: a NrlsInvoker instance, used in both read and write
    '''

    def __init__(self, wb, hdlr=None, **kwds):
        self._wb, self._sheet, self._nrls = wb, None, None
        self._nrm = kwds.get('_nrm', triml)
        '''
            self._nmps contains below maps:
                ._n2addr: name to address map. If this is a table item, _n2addr contains a _i2n map, which contains table's Id -> Name map
                ._addr2n: address to name map for those single value only
                ._reg2n: region((x0, y0), (x1, y1)) to name list, for those table values only. sorted by y0(because most x0 are the same)
        '''
        # hold the often used constans
        if hdlr and hasattr(hdlr, '_sheet'): # pylint: protected-access
            self._sheet = wb.sheets(getattr(hdlr, '_sheet').name)
            self._nmps, self._full_inited, self._cn_seqid, self._nrm, self._nrls = (getattr(hdlr, x) for x in '_nmps _full_inited _cn_seqid _nrm _nrls'.split())
        else:
            self._nmps = self._full_inited = None
            self._field_map(kwds.get('field_nls') or self._read_field_map())
            self._cn_seqid = kwds.get('cn_seqid', 'seqid')
        if not self._sheet:
            sn = kwds.get('sheet_name')
            if sn:
                self._sheet = wb.sheets(sn)
    @property
    def seqid(self):
        ''' name of the seqid
        '''
        return self._cn_seqid

    @seqid.setter
    def seqid(self, name):
        self._cn_seqid = self.normalize(name)

    @property
    def sheet(self):
        '''
        the target sheet I need to read_from/write_to
        '''
        return self._sheet

    @sheet.setter
    def sheet(self, sht):
        ''' @see :class: FormHandler.sheet
        '''
        self._sheet = sht

    def get(self, name, get_hotpoint=False):
        '''
        return a node related to given name(or address), for example, get('author')
        should return a Node who's get() will be Range("$L$2")
        '''
        if get_hotpoint:
            return self._get_hot_point(name)
        mp = self._nmps['_n2addr']
        nl = self.normalize(name).split(".")[0]
        if nl in mp:
            self._calc__n2addr(name)
            return Node(nl, self._sheet, mp[nl], self)
        return None

    def get_hotpoint(self, addr):
        '''
        convenience way for self.get(name, True)
        '''
        return self._get_hot_point(addr)

    def get_colnames(self, tblname, skip_seq=True):
        '''
        return the colnames of given table
        Args:
            tblname:    name of the table
        Returns:
            tuple(colname(String))
        '''
        cn_seqid = self.seqid if skip_seq else None
        return [x for x in self._nmps['_n2addr'][self.normalize(tblname)]['_i2n'].values()if x != cn_seqid]

    def read(self, names=None, nrls=None):
        '''
        read the data and return as a map
        For single values, it's key-value form. For table values, it's key-(namedlist, list) form

        Args:

            names=None: a tuple of name(string), specify the fields to get

            nrls:   A :class: NrlInvoker instance that will do normalization to the result map returned from my sheet.

        Returns:
            A map(name(string), value or tuple()) and a tuple(NamedList), refer to :class: BaseNrl for column names of the NamedList
        '''
        mp = self._nmps['_n2addr']
        ns = (self.normalize(name) for name in names) if names else mp
        rmp, cn_seqid = {}, self._cn_seqid
        for name in ns:
            nd, var = self.get(name), mp[name]
            if isinstance(var, dict):
                off, di = var['_org'], var['_dir']
                n2imp = {x[1]: x[0] - off[0 if di not in _get_const('_dir_ud') else 1] for x in var['_i2n'].items() if x[1] != cn_seqid}
                vvs = nd.get().value # all value in the table without seqid
                rmp[name] = self._transform(name, n2imp, vvs, di)
            else:
                rmp[name] = nd.get().value
        nrls = nrls or self._nrls
        logs = nrls.normalize(rmp)[1] if nrls else None
        return rmp, logs

    def write(self, mp=None, name=None, val=None, nrls=None):
        '''
        write by a map of data or by name + val pair. When name starts with '_', it will be ignored

        Args:

            mp:     a dict with many name-value pair. for example, {'author': 'zmFeng', 'createdate': datetime(2019, 1, 5), 'metal': [nl0, nl1...]}

            name:    the name of value to be set. Must be along with val
            val:    value of the name.

            nrls: A :class: NrlsInvoker instance
        '''
        fails = []
        def _add_fail(k, v):
            fails.append((k, v))
        nrls = nrls or self._nrls
        if mp:
            if nrls:
                nrls.normalize(mp)
            for n, v in mp.items():
                if n[0] == '_':
                    continue
                nd = self.get(n)
                if not nd:
                    _add_fail(n, v)
                # translate it to namedlist if it's not
                if isinstance(v, (tuple, list, NamedLists)):
                    lst = [x for x in v]
                    if lst and not isinstance(lst[0], NamedList):
                        lst = [x for x in NamedLists(lst)]
                    nd.value = lst
                else:
                    nd.value = v
        else:
            nd = self.get(name)
            if not nd:
                _add_fail(name, val)
            else:
                nrls = nrls or self._nrls
                if nrls:
                    nrls.normalize({name, val})
                nd.value = val
        return fails if fails else None

    def normalize(self, name):
        ''' normalize given name to the format I host the keys
        '''
        return self._nrm(name)

    @property
    def book(self):
        ''' the requested workbook '''
        return self._wb

    def _field_map(self, nls):
        _n2addr, _addr2n, mp = {}, {}, None
        for nl in nls:
            name, addr = (nl[x] for x in ('name', 'address'))
            if not self._sheet:
                self._sheet = self._wb.sheets(
                    addr[addr.find(']') + 1:addr.find('!')])
            idx, addr = name.find("."), addr[addr.find('!') + 1:]
            mp = _n2addr.setdefault(name[:idx], {}) if idx >= 0 else _n2addr
            if idx >= 0:
                name = name[idx + 1:]
            else:
                _addr2n[self.normalize(addr)] = name
            di = _d2d(nl.direction)
            mp[name] = (addr, di)
            #also put the index to name reserve lookup table
            if idx > 0:
                # get the adjusted directional from mp
                idx = addr2rc(addr)[0][1 if di in _get_const('_dir_ud') else 0]
                mp.setdefault('_i2n', {})[idx] = name
        self._nmps = {x[0]: x[1] for x in
            zip('_n2addr _addr2n _reg2n'.split(), (_n2addr, _addr2n, []))}

    def _read_field_map(self):
        '''
        read the field mapping from the requested workbook
        sht = self._wb.sheets('metadata')
        addr = apirange(sht.api.listObjects('m_fmp').Range)

        # don't use NamedRanges(addr, newinst=False) because the shape is obvisous
        # and NamedRanges is slow
        # for nl in NamedRanges(addr, newinst=False):
        return NamedLists(addr.value)
        '''
        return None

    def _calc__n2addr(self, name):
        '''
        do some range calculation based on the mappings
        '''
        mp = self._nmps['_n2addr'][self.normalize(name)]
        if not isinstance(mp, dict) or '_max_cnt' in mp:
            return
        _rg = self._sheet.range
        ud = _get_const('_dir_ud')
        cn_seqid = self._cn_seqid
        cn = cn_seqid if cn_seqid in mp else next(self._get_cns(mp))
        var, di = mp[cn]
        rng = self._sheet.range(var)
        if cn == cn_seqid:
            # don't use expand, quite slow
            # rng = _rg(rng, rng.end(di))# if di in _get_const('_dir_ul') else rng.expand(di)
            rng = _rg(rng.address + ':' + rng.end(di).address)
        else:
            # save time, 3 steps forward because in this case, won't be
            # less than 4 rows
            cn = nextcell(rng, di, detect_merge=False, steps=4)
            bi, his = _get_const('_dir_2_bdr')[di], []
            while cn.api.Borders(bi).LineStyle != LineStyle.xlLineStyleNone:
                cn = nextcell(cn, di, detect_merge=False)
                his.append(cn)
            rng = _rg(rng.address + ':' + his[-2].address)
        mc = rng.rows.count if di in ud else rng.columns.count
        mp['_max_cnt'], mp['_org'], mp['_dir'] = mc, addr2rc(var)[0], di
        mp['_region'] = var = self._calc_table_region(mp)
        # TODO:: sor tthe list by y0
        self._nmps['_reg2n'].append((addr2rc(var), self.normalize(name)))

    def get_region(self, name, skip_seq=False):
        '''
        return the region of given name.

        Args:
            name:       name to get

            skip_seq:   when given name is a table and seqid is provided, skip returning column of seqid
        '''
        n, sht = self.normalize(name), self._sheet
        mp = self._nmps['_n2addr'].get(n)
        if mp and isinstance(mp, tuple):
            return sht.range(mp[0])
        rn = '_region_nseq' if skip_seq else '_region'
        if rn not in mp:
            mp[rn] = self._calc_table_region(mp, skip_seq)
        return sht.range(mp[rn])

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

    def _calc_table_region(self, mp, skip_seq=False):
        ud, org = _get_const('_dir_ud'), mp['_org']
        di, mc = [mp[x] for x in ('_dir', '_max_cnt')]
        cols = [mp[x][0] for x in self._get_cns(mp) if not skip_seq or x != self._cn_seqid]
        idx = 1 if di in ud else 0
        cols = sorted([addr2rc(x)[0][idx] for x in cols])
        # for speed reason, don't use range(cell, cell), it's terribly slow(time spent: 0.338 -> 0.018)
        if idx:
            di, addr = 1 if di == 'down' else -1, org[0]
            addr = rc2addr((addr, cols[0]), (addr + di * (mc - 1), cols[-1]))
        else:
            di, addr = 1 if di == 'right' else -1, org[1]
            addr = rc2addr((cols[0], addr), (cols[-1], addr + di * (mc - 1)))
        return addr

    def _full_init(self):
        # only get the first cell to avoid multi-cell case
        if not self._full_inited:
            for x in [x[0] for x in self._nmps['_n2addr'].items() if isinstance(x[1], dict) and '_max_cnt' not in  x[1]]:
                self._calc__n2addr(x)
            self._full_inited = True

    def _get_hot_point(self, addr):
        '''
        return hotpoint of given address
        '''
        mp, ud = self._nmps['_addr2n'], _get_const('_dir_ud')
        di = mp.get(self.normalize(addr))
        if di:
            return di
        self._full_init()
        pt = addr2rc(addr)[0]
        for pts, name in self._nmps['_reg2n']:
            pt0, pt1 = pts
            # TODO:: find the exit point
            if all([pt0[x] <= pt[x] <= pt1[x] for x in range(2)]):
                mp = self._nmps['_n2addr'][self.normalize(name)]
                di, pt0 = 0 if mp['_dir'] in ud else 1, mp['_org']
                row = abs(pt0[di] - pt[di]) + 1
                return name, row, mp['_i2n'].get(pt[0 if di else 1])
        return None

    @staticmethod
    def _get_cns(mp):
        return (x for x in mp if x.find('_') < 0)

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
        cols = self._nmps.get('_table_key_cols')
        if not cols or tblname not in cols:
            return [x for x in nls if not x.isblank]
        cols = cols.get(tblname)
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

    def __init__(self, name, sht_tar, arg, hdlr):
        self._name, self._sheet, self._arg = name, sht_tar, arg
        self._hdlr = hdlr

    @property
    def name(self):
        ''' name of this Node
        '''
        return self._name

    @property
    def sheet(self):
        ''' worksheet of this Node
        '''
        return self._sheet

    @property
    def isTable(self):
        '''
        is this node a table
        '''
        return isinstance(self._arg, dict)

    def __getattribute__(self, name):
        if name == 'range':
            return self.get()
        if name == 'value':
            return self.get().value
        return super().__getattribute__(name)

    def __setattr__(self, name, value):
        if name == 'value':
            self._write(value)
        return super().__setattr__(name, value)

    @property
    def maxCount(self):
        '''
        the maximum record count that this node can hold
        '''
        return self._arg['_max_cnt'] if self.isTable else 1

    def get(self, idx=0, name=None, merged=False):
        '''
        return the excel range of idx_th row and name
        when this is not a table, any argument passed into will be ignored

        Args:
            only apply to table case, single value case will ignore any arguments

            idx=0:      0 to get the whole table range while <&gt> 0 to get given row
            name=None:  colname in the table to get from
            merged=False:  when there is merged cell, return the merged one

        '''
        if self.isTable:
            if idx < 0 or idx > self.maxCount:
                return None
            if idx == 0:
                return self._hdlr.get_region(self.name)
            di = self._arg['_dir']
            try:
                addr = self._arg[self._hdlr.normalize(name)][0]
                if idx > 1:
                    addr = nextrc(addr, di, idx - 1)
                return self.sheet.range(addr)
            except:
                return None
        rng = self.sheet.range(self._arg[0])
        if merged and rng and rng.api.mergecells:
            rng = apirange(rng.api.mergearea)
        return rng

    def _write(self, nls):
        '''
        write a collection of data to the given node
        it's test that only when there are more than 5 cells to write, fast-write
        will have benefit
        Args:
            nd: the Node item that this table is referring to
            nls: a tuple of namedlist item
        '''
        if not self.isTable:
            self.get().value = nls
            return
        self._hdlr.get_region(self._name, True).value = None # cleanup
        if not nls:
            return
        if isinstance(nls, NamedLists):
            nls = [nl for nl in nls]
        elif isinstance(nls, (tuple, list)) and not isinstance(nls[0], NamedList):
            nls = [nl for nl in NamedLists(nls)]

        # fast write works only for large data
        if len(nls) * len(nls[0].colnames) > 5:
            self._write_block(nls)
        else:
            for idx, nl in enumerate(nls):
                for cn in nl.colnames:
                    rng = self.get(idx + 1, cn)
                    if rng:
                        rng.value = nl[cn]


    def _write_block(self, nls):
        mp = self._arg
        org, cidx = mp['_org'], 1 if mp['_dir'] in _get_const('_dir_ud') else 0
        seqid = mp.get(self._hdlr.seqid)
        cnmp = {cn: addr2rc(mp[self._hdlr.normalize(cn)][0])[0][cidx] - org[cidx] -
            (1 if seqid else 0) for cn in nls[0].colnames}
        mcid = max(iter(cnmp.values())) + 1
        lsts = []
        for nl in nls:
            lst = [None] * mcid
            lsts.append(lst)
            for cn in nl.colnames:
                lst[cnmp[cn]] = nl[cn]
        di = mp['_dir']
        if seqid and addr2rc(seqid[0])[0] == org:
            org = nextrc(org, 'right' if di in _get_const('_dir_ud') else 'down')
        if di in ('up', 'left'):
            cnmp = mp['_max_cnt'] - len(lsts)
            org = nextrc(org, di, mp['_max_cnt'] - 1)
            if cnmp > 0:
                lsts.extend([[None] * mcid] * cnmp)
            if di == 'up':
                lsts = tuple(reversed(lsts))
            else:
                lsts = [tuple(reversed(x)) for x in tuple(zip(*lsts))]
        self.sheet.cells(*org).value = lsts
