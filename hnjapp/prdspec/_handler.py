#! coding=utf-8
'''
* @Author: zmFeng
* @Date: 2019-05-21 14:00:29
* @Last Modified by:   zmFeng
* @Last Modified time: 2019-05-21 14:00:29
services help to handle the product specification sheet
'''

from datetime import date
from difflib import SequenceMatcher
from numbers import Number
from os import path

from xlwings.constants import BordersIndex, LineStyle

from hnjapp import config
from hnjapp.c1calc import Writer, Utilz
from hnjapp.svcs.misc import StylePhotoSvc, LKSizeFinder, StyleFinder
from utilz import NamedList, NamedLists, stsizefmt, trimu, karatsvc, triml, na
from utilz.xwu import (addr2rc, apirange, appmgr, appswitch, esctext, fromtemplate,
                       insertphoto, nextcell, nextrc, rc2addr, usedrange)

from ._nrlib import (BaseNrl, DescNrl, FeatureNrl, FinishingNrl, JENrl, MetalNrl,
                     NWgtNrl, PartsNrl, SizeNrl, StoneNrl, _Utilz, _Tbl_Optr, thispath)

_nrm = _Utilz.nrm

class _SNS2Wgt(object):
    ''' given stone, shape and size, return weight
    '''
    def __init__(self):
        self._data = None

    def _load(self):
        with open(path.join(thispath, "sztbl.csv"), 'r') as fh:
            nls = [ln.split(",") for ln in fh if ln[0] != "#"]
        nls = NamedLists(nls)
        self._data = {trimu("%s,%s,%s" % (nl.stone, nl.shape, stsizefmt(nl.size, True))): float(nl.unitwgt) for nl in nls}

    def get(self, name, shape, sz):
        '''
        given name,shape and sz, return unit weight(in ct)
        '''
        if not self._data:
            self._load()
        key = trimu("%s,%s,%s" % (name, shape, stsizefmt(sz, True)))
        return self._data.get(key)

_sns2wgt = _SNS2Wgt()

class Handler(object):
    '''
    read data out from requested workbook
    '''

    def __init__(self, wb, hdlr=None):
        ''' init a reader from given workbook
        Args:
            wb: the workbook to process on
            hdlr: if exists, the settings will be inited from hdlr instead of loading from wb
        '''
        self._wb, self._sht_tar = wb, None
        # hold the often used constans
        if hdlr and hdlr._cnstmp:
            self._cnstmp = hdlr._cnstmp
            self._sht_tar = wb.sheets(hdlr._sht_tar.name)
            self._nmps, self._meta_mp, self._full_inited = (getattr(hdlr, x) for x in '_nmps _meta_mp _full_inited'.split())
        else:
            self._nmps = self._cnstmp = self._meta_mp = self._full_inited = None
            self._init_consts()
            self._read_field_map()

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

    @property
    def book(self):
        ''' the requested workbook '''
        return self._wb

    def _read_hint_defs(self):
        '''
        read meta data except the field mappings from the meta sheet
        '''
        if self._meta_mp:
            return
        self._meta_mp = {}
        for sn, mp in config.get('prodspec.meta_tables').items():
            sht = self._wb.sheets(sn)
            for meta_type, tblname in mp.items():
                addr = apirange(sht.api.listObjects(tblname).Range)
                self._meta_mp[_nrm(meta_type)] = {_nrm(nl['name']): nl for nl in NamedLists(addr.value)}
        # in the case of stone name, build an abbr -> name hints map
        mp = {_nrm(nl['hp.abbr']): _nrm(nl['name']) for nl in self._meta_mp[_nrm('stone.name')].values()}
        self._meta_mp[_nrm('stone.abbr')] = mp


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
        meta_type = _nrm(meta_type)
        mp = self._meta_mp[meta_type]
        if not mp:
            return 0, None
        ncand = _nrm(cand)
        if meta_type == _nrm('stone.name'):
            ncand = self._meta_mp[_nrm('stone.abbr')].get(ncand, ncand)
        nl = mp.get(ncand)
        if nl:
            return 1, _nrm(nl['name']) # TODO::return based on meta_type
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

    def _read_field_map(self):
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
        self._nmps['_table_key_cols'] = {_nrm(nl.tblname): nl.cols for nl in NamedLists(apirange(sht.api.listObjects('m_fkc').Range).value)}

    def _calc_n2addr(self, name):
        '''
        do some range calculation based on the mappings
        '''
        mp = self._nmps['n2addr'][_nrm(name)]
        if not isinstance(mp, dict) or '_max_cnt' in mp:
            return
        _rg = self._sht_tar.range
        ud = self._cnstmp['_dir_ud']
        cn_seqid = _Utilz.cn_seqid()
        cn = cn_seqid if cn_seqid in mp else next(self._get_cns(mp))
        var, di = mp[cn]
        rng = self._sht_tar.range(var)
        if cn == cn_seqid:
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
        self._nmps['reg2n'][_nrm(name)] = addr2rc(var)

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
        should return a Node who's get() will be Range("$L$2")
        '''
        if get_hotpoint:
            return self._get_hot_point(name)
        mp = self._nmps['n2addr']
        nl = _nrm(name).split(".")[0]
        if nl in mp:
            self._calc_n2addr(name)
            return Node(nl, self._sht_tar, mp[nl], self)
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
                mp = self._nmps['n2addr'][_nrm(name)]
                di, pt0 = 0 if mp['_dir'] in ud else 1, mp['_org']
                row = abs(pt0[di] - pt[di]) + 1
                return name, row, mp['_i2n'].get(pt[0 if di else 1])
        return None

    @staticmethod
    def _get_cns(mp):
        return (x for x in mp if x.find('_') < 0)

    def read(self, normalize=True, upd_src=True):
        '''
        read the data inside given workbook as a map
        For single values, it's key-value form. For table values, it's key-(namedlist, list) form
        Args:
            normalize: normalize the source data, this should always be true
                because there might be mal-form data
            upd_src: update the source(excel) for those been normalized
        '''
        rmp, cn_seqid = {}, _Utilz.cn_seqid()
        for name, var in self._nmps['n2addr'].items():
            nd = self.get(name)
            if isinstance(var, dict):
                off, di = var['_org'], var['_dir']
                n2imp = {x[1]: x[0] - off[0 if di not in self._cnstmp['_dir_ud'] else 1] for x in var['_i2n'].items() if x[1] != cn_seqid}
                vvs = nd.get().value # all value in the table without seqid
                rmp[name] = self._transform(name, n2imp, vvs, di)
            else:
                rmp[name] = nd.get().value
        if upd_src or normalize:
            rmp = _NRInvoker(self).normalize(rmp, normalize, BaseNrl.LEVEL_ADVICE)
        return rmp

    def _get_raw_colnames(self, tblname):
        cn_seqid = _Utilz.cn_seqid()
        return [x for x in self._nmps['n2addr'][_nrm(tblname)]['_i2n'].values()if x != cn_seqid]

    @staticmethod
    def from_jns(jns, tpl_file=None, **kwds):
        '''
        create map for writing back to excel from JO
        '''
        cnt = 0
        app, tk = appmgr.acq()
        sws = appswitch(app, {'EnableEvents': False})
        if not tpl_file:
            tpl_file = config.get("prodspec.template")
        hdlr = Handler(fromtemplate(tpl_file, app))
        rmp = FromJO(hdlr, kwds.get('hksvc'), kwds.get('cnsvc'), kwds.get('cache_sm')).read(jns)
        noimg = kwds.get('noimg')
        for jn, rst_hints in rmp.items():
            cnt += 1
            if cnt > 1:
                hdlr = Handler(fromtemplate(tpl_file, app), hdlr)
            nlzr = _NRInvoker(hdlr)
            mp = nlzr.normalize(rst_hints, False)
            for name, nls in mp[0].items():
                nd = hdlr.get(name)
                if not nd:
                    continue
                nd.value = nls
            styn = mp[0]["_styno"]
            jn = esctext(jn)
            # unprotect then protect
            if not noimg:
                hdlr._sht_tar.api.Unprotect()
                hdlr._insert_imgs(styn, jn)
                nlzr.update_xls(mp[1], hdlr, BaseNrl.LEVEL_ADVICE)
                hdlr._sht_tar.api.Protect(DrawingObjects=True, Contents=True, Scenarios=True)
            yield (jn, styn), hdlr.book, mp[0]
        appmgr.ret(tk)
        appswitch(app, sws)

    def _insert_imgs(self, styno, jono=None):
        stysvc = StylePhotoSvc.getInst()
        fns = stysvc.getPhotos(styno, hints=jono)
        if not fns:
            return
        pns = ['photo%d' % x for x in range(2)]
        for idx, pn in enumerate(pns):
            if idx >= len(fns):
                break
            rng = self.get(pn).get(getmerged=True)
            insertphoto(fns[idx], rng).name = pn

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

    def __init__(self, name, sht_tar, arg, hdlr):
        self._name, self._sht_tar, self._arg = name, sht_tar, arg
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
        return self._sht_tar

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

    def get(self, idx=0, name=None, getmerged=False):
        '''
        return the excel range of idx_th row and name
        when this is not a table, any argument passed into will be ignored
        Args(only apply to table case, single value won't accept any argument):
            idx: get the whole table range
            name: colname of in the table to get from
            getmerged: when there is merged cell, return the merged one, default is False
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

    def _write(self, nls, block_writing=None):
        '''
        write a collection of data to the given node
        it's test that only when there are more than 5 cells to write, fast-write
        will have benefit
        Args:
            nd: the Node item that this table is referring to
            nls: a tuple of namedlist item
            block_writing: using block writing, None for auto-detect, True to force block writing, False to use one-by one writing
        '''
        if not self.isTable:
            self.get().value = nls
            return
        if block_writing is None:
            # fast write will be fast only when there is enough cells
            block_writing = len(nls) * len(nls[0].colnames) > 5
        # logger.debug('using %s for table(%s)', 'Block-Writing' if block_writing else 'One-by-One-Writing', nd.name)
        if block_writing:
            self._write_block(nls)
        else:
            for idx, nl in enumerate(nls):
                for cn in nl.colnames:
                    rng = self.get(idx + 1, cn)
                    if not rng:
                        continue
                    rng.value = nl[cn]

    def _write_block(self, nls):
        mp = self._arg
        org, cidx = mp['_org'], 1 if mp['_dir'] in self._hdlr._cnstmp['_dir_ud'] else 0
        seqid = mp.get(_Utilz.cn_seqid())
        cnmp = {cn: addr2rc(mp[_nrm(cn)][0])[0][cidx] - org[cidx] -
            (1 if seqid else 0) for cn in nls[0].colnames}
        mcid = max(iter(cnmp.values())) + 1
        lsts = []
        for idx, nl in enumerate(nls):
            lst = [None] * mcid
            lsts.append(lst)
            for cn in nl.colnames:
                lst[cnmp[cn]] = nl[cn]
        di = mp['_dir']
        if seqid:
            if addr2rc(seqid[0])[0] == org:
                org = nextrc(org, 'right' if di in self._hdlr._cnstmp['_dir_ud'] else 'down')
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

class _NRInvoker(object):
    '''
    class help to validate the form's field one by one
    mainly make use of the sub classes of BaseNrl to complete the task
    '''

    def __init__(self, hdlr=None):
        """
        Args:
            hdlr: an Handler instance help to access the excel if there is
        """
        self._nrl_mp, self._hdlr = None, hdlr
        self._init_nrls()

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
        if self._nrl_mp:
            return
        self._nrl_mp = {}
        # maybe from the config file
        tu = BaseNrl.get_base_nrm('tu')
        for n in 'styno qclevel parent craft'.split():
            self._nrl_mp[_nrm(n)] = tu
        mp = {
            'docno': JENrl,
            'size': SizeNrl,
            'metal': MetalNrl,
            'finishing': FinishingNrl,
            'parts': PartsNrl,
            'stone': StoneNrl
        }
        for n, cn in mp.items():
            self._nrl_mp[_nrm(n)] = cn(name=n)
        tu = FeatureNrl(name='feature')
        for n in 'feature feature1'.split():
            self._nrl_mp[_nrm(n)] = tu
        n = _nrm('netwgt')
        self._nrl_mp[n] = NWgtNrl(name=n)
        n = _nrm('description')
        self._nrl_mp[n] = (BaseNrl.get_base_nrm('tu'), DescNrl(name=n), )

    def normalize(self, mp, upd_src=True, hl_level=BaseNrl.LEVEL_MINOR):
        '''
        normalize the result map
        Args:
            mp({string, BaseNrl}): a map generated by prdspec.Handler.read() method
            upd_src: update the source excel for those normalized
            hl_level: hight light those changes has high value than this
        '''
        logs, hdlr = [], self._hdlr
        for pp_name, vdrs in self._nrl_mp.items():
            vdrs = self._nrl_mp.get(pp_name)
            if not vdrs:
                continue
            if isinstance(vdrs, BaseNrl):
                vdrs = (vdrs, )
            for vdr in vdrs:
                if hdlr:
                    vdr.hdlr = hdlr
                logx = vdr.normalize(mp, pp_name)
                if logx:
                    logs.extend(logx)
        if logs:
            logs.insert(0, tuple(BaseNrl.nl.colnames))
            logs = self._merge_logs(NamedLists(logs))
            if hdlr and upd_src:
                self.update_xls(logs, hdlr, hl_level)
        return mp, logs

    @staticmethod
    def _merge_logs(logs):
        _mk = lambda x: (x.name, x.row, x.colname)
        mp = {}
        for log in logs:
            key = _mk(log)
            if key not in mp:
                mp[key] = log.clone(log.data)
            else:
                log0 = mp[key]
                log0.newvalue = log.newvalue
                lvls = log.level, log0.level
                if lvls[0] > lvls[1]:
                    log0.level = lvls[0]
                rmks = log.remarks, log.remarks
                rmks = [x for x in rmks if x]
                if rmks:
                    log0.remarks = ";".join(rmks)
        return tuple(mp.values())

    def update_xls(self, logs, hdlr, hl_level):
        '''
        high-light the given invalid items(a collection of self._nl_vld)
        '''
        for nl in logs:
            if not nl.name:
                continue
            nd = hdlr.get(nl.name)
            rng = nd.get(nl.row, nl.colname) if nl.row else nd.get()
            rng.value = nl.newvalue
            if nl.level >= hl_level:
                self._hl(rng, nl.level, nl.remarks)

class FromJO(object):
    '''
    read data from JO and return it as a ready-to-fill data set
    '''
    def __init__(self, hdlr, hksvc, cnsvc, cache_sm):
        '''
        need hdlr to read the field maps
        '''
        self._hdlr, self._hksvc, self._cnsvc = hdlr, hksvc, cnsvc
        self._mpr = _Tbl_Optr(hdlr)
        self._styfndr = StyleFinder(cache_sm) if cache_sm else None

    def read(self, jns):
        '''
        return the data of given jono
        Args:
            jns(tuple of String): the jO#s
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
        for row in raw[1:]:
            nl_src.setdata(row)
            jn = nl_src.jono # this jo# has a ' as prefix, for excel inserting
            if jn:
                self._mpr.mp = rmp[jn] = mp = {_nrm('docno'): jn}
                self._fetch_metal(nl_src)
                mp['_styno'] = styno = nl_src['styno'] # _styno for convenience only
                mp['_wgtinfo'] = jodtls[esctext(jn)]['wgt']
                if self._styfndr:
                    jn = LKSizeFinder(self._styfndr, stynos=(styno, )).find()
                    if jn:
                        mp[_nrm('size')] = ','.join(jn[styno])
                    jn = self._styfndr.find_by_feature(styno)
                    if jn:
                        mp[_nrm('_family')] = jn
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
            _FromJO_BL(self, mp, jodtls.get(esctext(mp[_nrm('docno')]))).bl()
        return rmp

    def _nrl_parts(self):
        tn, mp = _nrm('parts'), self._mpr.mp
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
        tn, mp = _nrm('metal'), self._mpr.mp
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
        tn, mp = _nrm('stone'), self._mpr.mp
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
                var = config.get("prodspec.fromjo.stshape_map").get(var)
                if var:
                    nl.shape = var
            if st == 'CZ' and nl.wgt and nl.wgt >= 1 or nl.wgt  and nl.qty and abs(nl.wgt - nl.qty) < 0.001: # CZ or fake weight from HK
                nl.wgt = 0
            else:
                nl.wgt = round(nl.wgt or 0, 3)
            nl.wgtunit = 'CT'
            if not nl.wgt:
                lvl, uwgt = config.get("prodspec.stsnscvt.level"), None
                if len(st) > 5 and self._cnsvc and lvl > 0:
                    uwgt = self._cnsvc.getavgpkwgt(st, nl.size)
                elif lvl > 1:
                    uwgt = _sns2wgt.get(st, nl.shape, nl.size)
                    if not uwgt and self._cnsvc and lvl > 2:
                        uwgt = self._cnsvc.getavgstwgt(st, nl.shape, nl.size)
                if uwgt:
                    nl.unitwgt = uwgt
            var = nl.setting
            nl.main = 'Y' if var.find('主') >= 0 else 'N'
            stmp = config.get("prodspec.fromjo.setting_map")
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
        exts = self._mpr.mp.get(_nrm('finishing'))
        exts = set(x.remarks for x in exts) if exts else {}
        vmp = None
        for fld, rmk, vc in (('f_pen', 'BY_PEN', 'VW'), ('f_chain', 'CHAIN', 'V?'), ('f_tt', 'VT/T', 'VX')):
            if rmk in exts:
                continue
            var = nl_src[fld]
            if not var:
                continue
            if vc == 'V?':
                mtls = self._mpr.mp.get(_nrm('metal'))
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
                elif not self._mpr.mp.get(_nrm('size')):
                    self._mpr.mp[_nrm('size')] = _Utilz.get_lksz(rmk)
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

    def __init__(self, frmjo, mp, jodtl):
        self._frmjo, self._mp, self._jodtl = frmjo, mp, jodtl
        self._hdlr = frmjo._hdlr
        self._mpr = _Tbl_Optr(self._hdlr, mp)

    def bl(self):
        self._mpr.set_feature('XSTYN', self._mp['_styno'])
        self._mpr.set_feature('SN', self._jodtl.get('sn') or na)
        bits = self._build_bits()
        #sometimes there is casting locket, but just let it be
        if bits["locket"] or bits['stp']:
            self._mpr.set_feature('REMARKS', '片厚_C')
        self._mp['hallmark'] = ";".join(bits['hm'])
        if bits['stp']:
            self._mp[_nrm('craft')] = 'MIXTURE' if bits['cst'] else 'STAMPING'
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
            idx = rmk.find('色')
            x = '_色绒布' if idx < 0 else rmk[idx - 1] + "色绒布"
            self._mpr.add_parts(x, 'PKM', 2)
            self._mpr.add_parts('膠片', 'PKM', 2)
            x = _nrm('size')
            if not self._mp.get(x):
                shp, self._mp[x] = '_HEART', '_HL:_MM'
            else:
                shp = config.get('locket.shapes').get(self._mp[x].split(':')[0], '_HEART')
            self._mpr.set_feature('KEYWORDS', '%s;LOCKET' % shp)
        if bits["cat"] == 'ring':
            self._hdlr.get('size').value = '_N'
        self._check_glue()
        self._check_cards(bits)
        self._check_fin_by_mtl()
        self._check_text()

    def _check_glue(self):
        sts = self._mp.get(_nrm('stone'))
        if not sts:
            return
        for nl in sts:
            if nl.name.find('GC') < 0: # stone name here not yet extended, still abbr
                continue
            nl.setting = 'GLUE'
            ss = [self._jodtl[x] for x in ('description', 'rmk')]
            idxs = [x.find('色') for x in ss]
            if any((x > 0 for x in idxs)):
                xtr = lambda idx: ss[idx][idxs[idx] - 1]
                clr = xtr(0) if idxs[0] > 0 else xtr(1)
            else:
                clr = '白'
            clr += '色泥膠'
            self._mpr.add_finishing('OP', na, '力架')
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
        mtls = self._mp.get(_nrm('metal'))
        if not mtls:
            return
        fins = self._mp.get(_nrm('finishing'))
        fins = set((nl['method'] for nl in fins)) if fins else set()
        for mtl in mtls:
            kt = mtl.karat
            if kt.find('BG') >= 0:
                if 'VY' not in fins:
                    self._mpr.add_finishing('V_Y', '0.125 MIC', 'FOR_PRT')
                    fins.add('VY')
            elif str(kt) == '925':
                if 'VW' not in fins:
                    self._mpr.add_finishing('V_W', '0.125 MIC', 'FOR_PRT')
                    fins.add('VW')
            elif kt == '9K' and self._jodtl.get('cstname') == 'GAM':
                if 'VY' not in fins:
                    self._mpr.add_finishing('VY', '0.125 MIC', 'FOR_PRT')
                    fins.add('VY')

    def _check_text(self):
        ''' check if there is text inside JO's remark, if yes, extract it to keywords
        '''
        txt = _Utilz.get_text(self._jodtl['rmk'])
        if txt:
            self._mpr.set_feature('TEXT', '_' + txt)

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
        # bits['locket'] = flag = desc.find('相盒') >= 0
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
                flag = x.find(var[1]) >= 0
                if flag:
                    bits[var[0]] = flag
        for var in (('vy', '電黃'), ('vw', '電白'), ('vr', '電玫瑰')):
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
        kts = [karatsvc.getkarat(x.karat) for x in self._mp[_nrm('metal')]]
        bits['tc'], cn = len(kts), self._jodtl['cstname']
        hms = bits['hm']
        for var in kts:
            hm = config.get('prodspec.metal.mark').get(var.name)
            if hm:
                if var.karat == 925:
                    bits['925'] = True
                if var.category != 'GOLD' and any((bits[x] for x in ('vy', 'vr'))):
                    # non gold, vgold, mark
                    x = config.get("prodspec.customer.vgmark").get(cn)
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
        hm = config.get('prodspec.customer.mark').get(cn)
        if hm:
            hms.append(hm)
        var = self._mp.get(_nrm('stone'))
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
