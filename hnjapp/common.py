# coding=utf-8
"""
 *@Author: zmFeng
 *@Date: 2018-06-14 16:20:38
 *@Last Modified by:   zmFeng
 *@Last Modified time: 2018-06-14 16:20:38
 """

import inspect
import logging
from os import path
from re import compile as compile_r
from csv import DictReader
from socket import gethostname, gethostbyname

from hnjcore import JOElement
from utilz import Config, karatsvc, trimu, stsizefmt, NamedList
try:
    import dbf
except:
    dbf = None

_logger = logging.getLogger("hnjapp")
thispath = path.abspath(path.dirname(inspect.getfile(inspect.currentframe())))
config = Config(path.join(thispath, "res", "conf.json"))
_dfkt = config.get("jono.prefix_to_karat")
_date_short = config.get("date.shortform")


def splitjns(jns):
    """ split the jes or runnings into 3 set
    jes/runnings/ids
    """
    if not jns:
        return None
    jes, rns, ids = set(), set(), set()
    for x in jns:
        if isinstance(x, JOElement):
            if x.isvalid:
                jes.add(x)
        elif isinstance(x, int):
            ids.add(x)
        elif isinstance(x, str):
            if x.find("r") >= 0:
                i0 = int(x[1:])
                if i0 > 0:
                    rns.add(i0)
            else:
                je = JOElement(x)
                if je.isvalid:
                    jes.add(je)
    return jes, rns, ids


def _getdefkarat(jn):
    """ return the jo#'s default main karat """
    return _dfkt.get(JOElement.tostr(jn)[0])


class _JO_Extractor(object):
    '''
    class help to extract Vermail/Stone data out from JO(HK)'s remark
    '''
    MISC_MIT = "_MIT_"
    MISC_LKSZ = "_LKSZ_"
    def __init__(self, keep_dtl=False):
        '''keep the detail as possible, use by invoker other than c1calc
        '''
        self._st_sns_abbr = None
        # vermail name and color/display mapping
        self._vc_mp = {x["name"]: x for x in config.get("vermail.defs")}
        self._ptns = {x[0]: compile_r(config.get(x[1])) for x in (('pkno', 'pattern.jo.pkno'), ('mit', 'pattern.mit'))}
        self._ptns["microns"] = [compile_r(x) for x in config.get("pattern.jo.microns")]

        # locket shape name and description detection
        cfg = config.get('locket.shapes')
        self._lk_d2n = {x[1]: x[0].strip() for x in cfg.items()}
        cfg = [x for x in zip(*(x for x in cfg.items()))]
        cfg = ['(^' + '$)|(^'.join(cfg[0]) + "$)", '(^' + ')|(^'.join(cfg[1]) + ")", ]
        self._ptns["locket"] = [compile_r(x) for x in cfg]
        # stone to default vx
        self._st2vx = config.get("vermail.stones")
        self._keep_dtl = keep_dtl

    def extract_micron(self, rmk):
        '''
        extract the micro from remark by C1
        '''
        rk = rmk.replace("。", ".")
        mts = [x.group(1) for x in (x.search(rk) for x in self._ptns["microns"]) if x and x.groups() and x.group(1)]
        if mts:
            for mt in mts:
                try:
                    return float(mt)
                except:
                    print(mt)
        return None

    def extract_vcs(self, jo, sts=None, wgts=None):
        '''
        extract vcs by rmk/stone/main karat and so on, return a set of vermail color as  string. But when the color is a vermeil item, it might be a tuple of string instead of a stirng. an example return set is:
        {('WHITE', ), ('YELLOW', ), ('VK', '1 MIC')}
        '''
        # first, from remark
        def _add_clr(st, clr):
            st.add((clr, None))
        rmk, vcs, var = jo.remark, set(), config.get("vermail.verb")
        if wgts:
            mtclrs = {karatsvc.getkarat(x.karat).color for x in wgts.wgts if x and x.karat}
        else:
            mtclrs = {}
        if rmk.find(var) >= 0:
            for var in rmk.split(var)[1:]:
                var = var[:6]
                cands = []
                for x in self._vc_mp.values():
                    if [1 for y in x["big5"] if 0 <= var.find(y) < 4]: #only first 3 characters count
                        cands.append(x)
                if cands:
                    if len(cands) > 1:
                        cands = sorted(cands, key=lambda x: x.get('priority', 0))
                    # TODO:: why prior is cands[-1]["name"]
                    x = cands[-1]["color"]
                    if x == 'VK' or var.find('咪') >= 0: # don't take care of V18K/VPT900 or alike, they are place holder only
                        var = self._extract_micron_h(var)
                        if var:
                            x = (x, var + " MIC")
                    if isinstance(x, tuple):
                        vcs.add(x)
                    else:
                        _add_clr(vcs, x)
        if wgts:
            kts = [m.karat for m in wgts.wgts if m]
            if 925 in kts and not vcs:
                _add_clr(vcs, karatsvc.COLOR_WHITE)
            if 9 in kts and jo.orderma.customer.name.strip() == 'GAM':
                _add_clr(vcs, karatsvc.COLOR_YELLOW)
            if 200 in kts and len(kts) > 1:
                _add_clr(vcs, karatsvc.getkarat(wgts.main.karat).color)
        if sts:
            for x in sts:
                vx = self._st2vx.get(x[:2])
                if vx and self._vc_mp[vx]["color"] not in mtclrs:
                    _add_clr(vcs, self._vc_mp[vx]["color"])
        return vcs

    def _extract_micron_h(self, mic):
        cand = []
        for ch in mic:
            if '0' <= ch <= '9' or mic == '.':
                cand.append(ch)
            else:
                break
        if not cand:
            ch = -1
            for idx, var in enumerate(('一壹', '二貳兩', '三叁', '四肆', '五伍', )):
                ch = var.find(mic[0])
                if ch >= 0:
                    ch = idx
                    break
            if ch >= 0:
                cand.append(str(ch + 1))
        return ''.join(cand) if cand else None

    def extract_st_sns(self, sns):
        """ extract stone and shape out of the sns(Stone&Shape)
        return a tuple of stone, shape
        """
        if not sns:
            return (None,) * 2
        if not self._st_sns_abbr:
            self._st_sns_abbr = {
                x[0]: x[1:] for x in config.get("jo.stone.abbr")
            }
        return self._st_sns_abbr.get(sns) or (sns[1:], sns[0])

    @property
    def nl_ji(self):
        return NamedList("stone shape stsize stqty stwgt setting szcal ms sto shpo".split())

    @staticmethod
    def is_main_stone(cat, stqty, szcalc):
        '''
        refer to Utilz.is_main_stone()
        '''
        return stqty == (1 if cat != "EARRING" else 2) and szcalc and szcalc >= "0300"

    def extract_jis(self, jos, hksvc):
        '''
        extract stones based JO and JOItems. Main/Side stone is calculated
        return a tuple of:
            Stones(tuple)
            Handler(Namedlist), that can be used to access items in the prior tuple
            misc items(list), for example, MIT/LocketSnS. Any non-stone item will be stored here
        '''
        # ms is mainstone sign, can be one of M/S/None
        jis = hksvc.getjis(jos) or {}
        nl = self.nl_ji
        mp = {}
        for jo in jos:
            sts, mns, miscs = [], [], []
            cat = Utilz.getStyleCategory(jo.style.name.value, jo.description)
            if jo.id not in jis:
                continue
            for ji in jis[jo.id]:
                self._extract_ji(ji, nl)
                if not nl.sto or nl.sto[0] == '_':
                    miscs.append(nl.data)
                    continue
                # main stone detection
                if self.is_main_stone(cat, nl.stqty, nl.szcal):
                    nl.ms = "M"
                    mns.append(nl.data)
                else:
                    nl.ms = "S" if nl.stqty else "X"
                sts.append(nl.data)
            if len(mns) > 1:
                # find out the actual MS, set others to S(ide):
                pk = nl.getcol("szcal")
                mns = sorted(mns, key=lambda x: x[pk], reverse=True)
                for x in mns[1:]:
                    nl.setdata(x)["ms"] = "S"
            def srt_key(data):
                nl.setdata(data)
                return "%s,%s,%s,%s" % ({"M": "0", "S": "1"}.get(nl.ms, "Z"), nl.sto,
                                        nl.shpo, nl.szcal)
            mp[jo.id] = sorted(
                sts, key=srt_key), nl, tuple(miscs) if miscs else []
        return mp

    def _fetch_jis(self, jns, hksvc):
        if isinstance(jns, str):
            jns = (jns, )
        if isinstance(jns[0], str):
            jos = hksvc.getjos(jns.keys())[0]
        else:
            jos = jns
        return hksvc.getjis(jos)

    def _extract_ji(self, ji, nl):
        """
        extract the JOItem to the given NamedList
        """
        st = trimu(ji.stname)
        if st in ('', '.', '..', '...'):
            st = None
        else:
            st = "".join([x for x in st if "A" <= x <= "Z"])
        sto, shp, shpo = (None,) * 3
        sz, qty, wgt = (ji.stsize, ji.qty, ji.wgt, )
        if st and self._ptns["mit"].search(st):
            sto, st = self.MISC_MIT, "MIT"
            sz = qty = None
        else:
            # stupid confusion between HK and py, HK clean stone field
            flag = [self._ptns['locket'][0].search(st) if st else None, self._ptns['locket'][1].search(ji.remark)]
            if any(flag):
                if self._keep_dtl:
                    sto, st = self.MISC_MIT, "MIT"
                else:
                    sto, st = self.MISC_LKSZ, flag[0].group() if flag[0] else self._lk_d2n.get(flag[1].group())
                qty = wgt = None
            else:
                pk = self._extract_pk(ji.remark)
                if pk:
                    st, shp = pk, None
                    sto, shpo = st[:2], st[2]
                else:
                    st, shp = sto, shpo = self.extract_st_sns(st)
        if sz == ".":
            sz = None
        if sz:
            sz = stsizefmt(sz, True)
            if sto == self.MISC_LKSZ:
                sz = sz + "MM"
        nl.setdata(
            [st, shp, sz, qty, wgt, ji.remark,
             stsizefmt(sz), None, sto, shpo])

    def _extract_pk(self, rmk):
        """
        extract the PK# from the remark
        @param rmk: the remark of cstbld table
        """
        mt = self._ptns['pkno'].search(rmk)
        if not mt:
            return None
        pts = [x for x in mt.groups()]
        if pts[2].isnumeric():
            pts[2] = "%04d" % int(pts[2])
        pts = "".join(pts)
        if self._keep_dtl:
            pts = pts + ";" + rmk[mt.span()[1]:].strip()
        return pts


class Utilz(object):
    '''
    shared functions for this package
    '''
    _jo_extr, _cat_def = _JO_Extractor(), None
    MISC_MIT = _JO_Extractor.MISC_MIT
    MISC_LKSZ = _JO_Extractor.MISC_LKSZ

    @classmethod
    def extract_micron(cls, rmk):
        return cls._jo_extr.extract_micron(rmk)

    @classmethod
    def extract_jis(cls, jos, hksvc):
        '''
        extract stones based JO and JOItems. Main/Side stone is calculated
        return a tuple of:
            Stones(tuple)
            Handler(Namedlist), that can be used to access items in the prior tuple
            Discarded items(set), maybe useful to the caller
        '''
        return cls._jo_extr.extract_jis(jos, hksvc)

    @classmethod
    def nl_ji(cls):
        return cls._jo_extr.nl_ji

    @classmethod
    def getStyleCategory(cls, styno, jodsc=None):
        '''
        return the style category, for example, RING/EARRING.
        In the case of bracelet/bangle, providing the style# only won't return accumulate result, you should use getCategory("B1234", "钻石手镯")
        @param styno(str): The style# or a p17code, or a description
        @param jodsc(str): The JO's description(in chinese, gbk or big5)
        '''
        if not styno:
            return None
        cls._load_categories()
        cand = []
        for key, mp in cls._cat_def.items():
            for x in mp["patterns"]:
                if x.search(styno):
                    cand.append((
                        key,
                        mp,
                    ))
                    break
        if not cand:
            if len(styno) == 17:
                if not cls._p17_dc:
                    cls._p17_dc = P17Decoder()
                styno = cls._p17_dc.decode(styno, "PRODTYPE") or styno
            #maybe it's a description only, try to detect it
            styno = (styno, jodsc)
            for key, mp, x in (
                (x[0], x[1], y)
                    for x in cls._cat_def.items()
                    for y in x[1].get("keywords") or x[1].get("keywords_opt")):
                for y in (var for var in styno if var):
                    if y.find(x) >= 0:
                        return key
            # logger.debug("no category defined for style(%s)" % styno[0])
            return None
        if len(cand) > 1 and jodsc:
            for key, mp, x in (
                (x[0], x[1], y)
                    for x in cand
                    for y in x[1].get("keywords") or x[1].get("keywords_opt")):
                if jodsc.find(x) >= 0:
                    return key
        return cand[0][0]

    @classmethod
    def _load_categories(cls):
        if cls._cat_def:
            return
        cls._cat_def, cfg = {}, config.get("style.categories")
        for mp in cfg:
            cls._cat_def[mp["category"]] = mpx = {}
            for var in mp.get("patterns"):
                mpx.setdefault("patterns", []).append(compile_r(var))
            for var in ("keywords", 'keywords_opt'):
                lst = mp.get(var)
                if not lst:
                    continue
                mpx[var] = tuple(lst)
                break

    @classmethod
    def extract_vcs(cls, jo, sts=None, wgts=None):
        return cls._jo_extr.extract_vcs(jo, sts, wgts)
    
    @classmethod
    def is_main_stone(cls, cat, stqty, szcalc):
        '''
        check if given stone in given cat should be treated as main stone
        @param cat: can be obtained by Utilz.getStyleCategory(jo.style.name.value, jo.description)
        @param szcalc: a size formatted by utilz.stsizefmt
        '''
        return cls._jo_extr.is_main_stone(cat, stqty, szcalc)


class P17Decoder():
    """
    classeto fetch the parts(for example, karat) out from a p17
    """

    def __init__(self):
        self._cats_ = self._getp17cats()
        self._ppart = None

    @classmethod
    def _getp17cats(cls):
        """return the categories of all the P17s(from database)
        @return: a map of items containing "catid/cat/digits. This module should not have db code, so hardcode here
        """
        rdr = path.join(path.dirname(__file__), "res", "pcat.csv")
        with open(rdr, "r") as fh:
            rdr = DictReader(fh, dialect='excel-tab')
            return {trimu(x["name"]): (x["cat"], x["digits"]) for x in rdr}

    @classmethod
    def _getdigits(cls, p17, digits):
        """ parse the p17's given code out
        @param p17: the p17 code need to be parse out
        @param digits: the digits, like "1,11"
        """
        rc = ""
        for x in digits.split(","):
            pts = x.split("-")
            rc += p17[int(x) - 1] if len(pts) == 1 else p17[int(pts[0]) - 1:(
                int(pts[1]))]
        return rc

    def _getpart(self, cat, code):
        """fetch the cat + code from database"""
        # todo:: no database now, try from csv or other or sqlitedb
        # "select description from uv_p17dc where category = '%(cat)s' and codec = '%(code)s'"
        if not self._ppart:
            fn = path.join(path.dirname(__file__), "res", "ppart.csv")
            with open(fn, "r") as fh:
                rdr = DictReader(fh)
                self._ppart = {x["catid"] + x["codec"]: x["description"].strip()\
                    for x in rdr}
        code = self._ppart.get(self._cats_[cat][0] + code, code)
        if not isinstance(code, str):
            code = code["description"]
        return code

    def decode(self, p17, parts=None):
        """parse a p17's parts out
        @param p17: the p17 code
        @param parts: the combination of the parts name delimited with ",". None to fetch all
        @return the actual value if only one parts, else return a dict with part as key
        """
        ns, ss = tuple(trimu(x) for x in parts.split(",")) if parts\
            else self._cats_.keys(), []
        for x in ns:
            ss.append((x,
                       self._getpart(x, self._getdigits(p17,
                                                        self._cats_[x][1]))))
        return ss[0][1] if len(ns) <= 1 else dict(ss)

def styno2unit(styno):
    '''
    an easy sty# to unit method, not accurate, for HK's legacy system
    '''
    return 'PR' if styno.find('E') >= 0 else 'PC'

def cmpdbfs(ut, dbf0, dbf1, skips):
    '''
    check if 2 given dbf files are of the same, use by unittest
    Args:
        ut(TestCase):   an TestCase instance
        dbf0(string):   the expected dbf file
        dbf1(string):   the dbf file to compare with the expected
        skips(Dict):    the colnames to skip, if provided, should be a dict
    '''
    lst0, lst1 = [_read_dbf(x, skips) for x in (dbf0, dbf1)]
    ut.assertEqual(len(lst0), len(lst1), 'the count of records')
    # sort the result to avoid index error
    lst0, lst1 = [sorted(x) for x in (lst0, lst1)]
    for idx, lst in enumerate(lst0):
        ut.assertListEqual(lst, lst1[idx], 'comparing %s' % lst)

def _read_dbf(fn, skips):
    tbl, data, var = dbf.Table(fn), [], None
    if not skips:
        skips = ()
    with tbl.open():
        cns = sorted(tbl.field_names)
        tbl.skip()
        while not tbl.eof:
            lst = [tbl.current_record[x] for x in cns if x not in skips]
            # set those Null to 0 to avoid comparing error
            var = [x.strip() if isinstance(x, str) else x for x in lst]
            var = [0 if x is None else x for x in var]
            data.append(var)
            tbl.skip()
    return data

def is_in_cn():
    """
    simple method to check if the host is now in china, just get the ip address
    HK's ip is sth. like 192. while chinese is 172.
    """
    ip = gethostbyname(gethostname())
    return ip.find('172') == 0
