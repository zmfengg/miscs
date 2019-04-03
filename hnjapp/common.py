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

from hnjcore import JOElement
from hnjcore.models.hk import JO, JOItem as JI
from sqlalchemy.orm import Query
from utilz import Config, karatsvc, trimu, stsizefmt, NamedList

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
    def __init__(self):
        self._st_sns_abbr = None
        # vermail name and color/display mapping
        self._vc_mp = {x["name"]: x for x in config.get("vermail.defs")}
        self._ptns = {x[0]: compile_r(config.get(x[1])) for x in zip(('pkno', 'mit'), ('pattern.jo.pkno', 'pattern.mit'))}
        self._ptns["microns"] = [compile_r(x) for x in config.get("pattern.jo.microns")]

        # locket shape name and description detection
        cfg = config.get('locket.shapes')
        self._lk_d2n = {x[1]: x[0] for x in cfg.items()}
        cfg = [x for x in zip(*(x for x in cfg.items()))]
        cfg = ['(^' + '$)|(^'.join(cfg[0]) + "$)", '(^' + ')|(^'.join(cfg[1]) + ")", ]
        self._ptns["locket"] = [compile_r(x) for x in cfg]
        # stone to default vx
        self._st2vx = config.get("vermail.stones")

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
        extract vcs by rmk/stone/main karat and so on, return a set of vermail color as  string
        '''
        # first, from remark
        rmk, vcs = jo.remark, set()
        idx = rmk.find(config.get("vermail.verb"))
        if idx >= 0:
            for var in rmk.split(config.get("vermail.verb"))[1:]:
                var = var[:6]
                cands = []
                for x in self._vc_mp.values():
                    if [1 for y in x["big5"] if var.find(y) >= 0]:
                        cands.append(x)
                if cands:
                    if len(cands) > 1:
                        cands = sorted(cands, key=lambda x: x.get('priority', 0))
                    vcs.add(cands[-1]["color"])
        #TODO: if there is CHAIN, maybe the chain need plating, this need invoking the extract_jis function, slow or duplicated?
        if wgts:
            kts = [m.karat for m in wgts.wgts if m]
            if 925 in kts and not vcs:
                vcs.add(karatsvc.COLOR_WHITE)
            if 9 in kts and jo.orderma.customer.name.strip() == 'GAM':
                vcs.add(karatsvc.COLOR_YELLOW)
            if 200 in kts and len(kts) > 1:
                vcs.add(karatsvc.getkarat(wgts.main.karat).color)
        if sts:
            for x in sts:
                vx = self._st2vx.get(x[:2])
                if vx:
                    vcs.add(self._vc_mp[vx]["color"])
        return vcs

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
        nl = NamedList("stone shape stsize stqty stwgt setting szcal ms sto shpo".split())
        _ms_chk = lambda cat, nl: nl.stqty == (1 if cat != "EARRING" else 2) and nl.szcal >= "0300"
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
                if _ms_chk(cat, nl):
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
        sto, shp, shpo = (None,) * 3
        st = "".join([x for x in st if "A" <= x <= "Z"])
        sz, qty, wgt = (ji.stsize, ji.qty, ji.wgt, )
        if self._ptns["mit"].search(st):
            sto, st = self.MISC_MIT, "MIT"
            sz = qty = None
        else:
            flag = [self._ptns['locket'][0].search(st), self._ptns['locket'][1].search(ji.remark)]
            if any(flag):
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
        mt = "".join(pts)
        return mt


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
