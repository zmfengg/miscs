'''
#! coding=utf-8
@Author:        zmFeng
@Created at:    2019-06-28
@Last Modified: 2019-06-28 2:04:48 pm
@Modified by:   zmFeng

'''

import re
from collections.abc import Sequence

from sqlalchemy import or_

from hnjcore import JOElement
from utilz import NA, ResourceCtx, splitarray

from ..common import config, splitjns


def jesin(jes, objclz):
    """ simulate a in operation for jo.name """
    if not isinstance(jes, (
            tuple,
            list,
    )):
        jes = list(jes)
    q = objclz.name == jes[0]
    for y in jes[1:]:
        q = or_(objclz.name == y, q)
    return q


#these 4 object for sqlalcehmy's query maker for ids/names
idsin = lambda ids, objclz: objclz.id.in_(ids)
idset = lambda ids: {y.id for y in ids}
namesin = lambda names, objclz: objclz.name.in_(names)
nameset = lambda names: {y.name for y in names}

def _getjos(self, objclz, q0, jns, extfltr=None):
    ss = splitjns(jns)
    if not (ss and any(ss)):
        return None
    jes, rns, ids = ss[0], ss[1], ss[2]
    rsts = [None, None, None]
    if ids:
        rsts[0] = self._getbyids(q0, ids, idsin, objclz, idset, extfltr)
    if rns:
        rsts[1] = self._getbyids(q0, rns, lambda x, y: y.running.in_(x), objclz,
                                 lambda x: set([y.running for y in x]), extfltr)
    if jes:
        rsts[2] = self._getbyids(q0, jes, jesin, objclz, nameset, extfltr)
    its, failed = dict(), []
    for x in rsts:
        if not x:
            continue
        if x[0]:
            its.update(dict([(y.id, y) for y in x[0]]))
        if x[1]:
            failed.extend(x[1])
    return list(its.values()), failed


class SvcBase(object):
    _querysize = 20

    def __init__(self, trmgr):
        self._trmgr = trmgr

    def sessmgr(self):
        return self._trmgr

    def sessionctx(self):
        return ResourceCtx(self._trmgr)

    def _getbyids(self, q0, objs, qmkr, objclz, smkr, extfltr=None):
        """
        get object by providing a list of vars, return tuple with valid object tuple and not found set
        """
        if not objs:
            return None
        if not isinstance(objs, Sequence):
            objs = tuple(objs)
        objss = splitarray(objs, self._querysize)
        al = []
        with self.sessionctx() as cur:
            for x in objss:
                q = q0.filter(qmkr(x, objclz))
                if extfltr is not None:
                    q = q.filter(extfltr)
                lst = q.with_session(cur).all()
                if lst:
                    al.extend(lst)
        if al:
            if len(al) < len(objs):
                a0 = set(objs)
                x = smkr(al)
                na = a0.difference(x)
            else:
                na = None
        else:
            na = set(objs)
        return al, na

def fmtsku(skuno):
    if not skuno:
        return None
    skuno = skuno.strip()
    if skuno.upper() == NA:
        return None
    return skuno


class SNFmtr(object):
    # _ptn_rmk = re.compile(r"\(.*\)")
    _ptn_rmk = re.compile(r"[\(（].*[\)）]")
    _voidset = set("SN;HB".split(";"))
    _crmp = {
        "小心形": "小心型",
        "細心": "小心型",
        "SN": "",
        "#": "",
        "，": ",",
        "。": ".",
        " ": ",",
        "/": ",",
        ".": ","
    }
    _snno_mp = config.get("snno.translation")

    @staticmethod
    def _splitsn(sn, parsemode):
        if not sn:
            return None
        sfx, ots = "", ""
        for x in sn:
            if ord(x) > 128:
                sfx += x
            else:
                ots += x
        je = JOElement(ots)
        if je.digit > 0:
            pfx, sfxx = je.name, None
            if parsemode == 2:
                sfxx = [x for x in je.suffix]
                if sfx:
                    sfxx.append(sfx)
            sn = tuple(pfx + x for x in sfxx) if sfxx else (pfx,)
        else:
            sn = (sn,)
        return sn

    @staticmethod
    def _split_sns(sns):
        lst, part, asc0, push = [], "", None, False
        for x in sns:
            # split them by the non-ascii for each part
            part = ""
            for ch in x:
                if ch == ' ':
                    continue
                asc = min(max(ord(ch), 250), 251) == 250
                if asc0 is not None:
                    push = asc0 ^ asc or ch == ','
                    if push and ch == ',':
                        ch = ""
                if push:
                    lst.append((part, asc0))
                    part, push = "", False
                asc0 = asc
                if ch:
                    part += ch
            if part:
                lst.append((part, asc0))
        return lst

    @classmethod
    def formatsn(cls, sn, parsemode=2, retuple=False):
        """
        parse/formatted/sort a sn string to tuple or a string
        Args:
            parsemode: #0 for keep SN like "BT1234ABC" as it was
                    #1 for set SN like "BT1234ABC" to BT1234
                    #2 for split SN like "BT1234ABC" to BT1234A,BT1234B,BT1234C
            retuple:   return the result as a tuple instead of string
        Returns:
            None if the sn# is invalid, or String or tuple based on Args(retuple)
        """
        if not sn:
            return None
        for x in cls._crmp.items():
            if sn.find(x[0]) < 0:
                continue
            sn = sn.replace(x[0], x[1])
        sn = re.sub(cls._ptn_rmk, ",", sn)
        sns = [x for x in sn.split(",") if x]
        fixed = []
        buff = []
        for idx, ch in enumerate(sns):
            ch = cls._snno_mp.get(ch)
            if ch:
                fixed.extend(ch.split(','))
                buff.append(idx)
        if buff:
            for idx in reversed(buff):
                del sns[idx]
        buff, qflag = [], False
        for s0, asc in cls._split_sns(sns):
            if not s0:
                continue
            if qflag and s0[0] == ')':
                qflag = False
                continue
            if not asc:
                #chinese, big5, need translation, if not translated, ignore it
                s0 = cls._snno_mp.get(s0)
                if s0:
                    fixed.extend(s0.split(','))
                    s0 = None
            if not s0 or s0 in buff:
                continue
            qflag = s0[-1] == '('
            if qflag:
                s0 = s0[:-1]
            if len(s0) > 2 and s0 != NA and s0.find('MM') < 0 and s0.find('CM') < 0:
                if parsemode:
                    buff.extend(cls._splitsn(s0, parsemode))
                else:
                    buff.append(s0)
        if fixed:
            if parsemode:
                # avoid case like HB411D kind of SN# injected by _snno_mp
                fixed = [y for x in fixed for y in cls._splitsn(x, parsemode)]
            buff.extend(fixed)
        if buff:
            # don't change the original order but remove the duplicated
            st, lst = set(), []
            for s0 in buff:
                if s0 in st:
                    continue
                lst.append(s0)
                st.add(s0)
            buff = lst
        return buff if retuple else ",".join(buff)

formatsn = SNFmtr.formatsn
