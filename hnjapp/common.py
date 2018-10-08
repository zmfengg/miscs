# coding=utf-8
"""
 *@Author: zmFeng
 *@Date: 2018-06-14 16:20:38
 *@Last Modified by:   zmFeng
 *@Last Modified time: 2018-06-14 16:20:38
 """


import inspect
import logging
import os
from numbers import Number

from hnjcore import JOElement

_logger = logging.getLogger("hnjapp")
_date_short = "%Y/%m/%d"
_dfkt = {"4": 925, "5": 925, "M": 8, "B": 9, "G": 10, "Y": 14, "P": 18}
thispath = os.path.abspath(os.path.dirname(inspect.getfile(inspect.currentframe())))


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
    if isinstance(jn, Number):
        jn = "%d" % int(jn)
    return _dfkt.get(jn[0])
