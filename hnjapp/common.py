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
from utilz import Config

from hnjcore import JOElement

_logger = logging.getLogger("hnjapp")
thispath = os.path.abspath(os.path.dirname(inspect.getfile(inspect.currentframe())))
config = Config(os.path.join(thispath, "res", "conf.json"))
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
