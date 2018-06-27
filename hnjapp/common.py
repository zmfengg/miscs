# coding=utf-8 
"""
 *@Author: zmFeng 
 *@Date: 2018-06-14 16:20:38 
 *@Last Modified by:   zmFeng 
 *@Last Modified time: 2018-06-14 16:20:38 
 """


import logging
from hnjcore import JOElement

_logger = logging.getLogger("hnjapp")
_date_short = "%Y/%m/%d"

def splitjns(jns):
    """ split the jes or runnings into 3 set
    jes/runnings/ids
    """
    if not jns:
        return
    jes, rns,ids= set(),set(),set()
    for x in jns:
        if isinstance(x, JOElement):
            jes.add(x)
        elif isinstance(x, int):
            ids.add(x)
        elif isinstance(x, str):
            if x.find("r") >= 0:
                i0 = int(x[1:])
                if i0 > 0: rns.add(i0)
            else:
                je = JOElement(x)
                if(je.isvalid):
                    jes.add(je)
    return jes, rns, ids