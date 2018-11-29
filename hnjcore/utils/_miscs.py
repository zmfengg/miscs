# coding=utf-8
'''
Created on 2018-05-18

@author: zmFeng
'''
from hnjcore import JOElement

__all__ = ["isvalidp17", "samekarat"]

_silveralphas = set("4 5".split())


def samekarat(srcje, tarje):
    """
    detect if the given 2 JOElement are of the same karat
    """
    jes = (
        srcje,
        tarje,
    )
    if not all(jes):
        return None
    jes = [x if isinstance(srcje, JOElement) else JOElement(x) for x in jes]
    return jes[0].alpha == jes[1].alpha or all(
        x.alpha in _silveralphas for x in jes)


def isvalidp17(p17):
    """
    minimum check if the given p17 code is a valid one
    """
    return isinstance(p17, str) and len(p17) == 17
    # and "0,1,2,3,4,9,C,P,W".find(p17[1:2]) >= 0
