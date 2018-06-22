# coding=utf-8
'''
Created on 2018-05-18

@author: zmFeng
'''

from os import path
import math
import sys
import os
import threading
import inspect
from sqlalchemy.orm import Session
from .common import _logger as logger
from utilz import appathsep, getfiles, deepget, daterange, \
    splitarray, ResourceMgr, ResourceCtx, SessionMgr, isnumeric

__all__ = ["splitarray","appathsep","deepget","getfiles", "daterange",\
    "samekarat","ResourceCtx","ResourceMgr", "SessionMgr","isnumeric"]

_silveralphas = set(("4", "5"))

def samekarat(srcje, tarje):
    """ detect if the given 2 JOElement are of the same karat """
    if not (srcje and tarje): return
    return srcje.alpha == tarje.alpha or (srcje.alpha in _silveralphas and tarje.alpha in _silveralphas)