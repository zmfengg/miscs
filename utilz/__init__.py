#! coding=utf-8
'''
* @Author: zmFeng
* @Date: 2018-06-16 14:18:50
* @Last Modified by:   zmFeng
* @Last Modified time: 2018-06-16 14:18:50
Some handy utils for most of my projects
'''

from . import odbctpl, xwu
from ._jewelry import Karat, KaratSvc, RingSizeSvc, stsizefmt
from .miscs import *
from .resourcemgr import ResourceCtx, ResourceMgr, SessionMgr

karatsvc = KaratSvc()
ringsizesvc = RingSizeSvc()
