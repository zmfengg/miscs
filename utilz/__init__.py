#! coding=utf-8 
'''
* @Author: zmFeng 
* @Date: 2018-06-16 14:18:50 
* @Last Modified by:   zmFeng 
* @Last Modified time: 2018-06-16 14:18:50 
Some handy utils for most of my projects
'''

from .resourcemgr import ResourceMgr, ResourceCtx, SessionMgr
from ._miscs import *
from ._jewelry import Karat, KaratSvc, RingSizeSvc
from . import odbctpl
from . import xwu

karatsvc = KaratSvc()
ringsizesvc = RingSizeSvc()