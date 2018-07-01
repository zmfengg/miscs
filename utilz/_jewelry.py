#! coding=utf-8 
'''
* @Author: zmFeng 
* @Date: 2018-07-01 11:38:33 
* @Last Modified by:   zmFeng 
* @Last Modified time: 2018-07-01 11:38:33 
'''
from collections import namedtuple
from csv import DictReader
from numbers import Number
from operator import attrgetter
from os import path
from .common import thispath

Karat = namedtuple("Karat","karat,name,fineness,category,color")

class KaratSvc(object):
    CATEGORY_GOLD = "GOLD"
    CATEGORY_SILVER = "SILVER"
    CATEGORY_COPPER = "COPPER"
    CATEGORY_BONDEDGOLD = "BG"

    _priorities = {CATEGORY_COPPER:-100,CATEGORY_SILVER:-50,CATEGORY_BONDEDGOLD:-10,CATEGORY_GOLD:10}
    """ class help to solve karat related issues """
    def __init__(self,fn = None):
        if not fn or not path.exists(fn):
            fn = path.join(thispath,"res","karats.csv")
        lst = []
        with open(fn,"r+t") as fh:
            rdr = DictReader(fh)
            for x in rdr:
                kt = x["karat"]
                if kt.isdigit(): kt = int(kt)
                fin = float(x["fineness"])
                if fin > 1.0: fin = fin / 100.0
                lst.append(Karat(kt,x["name"].strip(), \
                    fin,x["category"].strip(),x["color"].strip()))
        byid, byname, fingrp, fml = {},{},{},{}
        for x in lst:
            byid[x.karat] = x
            byname[x.name] = x
            fin = x.fineness            
            if fin < 100.0 and x.category == "GOLD":
                fingrp.setdefault(fin,[]).append(x)
        
        for x in fingrp.values():
            y = sorted(x, key = attrgetter("category","karat"))
            for ii in range(1,len(y)):
                fml[y[ii].karat] = y[0]
        
        self._byid, self._byname, self._byfamily, self._byfineness = byid, byname, fml, None
    
    @property
    def all(self):
        return self._byid.values()

    def __getitem__(self,key):
        return self.getkarat(key) 

    def getkarat(self, karat):
        """ return the karat object by id or by name
          for example, getkarat(8) or getkarat("8K")
        """
        if isinstance(karat,str):
            if karat.isdigit():
                karat = int(karat)
            else:
                karat = karat.upper().strip()
        for x in (self._byid, self._byname):
            if karat in x:
                return x[karat]        

    def getbyfineness(self,fineness):
        """ fineness must be an integer, the actual fineness * 1000, if not, I do it for you """
        if isinstance(fineness,Number):
            if fineness < 0: fineness = int(fineness * 1000)
            if not self._byfineness:
                self._byfineness = dict([(x.fineness * 1000,x) for x in self.all])
            if fineness in self._byfineness:
                return self.getfamily(self._byfineness[fineness])

    def getfamily(self,karat):
        """ the legacy karat issue: 9 -> 91 -> 98 10 -> 101 -> 108 ... """
        if not karat: return None
        if not isinstance(karat,Karat):
            karat = self.getkarat(karat)
            if not karat: return None
        if karat.karat in self._byfamily:
            karat = self._byfamily[karat.karat]        
        return karat

    def issamecategory(self,k0,k1):
        kx = [x if isinstance(x,Karat) else self[x] for x in (k0,k1)]
        if all(kx):
            return kx[0].category == kx[1].category

    def compare(self,k0,k1):
        if k0 is k1: return 0
        tcs = (k0,k1)
        cps = [self._priorities[x.category] for x in tcs]
        rc = cps[0] - cps[1]
        if rc == 0:
            rc = tcs[0].fineness - tcs[1].fineness
            if rc == 0: rc = k0.karat - k1.karat
        rc = 1 if rc > 0 else -1 if rc < 0 else 0
            
        return rc