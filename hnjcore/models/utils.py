# coding=utf-8
'''
Created on Apr 17, 2018



@author: zmFeng
'''

import csv
import os
from operator import attrgetter
import numbers
from collections import namedtuple

__all__ = ["JOElement","StyElement","KaratSvc","Karat","karatsvc"]

class JOElement(object):
    """ 
    representation of Alpha + digit composite key
    the constructor method can be one of:
    
    JOElement("A1234BC")    
    JOElement("A",123,"BC")
    JOElement(12345.0)
    
    JOElement(alpha = "A",digit = 123,suffix = "BC")
    """    
    __minlen__ = 5
    def __init__(self, *args, **kwargs):
        cnt = len(args)
        if(cnt == 1):
            self._parse_(args[0])
        elif(cnt >= 2):
            self.alpha = args[0].strip()
            self.digit = args[1]
            self.suffix = args[2].strip() if(cnt > 2) else "" 
        else:
            self._reset()
    
    def _parse_(self,jono):
        from numbers import Number        
        if not jono:
            self._reset()
            return
        stg, strs = 0, ["","",""]
        if isinstance(jono,Number): jono = "%d" % jono
        jono = jono.strip()
        for i in range(len(jono)):
            if(jono[i].isalpha()):
                if(stg == 0):
                    strs[0] = strs[0] + jono[i]
                else:
                    strs[2] = strs[2] + jono[i:]
                    break 
            elif(jono[i].isdigit()):
                if(not stg):                    
                    stg += 1
                    #first character is number, let it be alpha
                    if(len(strs[0]) == 0):
                        strs[0] = jono[i]
                        continue
                strs[1] = strs[1] + jono[i]
            else:
                break
        if(stg and strs[1].isdigit()):
            self.alpha = strs[0].strip()
            self.digit = int(strs[1])
            self.suffix = strs[2].strip()
        else:
            self._reset()
    
    def _reset(self):
        self.alpha = ""
        self.digit = 0
        self.suffix = "" 
            
    def __repr__(self, *args, **kwargs):
        return "JOElement(%s,%d,%s)" % (self.alpha,self.digit,self.suffix)
    
    def __str__(self, *args, **kwargs):
        if(hasattr(self,'digit')):
            return self.alpha + \
                (("%0" + str(self.__minlen__ - len(self.alpha)) + "d") % self.digit)
        else:
            return ""
        
    @property
    def value(self):
        return self.__str__()
    
    @property
    def name(self):
        return self.__str__()
    
    def isvalid(self):
        return bool(self.alpha) and bool(self.digit)
    
    def __composite_values__(self):
        return self.alpha,self.digit        
    
    def __eq__(self,other):
        return isinstance(other,JOElement) and \
            self.alpha == other.alpha and \
            self.digit == other.digit

    def __hash__(self):
        return hash((self.alpha,self.digit))

    def __ne__(self,other):
        return not self.__eq__(other)
    
    def __ge__(self,other):
        return isinstance(other,JOElement) and \
            self.alpha == other.digit and \
            self.digit >= other.digit

class StyElement(JOElement):
    def __composite_values__(self):
        pr = JOElement.__composite_values__(self)
        return pr[0],pr[1],self.suffix
    
    def __eq__(self, other):
        return JOElement.__eq__(self, other) and self.suffix == other.suffix

    def __hash__(self):        
        return hash((super(StyElement,self).__hash__(),self.suffix))

    def __str__(self, *args, **kwargs):
        val = super(StyElement,self).__str__(args,**kwargs)
        if val:
            val += self.suffix
        return val

Karat = namedtuple("Karat","karat,name,finess,category,color")

class KaratSvc(object):
    CATEGORY_GOLD = "GOLD"
    CATEGORY_SILVER = "SILVER"
    CATEGORY_COPPER = "COPPER"
    CATEGORY_BONDEDGOLD = "BG"

    """ class help to solve karat related issues """
    def __init__(self,fn = None):
        if not fn or not os.path.exists(fn):
            fn = os.path.dirname(__file__) + os.path.sep + "karats.csv"
        lst = []
        with open(fn,"r+t") as fh:
            rdr = csv.DictReader(fh)
            for x in rdr:
                kt = x["karat"]
                if kt.isdigit(): kt = int(kt)
                fin = float(x["finess"])
                if fin > 1.0: fin = fin / 100.0
                lst.append(Karat(kt,x["name"].strip(), \
                    fin,x["category"].strip(),x["color"].strip()))
        byid, byname, fingrp, fml = {},{},{},{}
        for x in lst:
            byid[x.karat] = x
            byname[x.name] = x
            fin = x.finess            
            if fin < 100.0 and x.category == "GOLD":
                fingrp.setdefault(fin,[]).append(x)
        
        for x in fingrp.values():
            y = sorted(x, key = attrgetter("category","karat"))
            for ii in range(1,len(y)):
                fml[y[ii].karat] = y[0]
        
        self._byid, self._byname, self._byfamily, self._byfiness = byid, byname, fml, None
    
    @property
    def all(self):
        return self._byid.values()

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

    def getbyfiness(self,finess):
        """ finess must be an integer, the actual finess * 1000, if not, I do it for you """
        if isinstance(finess,numbers.Number):
            if finess < 0: finess = int(finess * 1000)
            if not self._byfiness:
                self._byfiness = dict([(x.finess * 1000,x) for x in self.all])
            if finess in self._byfiness:
                return self.getfamily(self._byfiness[finess])

    def getfamily(self,karat):
        """ the legacy karat issue: 9 -> 91 -> 98 10 -> 101 -> 108 ... """
        if not karat: return None
        if not isinstance(karat,Karat):
            karat = self.getkarat(karat)
            if not karat: return None
        if karat.karat in self._byfamily:
            karat = self._byfamily[karat.karat]        
        return karat

karatsvc = KaratSvc()