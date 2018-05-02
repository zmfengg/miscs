# coding=utf-8
'''
Created on 2018-05-03

class to implement the PAJ Cost calculating

all cost related field are USD based
all weight related field are GM based

@author: zmFeng
'''

from collections import namedtuple
from _ast import arguments, Param
from sqlalchemy.ext.associationproxy import _NotProvided

#the fineness map for this calculation
_fineness = {8:33.3,81:33.3,88:33.3,9:37.5,91:37.5,98:37.5,10:41.7,101:41.7,108:41.7,14:58.5,141:58.5,148:58.5, \
    18:75.0,181:75.0,188:75.0,200:100,"925PAJ":95.0,"925":92.5}
#karat and weight
_KNW = namedtuple("_KNW","karat,wgt")

#mps string and the corresponding silver/gold value
class MPS():
    _slots = [None,None,None,None]
        
    def _parse(self,mps):
        mps = mps.strip().upper()
        tarmps = None
        for mp in mps.split(";"):
            ps = mp.split("=")
            if(len(ps) == 2 and ps[1].isdigit()):
                idx = 0 if(ps[0] == "S" or ps[0] == "SILVER") else 1 if(ps[0] == "G" or ps[0] == "GOLD") else -1
                if(idx >= 0):
                    self._slots[idx] = float(ps[1])
                    if(tarmps): tarmps += ";"
                    tarmps += ("S" if idx == 0 else "G") + "=" + ps[1]
        self._slots[3] = (tarmps != None)        
        if(tarmps): self._slots[2] = tarmps

    def __init__(self,mps):
        self._parse(mps)
    
    @property
    def isvalid(self):
        return self._slots[3] != None
    
    @property
    def gold(self):
        return self._slots[1] if self._slots[1] else 0
    
    @property
    def silver(self):
        return self._slots[0] if self._slots[0] else 0
    
    @property
    def mps(self):
        return self._slots[2]

Increment = namedtuple("Increment", "prdwgt,silver,gold,lossrate")

#the china cost related data    
PajChina = namedtuple("PajChina", "china,discount,increment,mps")

#product weight, of mainpart/auxpart/parts
class PrdWgt(namedtuple("PrdWgt", "main,aux,part")):
    @property
    def wgts(self):
        return (self.main,self.aux,self.part)


#constants
PAJCHINAMPS = MPS("S=30;G=1500")

def _getfiness(karat, vendor = None):
    """return the finenss of the given karat
    @param karat: the karat in numeric form, for example, 8 or 81
    @param vendor: PAJ or Non-PAJ or None  
    """
    
    lst = (karat, "%d%s" %(karat,vendor) if(vendor) else str(karat))
    rr = [_fineness[x] for x in lst if(_fineness.has_key(x))]
    if(len(rr) > 0): return rr[0] / 100.0

def _calcincr(prdwgt,vendor = "PAJ"):
    """ calculate the increment based on the product weight provided
        @param prdwgt:  weight of the product
        @param vendor: PAJ or Non-PAJ or None
    """
    kws = prdwgt.wgts
    s = g = 0;
    
    hix = [ii for ii in range(len(kws)) if kws[ii] and kws[ii].wgt > 0]
    mLossRate = max([1.1 if kws[x].karat == 925 else 1.06 for x in hix])
    for idx in hix:
        kw = kws[idx]
        #parts does not have loss
        lossRate = mLossRate if(idx < 3) else 1.0
        r0 = kw.wgt * _getfiness(kw.karat,vendor) * lossRate / 31.1035;
        if(kw.karat == 925):                
            s += r0;                    
        elif(kw.karat == 200):
            pass
        else:
            g += r0;

    return Increment(prdwgt,s,g,mLossRate)

def _calchina(prdwgt,refup,refmps,tarmps = None):
    """ calculate the China cost based on the provided arguments
    @param prdwgt: weights of the product
    @param refup: the reference unit price
    @param refmps: the reference mps of the @refup
    @param tarmps: the target MPS the need to be calculated. PAJ's china MPS is S=30;G=1500  
    """
    if(not( prdwgt and refup and refmps)): return None
    if(not tarmps): tarmps = PAJCHINAMPS
    if not (refup > 0 and refmps.isvalid() and tarmps.isvalid()): return None 
    
    #the discount ratio, when there is silver, follow silver, silver = 0.9 while gold = 0.85
    incr = _calcincr(prdwgt,"PAJ");
    kws = prdwgt.wgts
    dsc = 0.9 if kws[0].karat == 925 or (kws[1] and kws[1].karat == 925) else 0.85
    cn = refup / dsc + incr.gold * (tarmps.gold - refmps.gold) * 1.25 \
        + incr.silver * (tarmps.silver - refmps.silver) * 1.25 ;
    return PajChina(cn,dsc,incr,tarmps)
    

def _calctarget(cn,tarmps):
    """calculate the target unit price based on the data _NotProvided
    @param cn: the PAJChina cost
    @param tarmps: the target MPS      
    """
    r0 = cn.china + (tarmps.gold - cn.mps.gold) * cn.increment.gold * 1.25 \
        + (tarmps.silver - cn.mps.silver) * cn.increment.silver * 1.25;
    return round(r0 * cn.discount,2);

def _calcmtlcost(prdwgt,mps,lossrate,vendor):
    
    kws = [x for x in prdwgt.wgts if(x and x.wgt > 0)]
    r0 = 0;
    for x in kws:
        r0 += x.wgt * _getfiness(x.karat, vendor) * lossrate * \
            (mps.silver if x.karat == 925 else 0 if x.karat == 200 else mps.gold) / 31.1035        
    return r0