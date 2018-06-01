# coding=utf-8
'''
Created on 2018-05-18

@author: zmFeng
'''

from os import path
import math

_silveralphas = set(("4", "5"))

def splitarray(arr, logsize = 100):
    """split an array into arrays whose len is less or equal than logsize
    @param arr: the sequence object that need to split
    @param logsize: len of each sub-array's size  
    """    
    if not arr: return
    if not logsize: logsize = 100 
    return (arr[x * logsize:(x + 1) * logsize] for x in range(int(math.ceil(1.0 * len(arr) / logsize))))

def appathsep(fldr):
    """append a path sep into given path if there is not"""
    return fldr + path.sep if fldr[len(fldr) - 1:] != path.sep else fldr

def deepget(obj,names):
    """ get deeply from the object """
    rc = None
    for k in names.split("."):
        rc = rc.__getattribute__(
            k) if rc else obj.__getattribute__(k)
    return rc

def samekarat(srcje, tarje):
    """ detect if the given 2 JOElement are of the same karat """
    if not (srcje and tarje): return
    return srcje.alpha == tarje.alpha or (srcje.alpha in _silveralphas and tarje.alpha in _silveralphas)