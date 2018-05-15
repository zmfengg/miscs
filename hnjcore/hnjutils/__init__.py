'''
Created on Apr 19, 2018

@author: zmFeng
'''

from hnjutils import p17u
from . import odbctpl
from hnjutils import xwu
import math
from os import path

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