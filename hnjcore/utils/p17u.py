# coding=utf-8 
'''
Created on Apr 19, 2018

@author: zmFeng
'''


def isvalidp17(p17):
    """ check if the given p17 code is a valid one """
    return isinstance(p17,str) and len(p17) == 17 and "0,1,2,3,4,9,C,P,W".find(p17[1:2]) >= 0