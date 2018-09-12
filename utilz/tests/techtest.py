#! coding=utf-8 
'''
* @Author: zmFeng 
* @Date: 2018-07-04 08:46:52 
* @Last Modified by:   zmFeng 
* @Last Modified time: 2018-07-04 08:46:52 
for python's language/basic facility test, a practice farm
'''

import unittest
from unittest import TestCase
import re
from os import path
import os
import shutil
from utilz import getfiles

class TechTests(TestCase):
    def testRE(self):
        ptn = re.compile(r"C(\d{1})")
        s0 = "JMP12C1"
        mt = ptn.match(s0)
        self.assertFalse(bool(mt),"There should be no match")
        mt = ptn.search(s0)
        self.assertTrue(bool(mt),"There should be search")
        self.assertEqual(("1",),mt.groups(),"The so-call zero group")
        self.assertEqual("1",mt.group(1),"The so-call first group")
    
    def testSeveralForItr(self):
        rngs = ((1,3),(4,9))
        slots = [x for y in rngs for x in range(y[0],y[1])]
        #print(slots)
        self.assertEqual(7,len(slots))
        self.assertEqual(1,slots[0])
    
    def testFuncArgs(self):
        def sth0(a):
            return(a,)
        def sth1(a,*args):
            return (a,args)            
        def sth2(a,**kwds):
            return (a,kwds)
        def sth3(a, *args, **kwds):
            return (a,args, kwds)

        arr = sth0(5)
        self.assertTupleEqual((5,),arr, "single argument")
        arr = sth1(5,1,2,3)
        self.assertTupleEqual((5,(1,2,3)),arr, "single + positional argument")
        arr = sth2(5, nice = "to")
        self.assertTupleEqual((5,{"nice":"to"}),arr, "single + named argument")
        #this 2 argument error
        with self.assertRaises(TypeError):
            sth1(5, nice = "to")
        with self.assertRaises(TypeError):
            sth2(5, 1,2,3)
        #a full-blow
        arr = sth3(5,20, nice = "to")
        self.assertTupleEqual((5,(20,),{"nice":"to"}), arr)
        