#! coding=utf-8 
'''
* @Author: zmFeng 
* @Date: 2018-06-16 14:41:00 
* @Last Modified by:   zmFeng 
* @Last Modified time: 2018-06-16 14:41:00 
'''

import unittest
from unittest import TestCase
import random

from utilz.resourcemgr import ResourceCtx, ResourceMgr
from utilz._miscs import NamedList, NamedLists
from utilz.xwu import list2dict


class ResourceMgrTest(TestCase):
    def testResourceMgr(self):
        class A(object):
            def __init__(self,name):
                self.name = name
                print("resource %s created" % name)

            def _close(self):
                print("%s disposed" % self.name)

            def run(self):
                print("Your are making use of resource provided by(%s)" % self.name)
                return self.name
            
        def _newres():
            return A(str(random.randint(1,9999)))

        def _dispose(a):
            a._close()
        
        mgr = ResourceMgr(_newres,_dispose)
        with ResourceCtx(mgr) as r:
            with ResourceCtx(mgr) as r1:
                self.assertTrue(r == r1,"double fetch, new resource won't be return")
                self.assertTrue(r.run() == r.name,"Yes, it's the object expected")

        mgr1 = ResourceMgr(_newres,_dispose)
        with ResourceCtx([mgr,mgr1]) as r:
            r[0].run()
            r[1].run()

    def testNamedList(self):
        lsts = (["name","group","age"],
            ["Peter","Admin",30],
            ["Watson","Admin",45],
            ["Biz","Mail",20]
        )
        nl = NamedLists(lsts)
        lsts1 = [x for x in nl]
        self.assertEqual(3,len(lsts1),"len of list")
        idx = -1
        for x in lsts1:
            idx += 1
            if idx == 0:
                self.assertEqual("Peter",x.name,"the name property")
                self.assertEqual("Admin",x.group,"the group property")
                self.assertEqual(30,x.age,"the age property")
        nl = lsts1[0]
        self.assertEqual(lsts[1], nl.data,"title off, first row of data")
        self.assertEqual(lsts[1][0], nl[0],"access by index")
        sl = slice(1,None)
        self.assertEqual(lsts[1][sl], nl[sl],"access by slice")

        #now try the setter
        nl.name = "FF"
        self.assertEqual("FF",lsts[1][0],"They representing the same object")
        self.assertEqual("FF",nl.name,"They representing the same object")
        nl["name"] = "JJ"
        self.assertEqual("JJ",lsts[1][0],"They representing the same object")
        self.assertEqual("JJ",nl.name,"They representing the same object")

        #a smatter usage, use the nl to wrap a list
        nl.setdata(lsts[3])
        self.assertEqual(lsts[3][0],nl.name,"NamedList wrapping a list")

        idx = 0
        for x in NamedLists(lsts, newinst = False):
            idx += 1
            self.assertEqual(x.data,lsts[idx],"same object, newinst=False don't affect iterator")
        lsts1 = [x for x in NamedLists(lsts,newinst = False)]
        self.assertEqual(lsts[3],lsts1[0].data,"newinst=False affect list()")
        self.assertEqual(lsts[3],lsts1[1].data,"newinst=False affect list()")
        self.assertEqual(lsts[3],lsts1[2].data,"newinst=False affect list()")

