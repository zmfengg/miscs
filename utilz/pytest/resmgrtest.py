#! coding=utf-8 
'''
* @Author: zmFeng 
* @Date: 2018-06-16 14:41:00 
* @Last Modified by:   zmFeng 
* @Last Modified time: 2018-06-16 14:41:00 
'''

from .. import ResourceMgr, ResourceCtx
import unittest
from unittest import TestCase

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
        lsts = [("name","group","age"),
            ("Peter","Admin",30),
            ("Watson","Admin",45),
            ("Biz","Mail",20)
        ]
        nm = list2dict(lsts[0])
        nl = NameList(lsts[1:],nm)
        lsts1 = [x for x in nl]
        self.assertEqual(3,lsts1,"len of list")
        idx = -1
        for x in nl:
            idx += 1
            if idx == 0:
                self.assertEqual("Peter",x.name,"the name property")
                self.assertEqual("Admin",x.group,"the group property")
                self.assertEqual(30,x.age,"the age property")
        self.assertEqual(2,idx, "the count of iterator")