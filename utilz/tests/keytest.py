#! coding=utf-8 
'''
* @Author: zmFeng 
* @Date: 2018-06-16 14:41:00 
* @Last Modified by:   zmFeng 
* @Last Modified time: 2018-06-16 14:41:00 
'''

import random
import unittest
from os import path
from unittest import TestCase

from utilz import xwu
from utilz._miscs import (NamedList, NamedLists, appathsep, getfiles,Alias,list2dict, stsizefmt)
from utilz.resourcemgr import ResourceCtx, ResourceMgr

from . import logger, thispath
import datetime


class KeyTests(TestCase):
    """ key tests for this util funcs """

    def testResourceMgr(self):
        class A(object):
            def __init__(self,name):
                self.name = name
                logger.debug("resource %s created" % name)

            def _close(self):
                logger.debug("%s disposed" % self.name)

            def run(self):
                logger.debug("Your are making use of resource provided by(%s)" % self.name)
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

    @property
    def resfldr(self):
        return appathsep(appathsep(thispath) + "res")

    def testStsize(self):
        self.assertEqual("0500X0400X0300",stsizefmt("3x4x5mm") ,"Size format")
        self.assertEqual("0500X0400X0300",stsizefmt("3x4x5") ,"Size format")
        self.assertEqual("0530X0400X0350",stsizefmt("3.5x4.0x5.3") ,"Size format")
        self.assertEqual("0400",stsizefmt("4") ,"Size format")
        self.assertEqual("0530X0400X0350",stsizefmt("053004000350") ,"Size format")
        self.assertEqual("0530X0400X0350",stsizefmt("040005300350") ,"Size format")
        self.assertEqual("0530X0400X0350",stsizefmt("0400X0530X0350") ,"Size format")
        self.assertEqual("0400",stsizefmt("4m") ,"Size format")
        self.assertEqual("0400-0350",stsizefmt("4m-3.5m") ,"Size format")
        self.assertEqual("5X4X3",stsizefmt("3x4x5", True) ,"Size format")
        self.assertEqual("5X4X3",stsizefmt("0500X0400X0300", True) ,"Size format")
        self.assertEqual("5X4X3",stsizefmt("0300X0500X0400", True) ,"Size format")
    
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

        #now test a very often use ability, read data from (excel) and handle it
        #without NamedList(s), I have to use tr[map[name]] to get the value
        fn = getfiles(self.resfldr,"xls")[0]
        app = xwu.app(False)[1]
        try:
            wb = app.books.open(fn)
            sht = wb.sheets[0]
            rng = xwu.find(sht,"*Table")
            rng = rng.offset(1,0)
            rng = rng.expand("table")
            lst = [x for x in NamedLists(rng.value)]
            self.assertEqual(8,len(lst), "the count of data")
            self.assertEqual(1,lst[0].id, "the id of first Emp")
            #try build a dict
            lst = dict((x.id,x) for x in NamedLists(rng.value))
            self.assertEqual("Name 8",lst[8].name,"The Name")
            #now try an named-translation
            nl = NamedLists(rng.value,{"enter,":"Edate"})
            emp = nl.__next__()
            self.assertEqual(datetime.datetime(1998,1,3,0,0), emp["edate"],"get date use translated name")
        finally:
            if app: app.quit()
                
    def testAppathSep(self):
        fldr = thispath
        self.assertTrue(fldr[-1] != path.sep, "a path's name should not ends with path.sep")
        fldr = appathsep(fldr)
        self.assertTrue(fldr[-1] == path.sep, "with path.sep appended")

    def testGetFiles(self):
        fldr = appathsep(thispath) + "res"
        fns = getfiles(fldr,"xls",True)
        self.assertEqual("NamedList.xlsx",fns[0],"the only excel file there")
        fns = getfiles(fldr,"xls")
        self.assertEqual(appathsep(fldr) + "NamedList.xlsx",fns[0],"the only excel file there")
        fns = getfiles(fldr,nameonly = True)
        self.assertEqual(2,len(fns),"the count of files")
        fns = set([x for x in fns])
        self.assertTrue(u"厉害为国為幗.txt" in fns, "utf-8 based system can return mixing charset")

    def testAlias(self):
        class A(object):
            name,id,age = None,0,0
        al = Alias({"name":"nick"})
        it = A()
        al.settarget(it) 
        self.assertEqual(it.name, al.nick,"One object, 2 name or more")
        self.assertEqual(it, al.gettarget(),"return the object")