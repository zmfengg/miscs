#! coding=utf-8
'''
* @Author: zmFeng
* @Date: 2018-06-16 14:41:00
* @Last Modified by:   zmFeng
* @Last Modified time: 2018-06-16 14:41:00
'''

import datetime
import logging
import random
from functools import cmp_to_key
from os import listdir, path
from unittest import TestCase, main

from sqlalchemy import VARCHAR, Column, ForeignKey, Integer
from sqlalchemy.engine import create_engine
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import relationship
from xlwings.constants import LookAt

from utilz import imagesize, karatsvc, stsizefmt, xwu
from utilz._jewelry import RingSizeSvc
from utilz._miscs import (NamedList, NamedLists, appathsep, getfiles,
                          list2dict, lvst_dist)
from utilz.resourcemgr import ResourceCtx, ResourceMgr, SessionMgr

from .main import logger, thispath
from .._catalog import Page, PageSetting

#resfldr = appathsep(appathsep(thispath) + "res")
resfldr = path.join(thispath, "res")
class KeySuite(TestCase):
    """ key tests for this util funcs """

    def testResourceMgr(self):
        """ test for Resource manager, which should be invoked by context manager
        """
        class Res(object):
            """ simple resource class that dump action to logger for the manager"""
            def __init__(self, name):
                self.name = name
                logger.debug("resource %s created", name)

            def close(self):
                logger.debug("%s disposed", self.name)

            def run(self):
                """ simulate producer method that return value for the caller """
                logger.debug("Your are making use of resource provided by(%s)", self.name)
                return self.name

        def _newres():
            return Res(str(random.randint(1, 9999)))

        def _dispose(res):
            res.close()

        mgr = ResourceMgr(_newres, _dispose)
        with ResourceCtx(mgr) as r:
            with ResourceCtx(mgr) as r1:
                self.assertTrue(r == r1, "double fetch, new resource won't be return")
                self.assertTrue(r.run() == r.name, "Yes, it's the object expected")

        mgr1 = ResourceMgr(_newres, _dispose)
        with ResourceCtx([mgr, mgr1]) as r:
            r[0].run()
            r[1].run()

    def testStsize(self):
        """ test for stone size parser
        """
        self.assertEqual("N/A", stsizefmt("N/A"), "Not a valid stone size")
        self.assertEqual("N/A", stsizefmt("n/a"), "Not a valid stone size")
        self.assertEqual("0300", stsizefmt("3tk"), "Not a valid stone size")
        self.assertEqual("0500X0400X0300", stsizefmt("3x4x5mm"), "Size format")
        self.assertEqual("0500X0400X0300", stsizefmt("3x4x5"), "Size format")
        self.assertEqual("0530X0400X0350", stsizefmt("3.5x4.0x5.3"), "Size format")
        self.assertEqual("0400", stsizefmt("4"), "Size format")
        self.assertEqual("0530X0400X0350", stsizefmt("053004000350"), "Size format")
        self.assertEqual("0530X0400X0350", stsizefmt("040005300350"), "Size format")
        self.assertEqual("0530X0400X0350", stsizefmt("0400X0530X0350"), "Size format")
        self.assertEqual("0400", stsizefmt("4m"), "Size format")
        self.assertEqual("0400-0350", stsizefmt("4m-3.5m"), "Size format")
        self.assertEqual("5X4X3", stsizefmt("3x4x5", True), "Size format")
        self.assertEqual("5X4X3", stsizefmt("0500X0400X0300", True), "Size format")
        self.assertEqual("5X4X3", stsizefmt("0300X0500X0400", True), "Size format")

    def testList2Dict(self):
        """ test for _misc.list2Dict, maybe don't need to use this function separated,
        try to use NamedList instead
        """
        lst, alias = ("A", None, "", "bam", "Bam1"), {"namE": "A", "description": "b,am"}
        mp = list2dict(lst, alias=alias)
        self.assertEqual(0, mp.get("name"))
        self.assertEqual(3, mp.get("description"))
        self.assertEqual(1, mp.get(""))
        self.assertEqual(2, mp.get("1"))
        mp = list2dict(lst)
        self.assertEqual(0, mp.get("a"))
        self.assertEqual(1, mp.get(""))
        self.assertEqual(2, mp.get("1"))
        mp = list2dict(lst, alias=alias, div="_")
        self.assertEqual(0, mp.get("name"))
        self.assertEqual(1, mp.get(""))
        self.assertEqual(2, mp.get("_1"))
        # now try the non-normalized case
        mp = list2dict(lst, alias=alias, div="_", normalize=None)
        self.assertEqual(0, mp.get("namE"))
        self.assertEqual(2, mp.get("_1"))
        self.assertEqual(4, mp.get("Bam1"))

    def testNamedList(self):
        """ test for NamedList class, a good to use class for sequence data access by name type
        """
        lsts = (["Name", "group", "age"],
                ["Peter", "Admin", 30],
                ["Watson", "Admin", 45],
                ["Biz", "Mail", 20]
                )
        nl = NamedLists(lsts)
        lsts1 = [x for x in nl]
        self.assertEqual(3, len(lsts1), "len of list")
        idx = -1
        for x in lsts1:
            idx += 1
            if idx == 0:
                self.assertEqual("Peter", x.name, "the name property")
                self.assertEqual("Admin", x.group, "the group property")
                self.assertEqual(30, x.age, "the age property")
        nl = lsts1[0]
        self.assertTrue(hasattr(nl, "name"), "response to hasattr")
        self.assertEqual(lsts[1], nl.data, "title off, first row of data")
        self.assertEqual(lsts[1][0], nl[0], "access by index")
        self.assertTrue("name" in nl, "supports the in operation")
        sl = slice(1, None)
        self.assertEqual(lsts[1][sl], nl[sl], "access by slice")

        # now try the setter
        nl.name = "FF"
        self.assertEqual("FF", lsts[1][0], "They representing the same object")
        self.assertEqual("FF", nl.name, "They representing the same object")
        nl["name"] = "JJ"
        self.assertEqual("JJ", lsts[1][0], "They representing the same object")
        self.assertEqual("JJ", nl.name, "They representing the same object")

        # the column name<->id translate
        self.assertEqual(0, nl.getcol("Name"))
        self.assertEqual(0, nl.getcol("name "))
        self.assertEqual(2, nl.getcol("age"))
        self.assertTrue(nl.getcol("age_") is None)
        self.assertEqual(tuple("name,group,age".split(",")), nl.colnames)
        self.assertEqual((0, 1, 2), nl.colids)

        # a smatter usage, use the nl to wrap a list
        nl.setdata(lsts[3])
        self.assertEqual(lsts[3][0], nl.name, "NamedList wrapping a list")

        idx = 0
        for x in NamedLists(lsts, newinst=False):
            idx += 1
            self.assertEqual(x.data, lsts[idx], "same object, newinst=False don't affect iterator")
        lsts1 = [x for x in NamedLists(lsts, newinst=False)]
        self.assertEqual(lsts[3], lsts1[0].data, "newinst=False affect list()")
        self.assertEqual(lsts[3], lsts1[1].data, "newinst=False affect list()")
        self.assertEqual(lsts[3], lsts1[2].data, "newinst=False affect list()")

        lsts = (["Name", "group", "age"],
                ["Peter", "Admin", 30]
                )
        # now just a minor test on the normalize function, make sure non-normalized works
        nls = [x for x in NamedLists(lsts, normalize=None)]
        nl = nls[0]
        self.assertEqual("Peter", nl.get("Name"))
        self.assertTrue("Name" in nl.colnames)

        # now namedlist treating normal object, There is an object NamedList before
        # but finally merged into NamedList
        class Bean(object):
            """ sample bean like object, NamedList get its property directly """
            name, id, age = "Hello", 0, 0
        al = NamedList({"nick": "name"})
        it = Bean()
        al.setdata(it)
        # the getter
        self.assertEqual(it.name, al.nick, "One object, 2 name or more")
        self.assertEqual(it, al.data, "return the object")
        # the setter
        al.nick = "WXXX"
        self.assertEqual("WXXX", it.name)

        d0 = {"name": "David", "Age": 20}
        al = NamedList({"nick": "name"}, d0)
        self.assertEqual(al["name"], al["nick"])
        self.assertTrue("nick" not in d0)
        self.assertEqual(al.name, al.nick)

        nl = NamedList("id,name,agex,agey", [None]*4)
        nl.id, nl.agex = 1, 30
        self.assertEqual(1, nl.id)
        nl = nl._replace({"idx": "id,", "age": "agex"})
        self.assertEqual(1, nl.idx)
        self.assertEqual(30, nl.age)
        self.assertTrue("id" not in nl.colnames)
        self.assertEqual(2, nl.getcol("age"))

    def testAppathSep(self):
        """ tes for appathsep, early stage function of my python programming,
        it should be replaced by path.join """
        fldr = thispath
        self.assertTrue(fldr[-1] != path.sep, "a path's name should not ends with path.sep")
        fldr = appathsep(fldr)
        self.assertTrue(fldr[-1] == path.sep, "with path.sep appended")

    def testGetFiles(self):
        """ test for _misc.getfiles, a early stage funtion of my python programming """
        fldr = path.join(thispath, "res")
        fns = getfiles(fldr, "NamedL", True)
        self.assertEqual("NamedList.xlsx", fns[0], "the only excel file there")
        fns = getfiles(fldr, "List")
        self.assertEqual(appathsep(fldr) + "NamedList.xlsx", fns[0], "the only excel file there")
        fns = getfiles(fldr, nameonly=True)
        fnx = listdir(fldr)
        self.assertEqual(len(fnx), len(fns), "the count of files")
        fns = set(iter(fns))
        self.assertTrue(u"厉害为国為幗.txt" in fns, "utf-8 based system can return mixing charset")

    def testKaratSvc(self):
        """ the test for karat service """
        ks = karatsvc
        k0 = ks[9]
        k1 = ks["9K"]
        self.assertEqual(k0, k1, "same object return from byId/byName")
        k1 = ks["9KR"]
        self.assertEqual(k0.fineness, k1.fineness, "same fineness, different karat")
        k1 = ks.getfamily(k1)
        self.assertEqual(k0, k1, "9KR's family is 9K")
        self.assertTrue(ks.issamecategory(9, 91), "9K and 9KW are all gold")
        self.assertTrue(ks.issamecategory(9, "9KW"), "9K and 9KW are all gold")
        self.assertFalse(ks.issamecategory(9, 200), "gold is not bronze")
        self.assertTrue(ks.compare(k0, k0) == 0, "the same karat")
        self.assertTrue(ks.compare(k0, ks[200]) > 0, "Gold is larger than bronze")
        self.assertTrue(ks.compare(k0, ks[91]) < 0, "9K is smaller than 9KR")
        lst = [ks[9], ks[18], ks[200], ks[925]]
        lst = sorted(lst, key=cmp_to_key(ks.compare))
        self.assertEqual(ks[200], lst[0], "sort method")
        self.assertEqual(ks[925], lst[1], "sort method")
        self.assertEqual(ks[18], lst[-1], "sort method")

    def testRingSizeCvt(self):
        """ a size converter, maybe should be migrated to UOMConverter """
        rgsvc = RingSizeSvc()
        self.assertEqual("M", rgsvc.convert("US", "6", "UK"), "US#6 = UK#M")
        self.assertEqual("M", rgsvc.convert("US", "6", "AU"), "US#6 = UK#M, AU using UK")
        self.assertEqual("4 1/4", rgsvc.convert("EU", "47", "US"), "EU#47 = US#4 1/4")
        self.assertTrue(rgsvc.convert("EU", "A", "US") is None, "EU#A does not exist")
        self.assertAlmostEqual(47.0, rgsvc.getcirc("US", "4 1/4"), "the circumference of US#4 1/4 is 47.0mm")

    def testImagesize(self):
        """ the imagesize function(power by PIL) """
        fns = getfiles(path.join(thispath, "res"), "65x27")
        for fn in fns:
            self.assertEqual((65, 27), imagesize(fn), "the size of %s" % fn)
        # one special, the SOF C4 is used
        self.assertEqual((849, 826), imagesize(path.join(thispath, r"res\579616.jpg")))

    def testLvDist(self):
        """
        test the LevenshteinDistance function
        """
        self.assertEqual(0, lvst_dist("I'm", "I'm"), "same string")
        self.assertEqual(1, lvst_dist("I'mx", "I'm"), "same string")
        self.assertEqual(2, lvst_dist("'mI", "I'm"), "same string")


class XwuSuite(TestCase):
    """
    test suit for xwu funcitons
    """
    _hasxls = None
    _app, _tk = (None, ) * 2

    def setUp(self):
        """ self init, including app, tk """
        if self._hasxls is not None:
            return
        try:
            self._app, self._tk = xwu.appmgr.acq()
            self._hasxls = True
        except:
            self._app, self._tk = (None,) * 2
            self._hasxls = False
        if not self._hasxls:
            logger.debug("No excel is available")

    def tearDown(self):
        if self._hasxls:
            xwu.appmgr.ret(self._tk)
        return super().tearDown()

    @classmethod
    def tearDownClass(cls):
        if cls._hasxls:
            xwu.appmgr.ret(cls._tk)

    def fail_noexcel(self):
        """ raise error when no excel is found """
        self.fail("no excel was available, Pls. install one")

    def testappmgr(self):
        """ test for xwu.appmgr property """
        if not self._hasxls:
            self.fail_noexcel()
            return
        # xlwings.apps.count is not reliable, don't need to test

    def testxwuappswitch(self):
        """ test for xwu.appswitch function """
        if not self._hasxls:
            self.fail_noexcel()
            return
        app = self._app
        xwu.appswitch(app, True)
        os = xwu.appswitch(app)
        self.assertFalse(bool(os), "no changes need to be made")
        os = xwu.appswitch(app, False)
        self.assertTrue(len(os) > 0)
        xwu.appswitch(app, os)
        app.visible = False
        os = xwu.appswitch(app, {"visible": True})
        self.assertFalse(os["visible"])
        app.api.enableevents = True
        os = xwu.appswitch(app, {"visible": True, "enableevents": False})
        self.assertEqual(1, len(os))
        self.assertTrue(os["enableevents"])

    def testEscapettl(self):
        """ extract title data out from a excel's sheet header """
        ttls = ('2017&"宋体,Regular"年&"Arial,Regular"6&"宋体,Regular"月', '2017&"宋体,Regular"年&"Arial,Regular"&6 6&"宋体,Regular"月')
        exps = ("2017年6月", "2017年6月")
        for idx, it in enumerate(ttls):
            self.assertEqual(exps[idx], xwu.escapetitle(it), "the title")

    def testNamedList(self):
        """
        # now test a very often use ability, read data from (excel) and handle it
        # now think of use NamedRanges, better ability to detect even if there is merged range
        # without NamedList(s), I have to use tr[map[name]] to get the value
        """
        if not self._hasxls:
            self.fail_noexcel()
            return
        fn = getfiles(resfldr, "NamedList")[0]
        app = self._app
        wb = app.books.open(fn)
        sht = wb.sheets[0]
        rng = xwu.find(sht, "*Table")
        rng = rng.offset(1, 0)
        rng = rng.expand("table")
        lst = [x for x in NamedLists(rng.value)]
        self.assertEqual(8, len(lst), "the count of data")
        self.assertEqual(1, lst[0].id, "the id of first Emp")
        # try build a dict
        lst = dict((x.id, x) for x in NamedLists(rng.value))
        self.assertEqual("Name 8", lst[8].name, "The Name")
        # now try an named-translation
        nl = NamedLists(rng.value, {"Edate": "enter,"})
        emp = nl.__next__()
        self.assertEqual(datetime.datetime(1998, 1, 3, 0, 0), emp["edate"], "get date use translated name")

        # test the find's all function
        nl = xwu.find(sht, "Name", lookat=LookAt.xlPart, find_all=True)
        self.assertEqual(9, len(nl), "the are 9 items has name as part")

    def testDetectBorder(self):
        """ check the detect border function of xwu """
        if not self._hasxls:
            self.fail_noexcel()
            return
        app = self._app
        wb = app.books.open(path.join(thispath, "res", "getTableData.xlsx"))
        sht = wb.sheets["borderdect"]
        rng = xwu.find(sht, 1)
        rng = xwu.detectborder(rng)
        self.assertEqual("$B$2:$F$8", rng.address, "very regular region")
        rng = xwu.detectborder(xwu.find(sht, 2))
        self.assertEqual("$B$12:$G$19", rng.address, "mal-form shape")

    def testGetTableData(self):
        """ the gettabledata function under different conditions """
        if not self._hasxls:
            self.fail_noexcel()
            return
        app = self._app
        wb = app.books.open(path.join(thispath, "res", "getTableData.xlsx"))
        nmap = {"id": "id,", "9k": "9k,", "S950": "S950,"}
        sht = wb.sheets["gettabledata"]

        s0 = "No Merge,FirstMerge,NonFirstmerge,3Rows"
        #s0 = "FirstMerge,NonFirstmerge,3Rows"
        import time
        for name in s0.split(","):
            t0 = time.clock()
            rng = xwu.find(sht, name)
            nls = [x for x in xwu.NamedRanges(rng, skip_first_row=True, name_map=nmap)]
            #print("%s colnames are:(%s)" % (name, list(nls[0].colnames)))
            self.assertEqual(3, len(nls), "result count of %s" % name)
            self.assertEqual(2, nls[0]["9k"], "9K result of %s" % name)
            self.assertEqual(16, nls[2].s950, "S950 of %s" % name)
            print("using %f ms to perform %s" % (time.clock() - t0, name))
        # try a blank range, should return none
        nls = xwu.NamedRanges(sht.range(1000, 1000))
        self.assertIsNone(nls, "Nothing should be returned")


BaseClass = declarative_base()


class Mstr(BaseClass):
    """ master item """
    __tablename__ = "mstr"
    id = Column(Integer, primary_key=True)
    name = Column(VARCHAR(20))

    def __init__(self, name):
        self.name = name


class Dtl(BaseClass):
    """ detail item of the Master Item """
    __tablename__ = "dtl"

    def __init__(self, mstr):
        self.mstr = mstr
    id = Column(Integer, primary_key=True)
    pid = Column(Integer, ForeignKey('mstr.id'))
    mstr = relationship("Mstr")


class PKNAC(BaseClass):
    """ primary key not auto-increased test table """
    __tablename__ = "pknac"

    def __init__(self, id_):
        self.id = id_
    id = Column(Integer, primary_key=True, autoincrement=False)


class SessMgrSuite(TestCase):
    """ check if a sessionmgr will automatically rollback a transaction """
    #engine = create_engine('sqlite:///:memory:', echo=True)
    engine = create_engine('sqlite:///:memory:')
    Mstr.metadata.create_all(engine)
    sessmgr = SessionMgr(engine)

    def setup(self):
        """ class setup """
        logging.getLogger("sqlalchemy").setLevel(logging.DEBUG)

    @property
    def sessctx(self):
        """ section for with statement """
        return ResourceCtx(self.sessmgr)

    @property
    def newmstr(self):
        """ factory method for creating Mstr instances """
        return Mstr("fx")

    def testRollback(self):
        """ try the sqlalchemy's rollback function. I found in realtime app,
        it works sometimes only
        """
        mid = 0
        # by default, the session is not auto-commit, so it's rollbacked while exist
        with self.sessctx as cur:
            mstr = self.newmstr
            dtl = Dtl(mstr)
            cur.add(mstr)
            cur.add(dtl)
            cur.flush()
            mid = mstr.id
            did = dtl.id
        with self.sessctx as cur:
            mstr = cur.query(Mstr).filter(Mstr.id == mid).first()
            self.assertFalse(mstr, "The item should not be inserted")
            mstr = cur.query(Mstr).all()
            self.assertFalse(mstr, "There should be nothing in the db")
            dtl = cur.query(Dtl).filter(Dtl.id == did).first()
            self.assertFalse(dtl, "The item should not be inserted")

    def testCommit(self):
        """ test for sqlalchemy's commit function. it behaves as rollback(),
        it can pass this test, but in realtime app, sometimes goes wrong"""
        mid = 0
        with self.sessctx as cur:
            mstr = self.newmstr
            dtl = Dtl(mstr)
            cur.add(mstr)
            cur.add(dtl)
            cur.flush()
            mid, did = mstr.id, dtl.id
            cur.commit()
        with self.sessctx as cur:
            mstr = cur.query(Mstr).filter(Mstr.id == mid).first()
            self.assertEqual(self.newmstr.name, mstr.name, "Committed")
            mstr = cur.query(Mstr).all()
            self.assertEqual(1, len(mstr), "There should be only one item inside")
            dtl = cur.query(Dtl).all()
            self.assertEqual(1, len(dtl), "The count of detail")
            self.assertEqual(did, dtl[0].id)
            cur.delete(mstr[0])
            cur.commit()
            mstr = cur.query(Mstr).all()
            self.assertFalse(mstr, "nothing in the db")

    def testPkNotAutoInc(self):
        """ test for auto-increasement primary key, which need flush function """
        # is the non-autoincreased primary key object persistable?
        with self.sessctx as cur:
            pk = PKNAC(1)
            cur.add(pk)
            cur.flush()
            pk1 = cur.query(PKNAC).filter(PKNAC.id == pk.id).first()
            self.assertEqual(pk.id, pk1.id)
    
class CatalogTest(TestCase):
    """ class for catalog making """"
    def test_detectMargin(self):
        """
        the ability to test for height/width of a page can hold 
        """
        pass
        


if __name__ == "__main__":
    main()
