#! coding=utf-8
'''
* @Author: zmFeng
* @Date: 2018-06-16 14:41:00
* @Last Modified by:   zmFeng
* @Last Modified time: 2018-06-16 14:41:00
'''

import logging
import random
from datetime import date, datetime
from time import clock
from functools import cmp_to_key
from os import listdir, path, remove
from tempfile import gettempdir
from unittest import TestCase, main
from cProfile import Profile
from pstats import Stats
from io import StringIO, FileIO

from sqlalchemy import VARCHAR, Column, ForeignKey, Integer
from sqlalchemy.engine import create_engine
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import relationship
from xlwings.constants import LookAt

from utilz import getvalue, imagesize, iswritable, karatsvc, stsizefmt, xwu
from utilz._jewelry import RingSizeSvc
from utilz._miscs import (Config, NamedList, NamedLists, Salt, appathsep,
                          daterange, getfiles, list2dict, lvst_dist, monthadd,
                          shellopen, Number2Word)
from utilz.resourcemgr import ResourceCtx, ResourceMgr, SessionMgr
from utilz.exp import Exp, AbsResolver

from .main import logger, thispath
from time import clock
from itertools import product
from inspect import currentframe, getfile, getabsfile

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
                """ close the res """
                logger.debug("%s disposed", self.name)

            def run(self):
                """ simulate producer method that return value for the caller """
                logger.debug("Your are making use of resource provided by(%s)",
                             self.name)
                return self.name

        def _newres():
            return Res(str(random.randint(1, 9999)))

        def _dispose(res):
            res.close()

        mgr = ResourceMgr(_newres, _dispose)
        with ResourceCtx(mgr) as r:
            with ResourceCtx(mgr) as r1:
                self.assertTrue(r == r1,
                                "double fetch, new resource won't be return")
                self.assertTrue(r.run() == r.name,
                                "Yes, it's the object expected")

        # multi resources into one resource context
        mgr1 = ResourceMgr(_newres, _dispose)
        with ResourceCtx([mgr, mgr1]) as r:
            r[0].run()
            r[1].run()

        # Resource is None
        with ResourceCtx(None) as cur:
            self.assertTrue(cur is None)

        with ResourceCtx((None, None)) as curs:
            self.assertEqual(2, len(curs))
            self.assertTrue(curs[0] is None)
            self.assertTrue(curs[1] is None)

    def testShellOpen(self):
        '''
        test the shellopen function
        '''
        from sys import platform
        if platform.startswith("win"):
            root = path.join(thispath, "res")
            fns = listdir(root)
            self.assertEqual(1, shellopen([path.join(root, x) for x in fns]))
            self.assertEqual(2, shellopen([path.join(root, x) for x in fns[:2]], True))
            # maybe I should close the apps launched by shellopen
        else:
            self.assertEqual(-1, shellopen(thispath))
        del platform

    def testWritable(self):
        '''
        the iswritable test
        '''
        self.assertFalse(iswritable(None))
        # create a file in current folder(os.getcwd()), not always true
        # self.assertTrue(iswritable("abcdef"))
        # drive not exist
        self.assertFalse(iswritable(r"a:\bde"))
        tf = gettempdir()
        self.assertTrue(iswritable(tf))
        # folder not exist
        self.assertFalse(iswritable(path.join(tf, *"a b c d e".split())))
        fn = path.join(tf, str(random.randint(0, 99999)))
        try:
            with open(fn, "wt") as fp:
                fp.writelines("a b c d e".split())
            self.assertTrue(iswritable(fn))
        finally:
            remove(fn)

    def testStsize(self):
        """ test for stone size parser
        """
        self.assertEqual(None, stsizefmt(None), "An invalid size")
        self.assertEqual(None, stsizefmt("."), "An invalid size")
        self.assertEqual("N/A", stsizefmt("N/A"), "Not a valid stone size")
        self.assertEqual("N/A", stsizefmt("n/a"), "Not a valid stone size")
        self.assertEqual("0300", stsizefmt("3tk"), "Not a valid stone size")
        self.assertEqual("0500X0400X0300", stsizefmt("3x4x5mm"), "Size format")
        self.assertEqual("0500X0400X0300", stsizefmt("3x4x5"), "Size format")
        self.assertEqual("0530X0400X0350", stsizefmt("3.5x4.0x5.3"),
                        "Size format")
        self.assertEqual("0400", stsizefmt("4"), "Size format")
        self.assertEqual("0530X0400X0350", stsizefmt("053004000350"),
                        "Size format")
        self.assertEqual("0530X0400X0350", stsizefmt("040005300350"),
                        "Size format")
        self.assertEqual("0530X0400X0350", stsizefmt("0400X0530X0350"),
                        "Size format")
        self.assertEqual("0400", stsizefmt("4m"), "Size format")
        self.assertEqual("0400-0350", stsizefmt("4m-3.5m"), "Size format")
        self.assertEqual("5X4X3", stsizefmt("3x4x5", True), "Size format")
        self.assertEqual("5X4X3", stsizefmt("0500X0400X0300", True),
                        "Size format")
        self.assertEqual("5X4X3", stsizefmt("0300X0500X0400", True),
                        "Size format")
        self.assertEqual("1-0", stsizefmt("0~1", True))
        self.assertEqual("1", stsizefmt("1.0", True))
        self.assertEqual("1", stsizefmt("1", True))
        self.assertEqual("1.5", stsizefmt("1.5", True))
        self.assertEqual("00", stsizefmt("00", True))
        self.assertEqual("00-000", stsizefmt("000-00", True))
        self.assertEqual("0-0000", stsizefmt("0000-0", True))

    def testGetValue(self):
        """ the getvalue function for the dict, convenience way for upper/lower case """
        mp = {"abc": 123, "ABC": 456, "Abc": 457, "def": 567}
        self.assertEqual(123, getvalue(mp, "abc"), "return the extract one")
        self.assertEqual(456, getvalue(mp, "ABC"),
                         "return the extract one again")
        self.assertEqual(457, getvalue(mp, "Abc"),
                         "return the extract one again")
        self.assertEqual(567, getvalue(mp, "DEF"),
                         "get using the lower case in second attempt")
        self.assertEqual(123, getvalue(mp, "abc,def"),
                         "get using the lower case in second attempt")

    def testAppathSep(self):
        """ tes for appathsep, early stage function of my python programming,
        it should be replaced by path.join """
        fldr = thispath
        self.assertTrue(fldr[-1] != path.sep,
                        "a path's name should not ends with path.sep")
        fldr = appathsep(fldr)
        self.assertTrue(fldr[-1] == path.sep, "with path.sep appended")

    def testGetFiles(self):
        """ test for _misc.getfiles, a early stage funtion of my python programming """
        fldr = path.join(thispath, "res")
        fns = getfiles(fldr, "NamedL", True)
        self.assertEqual("NamedList.xlsx", fns[0], "the only excel file there")
        fns = getfiles(fldr, "List")
        self.assertEqual(
            appathsep(fldr) + "NamedList.xlsx", fns[0],
            "the only excel file there")
        fns = getfiles(fldr, nameonly=True)
        fnx = listdir(fldr)
        self.assertEqual(len(fnx), len(fns), "the count of files")
        fns = set(iter(fns))
        self.assertTrue(u"厉害为国為幗.txt" in fns,
                        "utf-8 based system can return mixing charset")

    def testKaratSvc(self):
        """ the test for karat service """
        ks = karatsvc
        k0 = ks[9]
        k1 = ks["9K"]
        self.assertEqual(k0, k1, "same object return from byId/byName")
        k1 = ks["9KR"]
        self.assertEqual(k0.fineness, k1.fineness,
                         "same fineness, different karat")
        k1 = ks.getfamily(k1)
        self.assertEqual(k0, k1, "9KR's family is 9K")
        self.assertTrue(ks.issamecategory(9, 98), "9K and 9KW are all gold")
        self.assertTrue(ks.issamecategory(9, "9KW"), "9K and 9KW are all gold")
        self.assertFalse(ks.issamecategory(9, 200), "gold is not bronze")
        self.assertTrue(ks.compare(k0, k0) == 0, "the same karat")
        self.assertTrue(
            ks.compare(k0, ks[200]) > 0, "Gold is larger than bronze")
        self.assertTrue(ks.compare(k0, ks[91]) < 0, "9K is smaller than 9KR")
        lst = [ks[9], ks[18], ks[200], ks[925]]
        lst = sorted(lst, key=cmp_to_key(ks.compare))
        self.assertEqual(ks[200], lst[0], "sort method")
        self.assertEqual(ks[925], lst[1], "sort method")
        self.assertEqual(ks[18], lst[-1], "sort method")
        self.assertEqual("GOLD", karatsvc.CATEGORY_GOLD, 'catetory gold')
        self.assertEqual("BONDEDGOLD", karatsvc.CATEGORY_BONDEDGOLD)
        self.assertEqual("BONDEDGOLD", karatsvc.CATEGORY_BG)
        with self.assertRaises(KeyError):
            self.assertEqual("BLACK", karatsvc.COLOR_BLACK)
        self.assertAlmostEqual(1.016, karatsvc.convert('8K', 1, '9K'), 3)
        self.assertTrue(karatsvc.convert('8K', 1, '24K') is None, '24K does not have density data')

    def testRingSizeCvt(self):
        """ a size converter, maybe should be migrated to UOMConverter """
        rgsvc = RingSizeSvc()
        self.assertEqual("M", rgsvc.convert("US", "6", "UK"), "US#6 = UK#M")
        self.assertEqual("M", rgsvc.convert("US", "6", "AU"),
                         "US#6 = UK#M, AU using UK")
        self.assertEqual("4 1/4", rgsvc.convert("EU", "47", "US"),
                         "EU#47 = US#4 1/4")
        self.assertTrue(
            rgsvc.convert("EU", "A", "US") is None, "EU#A does not exist")
        self.assertAlmostEqual(47.0, rgsvc.getcirc("US", "4 1/4"),
                               "the circumference of US#4 1/4 is 47.0mm")

    def testNumber2Words(self):
        '''
        the numeric to English(spell) function text
        '''
        n2w = Number2Word()
        nbrs = {
            0.01: "ONE CENT",
            0.1: "TEN CENTS",
            0.11: "ELEVEN CENTS",
            1: "ONE DOLLAR",
            1.1: "ONE DOLLAR AND TEN CENTS",
            1.11: "ONE DOLLAR AND ELEVEN CENTS",
            10: "TEN DOLLARS",
            10.1: "TEN DOLLARS AND TEN CENTS",
            10.11: "TEN DOLLARS AND ELEVEN CENTS",
            123: "ONE HUNDRED AND TWENTY THREE DOLLARS",
            123.4: "ONE HUNDRED AND TWENTY THREE DOLLARS AND FORTY CENTS",
            123.45: "ONE HUNDRED AND TWENTY THREE DOLLARS AND FORTY FIVE CENTS",
            2026.34: "TWO THOUSAND AND TWENTY SIX DOLLARS AND THIRTY FOUR CENTS",
            49129: "FORTY NINE THOUSAND ONE HUNDRED AND TWENTY NINE DOLLARS"
        }
        for n in nbrs.items():
            self.assertEqual(n[1], n2w.convert(n[0]))
        n2w = Number2Word(show_no_cents=True)
        self.assertEqual("ONE DOLLAR AND NO CENTS", n2w.convert(1))
        n2w = Number2Word(show_no_cents=True, join_ten=True)
        self.assertEqual("ONE DOLLAR AND FORTY-FIVE CENTS", n2w.convert(1.45))
        self.assertEqual("ONE DOLLAR AND FORTY FIVE CENTS", n2w.convert(1.45, join_ten=False))
        self.assertEqual("ONE DOLLAR AND FORTY-FIVE CENTS", n2w.convert(1.45), 'the join_ten option should have been restored')
        t = clock()
        for idx in range(1000):
            n2w.convert(idx + 0.11)
        t = clock() - t
        print("%f ms to run" % (t * 1000))


    def testImagesize(self):
        """ the imagesize function(power by PIL) """
        fns = getfiles(path.join(thispath, "res"), "65x27")
        for fn in fns:
            self.assertEqual((65, 27), imagesize(fn), "the size of %s" % fn)
        # one special, the SOF C4 is used
        self.assertEqual((849, 826),
                         imagesize(path.join(thispath, r"res\579616.jpg")))

    def testLvDist(self):
        """
        test the LevenshteinDistance function
        """
        self.assertEqual(0, lvst_dist("I'm", "I'm"), "same string")
        self.assertEqual(1, lvst_dist("I'mx", "I'm"), "same string")
        self.assertEqual(2, lvst_dist("'mI", "I'm"), "same string")

    def testDateX(self):
        """ test for the daterange/monthadd function """
        drs = daterange(1998, 1)
        self.assertEqual(date(1998, 1, 1), drs[0], "daterange's from")
        self.assertEqual(date(1998, 2, 1), drs[1], "daterange's to")
        drs = date(2018, 1, 1)
        self.assertEqual(date(2017, 12, 1), monthadd(drs, -1))
        self.assertEqual(date(2017, 1, 1), monthadd(drs, -12))
        self.assertEqual(date(2016, 11, 1), monthadd(drs, -14))
        self.assertEqual(date(2018, 2, 1), monthadd(drs, 1))
        self.assertEqual(date(2019, 1, 1), monthadd(drs, 12))
        # leap year test
        self.assertEqual(date(2018, 2, 28), monthadd(date(2018, 1, 31), 1))
        self.assertEqual(date(2018, 2, 28), monthadd(date(2018, 1, 29), 1))
        self.assertEqual(date(2016, 2, 29), monthadd(date(2016, 1, 29), 1))
        self.assertEqual(date(2016, 2, 29), monthadd(date(2016, 1, 30), 1))

    def testSalt(self):
        '''
        test the pwd's encode/decode function
        '''
        st = Salt()
        for idx in range(20):
            s0 = "A very long string %d" % idx
            salt = st.encode(s0)
            self.assertNotEqual(s0, salt)
            self.assertEqual(s0, st.decode(salt))

    def testPath(self):
        '''
        test the getpath function
        '''
        from utilz._miscs import getpath, getmodule
        from sys import modules
        self.assertEqual(modules[__package__], getmodule(__package__))
        self.assertEqual(path.dirname(getfile(currentframe())), getpath(), 'my working path')

class ConfigSuite(TestCase):
    '''
    tests for make use of the Config class
    '''
    _listener_hc, _new_value = 0, None
    _key = _old_value = None

    def testLoad(self):
        '''
        load from one file, then from another
        '''
        cfg = Config()
        cfg.load(path.join(thispath, "res", "conf_0.json"))
        self.assertEqual("value0", cfg.get("key0"))
        self.assertListEqual(["1", 2, "3"], cfg.get("keys2"))
        cfg.load(path.join(thispath, "res", "conf_1.json"))
        self.assertEqual("value0", cfg.get("key0"))
        self.assertListEqual(["1", "2", "3"], cfg.get("keys2"))
        self.assertIsNone(cfg.get("Value0"))

    def testListener(self):
        '''
        listener to setting changes
        '''
        cfg = Config()
        cfg.load(path.join(thispath, "res", "conf_0.json"))
        self.assertEqual("value0", cfg.get("key0"))
        cfg.addListener("keys2", self._listener)
        cfg.load(path.join(thispath, "res", "conf_1.json"))
        self.assertEqual(1, self._listener_hc)
        self.assertListEqual(["1", "2", "3"], self._new_value)
        cfg.set("keys2", "a")
        self.assertEqual(2, self._listener_hc)
        self.assertEqual("a", self._new_value)

    def _listener(self, key, old_value, new_value):
        self._listener_hc += 1
        self._new_value = new_value
        self._key, self._old_value = key, old_value


class NamedListSuite(TestCase):
    """ usages of namedlist class """

    def setUp(self):
        super().setUp()
        #Always new one item to avoid methods making changes to it
        self._lsts = (["Name", "group", "age"], ["Peter", "Admin", 30],
                      ["Watson", "Admin", 45], ["Biz", "Mail", 20])

    def testList2Dict(self):
        """
        test for _misc.list2Dict, but using NamedList is more straight forward
        """
        lst, alias = ("A", None, "", "bam", "Bam1"), {
            "namE": "A",
            "description": "b,am"
        }
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

    def testNonNewInst(self):
        """ Creating without NewInst, which save memory, BUT DANGEROUS """
        idx, lsts = 0, self._lsts
        #in the iter operation, newinst=True/False behaves the same, but
        #in fact, in the newinst=False case, there is only one NamedList
        #instance created, it's dangerous operation
        for nl in NamedLists(lsts, newinst=False):
            idx += 1
            self.assertEqual(
                nl.data, lsts[idx],
                "same object, newinst=False don't affect iterator")
        #but when stored in one array, problem exposed, one item only
        #and the data was pointed to the last
        nls = [x for x in NamedLists(lsts, newinst=False)]
        nl = nls[0]
        for x in nls[1:]:
            self.assertIs(nl, x, "they are in fact the same object")
        self.assertIs(lsts[-1], nl.data, "nl was pointed to the last one")

    def testNewInst(self):
        """
        most-common usage of NamedLists::wrap a list of list,
        whose first row is title, other rows are data. A similar
        class is xwu.NamedRanges, which wrap a excel range into
        NamedLists
        """
        idx, lsts = 0, self._lsts
        #iter operation, just the same as Non_NewInst
        for nl in NamedLists(lsts):
            idx += 1
            self.assertEqual(
                nl.data, lsts[idx],
                "same object, newinst=False don't affect iterator")

        nls = [x for x in NamedLists(lsts)]
        #safer than Non_NewInst, the first NL Points to first
        self.assertEqual(len(lsts) - 1, len(nls), "first row used as title")
        for nl in nls[1:]:
            self.assertIsNot(nls[0], nl, "different objects, not the same")
        nl = nls[0]
        self.assertEqual(lsts[1], nl.data, "title off, first row of data")
        self.assertEqual("Peter", nl.name, "the name property")
        self.assertEqual("Admin", nl.group, "the group property")
        self.assertEqual(30, nl.age, "the age property")

    def testAccess(self):
        """ after a NamedList was created use it to access the under-laying data """
        lsts = self._lsts
        nl = NamedLists(lsts).__next__()
        #get by name/idex/slice
        self.assertTrue(hasattr(nl, "name"), "response to hasattr")
        self.assertEqual(lsts[1][0], nl[0], "access by index")
        self.assertEqual(lsts[1][0], nl.name, "access by index")
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

        self.assertEqual(0, nl.getcol("Name"))
        self.assertEqual(0, nl.getcol("name "))
        self.assertEqual(2, nl.getcol("age"))
        self.assertTrue(nl.getcol("age_") is None)
        self.assertEqual(tuple("name,group,age".split(",")), nl.colnames)
        self.assertEqual((0, 1, 2), nl.colids)

    def testNonNormalize(self):
        """ normalize function, make sure non-normalized works """
        nls = [x for x in NamedLists(self._lsts, normalize=None)]
        nl = nls[0]
        self.assertEqual("Peter", nl.get("Name"))
        self.assertTrue("Name" in nl.colnames)
        self.assertTrue("name" not in nl.colnames)

    def testObjectAlias(self):
        """ access normal object """

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

    def testDictAlias(self):
        """ access a dict with alias name """
        d0 = {"name": "David", "Age": 20}
        al = NamedList({"nick": "name"}, d0)
        self.assertEqual(al["name"], al["nick"])
        self.assertTrue("nick" not in d0)
        self.assertEqual(al.name, al.nick)

    def testChangeName(self):
        """ after creation, change some colnames """
        # nl = NamedList("id,name,agex,agey", [None] * 4)
        nl = NamedList("id,name,agex,agey")
        nl.id, nl.agex = 1, 30
        self.assertEqual(1, nl.id)
        nl = nl._replace({"idx": "id,", "age": "agex"})
        self.assertEqual(1, nl.idx)
        self.assertEqual(30, nl.age)
        self.assertTrue("id" not in nl.colnames)
        self.assertEqual(2, nl.getcol("age"))


class XwuSuite(TestCase):
    """
    test suit for xwu funcitons
    """
    _hasxls = None
    _app, _tk = (None,) * 2

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
        super().tearDown()


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
        ttls = ('2017&"宋体,Regular"年&"Arial,Regular"6&"宋体,Regular"月',
                '2017&"宋体,Regular"年&"Arial,Regular"&6 6&"宋体,Regular"月')
        exps = ("2017年6月", "2017年6月")
        for idx, it in enumerate(ttls):
            self.assertEqual(exps[idx], xwu.escapetitle(it), "the title")

    def testNamedLists(self):
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
        rng = xwu.offset(rng, 1, 0)
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
        self.assertEqual(
            datetime(1998, 1, 3, 0, 0), emp["edate"],
            "get date use translated name")

        # test the find's all function
        nl = xwu.find(sht, "Name", lookat=LookAt.xlPart, find_all=True)
        self.assertEqual(9, len(nl), "the are 9 items has name as part")

    def testNextCell(self):
        '''
        check if next cell works
        '''
        fn = getfiles(resfldr, "getTableData")[0]
        app = self._app
        wb = app.books.open(fn)
        sht = wb.sheets[0]
        rng = xwu.nextcell(sht.cells(1, 1))
        self.assertEqual(sht.cells(1, 2), rng, 'Normal cell')
        rng = xwu.nextcell(sht.cells(10, 1))
        self.assertEqual(sht.cells(11, 2), rng, "Merged cell's right")
        rng = xwu.nextcell(sht.cells(10, 1), "up")
        self.assertEqual(sht.cells(9, 1), rng, "Merged cell's up")
        rng = xwu.nextcell(sht.cells(10, 1), "left")
        self.assertIsNone(rng, 'exceeds border')
        # self.assertEqual(sht.cells(9, 1), rng, "Merged cell's right")
        wb.close()

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
            nls = [
                x for x in xwu.NamedRanges(
                    rng, skip_first_row=True, name_map=nmap)
            ]
            #print("%s colnames are:(%s)" % (name, list(nls[0].colnames)))
            self.assertEqual(3, len(nls), "result count of %s" % name)
            self.assertEqual(2, nls[0]["9k"], "9K result of %s" % name)
            self.assertEqual(16, nls[2].s950, "S950 of %s" % name)
            print("using %f ms to perform %s" % (time.clock() - t0, name))
        # try a blank range, should return none
        nls = xwu.NamedRanges(sht.range(1000, 1000))
        self.assertIsNone(nls, "Nothing should be returned")

    def testCol(self):
        """ test for the column idx/name translation function
        """
        self.assertEqual(xwu.col('A'), xwu.col('a'))
        self.assertEqual(xwu.col('Xfd'), xwu.col('XFD'))
        mp = {'A': 1, 'AA': 27, 'AZA': 1353, 'XFD': 16384}
        for k, v in mp.items():
            self.assertEqual(v, xwu.col(k))
            self.assertEqual(k, xwu.col(v))
        self.assertEqual(((2, 1), ), xwu.addr2rc('$A$2'))
        self.assertEqual(((2, 1), (3, 2), ), xwu.addr2rc('$A$2:$B$3'))
    
    def testPerf(self):
        '''
        xlwings is infact very slow, test how slow it was
        Some conclusion:
            .cells(x, y) is fast
            .offset(x, y) is very slow, use cells(rng.row + x, rng.row + y) is much faster
        '''
        app = self._app
        wb = app.books.add()
        sht = wb.sheets[0]
        pf = Profile()
        def _run(uf, s, e):
            for i in range(s, e):
                rng = sht.cells(i + 1, 1)
                if uf:
                    xwu.offset(rng, 1, 1)
                else:
                    rng.offset(1, 1)
        tts = []
        for uf, s, e in ((True, 0, 5), (False, 30, 35)):
            pf.enable()
            _run(uf, s, e)
            pf.disable()
            tts.append(Stats(pf).total_tt)
        r = tts[1]/tts[0]
        print('api/udf = %4.2f' % r)
        self.assertTrue(r > 1, 'user-defined offset is faster than official API')
        wb.close()

    def testGetHidden(self):
        """ get the hidden row/columns inside a sheet """
        app = self._app
        app.visible = True
        wb = app.books.open(path.join(thispath, "res", "hidden_r_c.xlsx"))
        nl = NamedList("sn,row,exps")
        mp = (
            ("Spread", True, [(3, 10), (14, 18), (24, 9999)]),
            ("Header", True, [(1, 9), ]),
            ("NoHidden", True, None),
            ("AllHidden", True, [(1, 12), ]),
            ("Spread_Col", True, None),
            ("Spread_Col", False, [(3, 5), ]),
            ("Row_Col", True, [(2, 999), (1001, 9997), (9999, 9999), (10001, 10001)]),
            ("Row_Col", False, [(2, 56), (59, 66)]),
            ("Row_Col_Huge", True, [(1, 100000), (100002, 500000)])
        )
        # mp = (("AllHidden", True, [(1, 12), ]), )
        tc, loops = clock(), 2
        for idx in range(loops):
            print("doing loop %d" % idx)
            for val in mp:
                nl.setdata(val)
                print("doing sheet(%s)'s %s" % (nl.sn, "Row" if nl.row else "Col"))
                lsts, exps = xwu.hidden(wb.sheets(nl.sn), nl.row), nl.exps
                msg = "Sheet(%s), %s" % (nl.sn, "row" if nl.row else "col")
                if isinstance(exps, (tuple, list)):
                    self.assertListEqual(exps, lsts, msg)
                else:
                    self.assertEqual(exps, lsts, msg)
        tc = clock() - tc
        print("using %4.2fs for each loop, total loops = %d, total time = %4.2f" % (tc / loops, loops, tc, ))


    def testInsertPhoto(self):
        ''' check the insertphoto function '''
        fn = path.join(thispath, "res", "579616.jpg")
        app = self._app
        sws = xwu.appswitch(app, {'visible': True})
        wb = app.books.open(path.join(thispath, "res", "getTableData.xlsx"))
        sht = wb.sheets[0]
        shp = xwu.insertphoto(fn, sht.range("A1:F15"), margins=(2, 2))
        self.assertIsNotNone(shp, 'There must be a shape')
        self.assertAlmostEqual(2, shp.top, 2, 'the top')
        self.assertAlmostEqual(59.99, shp.left, 2, 'the left')
        self.assertAlmostEqual(204.03, shp.width, 2, 'the left')
        shp.delete()
        shp = xwu.insertphoto(fn, sht.range("A1:F15"), margins=(2, 2), alignment="L,M")
        self.assertAlmostEqual(2, shp.top, 2, 'the top')
        self.assertAlmostEqual(2, shp.left, 2, 'the left')
        shp = xwu.insertphoto(fn, sht.range("A1:F15"), margins=(2, 2), alignment="R,M")
        self.assertAlmostEqual(2, shp.top, 2, 'the top')
        self.assertAlmostEqual(117.97, shp.left, 2, 'the left')
        xwu.appswitch(self._app, sws)
    
    def testXwPerf(self):
        '''
        test for the performance issue of xlwings.
        After test, I found out that only when there are more than 5 cells, the one-by-one is slower than one
        '''
        app, tk = xwu.appmgr.acq()
        wb = app.books.add()
        ck = clock()
        sht = wb.sheets[0]
        rc, cc = 20, 6
        for row, col in product(range(rc), range(cc)):
            sht.cells[row, col].value = 1
        t0 = clock() - ck

        ck = clock()
        lst = [[1] * cc] * rc
        sht.cells(rc + 2, 1).value = lst
        t1 = clock() - ck
        print('cell-count=%d, one-by-one=%f, one = %f, obo/o = %f' % (rc * cc, t0, t1, t0 / t1))
        self.assertTrue(t0 > t1, 'one write faster than several writes')
        # app.visible = True
        xwu.appmgr.ret(tk)

class _DecHex(AbsResolver):
    '''
    class help translate Hex to decimal
    '''
    def __init__(self):
        s = '0123456789ABCDEF'
        self._h2d = {x[1]: x[0] for x in zip(range(16), s)}

    def resolve(self, arg):
        return self._h2d.get(arg) or arg

class ExpressionTest(TestCase):
    ''' test for the Exp/Resolver class
    '''

    def testEval(self):
        r = _DecHex()
        self.assertTrue(Exp("F", '>', 3).eval(r), "Hex(F) > 3")
        self.assertEqual(12, Exp("F", '-', 3).eval(r), "F - 3 == 12")
    
    def testChain(self):
        _add = lambda x, y: Exp(x, '+', y)
        _gt = lambda x, y: Exp(x, '>', y)
        exps = []
        exps.append(_gt(_add(3, 4), 6).chain('and', _gt(5, 3), _gt(4, 2)))
        self.assertTrue(exps[-1].eval(), '3+4 > 6 and 5 > 3 and 4 > 2')
        exps.append(_gt(_add(3, 4), 6).chain('and', _gt(5, 3), _gt(2, 4)))
        self.assertFalse(exps[-1].eval(), '3+4 > 6 and 5 > 3 and 2 > 4')
        exps.append(_gt(_add(3, 4), 6).chain('and', _gt(5, 3), _gt(4, 3)).or_(_gt(2, 0)))
        self.assertTrue(exps[-1].eval(), '(3+4 > 6 and 5 > 3 and 4 > 3) or (2 > 0)')
        exps.append(_add(2, 3).chain('add', 4, 5, _add(6, 7)))
        self.assertEqual(27, exps[-1].eval(), '2+3+4+5+6+7')
        for exp in exps:
            print(str(exp))

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
    _engine = create_engine('sqlite:///:memory:')
    Mstr.metadata.create_all(_engine)
    _sessmgr = SessionMgr(_engine)

    @classmethod
    def setUpClass(cls):
        super().setUpClass()
        logging.getLogger("sqlalchemy").setLevel(logging.DEBUG)

    @property
    def sessctx(self):
        """ section for with statement """
        return ResourceCtx(self._sessmgr)

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
            self.assertEqual(1, len(mstr),
                             "There should be only one item inside")
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
    """ class for catalog making """

    def test_detectMargin(self):
        """
        the ability to test for height/width of a page can hold
        """


if __name__ == "__main__":
    main()
