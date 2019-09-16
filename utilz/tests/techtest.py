#! coding=utf-8
'''
* @Author: zmFeng
* @Date: 2018-07-04 08:46:52
* @Last Modified by:   zmFeng
* @Last Modified time: 2018-07-04 08:46:52
for python's language/basic facility test, a practice farm
'''

import dbm
import gettext
import re
import subprocess as sp
import tempfile
from argparse import ArgumentParser
from cProfile import Profile
from datetime import date
from decimal import Decimal
from io import StringIO
from itertools import islice
from logging import Logger
from numbers import Number
from os import path, remove, walk
from pstats import Stats
from sys import platform
from unittest import TestCase, skip

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from bidict import ValueDuplicationError
from bidict import bidict as bd
from matplotlib.backends.backend_pdf import PdfPages

from utilz import getfiles, imagesize
from utilz.miscs import getpath, trimu

_logger = Logger(__name__)
try:
    import pytesseract as tesseract
    from cv2 import (THRESH_BINARY, GaussianBlur, imread,
                     imwrite, threshold)
    from PIL import Image
except ImportError:
    pass


@skip("TODO::")
class TesseractSuite(TestCase):
    '''
    tesseract tests
    '''
    #_srcfldr = r"p:\aa\x\org\jophotos"
    _srcfldr = r"p:\aa\x\org1\Smp"
    _cropfldr = r'd:\temp\crop'
    _ordbrd = (0.75, 0.1, 1, 0.2)
    _smpbrd = (0.75, 0.2, 1, 0.45)

    def testCrop_Gray(self):
        '''
        get metadata(like dpi), crop and convert to gray
        '''
        brd = self._smpbrd if self._srcfldr.lower().find("smp") >= 0 else self._ordbrd
        cnt = 0
        for fn in getfiles(self._srcfldr, ".jpg"):
            cnt += 1
            if cnt > 1E5:
                break
            orgsz = imagesize(fn)
            img = Image.open(fn)
            box = (orgsz[0] * brd[0], orgsz[1] * brd[1], orgsz[0]*brd[2], orgsz[1] * brd[3])
            img.load()
            dpi = img.__getstate__()[0].get("dpi")
            img = img.crop(box)
            tfn = path.join(self._cropfldr, path.basename(fn))
            if dpi:
                img.save(tfn, dpi=dpi)
            else:
                img.save(tfn)

    def testCV2(self):
        '''
        CV2's image adjustment for JO highlight
        '''
        dpi = None
        srcfn = r'd:\temp\CV2\0003.jpg'
        for fn in getfiles(r"d:\temp\cv2", ".jpg"):
            if fn.find("_") >= 0:
                continue
            if not dpi:
                img = Image.open(fn)
                img.load()
                dpi = img.__getstate__()[0].get("dpi")
                img.close()
            img = imread(fn, 0)
            img = GaussianBlur(img, (5, 5), 0)
            th1 = threshold(img, 160, 255, THRESH_BINARY)[1]
            fldr, bn, cnt = path.dirname(srcfn), path.splitext(path.basename(fn)), 0
            for x in (th1,):
                fn0 = path.join(fldr, "%s_%d%s" % (bn[0], cnt, bn[1]))
                imwrite(fn0, x)
                # because CV2 does not save metadata, while dpi is very important
                # use PIL's image to process it
                img = Image.open(fn0, mode="r")
                img.save(fn0, dpi=dpi)
                cnt += 1

    def testOCR(self):
        '''
        ocr test(with language specified)
        '''
        ptn = re.compile(r"N.\s?(\w*)")
        with open(path.join(self._cropfldr, "log.dat"), "wt", encoding="utf-8") as fh:
            for fn in getfiles(self._cropfldr, ".jpg"):
                img = Image.open(fn)
                s0 = tesseract.image_to_string(img, "eng")
                mt = ptn.search(s0)
                if not mt:
                    s0 = "JO#%s:%s" % (path.basename(fn), s0)
                else:
                    s0 = "JO#%s:%s" % (path.basename(fn), mt.group())
                fh.writelines(s0 + "\r\n")

class ArgParserTest(TestCase):
    '''
    Argument parser usage
    '''

    def testPositional_Optional(self):
        '''
        Positional argument and Optional Arguments.
        Argument does not have '-' inside the string, can be added whenever place(in the ArgumentParser instance). one command line can have one only?
        switch contains '-', can be short or long and can have several long name defined(maybe not necessary at all)
        '''
        ap = ArgumentParser("Positional+Optional", description="One pisitional, one optional", epilog="The optional has 3 names, the key one is \"what\"", add_help=True)

        # the positional arguments can be at whatever place
        # one argument can have more than one name. if so, the result name should be the first one with "--". The below example is "what"
        ap.add_argument("-w", "--what", "--what_what", help="what should be d0-d1", default="def_x")
        ap.add_argument("files", nargs="*", help="the files that need to be processed")
        np = ap.parse_args(["file1", "file2", "--what", "This-is-me"])
        self.assertEqual("This-is-me", np.what)
        self.assertListEqual(["file1", "file2"], np.files)
        np = ap.parse_args(["file1", "file2"])
        self.assertEqual("def_x", np.what)
        self.assertListEqual(["file1", "file2"], np.files)
        np = ap.parse_args(["file1", "file2", "--what_what", "This-is-me"])
        self.assertEqual("This-is-me", np.what)
        self.assertListEqual(["file1", "file2"], np.files)

        np = ap.parse_args(["file1", "file2", "file3", "-w", "This-is-me"])
        self.assertEqual("This-is-me", np.what)
        self.assertListEqual(["file1", "file2", "file3"], np.files)

        ap = ArgumentParser("Positional+Optional", description="Like above, but the key name chagned from \"what\" to \"what_what\"", add_help=True)
        ap.add_argument("files", nargs="*")
        # the result name here is "what_what"
        ap.add_argument("-w", "--what_what", "--what", help="what should be d0-d1", default="def_x")
        np = ap.parse_args(["file1", "file2"])
        self.assertEqual("def_x", np.what_what)

        # below statement show a help screen and throws exception, so ignore it
        ap.parse_args(["-h", ])
        print(np)
        return

    def test_2_positional(self):
        '''
        parser with 2 positional arguments
        '''
        ap = ArgumentParser("2 positional", description="Like above, but the key name chagned from \"what\" to \"what_what\"", add_help=True)
        ap.add_argument("domain", help="domain name")
        ap.add_argument("files", nargs="*", help="the files for the domain")
        np = ap.parse_args(["hnjchina", "file1", "file2"])
        self.assertEqual("hnjchina", np.domain)
        self.assertListEqual(["file1", "file2"], np.files)
        # inherits, but can not have help again, or it throws exception: argparse.ArgumentError
        ap = ArgumentParser("Descendant", parents=[ap, ], add_help=False)
        ap.add_argument("ext", help="the extension")
        np = ap.parse_args(["hnjchina", "file1", "file2", "exts"])
        self.assertEqual("hnjchina", np.domain)
        self.assertEqual("exts", np.ext)
        self.assertListEqual(["file1", "file2"], np.files)

    def testGetText(self):
        '''
        a i18n module by python
        '''
        gettext.bindtextdomain('utilz', r'd:\temp\abx')
        gettext.textdomain('utilz')
        _ = gettext.gettext
        print(_('This is a translatable string.'))
        print("hello")


class TechTests(TestCase):
    """
    class trying the technical test
    """

    def testTry(self):
        '''
        the try/catch mechanism
        '''
        flag = False
        try:
            print(1 / 0)
        except:
            flag = True
        self.assertTrue(flag, "exceptions occured")
        flag = False
        try:
            print(1 / 1)
        except:
            flag = True
        self.assertFalse(flag, "no exception occured")

    def testRE(self):
        """
        regexp tests
        """
        ptn = re.compile(r"C(\d{1})")
        s0 = "JMP12C1"
        mt = ptn.match(s0)
        self.assertFalse(bool(mt), "There should be no match")
        mt = ptn.search(s0)
        self.assertTrue(bool(mt), "There should be search")
        self.assertEqual(("1",), mt.groups(), "The so-call zero group")
        self.assertEqual("1", mt.group(1), "The so-call first group")

    def testProfiling(self):
        '''
        use the profile class to get performance report of some codes
        '''
        def _runner():
            v = 0
            for i in range(100):
                v += i
            return v
        pf = Profile()
        pf.enable()
        _runner()

        pf.disable()
        opt = StringIO()
        #opt = FileIO(r'd:\temp\pf.dat')
        sts = Stats(pf, stream=opt).sort_stats('cumulative').strip_dirs()
        # get the total-time used tu = sts.total_tt
        sts.print_stats(10)
        print(opt.getvalue())
        opt.close()

    def testSeveralForItr(self):
        """
        try a multi iteration in for statement
        """
        rngs = ((1, 3), (4, 9))
        slots = [x for y in rngs for x in range(y[0], y[1])]
        # print(slots)
        self.assertEqual(7, len(slots))
        self.assertEqual(1, slots[0])

    def testFuncArgs(self):
        """
        try the *arg and **kwds argument of a function
        """
        def sth0(a):
            return(a,)

        def sth1(a, *args):
            return (a, args)

        def sth2(a, **kwds):
            return (a, kwds)

        def sth3(a, *args, **kwds):
            return (a, args, kwds)

        arr = sth0(5)
        self.assertTupleEqual((5,), arr, "single argument")
        arr = sth1(5, 1, 2, 3)
        self.assertTupleEqual((5, (1, 2, 3)), arr, "single + positional argument")
        arr = sth2(5, nice="to")
        self.assertTupleEqual((5, {"nice": "to"}), arr, "single + named argument")
        # this 2 argument error
        with self.assertRaises(TypeError):
            sth1(5, nice="to")
        with self.assertRaises(TypeError):
            sth2(5, 1, 2, 3)
        # a full-blow
        arr = sth3(5, 20, nice="to")
        self.assertTupleEqual((5, (20,), {"nice": "to"}), arr)

    def testClassMethod(self):
        """
        static/class method can be accessed by
            .class of itself
            .instance of itself
            .class of child
            .instance of child
        Although they finally call to the same function, but they are not the referencely same
        """
        class A():
            '''
            parent class with static method
            '''
            @classmethod
            def sta(cls):
                return "sta"

            def inst(self):
                return "inst"

        class B(A):
            '''
            child class extends parent's inst() method
            '''
            def inst(self):
                return super().inst() + "_B"

        self.assertEqual(A.sta(), B().sta())
        self.assertEqual(A.sta(), B.sta())
        self.assertEqual(A.sta(), B().sta())
        self.assertEqual(A().inst() + "_B", B().inst())
        self.assertFalse(A.sta is A().sta)
        self.assertFalse(A.sta is B.sta)
        self.assertFalse(A().sta is B().sta)

    def testMArrayCreation(self):
        ''' create multiple array, refer to official doc's "built-in Types" FMI '''
        lsts = [[]] * 3
        lsts[0].append(1)
        # infact, the 3 array inside lsts refers the same array
        self.assertEqual(1, lsts[1][0])
        lsts = [[] for x in range(3)]
        lsts[0].append(1)
        # this is the correct way to create a 3 array array
        self.assertEqual(0, len(lsts[1]))

    def testAccessChain(self):
        '''
        all accesses are controlled __getattribute__(), below is the access proority(high to low):
        .class property
        .data descriptor
        .instance property
        .non-data descriptor
        .__getattr__() method
        .AttributeError raised
        descriptor only works when it's assigned to a class(not instance) while the attribute was not yet initialized
        '''
        class NullDev(object):
            ''' data descriptor that will get/set None '''

            def __init__(self, name):
                self.name = name

            def __get__(self, instance, owner):
                print("(%s) invoking get method, inst = %r, owner = %r" % (self.name, instance, owner))
                return None

            def __set__(self, instance, value):
                print("(%s) involing set method with value %r" % (self.name, value))
                instance.lst_data = value

        class Foo(object):
            cls_prop = "cls_prop"
            data_dsc = NullDev("init_clz_level")
            def __init__(self):
                self.inst_prop = "inst_prop"
                # assigning to instalce's won't bahaves like descriptor
                self.data_dsc1 = NullDev("init_inst_level")

            def doit(self):
                return "hello"

        fo = Foo()
        # descriptor's setter has instance, no owner because if you assign value to a descriptor in class
        # level, it will be killed
        self.assertFalse(hasattr(fo, 'lst_data'))
        fo.data_dsc = 'init the lst_data property by NullDev'
        self.assertTrue(hasattr(fo, 'lst_data'))

        Foo.data_dsc = 'x'
        self.assertEqual('x', Foo.data_dsc)
        self.assertEqual('x', fo.data_dsc)

        # restore the descriptor
        Foo.data_dsc = NullDev("reinit_clz_level")
        self.assertFalse(hasattr(Foo, "lst_data"), 'Existing NullDev was overrided by above statement')
        self.assertEqual("init the lst_data property by NullDev", fo.lst_data)
        fo.data_dsc = 'x'
        self.assertEqual('x', fo.lst_data)

        self.assertIs(NullDev, type(fo.data_dsc1), 'not behaves like a descriptor')
        fo.data_dsc1 = 'x'

        #assigning value to instance hides the one in class
        fo.doit = "y"
        self.assertEqual("y", fo.doit)
        del fo.doit
        # after removing the one in the instance, the one in class restore
        self.assertEqual("hello", fo.doit())

    def testComp(self):
        '''
        logical comparisons
        '''
        a = 3
        self.assertTrue(1 < a < 5) #continuous comparison
        self.assertTrue(a != 5) # same as "not (a == 5)" because not has lower priority in non-logical operation
        # find element inside sequence. sequence types are: tuple, list, range and the descestor.
        self.assertTrue('a' in 'abcea') # find sub-string in string
        self.assertTrue(1 in (2, 3, 1)) # find element
        self.assertTrue(1 in {1: 'a', 2: 'b'})
        self.assertEqual('T', a == 3 and 'T' or 'F') # same as 'T' if a == 3 else 'F'
        self.assertEqual('F', a != 3 and 'T' or 'F')
        self.assertTrue(isinstance(a, Number))
        self.assertTrue(isinstance(a, int))
        self.assertFalse(type(a) is type(object))
        self.assertIs(type(a), type(0))


    def testManyItf(self):
        '''
        ManyInterfaces class implements many built-in interfaces for study purpose
        '''
        mi = ManyInterfaces(f="k")
        # can make use of an iterator object without iter() function
        # self.assertListEqual([1, 2, 3], [x for x in iter(mi)])
        self.assertListEqual([1, 2, 3], [x for x in mi])
        # can not next() because the internal _iter not inited by the __iter__() method
        mi = ManyInterfaces()
        with self.assertRaises(StopIteration, msg='containor not activated'):
            next(mi)
        mi = ManyInterfaces()
        # self.assertEqual(1, iter(mi).next())
        mi += (2, 3, 4)
        self.assertListEqual([1, 2, 3, 2, 3, 4], mi.data, 'inplace add')
        mi = ManyInterfaces()
        gtr = mi.gtr
        self.assertEqual(1, gtr.send(None), 'the generator')

        mi = ManyInterfaces()
        self.assertListEqual([1, 2, 3], list(mi.lst_data), 'defined property')
        self.assertEqual("_data_len", mi.data_len, 'property of class')
        ld = _LenDescriptor()
        mi.data_len = ld
        self.assertIs(ld, mi.data_len, "once a attribute is already inside a instance or it's type, assigning descriptor won't bahave like descriptor")
        self.assertEqual("_data_len", ManyInterfaces.data_len, "class's attribute not changed")
        ManyInterfaces.data_len = ld
        self.assertEqual(3, mi.data_len, "new attribute in class works as promise")
        self.assertEqual('__getattribute__(y)', mi.y, 'by __getattribute__()')
        self.assertEqual('__getattr__(z)', mi.z, 'by __getattr__()')
        with self.assertRaises(AttributeError, msg='k not defined, and no __getattribute__/__getattr__ reponse') as err:
            print(mi.k)
        self.assertEqual('attribute k not defined in __getattr__()', err.exception.args[0])

    def testFlatten(self):
        ''' there often be chances lsts = [(1, 2, 3), (2, 5)] and you want them into
        one list of (1, 2, 3, 2, 5). this is called flattening
        I often forget this, so make it a test for reference.
        Maybe can be use together with itertools
        '''
        lsts = [(1, 2, 3), (2, 5)]
        from itertools import chain
        exps = (1, 2, 3, 2, 5)
        self.assertTupleEqual(exps, lsts[0] + lsts[1], 'one by one enumerate, stupid')
        self.assertTupleEqual(exps, tuple(y for x in lsts for y in x), 'list comprehensive, smart')
        self.assertTupleEqual(exps, tuple(chain(*lsts)), 'itertools, smarter')
        lsts = [[(1, 2, 3), (2, 5)], [(7, 8), (9, 10)], ]
        exps = (1, 2, 3, 2, 5, 7, 8, 9, 10)
        self.assertTupleEqual(exps, tuple(chain(*chain(*lsts))))
        self.assertTupleEqual(exps, tuple(z for x in lsts for y in x for z in y))

    def testTakeFromList(self):
        ''' take just part from a tuple
        '''
        lst = (1, 2, 3, 4, 5)
        first, *_ = lst
        self.assertEqual(1, first)
        first, *_, end = lst
        self.assertEqual((1, 5), (first, end))

    def testDbm(self):
        ''' Sometimes we need to store huge dict-like object to disk instead of memory, this is called persisting. According to https://docs.python.org/3/library/dbm.html, I can use dbm, which acts like a dict, the only difference is that the data is inside the disk instead of memory.
        Here using sqlite is too heavy.
        Below is a demo about how to use it
        '''
        dbfn = None
        with dbm.open(path.join(tempfile.gettempdir(), '__tmp__'), flag='c') as db:
            if not dbfn:
                dbfn = [db._datfile, db._dirfile, db._bakfile]
            db['a'] = 'b'
            db['b'] = 'c'
            with self.assertRaises(TypeError):
                db['c'] = 1
            self.assertTrue('a' in db)
            self.assertFalse('x' in db)
            self.assertListEqual([b'a', b'b'], [x for x in db])
        self.assertTrue(dbfn)
        for fn in dbfn:
            remove(fn)
        del tempfile, dbm, remove

    def testbidict(self):
        ''' usage of the bidict '''
        d = bd({'a': 1, 'b': 2, 'c': 3})
        self.assertEqual(1, d['a'])
        self.assertEqual('a', d.inverse[1])
        with self.assertRaises(ValueDuplicationError):
            d = bd({'a': 1, 'b': 2, 'c': 1})

    def testIter(self):
        ''' 3 methods to get the nth item from a iteration/generator
        '''
        mp, nth = {idx: val for idx, val in enumerate('abcdefghijk')}, 3

        # method 1: islice then next()
        it = islice(mp.keys(), nth, nth + 1)
        rt = next(it)
        self.assertEqual(3, rt, 'the most simple, maybe least memory consuming')

        # method 2: next() many times
        it = iter(mp.keys())
        i, rt = 0, None
        while i <= nth:
            rt = next(it)
            i += 1
        self.assertEqual(3, rt, 'not elegant and maybe slow')

        # method 3: to list then get the nth
        it = [x for x in mp.keys()]
        rt = it[nth]
        self.assertEqual(3, rt, 'maybe huge memory consuming')

    def testUnpack(self):
        ''' some ways to unpack
        '''
        f, s, *l = [x for x in range(10)]
        self.assertEqual(0, f)
        self.assertEqual(1, s)
        self.assertListEqual([x for x in range(2, 10)], l, 'what left is assigned to l')

        # the equipvalent in python 2 is
        it = iter(range(10))
        f, s, l = next(it), next(it), [x for x in it]
        self.assertEqual(0, f)
        self.assertEqual(1, s)
        self.assertListEqual([x for x in range(2, 10)], l, 'what left is assigned to l')

    def testShellX(self):
        ''' invoke the shell, capture the output
        '''
        from utilz._win32_reparse import clearTempFiles
        # print(users())
        # print(users('home'))
        clearTempFiles()


class _LenDescriptor(object):
    '''
    this is called data-descriptor, method/function,
    on the other hand, is called non-data-descriptor
    in this example, I won't store data myself, instead, use data in "instance" item
    '''
    def __get__(self, instance, owner):
        return len(instance.lst_data)

    def __set__(self, instance, value):
        raise AttributeError("instance's len property can not be set")

class ManyInterfaces(object):
    """
    this class try to implement many system-level interface for practice purpose
    """
    data_len = "_data_len"

    def __init__(self, *args, **kwds):
        ''' arguments as list and named-map '''
        # object does not support constructor arguments super().__init__(*args, **kwds)
        _logger.debug("args and kwds are: %s, %s" % (args, kwds))
        self._auto_iter = "auto_iter" in kwds
        self.lst_data = list(kwds.get("data", (1, 2, 3)))
        self._ptr = None

    @property
    def data(self):
        ''' the internal list '''
        return self.lst_data

    def __iter__(self):
        ''' __iter__ along with __next__ make an object iterable, that is, can be
        access using iter() method
        '''
        self._ptr = 0
        return self

    def __next__(self):
        ''' second interface of iterator '''
        if not self.lst_data or self._ptr is None or self._ptr >= len(self.lst_data):
            self._ptr = None
            raise StopIteration('eof reach or data not set')
        rc = self.lst_data[self._ptr]
        self._ptr += 1
        return rc

    @property
    def gtr(self):
        ''' a generator, acts just like a iterator '''
        for x in self.lst_data:
            yield x

    def __iadd__(self, other):
        ''' iadd/isub and so on, are inside operator module, measn in-place operation, '''
        if not isinstance(other, (tuple, list)):
            raise AttributeError("should provide tuple or list")
        if not self.lst_data:
            self.lst_data = []
        self.lst_data.extend(other)
        return self

    def __getattribute__(self, name):
        if name == "y":
            return '__getattribute__(y)'
        return super().__getattribute__(name)
        #raise AttributeError("attribute %s not defined in __getattribute__()" % name)

    def __getattr__(self, name):
        if name == 'z':
            return '__getattr__(z)'
        # return super().__getattr__(self, name)
        raise AttributeError('attribute %s not defined in __getattr__()' % name)

class SqlAlchemyURL(TestCase):
    '''
    url for create engine
    https://docs.sqlalchemy.org/en/13/core/engines.html
    '''
    pass

class PandasSuite(TestCase):
    ''' try pandas out
    '''

    @classmethod
    def setUpClass(cls):
        super().setUpClass()
        cls._dates_ = None

    @property
    def _dates(self):
        if self._dates_ is not None:
            return self._dates_
        def _probe_fmt(col):
            if col.find('-') > 0:
                fmt = '%Y-%m-%d %H:%M:%S' if col.find(':') > 0 else '%Y-%m-%d'
            else:
                fmt = '%H:%M:%S'
            return fmt
        def _dp(*dt):
            # will pass an array of array, each column as a array
            fmt, lsts = None, []
            for cols in dt:
                if isinstance(cols, str):
                    fmt = _probe_fmt(cols)
                    lsts.append(pd.datetime.strptime(cols, fmt))
                else:
                    lst, fmt = [], _probe_fmt(cols[0])
                    lsts.append(lst)
                    for col in cols:
                        lst.append(pd.datetime.strptime(col, fmt))
            return lsts
        s0 = 'id,name,height,bdate,btime,ts\n' +\
            '1,peter,120,2019-02-01,12:30:45,2019-03-02 9:30:00\n' +\
            '2,watson,125,2019-02-02,7:25:13,2019-03-02 9:32:00\n' +\
            '3,winne,110,2019-02-07,9:25:13,2019-03-04 5:10:00\n' +\
            '10,kate,100,2019-02-10,7:12:10,2019-03-04 5:10:00\n' +\
            '20,john,123,2019-03-18,9:25:10,2019-03-07 5:10:00\n'
        fh = StringIO(s0)
        self._dates_ = pd.read_csv(fh, parse_dates={'birthday': ['bdate', 'btime'], 'lastModified': ['ts']}, date_parser=_dp, converters={'name': trimu})
        return self._dates_

    def testSeries(self):
        ''' access Series by id/name or slice
        '''
        sts = self._dates
        sr = sts.id
        self.assertEqual(2, sts.columns.get_loc('id'), 'id was put to the 3 position')
        self.assertTrue(isinstance(sr, pd.Series), 'Return when access by colname')
        self.assertEqual(1, sr[0])
        self.assertEqual(5, len(sr))

        sr = sts.iloc[0]
        self.assertTrue(isinstance(sr, pd.Series), 'Return when access by row')
        self.assertEqual(1, sr.id, 'access by name')
        self.assertEqual(1, sr[2], 'access by index')

        # column of rows
        sr = sts.iloc[:2]
        self.assertListEqual([1, 2], [x for x in sr.id])

        sr = sts[['id', 'name']]
        self.assertTrue(isinstance(sr, pd.DataFrame))
        # can not use sts.loc['id', 'name'] or sts.loc[['id', 'name']]
        with self.assertRaises(KeyError):
            sts.loc['id', 'name']
        with self.assertRaises(KeyError):
            sts.loc[['id', 'name']]

    def testDataFrame(self):
        sts = self._dates
        for cn in ('birthday', 'lastModified'):
            self.assertTrue(cn in sts.columns)
        self.assertEqual(len(sts.columns), 4)
        # when iloc is not slice, the return item is a tuple
        # sr will be a tuple
        sr = sts.iloc[0].birthday
        self.assertTrue(isinstance(sr, list))
        self.assertTrue(isinstance(sr[0], pd.datetime))

        # sr will be a Series
        sr = sts.iloc[:2].id
        self.assertTrue(isinstance(sr, pd.Series))
        self.assertListEqual([1, 2], [x for x in sr])

        # instead of using iloc, get by colname then row, save writing time
        self.assertEqual('PETER', sts.name[0])
        self.assertEqual(1, sts.id[:2][0])

        sr = sts.loc[sts.id <= 2]
        sr = sts.loc[sts.id in (1, 2)]
        sr = sts.loc[~sts.id in (1, 2)]

    def testReadTable(self):
        ''' read table/csv differs only for the tab delimiter
        '''
        frm = pd.read_table(path.join(getpath(), 'res', 'pd_tbl.txt'), encoding='gbk')
        top = frm[frm.id > 5000]
    
    def testCSVRW(self):
        ''' csv's read/write ability
        '''
        d = date(2019, 2, 1)
        df = pd.DataFrame(((1, 'a', Decimal('0.1234'), d), ), columns='id name weight date'.split())
        fh = StringIO()
        df.to_csv(fh, index=None)
        fh = StringIO(fh.getvalue()) # reset the file pointer
        df1 = pd.read_csv(fh, parse_dates=['date'])
        self.assertEqual(d, df1.date[0], 'pd.timestamp can be compared directly to date')
        self.assertEqual(d, df1.date[0].date(), 'timestamp\'s date function')
        self.assertAlmostEqual(0.1234, df1.weight[0])

    def testDummyDF(self):
        ''' empty dataframe, empty/any/all usage
        '''
        df = pd.DataFrame(None, columns='a b c'.split())
        self.assertTrue(df.empty, 'dataframe is empty')
        sr = df.c
        self.assertTrue(sr.empty, 'Series is emptyj')
        df = pd.DataFrame(((None, None, None), ), columns='a b c'.split())
        self.assertFalse(df.empty, 'not empty dataFrame')
        with self.assertRaises(ValueError):
            self.assertFalse(df.any(), 'dataframe does not support any() function')
        with self.assertRaises(ValueError):
            self.assertFalse(df.all(), 'no non-empty data')
        sr = df.iloc[0]
        self.assertFalse(sr.any(), 'no non-empty data')
        self.assertFalse(sr.all(), 'all() function is wrong')
    
    def testModify(self):
        ''' append column assign() from existing column
        '''
        def _xx(row):
            # type(row) is a DataFrame
            return row.height * 2
        org = self._dates.copy()
        cns = org.columns
        df = org.assign(wgtx2=_xx)
        self.assertEqual(len(cns), len(df.columns) - 1)
        self.assertEqual(len(org.columns), len(cns), 'the original value not changed')
        df = org.assign(x=None)
        self.assertFalse(df.x.any())
        df = org.loc[org.id > 3]
        df.id[0] = 'xyx'
        self.assertEqual('xyx', df.id[0], 'yes, value in the view changed')
        self.assertNotEqual(df.id[0], org.loc[org.id > 3].id[0], 'but the original is not changed')
        df.loc[0, 'id'] = 'xyx'
        self.assertEqual('xyx', df.id[0], 'using loc can change it without warning')


class MplSuite(TestCase):
    ''' matplotlib tests
    '''

    @classmethod
    def setUpClass(cls):
        super().setUpClass()
        cls._show_plt = True
    
    def _show(self):
        if self._show_plt:
            plt.show()
        plt.close()

    def testQuickStart(self):
        ''' one axes set only
        '''
        ax = plt.subplot()
        t1 = np.arange(0.0, 2.0, 0.1)
        t2 = np.arange(0.0, 2.0, 0.01)
        l1, = ax.plot(t2, np.exp(-t2))
        l2, l3 = ax.plot(t2, np.sin(2 * np.pi * t2), '--o', t1, np.log(1 + t1), '.')
        l4, = ax.plot(t2, np.exp(-t2) * np.sin(2 * np.pi * t2), 's-.')
        l5, = ax.plot((0, 0.1, 0.3, 0.5, 0.1), label='a line')
        ax.legend((l2, l4, l5), ('oscillatory', 'damped', 'tuple'), loc='upper right', shadow=True)
        ax.set_xlabel('time')
        ax.set_ylabel('volts')
        ax.set_title('Damped oscillation')
        self._show()
    
    def testSubplotWithId(self):
        ''' invoke several subplot(), see the output
        '''
        # calling subplot() several times will have only one axes on top, so
        # need to do more
        # 2x2 can hold up to 4, so 224 is the upper limit
        for i in range(221, 225):
            ax = plt.subplot(i)
            print(id(ax))
            ax.plot([1, 2, 3])
            ax.set_xlabel('%d of' % i)
        with self.assertRaises(ValueError):
            ax = plt.subplot(226)
        self._show()

    def testsubplots(self):
        ''' 2x2 subplots, flat, constrained_layout
        '''
        fig, axs = plt.subplots(2, 2, constrained_layout=True)
        idx = 0
        # don't use "for ax in axs" because it return narray only, flat do it
        for ax in axs.flat:
            if idx == 2:
                ax.plot((1, 2, 3, 4, 5), (1, 2, 5, 3, 2), '--x', label='Line%d' % idx)
                ln = ax.plot((1, 2, 5, 3, 2), '--^', label="won't be shown")
                ax.legend((ln[0], ), ('use ln[0], not ln',))
                ax.set_xlabel('sequence') # or plt.xlabel()
                ax.set_ylabel('value') # or plt.ylabel()
                ax.set_title('%d axes' % idx)
            else:
                ax.plot((1, 2, 5, 3, 2), '--o', label='Line%d' % idx)
                ax.legend()
            ax.set_title("%d axes" % idx)
            idx += 1
        fig.suptitle('%d-row X %d-col figure' % axs.shape)
        self._show()
    
    def testMinorGridlines(self):
        fig = plt.figure(1)
        for idx in range(221, 225):
            ax = plt.subplot(idx)
            # below 2 lines can be before or after the grid commands
            for i in range(1, 30, 5):
                x = np.linspace(0, 10, 50)
                plt.plot(x, np.sin(x) * i, label='line %d' % i)
            # grids
            if idx == 221:
                ax.grid(b=True, which='major', color='k', linestyle='-.', alpha=0.8)
                ax.grid(b=True, which='minor', color='r')
                ax.legend()
            else:
                ax.grid(b=True, which='both', linestyle='-.')
            ax.minorticks_on() # must be call for each axes
        plt.show()

        
    def testMultiPage(self):
        ''' create a 2-page 2X2 axes pdf file
        '''
        fn = tempfile.TemporaryFile().name + '.pdf'
        with PdfPages(fn) as pdf:
            for cnt in range(2):
                fig, axs = plt.subplots(2, 2)
                fig.suptitle('page %d' % (cnt + 1))
                for ax in axs.flat:
                    for i in range(1, 30, 5):
                        x = np.linspace(0 + i, 10 + i, 50)
                        ax.plot(x, np.sin(x) * i, label='line %d' % i)
                    ax.grid(b=True, which='both', linestyle='-.')
                    ax.minorticks_on() # must be call for each axes
                pdf.savefig(fig)
        remove(fn)
