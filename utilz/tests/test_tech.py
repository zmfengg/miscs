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
import tempfile
from argparse import ArgumentParser
from cProfile import Profile
from datetime import datetime
from itertools import islice
import logging
from logging import Logger
from numbers import Number
from os import path, remove
from pstats import Stats
from unittest import TestCase, skip

from bidict import ValueDuplicationError
from bidict import bidict as bd
from sqlalchemy import create_engine
from sqlalchemy.ext.declarative.api import declarative_base
from sqlalchemy.sql.schema import Column, Index, ForeignKey
from sqlalchemy.sql.sqltypes import TIMESTAMP, VARCHAR, Integer, SmallInteger, DECIMAL, DateTime, Float
from sqlalchemy import text

from utilz import ResourceCtx, SessionMgr, getfiles, imagesize
from utilz.win32 import clearTempFiles

_logger = Logger(__name__)
try:
    import pytesseract as tesseract
    from cv2 import (THRESH_BINARY, GaussianBlur, imread,
                     imwrite, threshold)
    from PIL import Image
except ImportError:
    pass

String = VARCHAR

_base = declarative_base()

def logcfg(fn=None):
    '''
    config the loggers
    '''
    if not fn:
        fn = tempfile.TemporaryFile(prefix='test_tech').name + '.log'
    logging.basicConfig(
        level=logging.INFO,
        filename=fn)
    logging.getLogger("sqlalchemy").setLevel(logging.DEBUG)
    logging.getLogger().addHandler(logging.StreamHandler())
    logging.info("log with file(%s)" % fn)


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

class PajItem(_base):
    ''' A Paj's product, use for pajrdrs.PajUPTracker and pajrdrs.PajBomHdlr '''
    __tablename__ = "pajitem"
    id = Column(Integer, primary_key=True, autoincrement=True)
    __table_args__ = (Index('idx_pajinv_pcode', 'pcode', unique=True),)
    pcode = Column(VARCHAR(20))
    docno = Column(VARCHAR(20))
    createdate = Column(TIMESTAMP)
    tag = Column(SmallInteger)

class StoneOutMaster(_base):
    __tablename__ = 'stone_out_master'

    id = Column(Integer, primary_key=True, autoincrement=False)
    name = Column(Integer, nullable=False, name="bill_id")
    isout = Column(SmallInteger, nullable=False, name="is_out")
    joid = Column(Integer, nullable=False, name="jsid")
    qty = Column(DECIMAL, nullable=False)
    filldate = Column(DateTime, nullable=False, name="fill_date")
    packed = Column(SmallInteger, nullable=False, index=True)
    subcnt = Column(SmallInteger, nullable=False)
    workerid = Column(SmallInteger, server_default=text("0"), name="worker_id")
    lastuserid = Column(SmallInteger, server_default=text("0"))
    lastupdate = Column(DateTime)

class StoneOut(_base):
    __tablename__ = 'stone_out'

    id = Column(
        ForeignKey('stone_out_master.id'),
        primary_key=True,
        nullable=False,
        autoincrement=False)
    idx = Column(SmallInteger, primary_key=True, nullable=False)
    btchid = Column(Integer, nullable=False, index=True)
    workerid = Column(SmallInteger, nullable=False, name="worker_id")
    qty = Column(Integer, nullable=False, name="quantity")
    wgt = Column(Float, nullable=False, name="weight")
    checkerid = Column(SmallInteger, nullable=False, name="checker_id")
    checkdate = Column(DateTime, nullable=False, name="check_date")
    joqty = Column(SmallInteger, name="qty")
    printid = Column(Integer)
    lastuserid = Column(SmallInteger, server_default=text("0"))
    lastupdate = Column(DateTime)

class SqlAlchemyURL(TestCase):
    '''
    url for create engine
    https://docs.sqlalchemy.org/en/13/core/engines.html
    '''
    @classmethod
    def setUpClass(cls):
        super().setUpClass()
        fn = tempfile.TemporaryFile('w+t', prefix='cache').name + '.db'
        eng = create_engine("sqlite+pysqlite:///%s" % fn)  #, echo=True)
        PajItem.metadata.create_all(eng)
        cls._sm = SessionMgr(eng)
        logcfg()
        logging.getLogger().info('dbFile was %s' % fn)

    def testTrMgr(self):
        '''
            the add/flush/commit action failed on sybase, don't know if it works well under sqlite. If yes, sth. is wrong with the sybsae dialag.
            After test, rollback/add/add_all/flush/commit works
        '''
        with ResourceCtx(self._sm) as sess:
            sess.rollback()
            sess.query(PajItem.id == 2).delete()
            sess.commit()
            for idx in range(20):
                pi = PajItem()
                pi.docno = 'A1234'
                pi.pcode = '%04d' % idx
                pi.tag = 0
                pi.createdate = datetime.today()
                sess.add(pi)
            lst = []
            for idx in range(20, 50):
                pi = PajItem()
                pi.docno = 'BX1010'
                pi.pcode = 'PC%05d' % idx
                pi.tag = 1
                pi.createdate = datetime.today()
                lst.append(pi)
            sess.add_all(lst)
            sess.commit()

    def testStoneAdds(self):
        '''
        now test the SOM/SO add/commit actions, still works well. So, what's the problem.
        One more thing, when the With....statement was executed, there is no rollback() by SQLAlchemy engine, but in the case of sybase, there is
        a rollback() at the beginning. The behavior differs as below:
            DBMS    Begin   Rollback
            Sqlite  Yes     No
            Sybase  No      Yes
        '''
        lst = []
        lst1 = []
        with ResourceCtx(self._sm) as sess:
            for idx in range(1000000, 1000050):
                som = StoneOutMaster()
                som.id = som.name = idx
                som.lastupdate = som.filldate = datetime.today()
                som.isout = 200
                som.joid = 1010101
                som.lastuserid = 1
                som.packed = som.qty = som.subcnt = som.workerid = 0
                lst.append(som)
                for x in range(10):
                    so = StoneOut()
                    so.id = idx
                    so.idx = x
                    so.workerid = so.joqty = so.printid = so.checkerid = 0
                    so.lastupdate = so.checkdate = datetime.today()
                    so.btchid = 1020304
                    so.qty = 10101
                    so.wgt = 102.34
                    lst1.append(so)
            sess.add_all(lst)
            sess.add_all(lst1)
            sess.commit()
