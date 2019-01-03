#! coding=utf-8
'''
* @Author: zmFeng
* @Date: 2018-07-04 08:46:52
* @Last Modified by:   zmFeng
* @Last Modified time: 2018-07-04 08:46:52
for python's language/basic facility test, a practice farm
'''

import gettext
import re
from argparse import ArgumentParser
from os import path
from unittest import TestCase, skip
from logging import Logger
from numbers import Number

_logger = Logger(__name__)
try:
    import pytesseract as tesseract
    from cv2 import (ADAPTIVE_THRESH_GAUSSIAN_C, ADAPTIVE_THRESH_MEAN_C,
                     THRESH_BINARY, GaussianBlur, adaptiveThreshold, imread,
                     imwrite, threshold)
    from PIL import Image
except ImportError:
    pass

from utilz import getfiles, imagesize


@skip("TODO::")
class TesseractSuite(TestCase):
    #_srcfldr = r"p:\aa\x\org\jophotos"
    _srcfldr = r"p:\aa\x\org1\Smp"
    _cropfldr = r'd:\temp\crop'
    _ordbrd = (0.75, 0.1, 1, 0.2)
    _smpbrd = (0.75, 0.2, 1, 0.45)

    def testCrop_Gray(self):
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
            if False:
                th2 = adaptiveThreshold(img, 255, ADAPTIVE_THRESH_MEAN_C, THRESH_BINARY, 11, 2)
                th3 = adaptiveThreshold(img, 255, ADAPTIVE_THRESH_GAUSSIAN_C, THRESH_BINARY, 11, 2)
            for x in (th1,):
                fn0 = path.join(fldr, "%s_%d%s" % (bn[0], cnt, bn[1]))
                imwrite(fn0, x)
                # because CV2 does not save metadata, while dpi is very important
                # use PIL's image to process it
                img = Image.open(fn0, mode="r")
                img.save(fn0, dpi=dpi)
                cnt += 1

    def testOCR(self):
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

    def testParse(self):
        pass


@skip("still don't know how to make good use of it")
class ArgParserTest(TestCase):
    """
    test for the argument parser
    After many tests, know
    """

    def testSingle(self):
        ap = ArgumentParser("testPrg")  # , "usage of what?", "program try the argument parser", add_help=True)
        ap.add_argument("-w", "--date1[,date2]", default="def_x")
        ap.parse_args(["-h"])
        return
        #ap.add_argument("kill", default="def_bill")
        print(ap.parse_args(["-w", "kk"]))
        #print(ap.parse_args(["-xxx", "kk"]))
        # print(ap.parse_args(["-h"]))
        """
        print(ap.parse_args(["kill"]))
        print(ap.parse_args(["-w", "what what"]))
        print(ap.parse_args(["-w"]))
        print(ap.parse_args(["-w", "kill"]))
        """

    def testGetText(self):
        gettext.bindtextdomain('utilz', r'd:\temp\abx')
        gettext.textdomain('utilz')
        _ = gettext.gettext
        print(_('This is a translatable string.'))
        print("hello")


class TechTests(TestCase):
    """
    class trying the technical test
    """

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
            @classmethod
            def sta(cls):
                return "sta"

            def inst(self):
                return "inst"

        class B(A):
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
        ''' all accesses are controlled __getattribute__(), below is the access proority(high to low):
        .class property
        .data descriptor
        .instance property
        .non-data descriptor
        .__getattr__() method
        .AttributeError raised
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
                instance.fxfx = value

        class Foo(object):
            cls_prop = "cls_prop"
            data_dsc = NullDev("clz_level")
            def __init__(self):
                self.inst_prop = "inst_prop"                
                self.data_dsc1 = NullDev("inst_level")

            def doit(self):
                return "hello"

        fo = Foo()
        # descriptor's setter has instance, no owner because if you assign value to a descriptor in class
        # level, it will be killed
        Foo.data_dsc = 'x'
        self.assertEqual('x', Foo.data_dsc)
        self.assertEqual('x', fo.data_dsc)
        Foo.data_dsc = NullDev("clz_level")
        # it's restored
        self.assertFalse(hasattr(fo, 'fxfx'))
        # fo's fxfx property created by the __set__() method
        fo.data_dsc = 'x'
        self.assertEqual('x', fo.fxfx)
        self.assertTrue(fo.data_dsc != 'x', "it's not replaced, just the __set__() was called")
        fo.data_dsc1 = 'x'
        self.assertTrue('x' == fo.data_dsc1, "descriptor as instance property will be override")
        # so, don't use descriptor in instance, just in class
        fo.doit = "y"
        self.assertEqual("y", fo.doit)
        del fo.doit
        self.assertEqual("hello", fo.doit())

        

        mi = ManyInterfaces()
        self.assertListEqual([1, 2, 3], list(mi._lst_data), 'defined property')
        self.assertEqual("_data_len", mi.data_len, 'by descriptor, which overide any __getattxx__')
        mi.data_len = _LenDescriptor()
        self.assertEqual("_data_len", mi.data_len, 'by descriptor, which overide any __getattxx__')
        self.assertEqual('__getattribute__(y)', mi.y, 'by __getattribute__()')
        self.assertEqual('__getattr__(z)', mi.z, 'by __getattr__()')
        with self.assertRaises(AttributeError, msg='k not defined, and no __getattribute__/__getattr__ reponse') as err:
            print(mi.k)
        self.assertEqual('attribute k not defined in __getattr__()', err.exception.args[0])

    def testComp(self):
        a = 3
        self.assertTrue(1 < a < 5) #continuous comparison
        self.assertTrue(not a == 5) # same as "not (a == 5)" because not has lower priority in non-logical operation
        # find element inside sequence. sequence types are: tuple, list, range and the descestor.
        self.assertTrue('a' in 'abcea') # find sub-string in string
        self.assertTrue(1 in (2, 3, 1)) # find element
        self.assertTrue(1 in {1: 'a', 2: 'b'})
        self.assertEqual('T', a == 3 and 'T' or 'F') # same as 'T' if a == 3 else 'F'
        self.assertEqual('F', a != 3 and 'T' or 'F')
        self.assertTrue(isinstance(a, Number))
        self.assertTrue(isinstance(a, int))
        self.assertFalse(type(a) is type(object))
        self.assertTrue(type(a) is type(0))


    def testManyItf(self):
        ''' ManyInterfaces class implements many built-in interfaces for study purpose '''
        mi = ManyInterfaces(f="k")
        # can make use of an iterator object without iter() function
        # self.assertListEqual([1, 2, 3], [x for x in iter(mi)])
        self.assertListEqual([1, 2, 3], [x for x in mi])
        # can not next() because the internal _iter not inited by the __iter__() method
        mi = ManyInterfaces()
        with self.assertRaises(AttributeError, msg='containor not activated'):
            next(mi)
        mi = ManyInterfaces()
        self.assertEqual(1, iter(mi).__next__())
        mi += (2, 3, 4)
        self.assertListEqual([1, 2, 3, 2, 3, 4], mi.data, 'inplace add')
        mi = ManyInterfaces()
        self.assertEqual(1, mi.next(), 'the generator')

class _LenDescriptor():
    ''' this is called data-descriptor, method/function,
    on the other hand, is called non-data-descriptor
    '''
    def __get__(self, instance, owner):
        return len(instance._lst_data)
    
    def __set__(self, instance, value):
        raise AttributeError("instance's len property can not be set")

class ManyInterfaces(object):
    """ this class try to implement many system-level interface for practice purpose """
    data_len = "_data_len"
    
    def __init__(self, *args, **kwds):
        ''' arguments as list and named-map '''
        # object does not support constructor arguments super().__init__(*args, **kwds)
        _logger.debug("args and kwds are: %s, %s" % (args, kwds))
        self._auto_iter = "auto_iter" in kwds
        self._lst_data = list(kwds.get("data", (1, 2, 3)))
        self._iter = None

    @property
    def data(self):
        ''' the internal list '''
        return self._lst_data

    def __iter__(self):
        ''' __iter__ along with __next__ make an object iterable, that is, can be
        access using iter() method
        '''
        self._iter = iter(self._lst_data)
        return self

    def __next__(self):
        ''' second interface of iterator '''
        if not self._iter and self._auto_iter:
            self._iter = iter(self._lst_data)
        return self._iter.__next__()

    def next(self):
        ''' a generator, acts just like a iterator '''
        for x in self._lst_data:
            yield x

    def __iadd__(self, other):
        ''' iadd/isub and so on, are inside operator module, measn in-place operation, '''
        if not isinstance(other, (tuple, list)):
            raise AttributeError("should provide tuple or list")
        if not self._lst_data:
            self._lst_data = []
        self._lst_data.extend(other)
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
