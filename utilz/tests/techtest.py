#! coding=utf-8
'''
* @Author: zmFeng
* @Date: 2018-07-04 08:46:52
* @Last Modified by:   zmFeng
* @Last Modified time: 2018-07-04 08:46:52
for python's language/basic facility test, a practice farm
'''

import re
from os import path
from unittest import TestCase, skip
import gettext

import pytesseract as tesseract
from PIL import Image, ImageFile

from utilz import getfiles, imagesize
from argparse import ArgumentParser


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
        import cv2
        from cv2 import GaussianBlur

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
            img = cv2.imread(fn, 0)
            img = GaussianBlur(img, (5, 5), 0)
            th1 = cv2.threshold(img, 160, 255, cv2.THRESH_BINARY)[1]
            fldr, bn, cnt = path.dirname(srcfn), path.splitext(path.basename(fn)), 0
            if False:
                th2 = cv2.adaptiveThreshold(img, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY, 11, 2)
                th3 = cv2.adaptiveThreshold(img, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 11, 2)
            for x in (th1,):
                fn0 = path.join(fldr, "%s_%d%s" % (bn[0], cnt, bn[1]))
                cv2.imwrite(fn0, x)
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

class ArgParserTest(TestCase):
    """
    test for the argument parser
    After many tests, know 
    """
    def testSingle(self):
        ap = ArgumentParser("testPrg") #, "usage of what?", "program try the argument parser", add_help=True)
        ap.add_argument("-w", "--date1[,date2]", default="def_x")
        ap.parse_args(["-h"])
        return
        #ap.add_argument("kill", default="def_bill")
        print(ap.parse_args(["-w", "kk"]))
        #print(ap.parse_args(["-xxx", "kk"]))
        #print(ap.parse_args(["-h"]))
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
