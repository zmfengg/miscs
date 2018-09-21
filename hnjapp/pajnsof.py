#! coding=utf-8 
'''
* @Author: zmFeng 
* @Date: 2018-09-21 14:44:59 
* @Last Modified by:   zmFeng 
* @Last Modified time: 2018-09-21 14:44:59 
classes help to do to JO -> PAJ NSOF actions
'''

import cv2
from PIL import Image, ImageFilter
from cv2 import GaussianBlur
from tempfile import gettempdir
from utilz import imagesize, getfiles
from os import path

class JOImgOcr(object):
    """ class to do OCR
    """

    _ordbrd = (0.75, 0.1, 1, 0.2)
    _smpbrd = (0.75, 0.2, 1, 0.45)

    def beautify(self, fldr, newfile = True):
        """ generate a blur->thread image from the existing one, saved as oldfn_0.jpg, only the dpi meta data is kept.
        """
        dpi = None
        if path.isdir(fldr):
            fns = getfiles(fldr,".jpg")
        else:
            fns = (fldr,)
        rst = []
        root = path.dirname(fldr) if newfile else None
        for fn in fns:
            if fn.find("_") >= 0: continue
            if not dpi:
                img = Image.open(fn)
                img.load()
                dpi = img.__getstate__()[0].get("dpi")
                img.close()
            img = cv2.imread(fn,0)            
            img = GaussianBlur(img, (5,5), 0)
            img = cv2.threshold(img, 160, 255, cv2.THRESH_BINARY)[1]
            if newfile:
                bn = path.splitext(path.basename(fn))
                fn0 = path.join(root, "%s_0%s" % (bn[0], bn[1]))
            else:
                fn0 = fn
            cv2.imwrite(fn0, img)
            #CV2 does not save metadata, while dpi is very important for OCR
            #So use PIL's image to append the DPI
            img = Image.open(fn0, mode = "r")
            img.save(fn0, dpi = dpi)
            rst.append(fn0)
        return zip(fns,rst)
    
    def crop(self, fn):    
        """ crop the source file, save to the temp folder
        """
        #is it possible to detect the border?
        brd = []
        for ii in range(len(self._smpbrd)):
            brd.append(min(self._smpbrd[ii],self._ordbrd[ii]) if ii < 2 else max(self._smpbrd[ii],self._ordbrd[ii]))
        orgsz = imagesize(fn)
        img = Image.open(fn)
        box = (orgsz[0] * brd[0], orgsz[1] * brd[1], orgsz[0]*brd[2], orgsz[1] * brd[3])
        img.load()
        try:
            dpi = img.__getstate__()[0].get("dpi")
        except:
            dpi = (200,200)
        img1 = img.crop(box)
        img.close()
        img = img1
        tfn = path.join(gettempdir() ,path.basename(fn))
        img.save(tfn, dpi = dpi)
        img.close()
        self.beautify(tfn)
        return tfn
    
    def optbyPIL(self, img):
        #this don't work, at least I can not get prefered result, use cv2
        return img
        gfs = (ImageFilter.GaussianBlur(20), )#ImageFilter.MedianFilter(3))
        for fltr in gfs:
            img.filter(fltr)
        return img