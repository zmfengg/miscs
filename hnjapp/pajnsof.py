#! coding=utf-8
'''
* @Author: zmFeng
* @Date: 2018-09-21 14:44:59 
* @Last Modified by:   zmFeng 
* @Last Modified time: 2018-09-21 14:44:59 
classes help to do to JO -> PAJ NSOF actions
'''

from os import path
from tempfile import gettempdir

import cv2
import numpy as np
from cv2 import GaussianBlur
from PIL import Image, ImageFilter

from utilz import getfiles, imagesize

from .common import thispath


class JOImgOcr(object):
    """ class to do OCR
    """

    #_ordbrd = (0.7, 0.1, 1, 0.4); _smpbrd = (0.7, 0.2, 1, 0.45)
    _jn_brds = (0.7, 0.1, 1, 0.45)
    _imgtpls, _tpl_h_w, _imgwss = (None,) * 3
    _mt = cv2.TM_CCOEFF_NORMED

    def sharpen(self, fldr, newfile=True):
        """ generate a blur->thread image from the existing one, saved as oldfn_0.jpg, only the dpi meta data is kept.
        """
        dpi = None
        if path.isdir(fldr):
            fns = getfiles(fldr, ".jpg")
        else:
            fns = (fldr,)
        rst = []
        root = path.dirname(fldr) if newfile else None
        for fn in fns:
            if fn.find("_") >= 0:
                continue
            if not dpi:
                img = Image.open(fn)
                img.load()
                dpi = img.__getstate__()[0].get("dpi")
                img.close()
            img = cv2.imread(fn, 0)
            img = self._sharpen(img)
            if newfile:
                bn = path.splitext(path.basename(fn))
                fn0 = path.join(root, "%s_0%s" % (bn[0], bn[1]))
            else:
                fn0 = fn
            cv2.imwrite(fn0, img)
            # CV2 does not save metadata, while dpi is very important for OCR
            # So use PIL's image to append the DPI
            img = Image.open(fn0, mode="r")
            img.save(fn0, dpi=dpi)
            rst.append(fn0)
        return zip(fns, rst)

    def _sharpen(self, cv2img):
        return cv2.threshold(GaussianBlur(cv2img, (5, 5), 0), 160, 255, cv2.THRESH_BINARY)[1]
    
    def _getdpi(self, pilimg):
        """ get dpi from a PIL image, not cv2 image """
        #img.load()
        try:
            dpi = pilimg.__getstate__()[0].get("dpi")
        except:
            dpi = (200, 200)
        return dpi

    def _findtpl(self, imgsrc, imgtpl, method):
        res = cv2.matchTemplate(imgsrc, imgtpl, method)
        maxv, min_loc, max_loc = cv2.minMaxLoc(res)[1:]
        if maxv < 0.5:
            return
        return min_loc if method in (cv2.TM_SQDIFF, cv2.TM_SQDIFF_NORMED) else max_loc
        
    def extract(self, fn, tarfn=None, showframe=False):
        if not self._imgtpls:
            imgsrc = path.join(thispath, "res")
            self._imgtpls = [cv2.imread(path.join(imgsrc, fn)) for fn in ("JOTpl.jpg", "SmpTpl.jpg")]
            self._tpl_h_w = [x.shape[:-1] for x in self._imgtpls]
            self._imgwss = [cv2.cvtColor(cv2.imread(path.join(imgsrc, fn)), cv2.COLOR_BGR2GRAY) for fn in ("JOWs.jpg", "SmpWs.jpg")]
        imgsrc = cv2.imread(fn)
        #crop it to the JO# area
        brds = self._jn_brds
        hw0 = imgsrc.shape[:-1]
        imgsrc = imgsrc[int(brds[1]*hw0[0]):int(brds[3]*hw0[0]),int(brds[0]*hw0[1]):int(brds[2]*hw0[1])]
        #TODO::read dpi from cv2 instead of image
        dpi = (200, 200)        
        hw0, tarfn, img, idx = imgsrc.shape[:-1], 0, None, 0
        for idx in range(len(self._imgtpls)):
            top_left = self._findtpl(imgsrc, self._imgtpls[idx], self._mt)
            if not top_left:
                continue
            bottom_right = (hw0[1], top_left[1] + self._tpl_h_w[idx][0])
            if idx == 1:
                brds = (((top_left[0] - 55, top_left[1]), (top_left[0] - 5, bottom_right[1])),
                        ((top_left[0] + self._tpl_h_w[idx][1], top_left[1]), bottom_right))                
            else:
                brds = (((top_left[0] + self._tpl_h_w[idx][1], top_left[1]), bottom_right),)
            imgs = [imgsrc[x[0][1]:x[1][1], x[0][0]:x[1][0]] for x in brds]
            if len(brds) > 1:
                img = np.zeros((self._tpl_h_w[idx][0], sum(bd[1][0] - bd[0][0] for bd in brds), 3), np.uint8)
                ttlw = 0
                for it in zip(brds, imgs):
                    w = it[0][1][0] - it[0][0][0]
                    img[:self._tpl_h_w[idx][0], ttlw:w + ttlw] = it[1]
                    ttlw += w
            else:
                img = imgs[0]
            if showframe:
                for y in brds:
                    cv2.rectangle(imgsrc, y[0], y[1], 255, 2)
                cv2.imshow("tempRst",imgsrc)
                cv2.waitKey()
        if img is not None:
            img = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
            img = self._sharpen(img)
            if False and self._imgwss[idx] is not None:
                #can not find a good template, always return (0,0)
                top_left = self._findtpl(img, self._imgwss[idx], self._mt)
                if top_left and any(top_left):
                    img = img[0:top_left[1], 0: top_left[0]]
            bns = path.splitext(path.basename(fn))
            if not tarfn:
                tarfn = path.join(gettempdir(), bns[0] + "_cropped" + bns[1])
            cv2.imwrite(tarfn, img)
            if dpi:
                img = Image.open(tarfn, mode="r")
                img.save(tarfn, dpi=dpi)
        return tarfn
    
    def _optbyPIL(self, img):
        # this don't work, at least I can not get prefered result, use cv2
        return img
        gfs = (ImageFilter.GaussianBlur(20), )  # ImageFilter.MedianFilter(3))
        for fltr in gfs:
            img.filter(fltr)
        return img
