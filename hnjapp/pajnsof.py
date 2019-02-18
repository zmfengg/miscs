#! coding=utf-8
'''
* @Author: zmFeng
* @Date: 2018-09-21 14:44:59
* @Last Modified by:   zmFeng
* @Last Modified time: 2018-09-21 14:44:59
classes help to do to JO -> PAJ NSOF actions
'''

import io
from os import (path, remove, rename, makedirs, environ)
import re
import struct
from random import randint
from datetime import date
from tempfile import gettempdir

import numpy as np
import PyPDF2
import pytesseract as tess
from cv2 import (COLOR_BGR2GRAY, GaussianBlur, THRESH_BINARY, TM_CCOEFF_NORMED, TM_SQDIFF, TM_SQDIFF_NORMED,
                 cvtColor, imread, imshow, imwrite, matchTemplate, minMaxLoc, rectangle, threshold, waitKey)
from PIL import Image, ImageDraw, ImageFont

from hnjcore import JOElement
from utilz import (getfiles, lvst_dist)

from .common import (_logger as logger, thispath)


class JOImgOcr(object):
    """
    class to do OCR
    """

    def __init__(self):
        #_ordbrd = (0.7, 0.1, 1, 0.4); _smpbrd = (0.7, 0.2, 1, 0.45)
        self._jn_brds = (0.7, 0.1, 1, 0.45)
        self._imgtpls, _tpl_h_w, _imgwss, _tpl_b = (None,) * 4
        self._mt = TM_CCOEFF_NORMED
        self._jn_invalid = {x for x in "(Cc. |J%“¢<£"}
        self._jn_rpl = dict(x.split(",") for x in "b,6;E,8;),1".split(";"))
        self._dpi, self._rsz_height, self._ocr_using_tiff = (200, 200), 1170, False
        # for test only, 0 = pdf2img, 1 = jpg resize
        self._stage = 0

    def rename(self, pdf_fldr, tar_fldr):
        """
        extract image(jpg) from given pdf_fldr, rename them to JO#
        @param pdf_fldr: the folder contains the pdf files, I will choose the
        file with the most-updated date to rename
        @param tar_fldr: the folder to save the files to
        @return: None if everyting is OK. else A list with 2 list as element,
            first contains the un-renamed files(str),
            second contains the master JOs not found(str).
        """
        if self._stage < 1:
            # slow, in debug mode, don't do this again and again
            fns = self.pdff2img(pdf_fldr, tar_fldr)
        else:
            fns = getfiles(tar_fldr, ".jpg")
        fn_jn, jn_set, jn_fn, wrongs = {}, set(), {}, []
        import socket
        try:
            td = (date.today().strftime("%Y-%m-%d"), "%s@%s" % (environ["USERNAME"], environ["COMPUTERNAME"]))
        except KeyError:
            td = (date.today().strftime("%Y-%m-%d"), socket.gethostname())
        del socket
        for fn in fns:  # the first one is the cover sheet, won't extract anything
            jn = self.img2jn(fn)
            if jn is None:
                jn_set = jn_set.union(self._buildjnlist(fn))
                logger.debug("%d master JO#s extracted", len(jn_set))
                remove(fn)
            else:
                if jn not in jn_fn:
                    fn_jn[fn], jn_fn[jn] = jn, fn
                else:
                    # duplicated items found, let MatchHelper solve it
                    for x in zip(("_0_%s" % jn, "_1_%s" % jn), (jn_fn[jn], fn),):
                        fn_jn[x[1]], jn_fn[x[0]] = x
                    del jn_fn[jn]
        if not jn_set:
            jn_set = set()
        else:
            logger.debug("JO#s are: %s" % ",".join(jn_set))
        for jn in fn_jn.values():
            if jn not in jn_set:
                continue
            jn_set.remove(jn)
            del jn_fn[jn]
        # now jn_fn/jns contains the suspicious JO#s, do matching
        if jn_fn:
            mh = MatchHelper()
            act_cans = mh.solve(tuple(jn_set), tuple(jn_fn))
            if act_cans:
                for x in act_cans:
                    fn_jn[jn_fn[x[1]]] = "_" + x[0]
        jn_set = self._ren(fn_jn, td)
        return (tuple(x[0] for x in wrongs), tuple(jn_set)) if (wrongs or jn_set) else None

    def _ren(self, fn_jn, td):
        ''' do the fn -> jn renaming '''
        ext, jn_lst = None, sorted([x for x in fn_jn.values()])
        jn_idx = dict(zip(jn_lst, range(len(jn_lst))))
        for fn, jn in fn_jn.items():
            if not ext:
                ext = path.splitext(fn)[1]
                fns = path.dirname(fn)
            if self._stage < 2:
                # resize and chop
                img = Image.open(fn)
                img.load()
                o_sz = img.size
                o_sz = (int(o_sz[0] * self._rsz_height / o_sz[1]), self._rsz_height)
                img.thumbnail(o_sz)
                self._chop(img, "%s, %d of %d, by %s" % (td[0], jn_idx[jn] + 1, len(fn_jn), td[1]))
                img.save(fn, dpi=self._dpi)
            logger.debug("file(%s) renamed to %s" % (path.basename(fn), jn + ext))
            rename(fn, path.join(fns, jn + ext))
        return jn_lst


    def pdff2img(self, pdf_fldr, tar_fldr):
        """
        extract the images from folder contains pdf
        """
        pdfs, d0, fn, ptn = getfiles(pdf_fldr, "pdf"), 0, None, re.compile(r"\d{8}")
        var = {x: path.getmtime(x) for x in pdfs if ptn.search(path.basename(x))}
        for fd in var.items():
            if fd[1] > d0:
                d0, fn = fd[1], fd[0]
        fn = path.splitext(path.basename(fn))[0]
        fn = fn[:(fn.find("(") - 1)].strip()
        logger.debug("Begin to extract image, PDF family detected as %s", fn)
        pdfs, fns = sorted([x for x in pdfs if path.basename(x).find(fn) == 0]), []
        for d0 in pdfs:
            fns.extend(self.pdf2img(d0, tar_fldr, len(fns)))
        logger.debug("Totally %d images extracted from pdf family(%s)" % (len(fns), fn))
        return fns

    def crop_pdf_fldr(self, pdf_fldr, tar_fldr):
        """
        crop the pdf in the folder
        """
        fns = self.pdff2img(pdf_fldr, tar_fldr)
        if not fns:
            return None
        lst = []
        for fn in fns:
            fn = self._crop(fn)
            if fn:
                lst.append(fn)
        return lst

    def crop_imgs(self, fldr):
        """
        crop the images(jpg) inside fldr
        """
        fns = getfiles(fldr, ".jpg")
        if not fns:
            return None
        lst = []
        for fn in fns:
            fn = self._crop(fn)
            if fn:
                lst.append(fn)
        return lst

    def _buildjnlist(self, list_fn):
        """
        return a set of JO#s from the provided image file
        """

        img = imread(list_fn)
        img = cvtColor(img, COLOR_BGR2GRAY)
        try:
            img = self._sharpen(img)
            list_fn = path.join(gettempdir(), path.basename(list_fn))
            self._savecv2img(img, list_fn)
            txt = tess.image_to_string(list_fn, lang="eng")
        finally:
            remove(list_fn)
        # method1, the ocr recognizes it as one column
        lsts, stage = [[], []], re.compile(r"[0-9A-Z￥]\d{4,8}")
        for x in txt.split("\n"):
            paused = stage.findall(x)
            if len(paused) < 2:
                continue
            # only at the first 2 columns
            paused = x.split(" ")
            if not stage.search(paused[0]) and stage.search(paused[1]):
                continue
            for y in range(2):
                lsts[y].append(paused[y])
        if not (lsts[0] and len(lsts[0]) == len(lsts[1])):
            lsts, stage, paused = [[], []], -1, True
            for x in txt.split("\n"):
                if not x:
                    continue
                if stage == -1 and x.find("F)") > 0:
                    stage, paused = 0, False
                    continue
                elif stage == 0 and x.find("T)") > 0:
                    stage, paused = 1, False
                    continue
                if stage == -1:
                    continue
                if len(x) < 5:
                    if stage == 1:
                        break
                    paused = True
                    continue
                if paused:
                    continue
                lsts[stage].append(self._parsejo(x))
        txt = set()
        for x in zip(lsts[0], lsts[1]):
            if x[0] == x[1]:
                txt.add(x[0])
            for stage in range(JOElement(x[0]).digit, JOElement(x[1]).digit + 1):
                txt.add(JOElement("%s%d" % (x[0][0], stage)).value)
        return txt

    def sharpen(self, fldr, newfile=True):
        """
        generate a blur->thread image from the existing one, saved as oldfn_0.jpg, only the dpi meta data is kept.
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
                dpi = img.info.get('dpi')
                img.close()
            img = imread(fn, 0)
            img = self._sharpen(img)
            if newfile:
                bn = path.splitext(path.basename(fn))
                fn0 = path.join(root, "%s_0%s" % (bn[0], bn[1]))
            else:
                fn0 = fn
            imwrite(fn0, img)
            # CV2 does not save metadata, while dpi is very important for OCR
            # So use PIL's image to append the DPI
            img = Image.open(fn0, mode="r")
            img.save(fn0, dpi=dpi)
            rst.append(fn0)
        return zip(fns, rst)

    @staticmethod
    def _sharpen(cv2img, mode="normal"):
        # return cv2img
        if not mode or mode == "normal":
            gr, th = (5, 5), (160, 255)
        elif mode == "smpl":
            gr, th = (5, 5), (160, 255)
        else:
            gr, th = (5, 5), (140, 255)
        return threshold(GaussianBlur(cv2img, gr, 0), th[0], th[1], THRESH_BINARY)[1]
        return threshold(threshold(GaussianBlur(cv2img, gr, 0), th[0], th[1], THRESH_BINARY)[1], 127, 255, THRESH_BINARY)[1]

    def _getdpi(self, pilimg):
        """ get dpi from a PIL image, not cv2 image """
        # img.load()
        try:
            pilimg.info.get("dpi")
            #dpi = pilimg.__getstate__()[0].get("dpi")
        except:
            dpi = self._dpi
        return dpi

    @staticmethod
    def _findtpl(imgsrc, imgtpl, method):
        res = matchTemplate(imgsrc, imgtpl, method)
        maxv, min_loc, max_loc = minMaxLoc(res)[1:]
        if maxv < 0.5:
            return None
        return min_loc if method in (TM_SQDIFF, TM_SQDIFF_NORMED) else max_loc

    def _savecv2img(self, img, tarfn, dpi=None):
        imwrite(tarfn, img)
        if not dpi:
            dpi = self._dpi
        img = Image.open(tarfn, mode="r")
        img.save(tarfn, dpi=dpi)
        return tarfn

    def img2jn(self, jofn, tarfn=None):
        """
        provide the raw jophoto, return the JO# candidiates(as tuple) of it
        if tarfn is provided, the result image will be saved to that file
        @param:
        """
        flag = tarfn is None
        try:
            # in the case of new sample image, when the size >= 500K, it's a very
            # bad-quantity image(for example, B71586), I should optimize it first
            if path.getsize(jofn) > 1024 * 1024:
                img = imread(jofn)
                img = cvtColor(img, COLOR_BGR2GRAY)
                img = self._sharpen(img, "verybad")
                self._savecv2img(img, jofn)
            tarfn = self._crop(jofn, tarfn=tarfn)
            if not tarfn:
                return None
            tarfn, is_smp = tarfn
            jn = self._parsejo(tess.image_to_string(tarfn, lang="hnx" if is_smp else "eng"))#maybe a --psm 7 should be appended
            if not jn and is_smp:
                tarfn = self._crop(jofn, tarfn=tarfn, sharp_mode="verybad")[0]
                jn = self._parsejo(tess.image_to_string(tarfn, lang="eng"))
            return jn
        finally:
            # FIXME::
            flag = False
            if flag and tarfn and path.exists(tarfn):
                remove(tarfn)

    def _crop(self, srcfn, **kwds):
        """
        @param srcfn: the source image file
        @param tarfn: the target file to write to
        @param showui: show the cropping result
        @param sharp_mode: can be Noe of "normal/smpl/bad", default is normal
        crop the JO# part out from the JO image, return tuple(tar_fn:str, smpFlag:bool)
        """
        tarfn, showui, sharp_mode = tuple((kwds.get(x, None) for x in "tarfn,showui,sharp_mode".split(",")))
        if not self._imgtpls:
            imgsrc = path.join(thispath, "res")
            self._imgtpls = [imread(path.join(imgsrc, fn)) for fn in ("JOTpl.jpg", "SmpTpl.jpg")]
            self._tpl_b = imread(path.join(imgsrc, "B.jpg"))
            self._tpl_h_w = [x.shape[:-1] for x in self._imgtpls]
            self._imgwss = [cvtColor(imread(path.join(imgsrc, fn)), COLOR_BGR2GRAY)
                            for fn in ("JOWs.jpg", "SmpWs.jpg")]
        imgsrc = imread(srcfn)
        # crop it to the JO# area
        brds = self._jn_brds
        hw0 = imgsrc.shape[:-1]
        imgsrc = imgsrc[int(brds[1]*hw0[0]):int(brds[3]*hw0[0]), int(brds[0]*hw0[1]):int(brds[2]*hw0[1])]
        hw0, img, idx = imgsrc.shape[:-1], None, 0
        for idx in range(len(self._imgtpls)):
            top_left = JOImgOcr._findtpl(imgsrc, self._imgtpls[idx], self._mt)
            if not top_left:
                continue
            bottom_right = (hw0[1], top_left[1] + self._tpl_h_w[idx][0])
            if idx == 1:
                x = (top_left[0] - 55, top_left[1]), (top_left[0] - 5, bottom_right[1])
                # if there is a B, use prefix
                img = self._sharpen(imgsrc[x[0][1]:x[1][1], x[0][0]:x[1][0]], "smpl")
                if not JOImgOcr._findtpl(img, self._tpl_b, self._mt):
                    brds = (((top_left[0] + self._tpl_h_w[idx][1], top_left[1]), bottom_right),)
                else:
                    brds = (x, ((top_left[0] + self._tpl_h_w[idx][1], top_left[1]), (bottom_right[0] - 80, bottom_right[1])))
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
            break
        if img is None:
            return None
        if showui:
            for y in brds:
                rectangle(imgsrc, y[0], y[1], 255, 2)
            imshow("tempRst", imgsrc)
            waitKey()
        img = cvtColor(img, COLOR_BGR2GRAY)
        if idx > 0 and (not sharp_mode or sharp_mode == "normal"):
            sharp_mode = 'smpl'
        #logger.debug("file(%s) sharpen using mode %s" % (srcfn, sharp_mode))
        img = self._sharpen(img, sharp_mode)
        if False and self._imgwss[idx] is not None:
            # can not find a good template, always return (0,0)
            top_left = JOImgOcr._findtpl(img, self._imgwss[idx], self._mt)
            if top_left and any(top_left):
                img = img[0:top_left[1], 0: top_left[0]]
        if not tarfn:
            brds = path.splitext(path.basename(srcfn))
            tarfn = path.join(gettempdir(), brds[0] + "_cropped" + brds[1])
        # read dpi from cv2 instead of image?
        self._savecv2img(img, tarfn)
        return tarfn, idx > 0

    def _parsejo(self, txt):
        lst, idx, s0 = [], 0, [self._jn_rpl.get(x, x) for x in txt if x not in self._jn_invalid]
        for ch in s0:
            if ch == '¥':
                ch = "Y"
            if not ('A' <= ch <= 'Z' or '0' <= ch <= '9'):
                continue
            if idx == 0:
                if ch in ('S', 'C', 'M'):
                    continue
                elif ch in ('2', '3', '8', '6'):
                    ch = "B"
            if not ch:
                continue
            lst.append(ch)
            idx += 1
        lst = "".join(lst)
        # sometimes Y will be treated as '¥'Y+
        if lst[:2] == 'YY':
            return lst[1:]
        return lst

    """
    def _optbyPIL(self, img):
        # this don't work, at least I can not get prefered result, use cv2
        return img
        gfs = (ImageFilter.GaussianBlur(20), )  # ImageFilter.MedianFilter(3))
        for fltr in gfs:
            img.filter(fltr)
        return img
    """

    def pdf2img(self, fn, tar_fldr, start=0):
        r"""
        Thanks to https://stackoverflow.com/questions/2641770/extracting-image-from-pdf-with-ccittfaxdecode-filter
        Links:
        PDF format: http://www.adobe.com/content/dam/Adobe/en/devnet/acrobat/pdfs/pdf_reference_1-7.pdf
        CCITT Group 4: https://www.itu.int/rec/dologin_pub.asp?lang=e&id=T-REC-T.6-198811-I!!PDF-E&type=items
        Extract images from pdf: http://stackoverflow.com/questions/2693820/extract-images-from-pdf-without-resampling-in-python
        Extract images coded with CCITTFaxDecode in .net: http://stackoverflow.com/questions/2641770/extracting-image-from-pdf-with-ccittfaxdecode-filter
        TIFF format and tags: http://www.awaresystems.be/imaging/tiff/faq.html
        """
        def tiff_header_for_CCITT(width, height, img_size, CCITT_group=4):
            tiff_header_struct = '<' + '2s' + 'h' + 'l' + 'h' + 'hhll' * 8 + 'h'
            return struct.pack(tiff_header_struct,
                               b'II',  # Byte order indication: Little indian
                               42,  # Version number (always 42)
                               8,  # Offset to first IFD
                               8,  # Number of tags in IFD
                               256, 4, 1, width,  # ImageWidth, LONG, 1, width
                               257, 4, 1, height,  # ImageLength, LONG, 1, lenght
                               258, 3, 1, 1,  # BitsPerSample, SHORT, 1, 1
                               259, 3, 1, CCITT_group,  # Compression, SHORT, 1, 4 = CCITT Group 4 fax encoding
                               262, 3, 1, 0,  # Threshholding, SHORT, 1, 0 = WhiteIsZero
                               273, 4, 1, struct.calcsize(tiff_header_struct),  # StripOffsets, LONG, 1, len of header
                               278, 4, 1, height,  # RowsPerStrip, LONG, 1, lenght
                               279, 4, 1, img_size,  # StripByteCounts, LONG, 1, size of image
                               0  # last IFD
                               )
        fns = []
        with open(fn, 'rb') as fh:
            rdr = PyPDF2.PdfFileReader(fh)
            for page in rdr.pages:
                xobjs = page['/Resources']['/XObject'].getObject()
                for obj in [xobjs[x] for x in xobjs if xobjs[x]['/Subtype'] == '/Image']:
                    data, img_name = None, None
                    if obj['/Filter'] == '/CCITTFaxDecode':
                        data, o_sz = obj._data, (obj['/Width'], obj['/Height'])  # getData() does not work for CCITTFaxDecode
                        data = tiff_header_for_CCITT(o_sz[0], o_sz[1], len(
                            data), (4 if obj['/DecodeParms']['/K'] == -1 else 3)) + data
                        if self._ocr_using_tiff:
                            img_name = '.tiff'
                        else:
                            img_name = '.jpg'
                    elif obj["/Filter"] == '/DCTDecode':
                        # jpeg directly
                        img_name = ".jpg"
                        data = obj._data
                    if not img_name:
                        continue
                    start, o_sz, img_name = start + 1, img_name, path.join(tar_fldr, "%04d" % start) + img_name
                    if not path.exists(path.dirname(img_name)):
                        makedirs(path.dirname(img_name))
                    if o_sz == ".tiff":
                        with open(img_name, 'wb') as img:
                            img.write(data)
                    elif o_sz == ".jpg":
                        img = Image.open(io.BytesIO(data)).convert("RGB")
                        img.save(img_name, dpi=self._dpi)
                    fns.append(img_name)
        return fns

    def _chop(self, img, txt):
        ref_sz = 1654, 2340
        org_pt = 80, 120  # upper-right corner of the textbox, under ref_sz
        iz = img.size

        def _scale(val, direct=None):
            idx = 0 if direct == 'x' else 1
            return int(1.0 * val / ref_sz[idx] * iz[idx])
        cav = ImageDraw.Draw(img)
        fnt = ImageFont.truetype("arial.ttf", _scale(35, "y"))
        tz = cav.textsize(txt, fnt)
        tz = (iz[0] - tz[0] - _scale(org_pt[0], "x"), int(iz[1] - 0.5 * _scale(org_pt[1]) - 0.5 * tz[1]))
        cav.text(tz, txt, fill="black", font=fnt)


class MatchHelper(object):
    """
    don't use dict, use array instead
    """
    _lsts, _cost_array, _costs = (None, ) * 3

    def solve(self, acts, cands):
        if acts and cands and (len(acts) == len(cands)):
            self._lsts = [list(x) for x in (acts, cands)]
            # _costs holds the sorted unique cost
            self._cost_array, self._costs = None, None
            self._index()
            solved_lst = []
            for cost in self._costs:
                self._solve(solved_lst, cost)
            if solved_lst:
                solved_lst = tuple((x[0][1], x[1][1]) for x in solved_lst)
            return solved_lst
        return None

    def _index(self):
        if not self._lsts:
            return
        costs = set()
        self._cost_array = []
        for x in self._lsts[0]:
            row = []
            self._cost_array.append(row)
            for y in self._lsts[1]:
                y = lvst_dist(x, y)
                costs.add(y)
                row.append(y)
        self._costs = sorted(list(costs))

    @classmethod
    def _purify(self, lst, cost, solved_sets):
        return [it if it == cost and idx not in solved_sets[1] else 0 for idx, it in enumerate(lst)]

    def _calc_pendings(self, solved_lst, cost):
        """
        return tuple(pendings, solved_sets) where
        @pendings is tuple(tuple(act, cost), tuple(cand))
        @solved_sets is tuple(set(act_solved), set(cand_solved))
        """
        if solved_lst:
            solved_sets = []
            for idx in range(len(self._lsts)):
                solved_sets.append(set(x[idx][0] for x in solved_lst))
        else:
            solved_sets = (set(), set())
        pendings = tuple((idx, self._purify(it, cost, solved_sets)) for idx, it in enumerate(self._cost_array) if idx not in solved_sets[0])
        pendings = [x for x in ((x[0], int(sum(x[1]) / cost), x[1]) for x in pendings) if x[1] > 0]
        pendings = [sorted(pendings, key=lambda x: "%04d,%04d" % (x[1], x[0]))]
        pendings.append([x for x in range(len(self._cost_array[0])) if x not in solved_sets[1]])
        return pendings, solved_sets

    def _solve(self, solved_lst, cost):
        """
        put result to solved_set based on given cost
        solved_set contains tuple(a,b) where a is act and b is cand
        """
        pendings, solved_sets = self._calc_pendings(solved_lst, cost)
        # after the above, pendings[0] contains only the acts that is
        # not in solved_sets and contains cost == arg.cost
        for x in pendings[0][:]:
            y = x[2] if x[1] == 1 else self._purify(self._cost_array[x[0]], cost, solved_sets)
            y = [idx for idx, it in enumerate(y) if it]
            if len(y) == 1:
                y = tuple([x[0], x[1][x[0]]] for x in zip((x[0], y[0]), self._lsts))
                logger.debug("(%s) => (%s)" % tuple(self._lsts[x][y[x][0]] for x in range(len(self._lsts))))
                solved_lst.append(y)
                for it in zip(solved_sets, y):
                    it[0].add(it[1][0])
                # trim the act array, but don't touch the cand because this might be expensive
                pendings[0].remove(x)
            else:
                break
        if pendings[0]:
            for x in pendings[0]:
                pass
            print("pendings(acts) %s" % pendings[0])
        # TODO what's left need to analyse
