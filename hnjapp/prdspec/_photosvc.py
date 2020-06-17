'''
#! coding=utf-8
@Author:        zmfengg
@Created at:    2020-06-03
@Last Modified: 2020-06-03 5:16:26 pm
@Modified by:   zmfengg
Service for photo processing for the new style system
'''

from os import path, makedirs, remove
from shutil import move, copy
from tempfile import gettempdir
from xml.etree import ElementTree
from zipfile import ZipFile
from itertools import chain

from hnjapp.svcs.misc import StylePhotoSvc
from utilz import NamedLists, triml
from ..common import _logger as logger, config


class StylePhotoSvcX(StylePhotoSvc):
    XML_DM = "http://schemas.openxmlformats.org/drawingml/2006/main"
    XML_BM = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    XML_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    XML_XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
    _insts = {}

    @classmethod
    def getInst(cls, key=None):
        ''' return a pre-config instance
        Args:
            key(String): according to conf.json's key stylephoto.x, it can be one of:
            "h", "i", "new". Here "new" for the new styling system
            when ignored, it will be "h"
        '''
        if not key:
            key = config.get("stylephoto.default")
        key = triml(key)
        if key not in cls._insts:
            cfg = config.get("stylephoto.%s" % key)
            if not cfg:
                return None
            cls._insts[key] = StylePhotoSvc(cfg["root"], cfg["level"])
        return cls._insts.get(key)


    def extract(self, fn, styno):
        '''
        exact the photo inside a worksheet as style photos, but I find out that even using xlsx
        format to store an image and then extract it using unzip method, the photo's resolution
        was changed, excel shrink it. SO, this function is discarded. Of course, using the copy/
        paste method gets a even lower solution one.
        Finally I found out that you can setup a file that won't compress any images by choose
        tools/compress pictures/options and uncheck "Automatically perform basic..."
        So I need to make sure the excel file to be xlsx(infact, open document format), then I'll
        use the zipfile/xml to parse the image out and then copy the image file out.
        '''

        fns = None
        with ZipFile(fn) as zf:
            nls = self._find_drawing(zf)
            if nls:
                fns = self._save_drawings(zf, nls, styno)
        return fns

    def _find_sheet_id(self, zf):
        with zf.open('xl/workbook.xml') as fh:
            doc = ElementTree.fromstring(fh.read())
            nodes = self._get(doc, './a:sheets/a:sheet', 'r:id name state'.split(), {'a': self.XML_BM, 'r': self.XML_REL})
            nmp = {}
            for node in (x for x in nodes if x['state'] != 'hidden'):
                nmp[node['name']] = node['r:id'][3:]
        return nmp['Main']

    def _find_drawing(self, zf):
        nsMp = {'a': self.XML_DM, 'r': self.XML_REL, 'xdr': self.XML_XDR}
        idx = self._find_sheet_id(zf)
        fn = None
        with zf.open('xl/worksheets/_rels/sheet%s.xml.rels' % idx) as fh:
            doc = ElementTree.fromstring(fh.read())
            for node in doc:
                if node.attrib['Type'].split('/')[-1] == 'drawing':
                    fn = node.attrib['Target'].split('/')[-1]
                    break
        if fn:
            fn_xml = 'xl/drawings/%s' % fn
            fn_res = 'xl/drawings/_rels/%s.rels' % fn
            fnMp = {}
            with zf.open(fn_xml) as fh:
                doc = ElementTree.fromstring(fh.read())
                for node in doc.findall('./xdr:twoCellAnchor', nsMp):
                    tl = self._get(node, './xdr:from',
                                   'xdr:row xdr:col'.split(), nsMp, 'data')
                    br = self._get(node, './xdr:to', 'xdr:row xdr:col'.split(),
                                   nsMp, 'data')
                    loc = [[int(x) for x in x.values()] for x in (tl, br)]
                    loc = (tuple(chain(*loc)),)
                    n_n = self._get(node, './xdr:pic/xdr:nvPicPr/xdr:cNvPr',
                                    'name descr'.split(), nsMp)
                    n_n = [n_n[x] for x in 'name descr'.split()]
                    tId = self._get(node, './xdr:pic/xdr:blipFill/a:blip',
                                    'r:embed', nsMp)
                    fnMp[tId] = [n_n, loc]
            with zf.open(fn_res) as fh:
                doc = ElementTree.fromstring(fh.read())
                for node in doc:
                    tId = node.get("Id")
                    if tId in fnMp:
                        fnMp[tId].append((node.get('Target').replace(
                            '..', 'xl'),))
            lst = [x for x in fnMp.values()]
            lst = [tuple(chain(*x)) for x in lst]
            # now lst's element is [ShapeName, orgImgFileName, locationArr, pathInXlsx]
            # sort the results by location so that you know the priority
            lst = sorted(lst, key=lambda x: x[2])
            lst.insert(0, 'shpname orgfn loca fn'.split())
            return NamedLists(lst)
        return None

    def make_name(self, styno, idx):
        '''
        make a good file name of given sty# and index
        Args:
            idx:    0 basic
        '''
        # styno + ("" if idx == 0 else '_%02d' % idx) + '.jpg'
        return styno + ('_%02d' % idx) + '.jpg'

    def _save_drawings(self, zf, nls, styno):
        tmpRoot = gettempdir()
        fns = []
        for idx, nl in enumerate(nls):
            fn0 = zf.extract(nl['fn'], tmpRoot)
            fn = path.join(self._build_root(styno), self.make_name(styno, idx))
            fns.append((fn0, fn))
        same_drv = path.splitdrive(tmpRoot)[0] == path.splitdrive(fn)[0]
        for idx, fn in enumerate(fns):
            # rename existing file if there is, the do the move
            rt = path.dirname(fn[-1])
            if path.exists(fn[-1]):
                move(fn[-1], self._make_bak_fn(fn[-1]))
            if not path.exists(rt):
                makedirs(rt)
            if same_drv:
                move(fn[0], fn[-1])
            else:
                copy(fn[0], fn[-1])
                remove(fn[0])
        return [x[1] for x in fns]

    def _make_bak_fn(self, fn):
        root = path.dirname(fn)
        bn = path.basename(fn)
        bn, ext = path.splitext(bn)
        cnt = -1
        while path.exists(fn):
            cnt += 1
            fn = path.join(root, bn + '_bak%02d' % cnt + ext)
        return fn

    def _get(self, node, xpath, flds, nsMp, get_type='attrib'):
        if isinstance(flds, str):
            flds = (flds,)
        lst = []
        flds0 = flds
        flds = [self._fmt(x, nsMp) for x in flds]
        fc = len(flds)
        for nd in node.findall(xpath, nsMp):
            if get_type == 'attrib':
                if fc > 1:
                    lst.append({flds0[idx]: nd.get(x) for idx, x in enumerate(flds)})
                else:
                    lst.append(nd.get(flds[0]))
            else:
                if fc > 1:
                    mp = {}
                    lst.append(mp)
                    for idx, fld in enumerate(flds):
                        n0 = nd.find(fld)
                        mp[flds0[idx]] = None if n0 is None else n0.text
                else:
                    n0 = nd.find(flds[0])
                    lst.append(None if n0 is None else n0.text)
        return lst if len(lst) > 1 else lst[0]

    @classmethod
    def _fmt(cls, theStr, nsMp):
        strs = theStr.split(':')
        ln = len(strs)
        if ln > 1:
            for idx, s0 in enumerate(strs):
                idx1 = s0.find('/') + 1
                strs[idx] = (
                    s0[:idx1] + '{%(' + s0[idx1:] + ')s}'
                ) if idx1 > 0 else s0 if idx % 2 else '{%(' + s0 + ')s}'
            return ''.join(strs) % nsMp
        return theStr
