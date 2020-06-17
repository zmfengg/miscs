#! coding=utf-8
'''
* @Author: zmFeng
* @Date: 2019-01-17 10:30:35
* @Last Modified by:   zmFeng
* @Last Modified time: 2019-01-17 10:30:35
* holds the misc services except db ones
'''
import imghdr
from datetime import datetime, timedelta
from email.message import EmailMessage
from email.policy import SMTP
from os import path, remove
from re import compile as compile_r
from tempfile import gettempdir
from zipfile import ZipFile

from PIL import Image
from sqlalchemy import and_
from sqlalchemy.orm import Query

from ._common import jesin
from hnjcore import JOElement
from hnjcore.models.cn import JO, MM, MMMa, Style
from hnjcore.models.hk import JO as JOHK, PajShp
from utilz import NamedList, getvalue, trimu, triml, xwu, na, daterange, splitarray
from utilz.resourcemgr import ResourceCtx
from hnjapp.localstore import Codetable
from ._common import SvcBase

from ..common import _logger as logger, config, Utilz

try:
    from os import scandir
except:
    from scandir import scandir


class StylePhotoSvc(object):
    '''
    service for getting style photo
    '''
    TYPE_STYNO = "styno"
    TYPE_JONO = "jono"
    _insts, _p17_dc = {}, None

    def __init__(self, root=r"\\172.16.8.91\Jpegs\style", level=3):
        self._root = root
        self._min_level = 2
        self._level = max(min(5, level), self._min_level)        

    def _build_root(self, styno):
        parts = [styno[:x] for x in range(self._min_level, self._level + 1)]
        return path.join(self._root, *parts)

    @classmethod
    def default(cls):
        '''
        default instance, can use code from config file if necessary
        '''
        return cls.getInst(config.get("stylephoto.default") or "h")

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

    @classmethod
    def isGood(cls, img, min_grey=240):
        '''
        detect if an given PIL image is a good(dim greater then 1000*1000, top left is near white). Return True if it is
        @param img: An opened PIL image or an URL
        @param min_grey: the min rgb level that will be treated as white
        '''

        open_byme = False
        if isinstance(img, str):
            img, open_byme = Image.open(img), True
            img.load()
        if img.mode != "RGB":
            img = img.convert("RGBA")
        # flag = img.size >= (1000, 1000)
        flag = img.size >= tuple(config.get("stylephoto.good_img.min_dim"))
        if flag:
            # get the bottom-left:
            y = img.height - 3
            for x in range(0, max(30, int(img.width / 2))):
                px = img.getpixel((x, y))[:3]
                # color close to white, but can be non-pure-white(255, 255, 255)
                if [1 for y in px if y < min_grey or y != px[0]]:
                    flag = False
                    break
        if open_byme:
            del img
        return flag

    def getPhotos(self, styno, atype="styno", hints=None, **kwds):
        ''' return the valid photos of given style or jo#
        Args:
            styno: the styno or jono. when it's a jono, atype should be of "jono"
            atype: argument type, can be one of StylePhotoSvc.TYPE_JONO/TYPE_STYNO
        return a list of files sorted by below criterias:
                hints: hit(DESC)
                modified date(DESC)
            hints(String): None or a string using "," as separator
            kwds: engine -> an object help to convert JO# to Sty#
        '''
        if not styno:
            return None
        jo = jono = None
        if atype == self.TYPE_JONO:
            eng = getvalue(kwds, "engine cache_db")
            if not eng:
                # no helper to convert JO# to sty#
                return None
            jono, jo = styno, self._get_jo(styno)
            styno = jo.style.name.value
            # find the JO# with same SKU#
            hints = (hints + "," + jono) if hints else jono
        if hints:
            hints = hints.split(",")
        root = self._build_root(styno)
        if not path.exists(root):
            return None
        styno = trimu(styno)
        ln, lst = len(styno), []
        fns = [
            x for x in scandir(root)
            if x.is_file() and trimu(x.name[:ln]) == styno
        ]
        if not fns:
            return None
        if len(fns) > 1:
            for fn in fns:
                styno = fn.name
                if '0' <= styno[ln] <= '9':
                    continue
                lst.append((
                    fn,
                    self._match(root, styno, ln, hints),
                ))
            if not lst:
                return None
        else:
            return [
                fns[0].path,
            ]
        return [
            x[0].path for x in sorted(lst, key=lambda x: x[1], reverse=True)
        ]

    def _get_jo(self, jono):
        #TODO
        return jono

    def _match(self, root, fn, ln, hints):
        ''' if found in hints, result is positive,
        else return the days to current days as negative
        so that the call can sort the result
        '''
        if hints:
            cand = trimu(fn[ln:fn.rfind('.')])
            if cand:
                for x in hints:
                    if cand.find(trimu(x)) >= 0:
                        return 100
        try:
            cand = datetime.fromtimestamp(path.getmtime(path.join(root, fn)))
        except OSError:
            return -100
        return (cand - datetime.today()).days

    def getJOPhotos(self, jonos, cnsvc=None):
        '''
        return the photo status of given list of jono, for further info., refer to self.get_sample_wo_img()
        @param jonos: can be one of
            .A list JO#(as string), when the argument is of this form, @cnsvc will also be provided
            .A string, same as above
            .A list of (JO#, Sty#, OrdType, Doc#) tuple, this form don't need the cnsvc
        @return:
            a map of below structure:
            {
                "missing": [((jono, styno, ordertype, doc#), Null_Or_Draft_File_Path), ],
                "good": [((jono, styno, ordertype, doc#), Null_Or_Draft_File_Path), ],
                "need_process": [((jono, styno, ordertype, doc#), File_Path_Of_The_Candidate), ]
            }
        '''
        if not isinstance(jonos, (list, tuple)):
            if not isinstance(jonos, str):
                return None
            jonos = (jonos,)
        if not isinstance(jonos[0], (tuple, list)) or len(jonos[0]) != 4:
            if not cnsvc:
                return None
            jonos = cnsvc.getjobyjonos(jonos)[0]
            jonos = {x.style.name.value: x.name.value for x in jonos}
            jonos = [(x[1], x[0]) for x in jonos]
        fn_mp, nl, rst = {}, NamedList(("jono,styno,ordertype,docno".split(","))), {}
        mk_key = lambda x: (x.jono, x.styno, x.ordertype, x.docno)
        ts = datetime(*config.get("stylephoto.good_img.from")).timestamp()
        for jo in jonos:
            # maybe I should cache the result for each style photo based on mtime
            nl.setdata(jo)
            styn = nl.styno
            fns = fn_mp.get(styn, 0)
            if fns == 0:
                fns = self.getPhotos(styn)
                if fns:
                    # photos older than 2016 should be removed from the candidate result
                    flag = [x for x in filter(lambda x: path.getmtime(x) > ts, fns)]
                    if flag:
                        fns = flag
                    fns = sorted(fns, key=path.getsize, reverse=True)
                    if not flag:
                        # no image since 2016, only get the last 5(if there is) to reduce process time
                        fns = fns[:5]
                fn_mp[styn] = fns
            self._classify(mk_key, nl, fns, rst)
        # sort the results by JO#
        rst = {x[0]: sorted(x[1], key=lambda x: x[0]) for x in rst.items()}
        return rst

    @classmethod
    def _classify(cls, mk_key, nl, fns, rst):
        if fns:
            fn = fns[0]
            # for most case, design draft is less than 40K
            if path.getsize(fn) > config.get("stylephoto.max_draft_size"):
                max_sz, mx_fn = (0, 0), None
                for fn in fns:
                    img = Image.open(fn)
                    img.load()
                    if img.size > max_sz:
                        max_sz, mx_fn = img.size, fn
                    flag = cls.isGood(img, 250)
                    del img
                    if flag:
                        break
                if flag:
                    rst.setdefault("good", []).append((mk_key(nl), fn))
                else:
                    rst.setdefault("need_process" if max(max_sz) >= max(config.get("stylephoto.good_img.min_dim")) else "missing", []).append((mk_key(nl), mx_fn, ))
                return
        rst.setdefault("missing", []).append((mk_key(nl),
                                                fns[0] if fns else None))
    @classmethod
    def getCategory(cls, styno, jodsc=None):
        '''
        return the style category, for example, RING/EARRING.
        In the case of bracelet/bangle, providing the style# only won't return accumulate result, you should use getCategory("B1234", "钻石手镯")
        @param styno(str): The style# or a p17code, or a description
        @param jodsc(str): The JO's description(in chinese, gbk or big5)
        '''
        return Utilz.getStyleCategory(styno, jodsc)

class PhotoQueryAsMail(object):
    '''
    class to perform query to db using newSample/Most-Recent/JOList
    then send mail to given list
    '''

    def __init__(self, pssvc, cnsvc=None):
        self._root = gettempdir()
        self._pssvc, self._cnsvc = pssvc, cnsvc

    def q_new_sample(self, date_frm=None, days=20, **mail_hds):
        '''
        Query for New/QC sample of given date range and return an eml file(Outlook(express) can open it)
        @param cnsvc: A CNSvc
        @param date_frm: the start date(included) for the query
        @param days: the days after the date_frm(excluded) for the query
        @param mail_hds: the From/To/Title/Content inside the map, if a "target" is found, the result eml file would be placed there.
        Special items that can be used:
            "ignores":tuple: names inside the tuple won't be built into the msg file
            "hksvc":dbsvcs.HKSvc item help to resolve JO to Pcode
        @return:
            A ready to mail file

        '''
        if not date_frm:
            date_frm = datetime.today() - timedelta(days=30 * 6)
        date_to = date_frm + timedelta(days=days)
        jos = Query(JO).join(MM).join(MMMa).filter(JO.ordertype.in_((
            'N',
            'Q',
        ))).filter(and_(JO.deadline >= date_frm, JO.deadline < date_to)).filter(
            and_(JO.id > 0, JO.tag > -10))
        with self._cnsvc.sessionctx() as cur:
            jos = jos.with_session(cur).all()
            if not jos:
                return None
            jos = self._jos_2_q(jos)
        jos = self._pssvc.getJOPhotos(jos)
        return self._build_msg(jos, **mail_hds)

    def q_new_sample_rt(self, stsm, **mail_hds):
        '''
        send new sample request to related staff. The affected new sample counts from the last date sent to today - 42. if the query is too short, won't send request
        '''
        cd_key = "new_sample_request_rt"
        cdsvc = CDSvc(stsm)
        cd = cdsvc.get(cd_key)
        df = cd[0].date0 if cd else datetime.date(2018, 2, 1)
        if (datetime.today() - df).days < 42 + 30:
            return None
        df, dt = daterange(df.year, df.month, df.day)
        msg = self.q_new_sample(df, (dt - df).days, **mail_hds)
        cd = cd[0] if cd else cdsvc.newInstance()
        if cd.id:
            cd.coded2, cd.coded1 = cd.coded1, cd.coded0
        cd.coded0 = dt
        # TODO:: cdsvc.save(cd)


    def q_jns(self, jonos, **mail_hds):
        r'''
        Query by JO#s and return an eml file(using Outlook can read it)
        @param jonos: Can be one of:
            .A single JO#(String)
            .A tuple if JO#(String)
            .A a file contains JO#. JO#'s pattern(^[A-Z45]\d{4,6}$) will be used to filter the JO#. Now accept txt and excel file. While using excel, a JO# should be of string format.
        @param mail_hds: the From/To/Title/Content inside the map, if a "target" is found, the result eml file would be placed there
        '''
        if isinstance(jonos, str):
            if path.exists(jonos):
                jonos = JOElement.from_file(jonos)
            else:
                jonos = (jonos, )
        with self._cnsvc.sessionctx():
            jonos = self._cnsvc.getjos(jonos)
            if not jonos:
                return None
            fn = None
            if jonos[0]:
                fn = self._jos_2_q(jonos[0])
        if jonos[1]:
            if fn is None:
                fn = []
            for x in jonos[1]:
                fn.append((x, "_InvalidJO#_", "N/A"))
        return self._build_msg(self._pssvc.getJOPhotos(fn), **mail_hds)

    def q_styns(self, stynos, **mail_hds):
        r'''
        Query by Sty#s and return an eml file(using Outlook can read it)
        @param stynos: Can be one of:
            .A single JO#(String)
            .A tuple if JO#(String)
            .A a file contains Sty#. Sty#'s pattern(^[A-Z]{1,2}\d{3,6}$) will be used for Sty# finding. Now accept txt and excel file. While using excel, a Sty# should be of string format
        @param mail_hds: the From/To/Title/Content inside the map, if a "target" is found, the result eml file would be placed there
        '''
        # ptn = compile_r(r"^[A-Z]{1,2}\d{3,6}$")
        ptn = compile_r(config.get("pattern.styno"))
        if isinstance(stynos, str):
            if path.exists(stynos):
                stynos = JOElement.from_file(stynos, vdl=lambda x: isinstance(x, str) and ptn.search(x))
            else:
                stynos = (stynos, )
        with self._cnsvc.sessionctx() as cur:
            mp = {}
            for pt in splitarray(stynos, 20):
                q = Query(JO).join(Style).filter(jesin([JOElement(x) for x in pt], Style))
                q = q.with_session(cur).all()
                if not q:
                    continue
                q = sorted(q, key=lambda jo: (jo.style.name.value, jo.deadline))
                mp.update({x.style.name.value: x for x in q})
            q = self._jos_2_q(mp.values())
        return self._build_msg(self._pssvc.getJOPhotos(q), **mail_hds)

    @staticmethod
    def _jos_2_q(jos):
        '''
        return a (sty#, jo#, ordertype, doc#) tuple
        '''
        mp = {x.style.name.value: (x.name.value, x.ordertype, x.docno) for x in jos}
        return [(x[1][0], x[0], x[1][1], x[1][2]) for x in mp.items()]

    def _build_msg(self, mp, **kwds):
        '''
        create a mail file based on the provided mp.
        @param mp: a map's format as described in StylePhotoSvc.getJOPhoto#jonos
        '''
        t_fns, tr = [], gettempdir()
        t_fns.append(self._build_wb(mp, tr, kwds.get("hksvc")))
        ignores = kwds.get("ignores", [])
        for key in "missing need_process good".split():
            if key in ignores:
                continue
            lst = mp.get(key)
            if not lst:
                continue
            t_fns.append(path.join(tr, key + ".zip"))
            with ZipFile(t_fns[-1], mode="w") as tzip:
                for var in lst:
                    if not var[1]:
                        continue
                    tzip.write(var[1], path.basename(var[1]))
        msg = self._new_msg(**kwds)
        for var in t_fns:
            with open(var, 'rb') as lst:
                img_data = lst.read()
            lst = path.splitext(var)[1].lower()
            # maybe using mimetypes help alot
            if lst in ('jpg', 'png', 'bmp'):
                mtp, stp = "image", imghdr.what(None, img_data)
            elif lst.find("xl") >= 0:
                mtp, stp = 'application', "excel"
            elif lst == "zip":
                mtp, stp = "application", "zip"
            else:
                mtp, stp = "application", "octet-stream"  # "unknown" or "rfc822" also ok
            msg.add_attachment(
                img_data, maintype=mtp, subtype=stp, filename=path.basename(var))
        var = kwds.get("target", path.join(self._root, "out.eml"))
        with open(var, 'wb') as lst:
            lst.write(msg.as_bytes(policy=SMTP))
        for lst in t_fns:
            remove(lst)
        return var

    @staticmethod
    def _doc_2_request(key, docno, jn, hksvc):
        if key != 'missing':
            return ""
        rc = None
        if docno.find("PAJ") >= 0:
            rc = ",向PAJ索图"
            if hksvc:
                with hksvc.sessionctx() as cur:
                    q = cur.query(PajShp.pcode).join(JOHK).filter(JOHK.name == JOElement(jn)).first()
                    if q:
                        rc = rc + "(%s)" % q[0]
        else:
            rc = ",向HK索图(C1)"
        return rc


    def _build_wb(self, mp, root, hksvc=None):
        app, tk = xwu.appmgr.acq()
        fn = None
        try:
            wb, idx, ttl = app.books.add(
            ), 0, "工单号 款号 类型 备注".split()

            for key, var in {
                    "missing": ("无图", "草图"),
                    "good": ("已修图", "高质量图"),
                    "need_process": ("待修图", "请修图")
            }.items():
                lst = mp.get(key)
                if not lst:
                    continue
                # maybe need to know its Paj or C1
                lst1 = [ttl]
                lst1.extend([("'" + JOElement.tostr(x[0][0]), x[0][1], x[0][2],
                              (var[1] if x[1] else "无图") + self._doc_2_request(key, x[0][3], x[0][0], hksvc)) for x in lst])
                if idx == len(wb.sheets):
                    lst = wb.sheets.add()
                lst = wb.sheets[idx]
                lst.cells(1, 1).value, lst.name = lst1, var[0]
                lst.autofit("c")
                xwu.freeze(lst.cells(2, 2))
                idx += 1
            fn = path.join(root, "result")
            wb.save(fn)
            fn = wb.fullname
        finally:
            if wb:
                wb.close()
            if tk:
                xwu.appmgr.ret(tk)
        return fn

    @staticmethod
    def _new_msg(**kwds):
        msg = EmailMessage()
        msg['Subject'] = kwds.get("subject", 'Image query result')
        msg['From'] = kwds.get("from", "zmfeng <zmfeng@hnjhk.com>")
        msg['To'] = kwds.get("to", "hnjchina@gmail.com, zmfeng@hnjhk.com")
        # pack them all into a MIME msg for e-mail
        msg.preamble = msg["Subject"]
        return msg

class CDSvc(SvcBase):
    '''
    codetable service for local store
    '''

    def __init__(self, trmgr):
        super().__init__(trmgr)

    def newInstance(self):
        '''
        create a default instance that is ready to insert
        '''
        cd = Codetable()
        cd.codec0 = cd.codec1 = cd.codec2 = cd.description = na
        cd.coden0 = cd.coden1 = cd.coden2 = cd.tag = 0
        cd.coded0 = cd.coded1 = cd.coded2 = cd.createdate = cd.lastmodified = datetime.today()
        return cd


    def get(self, name):
        '''
        get by name
        '''
        with self.sessionctx() as cur:
            return Query(Codetable).filter(Codetable.name == name).with_session(cur).all()
