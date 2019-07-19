'''
#! coding=utf-8
@Author:        zmFeng
@Created at:    2019-07-10
@Last Modified: 2019-07-10 8:54:43 am
@Modified by:   zmFeng

excel form hander test

'''

from datetime import datetime
from json import load as load_json
from os import path
from random import randint
from unittest import TestCase

from utilz.miscs import NamedLists, getpath, triml
from utilz.xwu import FormHandler, appmgr, fromtemplate, NrlsInvoker, BaseNrl, SmpFmtNrl

thispath = getpath()

class FormTestSuite(TestCase):

    @classmethod
    def setUpClass(cls):
        super().setUpClass()
        cls._app, cls._tk = appmgr.acq()
        cls._dispose = True
        cls._fmt = triml
        cls._books = []
        cls._data = FormTestSuite._read_data()

    @staticmethod
    def _read_data():
        rt = {}
        mp = None
        none_s = lambda s: None if isinstance(s, str) and s == '_NONE_' else s
        with open(path.join(thispath, 'res', 'formread.json'), encoding="utf-8") as fp:
            mp = load_json(fp)
            if not mp:
                return None
        mp0 = {}
        for k, v in mp['data'].items():
            if isinstance(v, list):
                for x in v[1:]:
                    for idx, s in enumerate(x):
                        s1 = none_s(s)
                        if s1 != s:
                            x[idx] = s1
                mp0[k] = [x for x in NamedLists(v)]
            else:
                if k.find('date') >= 0 or k.find('modified') >= 0:
                    v = datetime.strptime(v, '%Y/%m/%d')
                mp0[k] = none_s(v)
        rt['data'] = mp0
        mp0 = NamedLists([[none_s(y) for y in x.split(',')] for x in mp['fields']])
        rt['fields'] = [x for x in mp0]
        return rt

    @classmethod
    def tearDownClass(cls):
        if cls._dispose and cls._tk:
            if cls._books:
                for wb in cls._books:
                    try:
                        wb.close()
                    except:
                        pass
            appmgr.ret(cls._tk)
        elif cls._app:
            cls._app.visible = True
        super().tearDownClass()

    @property
    def cols(self):
        return FormTestSuite._data['fields']

    @classmethod
    def _new_wb(cls):
        wb = fromtemplate(path.join(thispath, 'res', 'form.xltx'))
        cls._books.append(wb)
        return wb

    @classmethod
    def _open_wb(cls, fn):
        wb = cls._app.books.open(path.join(thispath, 'res', fn))
        cls._books.append(wb)
        return wb

    @property
    def _read_exp(self):
        return FormTestSuite._data['data']

    def testGet(self):
        ''' test for the FormHandler.get() function
        '''
        fh = FormHandler(self._open_wb('FormRead.xlsx'), field_nls=self.cols)
        nd = fh.get('author')
        self.assertEqual('$L$2', nd.range.address, "single item's address")
        self._test_get_table(fh)
        self._test_get_hot_point(fh)

    def _test_get_table(self, rdr):
        # a normal, left to right, top to down table
        tn = 'metal'

        nd = rdr.get(tn)
        self.assertEqual('$B$18',
                         nd.get(1, 'karat').address,
                         'the address of the first karat in %s' % tn)

        self.assertTrue(nd.isTable, '%s is a table' % tn)
        self.assertEqual('$D$20',
                         nd.get(3, 'remarks').address,
                         'the remarks of the 3rd remarks in %s' % tn)
        self.assertEqual(5, nd.maxCount,
                         'the maximum records %s table can hold' % tn)
        self.assertIsNone(nd.get(10, 'remarks'), 'maximum record count reached')
        self.assertEqual('$A$18:$D$22',
                         nd.range.address,
                         'address of the whole  %s table' % tn)

        tn = 'stone'
        nd = rdr.get(tn)
        self.assertEqual('$D$40',
                         nd.get(7, 'name').address,
                         'the addr of the 7th name in %s' % tn)
        self.assertEqual('$A$34:$M$43',
                         nd.range.address,
                         'address of the whole  %s table' % tn)

        # no seqid, detected by border
        tn = 'feature'
        nd = rdr.get(tn)
        self.assertTrue(nd.isTable, '%s is a table' % tn)
        self.assertEqual('$A$47',
                         nd.get(1, 'catid').address,
                         'the address of the first catid in %s' % tn)
        self.assertEqual('$C$53',
                         nd.get(7, 'value').address,
                         'the address of the 7th value in %s' % tn)
        self.assertEqual(7, nd.maxCount,
                         'the maximum records %s table can hold' % tn)
        self.assertIsNone(nd.get(10, 'value'), 'maximum record count reached')
        self.assertEqual('$A$47:$C$53',
                         nd.range.address,
                         'address of the whole  %s table' % tn)

        # to left one
        tn = 'l-type'
        nd = rdr.get(tn)
        self.assertTrue(nd.isTable, '%s is a table' % tn)
        self.assertEqual('$L$47',
                        nd.get(1, 'name').address, 'the 1st name in %s' % tn)
        self.assertEqual('$K$49',
                        nd.get(2, 'sth').address, 'the 2nd Sth. in %s' % tn)
        self.assertEqual('$J$46:$L$49',
                        nd.range.address,
                        'address of the whole  %s table' % tn)

        # to right one
        tn = 'r-type'
        nd = rdr.get(tn)
        self.assertEqual('$E$47',
                         nd.get(1, 'name').address, 'the 1st name in %s' % tn)
        self.assertEqual('$F$48',
                         nd.get(2, 'value').address, 'the 2nd value in %s' % tn)
        self.assertEqual('$E$46:$F$48',
                         nd.range.address,
                         'address of the whole  %s table' % tn)

        # to up one
        tn = 'u-type'
        nd = rdr.get(tn)
        self.assertTrue(nd.isTable, '%s is a table' % tn)
        self.assertEqual('$E$51',
                         nd.get(1, 'name').address, 'the 1st name in %s' % tn)
        self.assertEqual('$F$50',
                         nd.get(2, 'value').address, 'the 2nd value in %s' % tn)
        self.assertEqual('$D$50:$F$51',
                         nd.range.address,
                         'address of the whole  %s table' % tn)

    def _test_get_hot_point(self, rdr):
        # test the hotpoint ability
        g = rdr.get
        e = self.assertEqual
        fmt = FormTestSuite._fmt
        # single of hot point
        self.assertIsNone(g('$K$2', True), 'not in any hot point')
        e(fmt('createdate'), g('$L$3', True), 'createdate')
        # table of hot point
        e((fmt('parts'), 2, fmt('karat')), g('$I$27', True), '2nd karat of parts')
        e((fmt('parts'), 1, fmt('matid')), g('$M$26', True), '1st material Id of parts')
        self.assertIsNone(g('$M$31', True), 'not in any hot point')
        e((fmt('r-type'), 2, fmt('value')), g('$F$48', True), '2nd value of r-type')
        e((fmt('l-type'), 1, fmt('sth')), g('$L$49', True), '1st sth of l-type')
        e((fmt('l-type'), 2, fmt('value')), g('$K$48', True), '2nd value of l-type')

        e((fmt('feature'), 3, fmt('catid')), g('$A$49', True), '3rd catid of feature')
        e((fmt('feature'), 7, fmt('value')), g('$C$53', True), '7th value of feature')

    def testRead(self):
        ''' try to read node from handler, set value
        '''
        wb = FormTestSuite._open_wb('FormRead.xlsx')
        fh = FormHandler(wb, field_nls=self.cols)
        mp = fh.read()[0]
        exp = self._read_exp
        self._map_eq(exp, mp)

    def testReadNR(self):
        ''' read with Nrls specified; logs merged into one;
        '''
        wb = FormTestSuite._open_wb('FormRead.xlsx')
        fh = FormHandler(wb, field_nls=self.cols)
        doc, dsc = [fh.normalize(x) for x in ('docno', 'description')]
        class JENrl(SmpFmtNrl):

            @classmethod
            def je(cls, jn):
                return '%s%04d' % (jn[0], int(jn[1:]))

            def __init__(self, **kwds):
                kwds['name'] = 'JO# Format'
                super().__init__(JENrl.je, **kwds)

        nrls = NrlsInvoker({
            dsc: BaseNrl.get_base_nrm('tu'),
            doc: (BaseNrl.get_base_nrm('tr'), JENrl())
        })
        mp, logs = fh.read((dsc, doc), nrls)
        self.assertEqual('THIS IS A VERY LONG LONG LONG LONG LONG DESCRIPTION ABOUT THIS ITEM', mp[dsc])
        self.assertEqual('Y0015', mp[doc])

        self.assertEqual(2, len(logs), 'one advice was created')
        nl = logs[0]
        self.assertEqual(BaseNrl.LEVEL_RPLONLY, nl.level)
        self.assertEqual('THIS Is A VERY LONG LONG LONG LONG LONG DESCRIPTION ABOUT THIS ITEM ', nl.oldvalue, 'old value of log')
        self.assertEqual('THIS IS A VERY LONG LONG LONG LONG LONG DESCRIPTION ABOUT THIS ITEM', nl.newvalue, 'new value of log')
        self.assertEqual(BaseNrl.get_base_nrm('tu')._name, nl.remarks)
        nl = logs[1] # a merged one, trim and customer formatter
        self.assertEqual(BaseNrl.get_base_nrm('tr')._name + ';JO# Format', nl.remarks)

    def _map_eq(self, exp, mp):
        self.assertEqual(len(exp), len(mp), 'the same count of records')
        for k, v in exp.items():
            if isinstance(v, list):
                r = [x for x in mp[k]]
                self.assertEqual(len(v), len(r), 'count of table(%s)' % k)
                for idx, val in enumerate(v):
                    r1 = r[idx]
                    for cn in val.colnames:
                        self.assertEqual(val[cn], r1[cn], '%s.%s of %d item' % (k, cn, idx))
            else:
                self.assertEqual(v, mp[k], 'value of (%s)' % k)

    def testWrite(self):
        ''' test the write function of FormHandler
        '''
        wb = self._new_wb()
        fh = FormHandler(wb, field_nls=self.cols)
        fh.write(self._read_exp)
        self._app.visible = True
        mp = fh.read()[0]
        self._map_eq(self._read_exp, mp)

        wb = self._new_wb()
        fh = FormHandler(wb, field_nls=self.cols)
        k, val = 'author', 'Name %d' % randint(0, 1023)
        fh.write(name=k, val=val)
        node = fh.get(k)
        self.assertEqual(val, node.value, 'single item')
        fails = fh.write(name=val, val=k)
        self.assertTupleEqual((val, k), fails[0], 'some record could not be written')

        # make sure writing table will erase data already in the table
        node = fh.get('metal')
        node.get(5, 'karat').value = '9K'
        mtls1 = node.get().value
        self.assertEqual('9K', mtls1[4][1])
        mtls = self._read_exp['metal']
        node.value = mtls
        mtls1 = fh.read()[0]['metal']
        for idx, arr in enumerate(mtls):
            self.assertListEqual(arr.data, mtls1[idx].data)

        # the ability to read given names
        mp = fh.read(('metal', 'author'))[0]
        self.assertEqual(2, len(mp))
        mtls1 = mp['metal']
        for idx, arr in enumerate(mtls):
            self.assertListEqual(arr.data, mtls1[idx].data)
