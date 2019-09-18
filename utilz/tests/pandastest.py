'''
#! coding=utf-8
@Author:        zmFeng
@Created at:    2019-09-16
@Last Modified: 2019-09-16 8:19:24 am
@Modified by:   zmFeng
pandas usages
'''

from datetime import date
from decimal import Decimal
from io import StringIO
from os import path
from unittest import TestCase

import numpy as np
import pandas as pd

from utilz.miscs import getpath, trimu


class _Base(TestCase):
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


class IndexSuite(_Base):
    ''' index is everywhere, axes of Series/DataFrame ...
    '''

    def testCreate(self):
        ''' create an index instance
        '''
        idx = pd.Index(range(10))
        self.assertEqual(10, len(idx))
        self.assertEqual(3, idx.get_loc(3))
        idx = pd.Index(tuple('abcdefg'))
        self.assertEqual(1, idx.get_loc('b'))
    
    def testAccess(self):
        ''' access by default indexer by numeric or get_loc() by name
        '''
        idx = pd.Index(tuple('abcdefg'))
        self.assertEqual(len('abcdefg'), len(idx))
        self.assertEqual('a', idx[0], 'by indexer')
        self.assertEqual(0, idx.get_loc('a'), 'by name')
        idx = pd.Index(pd.date_range(date(2018, 1, 1), periods=5, freq='d'))
        self.assertEqual(date(2018, 1, 1), idx[0])
        self.assertEqual(0, idx.get_loc(date(2018, 1, 1)), "name does not have to be a string")
        self.assertTrue(idx.contains(date(2018, 1, 1)))

class SeriesSuite(_Base):
    ''' try pandas out
    '''

    def testCreate(self):
        ''' ways to create a Series
        '''
        sr = pd.Series((1, 2, 3, 4), index=tuple('abcd'))
        self.assertEqual(4, len(sr))

    def testIndexer(self):
        ''' loc[], iloc[] and [] indexer for Series
        '''
        sr = pd.Series((1, 2, 3, 4), index=tuple('abcd'))
        self.assertEqual(4, len(sr), 'length of a Series')
        self.assertEqual(1, sr['a'], 'default indexer')
        self.assertEqual(1, sr.a, 'attribute access')
        self.assertEqual(1, sr.loc['a'], 'loc access by column name')
        self.assertEqual(1, sr.iloc[0], 'loc access by index')
        self.assertTrue(sr[['a', 'b']].equals(pd.Series([1, 2], index=list('ab'))), 'select more than one column using default indexer, return series')
        self.assertTrue(sr.loc[['a', 'b']].equals(pd.Series([1, 2], index=list('ab'))), 'select more than one column using loc, return series')
        sr = pd.Series((1, 2, 3, 4))
        with self.assertRaises(KeyError):
            print(sr[('a', 'b')], "don't use tuple, use list instead")
        with self.assertRaises(KeyError):
            print(sr['a'])
        with self.assertRaises(KeyError):
            print(sr['0'])
        self.assertEqual(1, sr[0], 'loc access by column name/index is the same in a non-name Series')


class DataFrameSuite(_Base):
    ''' the data frame suite
    '''

    def testCreate(self):
        ''' create from array with/without index/name
        '''
        df = pd.DataFrame(None, columns=tuple('abcd'))
        self.assertEqual(0, len(df), 'blank frame with column names')
        # without column, range() as default index
        df = pd.DataFrame(np.random.random((6, 4)))
        self.assertEqual(6, len(df), 'row count')
        self.assertEqual(4, len(df.iloc[0]), 'column count')
        # with columns, range() as default index
        df = pd.DataFrame(np.random.random((6, 4)), columns=tuple('abcd'))
        self.assertTrue(0 in df.d.index)
        self.assertTrue(6 not in df.d.index)
        self.assertEqual('d', df.d.name, 'default indexer return Series whose name is the column-name')
        # with columns, range(1, 7) as indexer
        df = pd.DataFrame(np.random.random((6, 4)), index=range(1, 7), columns=tuple('abcd'))
        self.assertTrue(0 not in df.a.index)
        self.assertTrue(6 in df.a.index)

        # create from some rows from an existing df
        lst = [df.iloc[x].values for x in (0, 2)]
        dfx = pd.DataFrame(lst, columns=df.columns)
        self.assertEqual(2, len(dfx))
        self.assertEqual(df.a[1], dfx.a[0])

        # string as indexer
        df = pd.DataFrame(np.random.random((2, 2)), index='kate peter'.split(), columns=list('ab'))
        self.assertTrue('kate' in df.a.index)
        self.assertEqual(df.loc['kate', 'a'], df.a[0])
        self.assertEqual(df.loc['kate', 'a'], df.a['kate'])

    def testIndexer(self):
        ''' access using attribute/loc/iloc/[]
        '''
        # with columns, range() as default index
        df = pd.DataFrame(np.random.random((6, 4)), columns=tuple('abcd'))
        self.assertTrue(0 in df.d.index, 'attribute accesser')
        self.assertTrue(6 not in df.d.index)
        self.assertEqual('d', df.d.name, 'attribute accestor return a Series whose name is the column-name')
        self.assertTrue(isinstance(df.iloc[0], pd.core.series.Series), 'iloc, single element, returned as an Series instance')
        self.assertTrue(isinstance(df.iloc[:2], pd.core.frame.DataFrame), 'iloc with slice, returns a DataFrame instance')

        # with columns, range(1, 7) as indexer
        df = pd.DataFrame(np.random.random((6, 4)), index=range(1, 7), columns=tuple('abcd'))
        self.assertTrue(0 not in df.a.index)
        self.assertTrue(6 in df.a.index)
        with self.assertRaises(KeyError):
            print(df.a[0])
        self.assertEqual(df.a[1], df.iloc[0].a, 'iloc is always from 0')

        # string as indexer
        df = pd.DataFrame(np.random.random((2, 2)), index='kate peter'.split(), columns=list('ab'))
        self.assertTrue('kate' in df.a.index)
        self.assertEqual(df.loc['kate', 'a'], df.a[0])
        self.assertEqual(df.loc['kate', 'a'], df.a['kate'])

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

    def testQuery(self):
        ''' common use query
        '''
        sts = self._dates
        lst = (1, 2)
        df = sts.loc[sts.id <= 2]
        df = sts.loc[sts.id.isin(lst)]
        df = sts.query('id not in @lst')
        df = sts.loc[~sts.id.isin((1, 2))] # not in

    def testBoolean(self):
        ''' test the union/intersection funciton
        '''
        df = pd.read_excel(r'd:\temp\syn.xlsx')
        lst = []
        for loc in df.location.unique():
            lst.append(df.loc[df.location == loc].pcode.unique())
        df, dfs = lst[0], set()
        for df1 in lst[1:]:
            df2 = pd.np.intersect1d(df, df1)
            if len(df2) == 0:
                dfs.update([x for x in df])
                df = df1
            else:
                df = df2
        dfs.update([x for x in df])
        print(dfs)


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
