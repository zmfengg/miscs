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

from numpy.random import random
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
    ''' index is everywhere, axes of Series/DataFrame, just too complex ...
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

    def testCreateFromDraft(self):
        ''' create from array with/without index/name
        '''
        # Dummy, no data, just columns
        df = pd.DataFrame(None, columns=tuple('abcd'))
        self.assertEqual(0, len(df), 'blank frame with column names')

        # without column, range() as default index
        df = pd.DataFrame(random((6, 4)))
        self.assertEqual(6, len(df), 'row count')
        self.assertEqual(4, len(df.iloc[0]), 'column count')

        # with columns, range() as default index
        df = pd.DataFrame(random((6, 4)), columns=tuple('abcd'))
        self.assertTrue(0 in df.d.index)
        self.assertTrue(6 not in df.d.index)
        self.assertEqual('d', df.d.name, 'default indexer return Series whose name is the column-name')

        # with columns, range(1, 7) as indexer
        df = pd.DataFrame(random((6, 4)), index=range(1, 7), columns=tuple('abcd'))
        self.assertTrue(0 not in df.a.index)
        self.assertTrue(6 in df.a.index)

        # string as indexer
        df = pd.DataFrame(random((2, 2)), index='kate peter'.split(), columns=list('ab'))
        self.assertTrue('kate' in df.a.index)
        self.assertEqual(df.loc['kate', 'a'], df.a[0])
        self.assertEqual(df.loc['kate', 'a'], df.a['kate'])
    
    def testCreateFromExisting(self):
        ''' create from an existing df
        '''
        df = pd.DataFrame(random((6, 4)), columns=tuple('abcd'))
        # create from some rows from an existing df
        dfx = pd.DataFrame([df.iloc[x] for x in (0, 2)]) # standard
        lst = [df.iloc[x].values for x in (0, 2)] #stupid
        dfx = pd.DataFrame(lst, columns=df.columns)
        self.assertEqual(2, len(dfx))
        self.assertEqual(df.a[0], dfx.a[0])

        # create from some columns from existing df as columns
        dfx = pd.concat([df.a, df.b], keys=['na', 'nb'], axis=1, ignore_index=False)
        self.assertTrue('na' in dfx.columns)
        self.assertEqual(df.iloc[0].a, dfx.iloc[0].na)
        # column name ignored
        dfx = pd.concat([df.a, df.b], keys=['na', 'nb'], axis=1, ignore_index=True)
        self.assertFalse('na' in dfx.columns)
        # create from some columns from existing df as rows
        dfx = pd.concat([df.a, df.b], keys=['na', 'nb'], axis=0, ignore_index=True)
        self.assertEqual(len(dfx), len(df) * 2, 'appended as row, not column')
        self.assertEqual(df.iloc[0].b, dfx[len(df)])

        # concate 2 df with same column names
        dfx = pd.concat([df, df], ignore_index=True)
        self.assertTrue('a' in dfx.columns, 'index ignored but colname kept')
        self.assertEqual(2 * len(df), len(dfx), 'merged')

        # concate 2 df with different column names
        dfx = pd.concat([pd.DataFrame(random((6, 4)), columns=tuple('abcd')), \
            pd.DataFrame(random((6, 4)))])
        self.assertEqual(8, len(dfx.columns))
        # concate 2 df with some same column names
        dfx = pd.concat([pd.DataFrame(random((6, 4)), columns=tuple('abcd')), \
            pd.DataFrame(random((6, 4)), columns=tuple('abce'))])
        self.assertEqual(5, len(dfx.columns))

    def testMakeDict(self):
        ''' create dict from 2 columns
        '''
        lst = [('A', 1), ('B', 2)]
        df = pd.DataFrame(lst, columns=('name', 'id'))
        dct = dict(zip(df.name, df.id))
        self.assertEqual(1, dct['A'])

    def testChangeIdxAndCol(self):
        ''' change a df's index/colname
        '''
        df = pd.DataFrame(random((6, 4)))
        self.assertTrue(0 in df.columns)
        df.columns = pd.Index(tuple('abcd'))
        self.assertTrue('a' in df.columns)

    def testIndexer(self):
        ''' access dataframe using attribute/loc/iloc/[]
        '''
        # with columns, range() as default index
        df = pd.DataFrame(random((6, 4)), columns=tuple('abcd'))
        self.assertTrue(0 in df.d.index, 'attribute accesser')
        self.assertTrue(6 not in df.d.index)
        self.assertEqual('d', df.d.name, 'attribute accestor return a Series whose name is the column-name')
        self.assertTrue(isinstance(df.iloc[0], pd.core.series.Series), 'iloc, single element, returned as an Series instance')
        self.assertTrue(isinstance(df.iloc[:2], pd.core.frame.DataFrame), 'iloc with slice, returns a DataFrame instance')

        # with columns, range(1, 7) as indexer
        df = pd.DataFrame(random((6, 4)), index=range(1, 7), columns=tuple('abcd'))
        self.assertTrue(0 not in df.a.index)
        self.assertTrue(6 in df.a.index)
        with self.assertRaises(KeyError):
            print(df.a[0])
        self.assertEqual(df.a[1], df.iloc[0].a, 'iloc is always from 0')

        # string as indexer
        df = pd.DataFrame(random((2, 2)), index='kate peter'.split(), columns=list('ab'))
        self.assertTrue('kate' in df.a.index)
        self.assertEqual(df.loc['kate', 'a'], df.a[0])
        self.assertEqual(df.loc['kate', 'a'], df.a['kate'])
        self.assertEqual(df.iloc[0].a, df.a['kate'])
        idx = df.index
        self.assertEqual(idx[0], 'kate')
        self.assertEqual(idx.get_loc('kate'), 0, 'yes, index, name and idx using default indexer and get_loc')

        # index from no-column/duplicate index case
        # in fact, no columns is numeric column, no column is impossible
        df = pd.DataFrame(random((6, 4)))
        df.loc[0, 0] = 0 #
        df = pd.concat([df, df], sort=False)
        self.assertEqual([0.0, 0.0], df.loc[0, 0].values.tolist(), 'after concate without ignoring the index, index duplicated')

        # multiple find in one column
        # df.btchno is Series instead of string, so find failed. use apply instead
        # this technique can be used to find complex things
        # remember to delete the created 'flag' field
        df = pd.read_table(path.join(getpath(), 'res', 'pd_tbl.txt'), encoding='gbk')
        with self.assertRaises(Exception):
            df.loc[[1 for x in ('19', '18') if df.btchno.find(x) >= 0]]
        df['flag'] = df.pkno.apply(lambda x: sum([1 for y in ('PLR', 'RQP') if x.find(y) >= 0]))
        self.assertEqual(230, len(df.loc[df.flag > 0]))
        self.assertTrue('flag' in df.columns)
        del df['flag']
        self.assertTrue('flag' not in df.columns)

        # df.loc[sth].btchno[0] to get the first btchno might failed, use
        # df.loc[sth].iloc[0]btchno is the correct way
        x = df.loc[df.pkno == 'SCO00170']
        self.assertEqual(2, len(x), 'the count of this package is 2')
        with self.assertRaises(KeyError):
            self.assertEqual('1705009', x.btchno[0])
        self.assertEqual('1705009', x.iloc[0].btchno)
        # if really want to access by df.btchno[0], need to reset the index
        x = x.reset_index()
        self.assertEqual('1705009', x.btchno[0])

    def testQuery(self):
        ''' common use query
        '''
        sts = self._dates
        lst = (1, 2)
        df = sts.loc[sts.id <= 2]
        df = sts.loc[sts.id.isin(lst)]
        df = sts.query('id not in @lst')
        df = sts.loc[~sts.id.isin((1, 2))] # not in
        # https://pandas.pydata.org/pandas-docs/stable/reference/api/pandas.Series.all.html
        sr = pd.Series((None, pd.np.nan, 1, 2))
        self.assertTrue(sr.all(skipna=False)) # None is not na
        self.assertTrue(sr.all(skipna=True)) # quite strange
        self.assertEqual(2, sum(sr.isna()), 'the correct method is to use isna() + sum')

    def testReadTable(self):
        ''' read table/csv differs only for the tab delimiter
        '''
        frm = pd.read_table(path.join(getpath(), 'res', 'pd_tbl.txt'), encoding='gbk')
        top = frm[frm.id > 5000]

    def testCSVRW(self):
        ''' csv's read/write ability, use StringIO as mock file
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
        self.assertTrue(sr.all(), 'all() function is wrong')

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
        # df is a view, don't try to set data in a view
        df.iloc[0].id = 'xyx'
        self.assertNotEqual('xyx', df.iloc[0].id, 'value in the view not changed')
        # but how to query then change?, maybe use the returned df's index is a good point
        # https://stackoverflow.com/questions/17729853/replace-value-for-a-selected-cell-in-pandas-dataframe-without-using-index
        # 3 ways to change value of a dataFrame
        org.id[org.id == 3] = 30
        self.assertFalse(org.loc[org.id == 30].empty)
        # only these two without warning
        org.loc[org.id == 30, 'id'] = 3
        self.assertFalse(org.loc[org.id == 3].empty)
        org.id.replace(3, 30, inplace=True)
        self.assertFalse(org.loc[org.id == 30].empty)
        self.assertTrue(org.loc[org.id == 3].empty)
        df.loc[0, 'id'] = 'xyx'
        self.assertEqual('xyx', df.id[0], 'using loc can change it without warning')

    def testComputeField(self):
        ''' sometimes you need to new compute column based on existing columns
        https://stackoverflow.com/questions/21702342/creating-a-new-column-based-on-if-elif-else-condition
        '''
        frm = pd.read_table(path.join(getpath(), 'res', 'pd_tbl.txt'), encoding='gbk')
        def _f(row):
            return row.qty >= 50 and row.id < 440
        # Dataframe apply(), the argument is of Series
        frm['flag'] = frm.apply(_f, axis=1)
        # Series apply(), the argument is of Scalar
        frm['flag1'] = frm.qty.apply(lambda x: x >= 50)
        self.assertTrue(frm.loc[frm.id == 275].flag.bool())
        self.assertFalse(frm.loc[frm.id == 425].flag.bool())
        self.assertTrue(frm.loc[frm.id == 275].flag.bool())
        self.assertFalse(frm.loc[frm.id == 470].flag.bool())
        # instead of creating new field, you can overwrite old field
        frm.btchno = frm.btchno.apply(lambda s: 'C%s' % s)
        self.assertEqual(frm.loc[frm.id == 275].btchno[0], 'C1710036')
        self.assertEqual(frm.loc[frm.id == 275].iloc[0].btchno, 'C1710036')
    
    def testGroupBy(self):
        ''' group the result by some columns and the merge them back
        '''
        frm = pd.read_table(path.join(getpath(), 'res', 'pd_tbl.txt'), encoding='gbk')
        # group by and merge, now try to get batches and sumQty of some pkno
        g = frm.groupby(['pkno', 'type'])
        df0 = g['btchno'].apply(','.join).reset_index() # join return series with multi-index
        df0['btchno'] = df0.btchno.apply(lambda x: x.split(','))
        df0['qty'] = g['qty'].sum().values #g['qty'].sum() return Series
        x = df0.loc[(df0.pkno == 'AMM00362') & (df0.type == '补烂')]
        self.assertTrue(len(x) == 1)
        self.assertAlmostEqual(x.iloc[0].qty, 1)
        x = df0.loc[(df0.pkno == 'AMM00362') & (df0.type == '配出')]
        self.assertAlmostEqual(x.iloc[0].qty, 120)

    def testAppendRows(self):
        df = pd.DataFrame(columns=list('abc'))
        lst = [[1, 2, 3], ] * 3
        lst.insert(0, list('abc')) # the first element should be the title
        # https://thispointer.com/python-pandas-how-to-add-rows-in-a-dataframe-using-dataframe-append-loc-iloc/
        df1 = df.append(lst) #directly append list, but failed. 3 more columns appended
        self.assertEqual(4, len(df1))
        # 3 columns(0, 1, 2) were appended instead of 3 rows
        self.assertTrue(0 in df1.columns)
        self.assertFalse(0 in df.columns)
        self.assertEqual(1, df1[0][1])
        self.assertEqual(3, df1[2][3])

        # the actual append action should be sth. like below
        lst = [[1, 2, 3], ] * 3
        df1 = pd.DataFrame(lst, columns=df.columns)
        df1 = pd.concat((df, df1))
        self.assertEqual(3, len(df1))
        self.assertEqual(1, df1['a'][0])

    def testSorting(self):
        df = pd.DataFrame([[1, 2, 3], [1, 0, 4], [0, 1, 2], [-1, 1, 2]], columns=list('abc'))
        df1 = df.sort_values(['a'])
        self.assertEqual(-1, df1.iloc[0].a)
        self.assertEqual(1, df1.a[0], 'the index not changed, so a[0] still points to the original first')
        df1 = df.sort_values(['a', 'b'])
        self.assertEqual(2, df1.iloc[-1].b)
