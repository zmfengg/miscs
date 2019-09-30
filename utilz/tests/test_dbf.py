'''
#! coding=utf-8
@Author:        zmFeng
@Created at:    2019-09-19
@Last Modified: 2019-09-19 10:40:58 am
@Modified by:   zmFeng
read/write dbf test
'''

from unittest import TestCase, skipIf
try:
    from dbfread import DBF
except:
    DBF = None
from ..simpledbf import Dbf5
from ..miscs import triml

@skipIf(DBF is None, 'dbfread lib was not installed')
class DbfReadSuite(TestCase):
    ''' try the dbfread lib, need hacking, so I turn to hack simpledbf
    '''

    @classmethod
    def setUpClass(cls):
        super().setUpClass()
        cls.__dbf = None

    @property
    def _dbf(self):
        if not self.__dbf:
            # opening a dbf is not slow using lazy-loading
            self.__dbf = DBF(r'p:\aa\bc\stocks.DBF')
        return self.__dbf

    def testGetFields(self):
        dbf = self._dbf
        print(dbf.fields)
        # seems it can only do progressive read, don't call th len() of it
        # very slow, also dbf[-1] is not supported
        print(len(dbf.records))

class SimpleDbfSuite(TestCase):
    ''' a modified one that supports reading the last x records
    '''
    def testReadLastX(self):
        fn = r'p:\aa\bc\stocks.DBF'
        # fn = r'p:\aa\bc\stocks_tiny.DBF'
        for n in (-1, 5):
            dbf = Dbf5(fn, last_n=n, cn_fmtr=triml)
            df = dbf.to_dataframe()
            self.assertEqual(abs(n), len(df), 'the last %d records' % n)
        self.assertEqual(51, len(Dbf5(r'p:\aa\bc\stocks_tiny.DBF', last_n=500).to_dataframe()))
