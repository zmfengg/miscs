# coding=utf-8
'''
Created on Apr 17, 2018



@author: zmFeng
'''

from numbers import Number

from utilz import Karat, KaratSvc, karatsvc, xwu
from os import path

__all__ = ["JOElement", "StyElement", "KaratSvc", "Karat", "karatsvc"]


class JOElement(object):
    """
    representation of Alpha + digit composite key
    the constructor method can be one of:
    
    JOElement("A1234BC")
    JOElement("A",123,"BC")
    JOElement(12345.0)

    JOElement(alpha = "A",digit = 123,suffix = "BC")

    Args:
    length=5: the default mininum length of the alpha + digit result
    """
    __minlen__ = 5

    @staticmethod
    def tostr(jn):
        ''' conver the JO# to string if is a number, this occur in some excel data '''
        return "%d" % int(jn) if isinstance(jn, Number) else jn

    @staticmethod
    def from_file(file, vdl=None):
        '''
        extract JO# from given file
        @param file: the full-path file
        @param vdl: the validator that let pass or not, if nothing is provided, JOElement's own validator will be used
        '''
        if not path.exists(file):
            return None
        if not vdl:
            vdl = lambda x: JOElement(x).isvalid()
        fn = path.splitext(file)[1].lower()
        if fn.find("xl") >= 0:
            app, tk = xwu.appmgr.acq()
            fn = app.books.open(file)
            try:
                lst = []
                for sht in fn.sheets:
                    var = [y for x in xwu.usedrange(sht).value for y in x if y and isinstance(y, str) and vdl(y)]
                    if var:
                        lst.extend(var)
            finally:
                if fn:
                    fn.close()
                if tk:
                    xwu.appmgr.ret(tk)
        else:
            with open(file, 'rt') as fn:
                lst = [y for x in fn for y in x.split() if vdl(y)]
        return lst

    def __init__(self, *args, **kwds):
        cnt = kwds.get('length')
        if cnt:
            self._minlen = cnt
        cnt = len(args)
        if cnt == 1:
            self._parse_(args[0])
        elif cnt >= 2:
            self.alpha = args[0].strip()
            self.digit = args[1]
            self.suffix = args[2].strip() if (cnt > 2) else ""
        else:
            self._reset()

    def _parse_(self, jono):
        if not jono:
            self._reset()
            return
        stg, strs = 0, ["", "", ""]
        jono = JOElement.tostr(jono).strip()
        for idx, ch in enumerate(jono):
            if ch.isalpha():
                if stg == 0:
                    strs[0] = strs[0] + ch
                else:
                    strs[2] = strs[2] + jono[idx:]
                    break
            elif ch.isdigit():
                if not stg:
                    stg += 1
                    #first character is number, let it be alpha
                    if not strs[0]:
                        strs[0] = ch
                        continue
                strs[1] = strs[1] + ch
            else:
                break
        if stg and strs[1].isdigit():
            self.alpha = strs[0].strip().upper()
            self.digit = int(strs[1])
            self.suffix = strs[2].strip().upper()
        else:
            self._reset()

    def _reset(self):
        self.alpha = ""
        self.digit = 0
        self.suffix = ""

    def __repr__(self, *args, **kwargs):
        return "JOElement(%s,%d,%s)" % (self.alpha, self.digit, self.suffix)

    def __str__(self, *args, **kwargs):
        if hasattr(self, 'digit'):
            ml = self._minlen if hasattr(self, '_minlen') else self.__minlen__
            return self.alpha + (
                ("%0" + str(ml - len(self.alpha)) + "d") %
                self.digit)
        return ""

    @property
    def value(self):
        """ a well-formatted alpha+digit """
        return self.__str__()

    @property
    def name(self):
        """ a well-formatted alpha+digit, the same as @property(value) """
        return self.__str__()

    def isvalid(self):
        """ is alpha and digit set """
        return bool(self.alpha) and bool(self.digit)

    def __composite_values__(self):
        return self.alpha, self.digit

    def __eq__(self, other):
        return isinstance(
            other, JOElement
        ) and self.alpha == other.alpha and self.digit == other.digit

    def __lt__(self, other):
        if isinstance(other, JOElement):
            flag = self.alpha < other.alpha
            if not flag:
                if self.alpha == other.digit:
                    flag = self.digit < other.digit
            return flag
        raise ValueError(
            "given object(%r) is not a %s" % (other, type(self).name))

    def __hash__(self):
        return hash((self.alpha, self.digit))

    def __ne__(self, other):
        return not self.__eq__(other)

    def __ge__(self, other):
        return isinstance(
            other, JOElement
        ) and self.alpha == other.digit and self.digit >= other.digit


class StyElement(JOElement):
    """ JOElement with suffix """
    def __composite_values__(self):
        pr = JOElement.__composite_values__(self)
        return pr[0], pr[1], self.suffix

    def __eq__(self, other):
        return JOElement.__eq__(self, other) and self.suffix == other.suffix

    def __hash__(self):
        return hash((super(StyElement, self).__hash__(), self.suffix))

    def __str__(self, *args, **kwargs):
        val = super(StyElement, self).__str__(args, **kwargs)
        if val:
            val += self.suffix
        return val
