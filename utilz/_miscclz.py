'''
#! coding=utf-8
@Author:        zmFeng
@Created at:    2019-07-19
@Last Modified: 2019-07-19 2:25:18 pm
@Modified by:   zmFeng
some class splitted from misc.py because file too large
'''

from base64 import b64decode, b64encode
from random import randint
from numbers import Number
from json import load as load_json

class Salt(object):
    '''
    a simple hash class for storing not human-readable senstive data. Don't call me crypto because crypto is not revisable but I can
    '''
    def __init__(self, key_mp=None):
        self._key_mp = key_mp or {"A": 2, "C": 10, "D": 5, "E": 18, "F": 0, "G": 18, "H": 6, "I": 9, "J": 4, "K": 1, "L": 3, "N": 15, "O": 9, "P": 8, "Q": 10, "R": 12, "S": 8, "T": 17, "U": 5, "V": 2, "X": 11, "Y": 13, "a": 0, "b": 18, "c": 17, "f": 12, "l": 19, "p": 2, "q": 12, "s": 5, "t": 8, "u": 15, "v": 6, "*": 7, "w": 10, "x": 12, "y": 12, "z": 8, "=": 19, "|": 7, "`": 3}
        self._keys = [x for x in self._key_mp.keys()]
        self._key_ln = len(self._keys)

    def encode(self, src):
        '''
        encode the source using b64 while appending sth. to the suffix and suffix
        '''
        rc = b64encode(src.encode()).decode()
        ptr = randint(0, self._key_ln - 1)
        hdl = self._key_mp[self._keys[ptr]]
        salt, idx = "".join([self._keys[randint(0, self._key_ln - 1)] for x in range(hdl)]), hdl % 3
        if idx == 0:
            rc = salt + rc
        elif idx == 1:
            rc = rc[:len(rc)//2] + salt + rc[len(rc)//2:]
        else:
            rc = rc + salt
        return rc + self._keys[ptr]

    def decode(self, cookie):
        '''
        revise an encoded item
        '''
        hdl = cookie[-1]
        if hdl not in self._key_mp:
            raise AttributeError("cookie(%s) not encoded by me")
        hdl = self._key_mp[hdl]
        idx = hdl % 3
        if idx == 0:
            rc = cookie[hdl:-1]
        elif idx == 1:
            rc = len(cookie) - 1 - hdl
            rc = cookie[:rc // 2] + cookie[rc // 2 + hdl:-1]
        else:
            rc = cookie[:len(cookie) - 1 - hdl]
        return b64decode(rc).decode()

class Config(object):
    '''
    A dict like config storage, different call can get/set changes here. Also have the ability for change listener to monitor setting changes.
    It's advised for the consumer for this class to have it's name space.
    Also, the key won't be normalized, the consumer take control if it
    By default, this module contains one Config instance for convenience. You can  store settings directly to this instance.

    example of boot strap:
        from utilz import config
        ...
        if not config.get("_MY_SIGNATURE_"):
            config.load(json_file)
        ...
        config.get("a")

    example of put your own Config:
        from utilz import config
        if not config.get("_MY_SIGNATURE_"):
            config.set("_MY_SIGNATURE_", Config(json_file))
        ...
        config.get("_MY_SIGNATURE_").get("a")
    '''

    def __init__(self, json_file=None):
        self._dict, self._listeners = {}, {}
        if json_file:
            self.load(json_file)

    def get(self, key, df=None):
        '''
        return the given setting of given key
        '''
        return self._dict.get(key, df) if self._dict else df or None

    def set(self, key, new_value):
        '''
        set value to specified key
        '''
        old_val = self._dict.setdefault(key, new_value)
        lstrs = self._listeners.get(key)
        if not lstrs:
            return
        for lstr in (x for x in lstrs if x):
            try:
                lstr(key, old_val, new_value)
            except:
                pass

    def addListener(self, key, chg_listener):
        '''
        monitor the setting changes
        @param key: the key or keys that the chg_listener need to monitor
        @param chg_listener:
            A method that should have this form: method(key, old_value, new_value) and return value
        '''
        lst = self._listeners.setdefault(key, [])
        if chg_listener not in lst:
            lst.append(chg_listener)

    def removeListener(self, key, listener):
        '''
        remove the listener added to me
        '''
        lst = self._listeners.get(key)
        if not lst:
            return None
        if listener not in lst:
            return None
        lst.remove(listener)
        return listener

    def load(self, json_file, refresh=False):
        '''
        load setting from the given fn(json file)
        @param json_file:
            the file to load data from, or a dict already contains data
        @param refresh:
            clear existing settings(if there is)
        '''
        if refresh or self._dict is None:
            self._dict = {}
        try:
            if isinstance(json_file, dict):
                mp = json_file
            else:
                with open(json_file, encoding="utf-8") as fp:
                    mp = load_json(fp)
                    if not mp:
                        return
            di = {x: y for x, y in mp.items() if x not in self._listeners}
            self._dict.update(di)
            di = {x: y for x, y in mp.items() if x in self._listeners}
            if not di:
                return
            for key, val in di.items():
                self.set(key, val)
        except:
            pass

class Number2Word(object):
    '''
    number to English word
    Usage:
    n2w = Number2Word()
    s = n2w.convert(123.45)
    s = n2w.convert(123.45, join_hund=True, join_ten=True)
    references:
    https://en.wikipedia.org/wiki/Long_and_short_scales#Long_scale
    https://lingojam.com/NumbersToWords
    accept below constructor arguments, all default is False:
    @param join_hund: create an "AND" between hundred and ten
    @param join_ten: create an "-" between ten and digit
    @param show_no_cents: show "And No Cents" when there no cent
    @param case: one of U[pper]/L[ower]
    '''
    def __init__(self, **kwds):
        self._join_ten = self._join_hund = self._show_no_cents = None
        self._case = "U"
        if not kwds:
            kwds = {}
        for x in "join_hund".split():
            if x not in kwds:
                kwds[x] = True
            self._config(kwds)
        ss = ["One Two Three Four Five Six Seven Eight Nine Ten Eleven Twelve Thirteen Fourteen Fifteen Sixteen Seventeen Eighteen Nineteen Twenty Thirty Forty Fifty Sixty Seventy Eighty Ninety", ", Thousand , Million , Billion , Trillion ", "And", "Hundred", "Dollar", "Cent", "No", "s"]
        if self._case in ('U', 'L'):
            for idx, s0 in enumerate(ss):
                s0 = s0.upper() if self._case == 'U' else s0.lower()
                ss[idx] = s0
        self._digits, self._pwrs = ss[0].split(), ss[1].split(',')
        self._ttl = {x[0]: x[1] for x in zip(('and', 'hundred', 'dollar', 'cent', 'no', 'cpl',), ss[2:])}

    def _config(self, kwds):
        self._join_hund, self._join_ten, self._show_no_cents = [kwds.get(x, False) for x in "join_hund join_ten show_no_cents".split()]

    def convert(self, theVal, **kwds):
        '''
        translate the number to Words
        you can provide convert option to override the ones in the constructor, except for case(U/L)
        '''
        # hold current settings
        if kwds:
            cents = {x[0]: x[1] for x in zip("join_hund join_ten show_no_cents".split(), (self._join_ten, self._join_hund, self._show_no_cents, ))}
            self._config(kwds)
            kwds = cents
        theVal, cents = str(theVal) if isinstance(theVal, Number) else theVal.strip(), None
        idx = theVal.find(".")
        if idx >= 0:
            cents, theVal = self._ten((theVal[idx+1:] + "00")[:2]), theVal[:idx]
        dollars, cnt, segs = "", -1, []
        while theVal:
            segs.append(theVal[-3:])
            theVal = theVal[:-3] if len(theVal) > 3 else None
        for cnt, s0 in enumerate(reversed(segs)):
            s0 = self._hund(s0, dollars)
            if s0:
                dollars += s0 + self._pwrs[len(segs) - cnt - 1]
        dc = lambda cnt, unit: ("%s %s%s" % (self._ttl['no'], unit, self._ttl['cpl']) if self._show_no_cents else "") if not cnt else "%s %s" % (self._dig("1"),  unit) if cnt == self._dig("1") else "%s %s%s" % (cnt, unit, self._ttl['cpl'])
        dollars, cents = dc(dollars, self._ttl["dollar"]), dc(cents, self._ttl['cent'])
        if dollars and (cents or self._show_no_cents):
            cents = (" %s " % self._ttl['and']) + cents
        if kwds:
            self._config(kwds)
        return dollars + cents

    def _hund(self, txt, pfx=None):
        if not int(txt):
            return ""
        txt = ("000" + txt)[-3:]
        h = self._dig(txt[0]) + " %s" % self._ttl['hundred'] if txt[0] != "0" else ""
        t = self._ten(txt[1:])
        if h:
            t = (" %s " % self._ttl['and'] if self._join_hund else " ") + t
        elif pfx:
            t = ("%s " % self._ttl['and'] if self._join_hund else "") + t
        return h + t

    def _ten(self, txt):
        s0 = txt[0]
        if s0 == "0":
            s0 = self._dig(txt[-1])
        elif s0 == "1":
            s0 = self._digits[int(txt) - 1]
        else: # If value between 20-99...
            s1, idx = self._dig(txt[-1]), int(txt[0]) + 17
            s0 = self._digits[idx]
            if s1:
                s0 += ("-" if self._join_ten else " ")  + s1
        return s0

    def _dig(self, txt):
        txt = int(txt)
        return self._digits[txt - 1] if txt else ""

class Literalize(object):
    ''' just like the dec/hex expression, but the chars for each digit can be customized
    '''

    def __init__(self, chars, digits=4, expand=True):
        '''
        Args
            chars:  the valid characters
            digits: the maximum count of digits
            expand: append zeros to the front when digits not enough
        '''
        self._chars, self._digits = chars, max(digits, 1)
        self._expand = expand
        self._idx = self._index()
        self._radix = len(chars)
        self._powers = None

    @property
    def charset(self):
        ''' the charsets used by me
        '''
        return self._chars

    @property
    def radix(self):
        ''' in fact, length of the charset
        '''
        return self._radix

    def diff(self, cha, chb):
        ''' return index(cha) - index(chb)
        '''
        return self.valueOf(cha) - self.valueOf(chb)

    def valueOf(self, char):
        ''' the integer value of given character

        Args:

            char(string):   the character (inside the charset of me) to get value for

        throws:

            KeyError when the char is not inside my charset

        '''
        return self._idx[0][char]

    def charOf(self, val):
        ''' the character of given integer value

        Args:

            val(integer):   the integer value to get character for

        throws:

            KeyError when the val is out of bounded

        '''
        return self._idx[1][val]

    def isvalid(self, chars):
        ''' check if all the characters are inside my charset

        Args:

            chars(string): the string for detection

        return:
            True when all characters are inside my charset

        '''

        n2i = self._idx[0]
        return not [x for x in chars if x not in n2i]

    def _validate(self, chars):
        if not self.isvalid(chars):
            raise TypeError('invalid character found in %s, valid charset is %s' % (chars, self._chars))

    @property
    def digits(self):
        ''' the max number of digits that I can generate
        '''
        return self._digits

    
    @property
    def first(self):
        ''' The first char in the charset
        '''
        return self._idx[1][0]


    @property
    def last(self):
        ''' the last char in the charset
        '''
        return self._idx[1][self._radix-1]


    def _fill(self, x):
        return self.first * (self._digits - len(x)) + x if self._expand else x

    def next(self, current=None, frm=0, steps=1):
        ''' given current value and chars, return the next one
        for example:
            ni = NextI('ABCDEF', expand=False)
            ni.next('') == 'A'
            ni.next('F') == '10'
            ni.next('AFDA') == 'AFDB'
            ni.next('01DA') == '01DB'
            ni.next('0000') => TypeError
            ni.next('FFFF') => OverflowError
        '''
        if not current:
            return self._fill(self.first)
        self._validate(current)
        ln = len(current)
        length = max(ln, self._digits)
        idx = frm - 1
        nc = self._next(current[idx], steps=steps)
        lsts = [x for x in current]
        if nc:
            lsts[idx] = nc
            return self._fill(''.join(lsts))
        while -idx < length:
            if -idx == ln:
                lsts.insert(0, self.first)
                ln += 1
            lsts[idx] = self.first
            idx -= 1
            nc = self._next(lsts[idx], steps=steps)
            if nc:
                lsts[idx] = nc
                return self._fill(''.join(lsts))
        raise OverflowError('due to charset(%s), \'%s + steps(%d) \' reaches the end' % (self._chars, current, steps))

    def toInteger(self, literal):
        ''' convert given literal to decimal
        Args:
            literal: the string to convert, for example, FF
        Returns:
            the integer presence of given literal
        '''

        self._validate(literal)
        ln = len(literal) + 1
        if not self._powers:
            self._powers = []
        rc = len(self._powers)
        if rc < ln:
            for x in range(rc, ln):
                self._powers.append(self._radix ** x)
        rc = 0
        for idx in range(1, ln):
            rc += self._idx[0][literal[-idx]] * self._powers[(idx - 1)]
        return rc

    def fromInteger(self, val):
        ''' convert given val to my string
        '''
        lst = []
        t = val
        while t > 0:
            lst.append(self._idx[1][t % self._radix])
            t = int(t / self._radix)
        return self._fill(''.join(reversed(lst)))

    def _index(self):
        n2i = {char: idx for idx, char in enumerate(self._chars)}
        i2n = {idx: char for idx, char in enumerate(self._chars)}
        return (n2i, i2n, )

    def _next(self, char, steps=1):
        if not char:
            return self.first
        idx = self._idx
        i = idx[0].get(char)
        if i is None:
            raise TypeError('%s not in sequence %s' % (char, self._chars))
        i += steps
        return idx[1].get(i)

class Segments(object):
    '''
    given an A * A area and the length(size) of a segment, split the area
    into segments. Segment contains length of sector(s).

    First sector of a segment is the segment's header.

    A sector is a tuple of 2 integer, tuple[0] for row index and tuple[1] for column index

    Resources defined:

        size:   length of the area.

        segment_size=20: the length of a segment

        row_first=True: segments spawn by row then column

    row_first using below pattern:
        figure(assume size=5, segment_size=3):
        00011

        00011

        00011

        00022

        00022

    here 0 and stands for level 0&1, 2 for span
    '''
    def __init__(self, size, segment_size, row_first=True):
        self._size = size
        self._row_first = row_first
        self._calc = _SpanCalc(size, segment_size)

    def next(self, addr, segment=True):
        ''' get the next segment or sector, based on what ${segment} specified
        Args:

            addr:       the current address

            segment=True:    True for segment, False for sector

        throws:

            OverflowError if current is already at the end
        '''
        self._validate(addr)
        if segment:
            if addr:
                addr = self.range(addr)[0]
            return self._transpose(self._new_segment(self._transpose(addr)))
        return self._transpose(self._new_sector(self._transpose(addr)))

    def _validate(self, addr):
        if addr:
            if any(x < 0 or x >= self._size for x in addr):
                raise OverflowError('given address is out of border')

    def _transpose(self, addr):
        if not addr:
            return addr
        return addr if self._row_first else (addr[1], addr[0])

    def all(self, dump=None):
        ''' return all the segments and sectors, each segments as a list
        '''

        sgs = [self.sectors(sg) for sg in self.segments]
        if dump:
            chcnt = '%0' + str(len(str(self._size))) + 'd'
            _ts = lambda arr: '.'.join(chcnt % x for x in arr)
            sgc = self.capacity
            stc = sgc * self._calc.size
            rt = stc / self._size ** 2 * 100
            print('Size=%d, SegSz=%d, RowFirst=%s, SegCnt=%d, SectCnt=%d, Ratio=%4.2f%%' % (self._size, self._calc.size, 'True' if self._row_first else 'False', sgc, stc, rt), file=dump)
            for scts in sgs:
                print('%s: %s' % (_ts(scts[0]), ', '.join(_ts(x) for x in scts[1:])), file=dump)
        return sgs

    @property
    def segments(self):
        ''' return the segments available
        '''
        lst, seg = [], None
        while True:
            try:
                seg = self.next(seg)
                lst.append(seg)
            except OverflowError:
                break
        return lst

    def sectors(self, addr):
        ''' return all the sectors(from head to tail) of given address
        '''
        self._validate(addr)
        sect = self.range(addr)[0]
        lst = [sect, ]
        while True:
            try:
                sect = self.next(sect, False)
                lst.append(sect)
            except OverflowError:
                break
        return lst

    @property
    def capacity(self):
        ''' the count of segments that I can hold
        '''
        return self._size ** 2 // self._calc.size

    def range(self, addr):
        ''' return the header/tail sector of given address
        '''
        self._validate(addr)
        if self._calc.size == 1:
            return (addr, ) * 2
        addr = self._transpose(addr)
        lvl = self._calc.level(addr)
        ss = self._calc.size
        if lvl == 0:
            rc = (addr[0], addr[1] // ss * ss)
            rc1 = (rc[0], rc[1] + ss - 1)
        elif lvl == 1:
            rc = (addr[0] // ss * ss, addr[1])
            rc1 = (rc[0] + ss - 1, rc[1])
        else:
            rc, rc1 = self._calc.range(addr, True)
        return tuple(self._transpose(x) for x in (rc, rc1))

    def _next_axis(self, current=None, steps=None):
        if current is None:
            return 0
        if not steps:
            steps = 1
        if current + steps >= self._size:
            raise OverflowError()
        return current + steps

    def _new_segment(self, addr):
        if not addr:
            return 0, 0
        g1, calc, rc = self._next_axis, self._calc, None
        mx = calc.client_org - calc.size + 1
        verxi = addr[1]
        if verxi < mx:
            # size level 0
            try:
                rc = g1(addr[0]), addr[1]
            except OverflowError:
                # next span of level 0
                if verxi + calc.size < self._size:
                    rc = 0, g1(addr[1], steps=calc.size)
        if rc:
            return rc
        verxi = addr[0]
        if verxi < mx:
            # level 1
            if verxi + calc.size * 2 < self._size:
                rc = g1(addr[0], steps=calc.size), addr[1]
            else:
                try:
                    rc = 0, g1(addr[1])
                except OverflowError:
                    # header for level 2
                    if calc.client_area < calc.size:
                        raise OverflowError('no chance to enter level 2')
                    rc = (calc.client_org, ) * 2
        return rc or calc.add_span(addr)

    def _new_sector(self, addr):
        if not addr:
            return 0, 0
        nxt, calc = self._next_axis, self._calc
        if calc.size == 1:
            raise OverflowError('segment of size = 1 should not have sector')
        lvl = calc.level(addr)
        def _eo_sgt(idx):
            if idx and (idx + 1) % calc.size == 0:
                raise OverflowError('level 0 ends')
            return idx
        if lvl == 0:
            return addr[0], nxt(_eo_sgt(addr[1]))
        if lvl == 1:
            return nxt(_eo_sgt(addr[0])), addr[1]
        return calc.next_sector(addr)


class _SpanCalc(object):
    ''' help Partition to calculate the span locaions

    Args:

        size:   length of the axis

        segment_size: length of a segment
    '''

    def __init__(self, size, segment_size):
        self._size = segment_size
        self._cache = mp = {}
        x = size // segment_size * segment_size
        mp['_client_org'] = x
        if x < size:
            x = size - x
            mp['_client_height'] = x
            mp['_client_area'] = x * x
        else:
            mp['_client_height'] = mp['_client_area'] = 0

    @staticmethod
    def dim_convert(addr, org, hgt):
        ''' convert between 2d and 1d
        '''
        if isinstance(addr, (tuple, list)):
            if addr[0] < org:
                org = 0
            return (addr[0] - org) * hgt + addr[1] - org
        return addr // hgt + org, addr % hgt + org


    def _get_cache(self, key, calc):
        if key not in self._cache:
            self._cache[key] = calc()
        return self._cache[key]

    @property
    def size(self):
        ''' span size
        '''
        return self._size

    @property
    def client_height(self):
        ''' height of the span area
        '''
        return self._cache['_client_height']

    @property
    def client_area(self):
        ''' the size of client area
        '''
        return self._cache['_client_area']

    @property
    def client_org(self):
        ''' original point of the span area
        '''
        return self._cache['_client_org']

    def level(self, addr):
        ''' return the level of given address
        '''
        org = self.client_org
        if addr[1] < org:
            return 0
        if addr[0] < org:
            return 1
        return 2

    def next_sector(self, addr):
        ''' calc the next sector of given addr, assume addr is in client_area
        Args:

            addr: a formatted address
        '''
        rng = self.range(addr)
        r_c = self._add(addr, 1)
        if self._dim_convert(r_c) <= self._dim_convert(rng[1]):
            return r_c
        raise OverflowError('end of span')

    def range(self, addr, offset=False):
        ''' return the head/tail sector of given addr
        '''
        ln = self._dim_convert(addr)
        sz = self.size
        hdr = ln // sz
        tail = (hdr + 1) * sz - 1
        if tail >= self.client_area:
            raise OverflowError()
        cv = lambda x: self._dim_convert(x, offset)
        return (cv(hdr * sz), cv(tail), )

    def _dim_convert(self, addr, offset=False):
        ''' convert between one-dim and 2 dim
        '''
        if isinstance(addr, (tuple, list)):
            org = self.client_org
        else:
            org = self.client_org if offset else 0
        return _SpanCalc.dim_convert(addr, org, self.client_height)

    def add_span(self, addr):
        ''' add one span to given ver
        '''
        if not self.client_height:
            raise OverflowError('perfect fit, no level 2 needed')
        if self._dim_convert(addr) + self.size * 2 > self.client_area:
            raise OverflowError('level 2 overflow')
        return self._add(addr)

    def _add(self, r_c, steps=None, offset=True):
        if not steps:
            steps = self.size
        return self._dim_convert(self._dim_convert(r_c) + steps, offset)

class NumericRange(object):
    ''' given numeric a and b, split (b - a) into n segments, then query the span that a numer belongs to

    sometimes you have a set of numeric range from a to b, for performance issue, 
    you split it into n groups, then for each group you do some issue.

    Now you have a number x, you need to know which group it belongs to.

    nr = NumericRange(1, 100, 10)
    idx, rg = nr.range(0)
    idx, rg = nr.range(10)
    idx, rg = nr.range(32.5)
    idx, rg = nr.range(100.1)

    '''
    def __init__(self, a, b, spancnt=10, step=None):
        if a > b:
            a, b = b, a
        self._a, self._b = a, b
        if step:
            self._n, self._step = int((b - a + step - 1) // step), step
        else:
            self._n, self._step = spancnt, (b - a) / spancnt
    
    @property
    def step(self):
        ''' return step of each range
        '''
        return self._step

    def range(self, x):
        '''
        Returns:
            the span id and range of given x

            but:
                when x not in [a, b], OverflowError will be raised
        '''
        if x < self._a or x > self._b:
            raise OverflowError()
        cnt = int((x - self._a) // self._step)
        tmp = self._a + cnt * self._step
        return cnt, (tmp, tmp + self._step)
