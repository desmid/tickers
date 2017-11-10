###########################################################################
import re

class cellrange(object):
    def __init__(self, name=''):
        self.set_by_name(name)

    def set_by_posn(self, start_col=0, start_row=0, end_col=0, end_row=0):
        if start_col < 0 or start_row < 0 or end_col < 0 or end_row < 0:
            raise TypeError("cell positions cannot be negative")
        self.start_col = start_col
        self.start_row = start_row
        self.end_col = end_col
        self.end_row = end_row

    def set_by_name(self, name):
        ((self.start_col, self.start_row), (self.end_col, self.end_row)) \
            = range2posn(name)

    def name2posn(self, name): return name2posn(name)
    def posn2name(self, pos): return posn2name(pos)

    def range2posn(self, name): return range2posn(name)

    def __eq__(self, other):
        if isinstance(other, self.__class__):
            return self.__dict__ == other.__dict__
        return False

    def __ne__(self, other):
        return not self.__eq__(other)

###########################################################################
# static
###########################################################################
ORD_A = ord('A')

def name2posn(n=''):
    """Convert spreadsheet cell names ('A1', AZ2', etc.) and return a
    pair of 0-based positions as a tuple: (column, row).

    - Ignores $ signs.
    - Harmlessly allows row 0 (treats as row 1).

    Example usage and return values:

    name2posn('')     =>  (0,0)
    name2posn('0')    =>  (0,0)
    name2posn('A')    =>  (0,0)
    name2posn('A0')   =>  (0,0)
    name2posn('A1')   =>  (0,0)

    name2posn('B')    =>  (1,0)
    name2posn('B0')   =>  (1,0)
    name2posn('B1')   =>  (1,0)
    name2posn('B2')   =>  (1,1)
    name2posn('Z')    =>  (25,0)
    name2posn('AA')   =>  (26,0)
    name2posn('AZ')   =>  (51,0)
    name2posn('BA')   =>  (52,0)
    name2posn('ZZ')   =>  (701,0)
    name2posn('AAA')  =>  (702,0)

    """
    if isinstance(n, int):
        if n < 0: n = 0
        return (n, n)
    if not isinstance(n, str):
        return (0, 0)
    n = re.sub('[^A-Z0-9:]', '', str(n).upper())
    #0-based arithmetic
    if n is '': n = '1'
    c, r = 0, 0
    while len(n) > 0:
        v = ord(n[0])
        if v < ORD_A:
            r = int(n)
            break
        c = 26*c + v-ORD_A+1
        n = n[1:]
    if c > 0: c -= 1
    if r > 0: r -= 1
    return (c, r)

def posn2name(p):
    """Convert spreadsheet cell positions as a pair (column, row) and
    return a cell name as a string: 'A1'

    Example usage and return values:
 
    posn2name((0,0))    =>  'A1'
    posn2name((0,1))    =>  'A2'
    posn2name((1,0))    =>  'B1'
    posn2name((1,1))    =>  'B2'

    posn2name((25,0))   =>  'Z1'
    posn2name((26,0))   =>  'AA1'
    posn2name((27,0))   =>  'AB1'

    posn2name((51,0))   =>  'AZ1'
    posn2name((52,0))   =>  'BA1'
    posn2name((53,0))   =>  'BB1'

    posn2name((701,0))  =>  'ZZ1'
    posn2name((702,0))  =>  'AAA1'
    posn2name((703,0))  =>  'AAB1'
    """
    try:
        col,row = p
    except:
        raise TypeError("cell positions must an int pair")
    if not (isinstance(col, int) and isinstance(row, int)):
        raise TypeError("cell positions must an int pair")
    if col < 0 or row < 0:
        raise TypeError("cell positions cannot be negative")
    col += 1
    name = ""
    while col > 0:
        col, remainder = divmod(col-1, 26)
        name = chr(ORD_A + remainder) + name
    return name + str(row+1)

def range2posn(n=''):
    """Convert spreadsheet cell name ranges ('A1:A10', B10:G100', etc.)
    and return a pair of pairs of 0-based positions as a tuple:
    ((start_column, start_row), (end_column, end_row)).

    - Ignores $ signs.
    - Harmlessly allows row 0 (treats as row 1).

    Example usage and return values:

    range2posn()          =>  ((0,0), (0,0))
    range2posn('')        =>  ((0,0), (0,0))
    range2posn('A:Z')     =>  ((0,0), (25,0))
    range2posn('A1:A99')  =>  ((0,0), (0,98))
    range2posn('A1:C3')   =>  ((0,0), (2,2))

    """
    if isinstance(n, int):
        n = name2posn(n)
        return (n, n)
    if not isinstance(n, str):
        n = name2posn(n)
        return (n, n)
    try:
        c,r = n.split(':', 1)
    except:
        c = n.split(':')[0]
        r = c
    if c == '': c = '0'
    if r == '': r = '0'
    return (name2posn(c), name2posn(r))

###########################################################################
