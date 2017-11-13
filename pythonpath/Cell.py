###########################################################################
import logging
Logger = logging.getLogger('LoadPrices')
Logger.debug("Load: Cell")

###########################################################################
import re

###########################################################################
def name2posn(n=''):
    """Return spreadsheet cell names ('A1', AZ2', etc.) as a tuple of
    0-based positions.

    - Ignores $ signs.
    - Harmlessly allows row 0 (treats as row 1).

    Example usage and return values:

    name2posn('')    =>  (0,0)
    name2posn('0')   =>  (0,0)
    name2posn('A')   =>  (0,0)
    name2posn('A0')  =>  (0,0)
    name2posn('A1')  =>  (0,0)

    name2posn('B')   =>  (1,0)
    name2posn('B0')  =>  (1,0)
    name2posn('B1')  =>  (1,0)
    name2posn('B2')  =>  (1,1)
    name2posn('Z')   =>  (25,0)
    name2posn('AA')  =>  (26,0)
    name2posn('AZ')  =>  (51,0)
    name2posn('BA')  =>  (52,0)
    name2posn('ZZ')  =>  (701,0)
    name2posn('AAA') =>  (702,0)
    """
    if isinstance(n, int):
        if n < 0: n = 0
        return (n,n)
    if not isinstance(n, str): return (0,0)
    n = re.sub('[^A-Z0-9:]', '', str(n).upper())
    #0-based arithmetic
    if n is '': n = '1'
    c,r,ordA = 0,0,ord('A')
    while len(n) > 0:
        v = ord(n[0])
        if v < ordA:
            r = int(n)
            break
        #print(n, v, c, r)
        c = 26*c + v-ordA+1
        #print(n, v, c, r)
        n = n[1:]
    if c > 0: c -= 1
    if r > 0: r -= 1
    return (c,r)

def range2posn(n=''):
    """Return spreadsheet cell name ranges ('A1:A10', B10:G100', etc.) as a
    tuple of pairs of 0-based positions.

    - Ignores $ signs.
    - Harmlessly allows row 0 (treats as row 1).

    Example usage and return values:

    range2posn()          =>  ((0,0), (0,0))
    range2posn('')        =>  ((0,0), (0,0))
    range2posn('A:Z')     =>  ((0,0), (25,0))
    range2posn('A1:A99')  =>  ((0,0), (0,98))
    range2posn('A1:C3')   =>  ((0,0), (2,2))
    """
    if isinstance(n, int): return (name2posn(n), name2posn(n))
    if not isinstance(n, str): return (name2posn(n), name2posn(n))
    try:
        c,r = n.split(':')
    except:
        c = n.split(':')[0]
        r = c
    if c == '': c = '0'
    if r == '': r = '0'
    return (name2posn(c), name2posn(r))

###########################################################################
