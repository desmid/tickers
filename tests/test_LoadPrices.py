import unittest
import re

def name2posn(n=''):
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

class test_range2posn(unittest.TestCase):

    def test_range2posn_with_notintstr(self):
        self.assertEqual(range2posn(), ((0,0), (0,0)))
        self.assertEqual(range2posn(None), ((0,0), (0,0)))
        self.assertEqual(range2posn(False), ((0,0), (0,0)))
        self.assertEqual(range2posn(''), ((0,0), (0,0)))

    def test_range2posn_with_A_Z(self):
        self.assertEqual(range2posn('A:Z'), ((0,0), (25,0)))

    def test_range2posn_with_A1_A99(self):
        self.assertEqual(range2posn('A1:A99'), ((0,0), (0,98)))

    def test_range2posn_with_A1_C3(self):
        self.assertEqual(range2posn('A1:C3'), ((0,0), (2,2)))

class test_name2posn(unittest.TestCase):

    def test_name2posn_with_notintstr(self):
        self.assertEqual(name2posn(), (0,0))
        self.assertEqual(name2posn(None), (0,0))
        self.assertEqual(name2posn(False), (0,0))
        self.assertEqual(name2posn(''), (0,0))

    def test_name2posn_with_punctuation(self):
        self.assertEqual(name2posn(' $	'), (0,0))

    def test_name2posn_with_int(self):
        self.assertEqual(name2posn(0), (0,0))
        self.assertEqual(name2posn(1), (1,1))
        self.assertEqual(name2posn(2), (2,2))
        self.assertEqual(name2posn(-2), (0,0))

    def test_name2posn_dollars(self):
        self.assertEqual(name2posn('$0'), (0,0))
        self.assertEqual(name2posn('$A'), (0,0))
        self.assertEqual(name2posn('$A$0'), (0,0))
        self.assertEqual(name2posn('$A$1'), (0,0))

    def test_name2posn_A1(self):
        self.assertEqual(name2posn('0'), (0,0))
        self.assertEqual(name2posn('A'), (0,0))
        self.assertEqual(name2posn('A0'), (0,0))
        self.assertEqual(name2posn('A1'), (0,0))

    def test_name2posn_B1(self):
        self.assertEqual(name2posn('B'), (1,0))
        self.assertEqual(name2posn('B0'), (1,0))
        self.assertEqual(name2posn('B1'), (1,0))

    def test_name2posn_B2(self):
        self.assertEqual(name2posn('B2'), (1,1))

    #boundaries 1 to 2 chars
    def test_name2posn_Z1(self):
        self.assertEqual(name2posn('Z'), (25,0))
    def test_name2posn_AA(self):
        self.assertEqual(name2posn('AA'), (26,0))

    #boundaries within 2 chars
    def test_name2posn_AZ(self):
        self.assertEqual(name2posn('AZ'), (51,0))
    def test_name2posn_BA(self):
        self.assertEqual(name2posn('BA'), (52,0))

    #boundaries 2 to 3 chars
    def test_name2posn_ZZ(self):
        self.assertEqual(name2posn('ZZ'), (701,0))
    def test_name2posn_AAA(self):
        self.assertEqual(name2posn('AAA'), (702,0))
