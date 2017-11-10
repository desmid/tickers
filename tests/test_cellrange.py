import unittest

from cellrange import cellrange, name2posn, range2posn, posn2name

class test_name2posn(unittest.TestCase):
    def test_name2posn_with_empty(self):
        self.assertEqual(name2posn(), (0,0))
        self.assertEqual(name2posn(None), (0,0))
        self.assertEqual(name2posn(False), (0,0))
        self.assertEqual(name2posn(''), (0,0))
        self.assertEqual(name2posn(' '), (0,0))

    def test_name2posn_with_int(self):
        self.assertEqual(name2posn(0), (0,0))
        self.assertEqual(name2posn(1), (1,1))
        self.assertEqual(name2posn(2), (2,2))
        self.assertEqual(name2posn(-2), (0,0))

    def test_name2posn_dollars(self):
        self.assertEqual(name2posn(' $	'), (0,0))
        self.assertEqual(name2posn('$0'), (0,0))
        self.assertEqual(name2posn('$A'), (0,0))
        self.assertEqual(name2posn('$A$0'), (0,0))
        self.assertEqual(name2posn('$A$1'), (0,0))

    def test_name2posn_A1_forgiving(self):
        self.assertEqual(name2posn('0'), (0,0))
        self.assertEqual(name2posn('00'), (0,0))
        self.assertEqual(name2posn('A0'), (0,0))
        self.assertEqual(name2posn('01'), (0,0))
        self.assertEqual(name2posn('02'), (0,1))

    def test_name2posn_A(self):
        self.assertEqual(name2posn('A'), (0,0))
        self.assertEqual(name2posn('A1'), (0,0))
        self.assertEqual(name2posn('A2'), (0,1))

    def test_name2posn_B(self):
        self.assertEqual(name2posn('B'), (1,0))
        self.assertEqual(name2posn('B1'), (1,0))
        self.assertEqual(name2posn('B2'), (1,1))

    #boundaries 1 to 2 chars
    def test_name2posn_Z1(self):
        self.assertEqual(name2posn('Z'), (25,0))
    def test_name2posn_AA(self):
        self.assertEqual(name2posn('AA'), (26,0))
    def test_name2posn_AB(self):
        self.assertEqual(name2posn('AB'), (27,0))

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

###########################################################################
class test_posn2name(unittest.TestCase):
    def test_posn2name_with_empty(self):
        with self.assertRaises(TypeError):
            posn2name()
        with self.assertRaises(TypeError):
            posn2name(False)
        with self.assertRaises(TypeError):
            posn2name(None)
        with self.assertRaises(TypeError):
            posn2name('')
        with self.assertRaises(TypeError):
            posn2name(0)

    def test_posn2name_with_not_int(self):
        with self.assertRaises(TypeError):
            posn2name( ('a',1) )
        with self.assertRaises(TypeError):
            posn2name( (1,'a') )
        
    def test_posn2name_with_negative(self):
        with self.assertRaises(TypeError):
            posn2name( (-1,0) )
        with self.assertRaises(TypeError):
            posn2name( (0,-1) )
        
    def test_posn2name_with_pair_A_to_Z(self):
        self.assertEqual(posn2name((0,0)), 'A1')
        self.assertEqual(posn2name((0,1)), 'A2')
        self.assertEqual(posn2name((1,0)), 'B1')
        self.assertEqual(posn2name((1,1)), 'B2')

    def test_posn2name_with_pair_Z_to_AA_AB(self):
        self.assertEqual(posn2name((25,0)), 'Z1')
        self.assertEqual(posn2name((26,0)), 'AA1')
        self.assertEqual(posn2name((27,0)), 'AB1')

    def test_posn2name_with_pair_AZ_to_AA_to_AB(self):
        self.assertEqual(posn2name((51,0)), 'AZ1')
        self.assertEqual(posn2name((52,0)), 'BA1')
        self.assertEqual(posn2name((53,0)), 'BB1')

    def test_posn2name_with_pair_ZZ_to_AAA_to_AAB(self):
        self.assertEqual(posn2name((701,0)), 'ZZ1')
        self.assertEqual(posn2name((702,0)), 'AAA1')
        self.assertEqual(posn2name((703,0)), 'AAB1')

###########################################################################
class test_range2posn(unittest.TestCase):
    def test_range2posn_with_empty(self):
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

###########################################################################
class test_cellrange(unittest.TestCase):
    def test_empty(self):
        c = cellrange()
        self.assertEqual(c.start_col, 0)
        self.assertEqual(c.start_row, 0)
        self.assertEqual(c.end_col, 0)
        self.assertEqual(c.end_row, 0)

    def test_set_by_posn(self):
        c = cellrange()
        c.set_by_posn(1,2,3,4)
        self.assertEqual(c.start_col, 1)
        self.assertEqual(c.start_row, 2)
        self.assertEqual(c.end_col, 3)
        self.assertEqual(c.end_row, 4)

    def test_set_by_posn_not_negative(self):
        c = cellrange()
        with self.assertRaises(TypeError):
            c.set_by_posn(-1,2,3,4)
        with self.assertRaises(TypeError):
            c.set_by_posn(1,-2,3,4)
        with self.assertRaises(TypeError):
            c.set_by_posn(1,2,-3,4)
        with self.assertRaises(TypeError):
            c.set_by_posn(1,2,3,-4)

    def test_set_by_name(self):
        c = cellrange()
        c.set_by_name('B2')
        self.assertEqual(c.start_col, 1)
        self.assertEqual(c.start_row, 1)
        self.assertEqual(c.end_col, 1)
        self.assertEqual(c.end_row, 1)

    def test_named_A1(self):
        c = cellrange('A1')
        self.assertEqual(c.start_col, 0)
        self.assertEqual(c.start_row, 0)
        self.assertEqual(c.end_col, 0)
        self.assertEqual(c.end_row, 0)

    def test_named_Z100(self):
        c = cellrange('Z100')
        self.assertEqual(c.start_col, 25)
        self.assertEqual(c.start_row, 99)
        self.assertEqual(c.end_col, 25)
        self.assertEqual(c.end_row, 99)

    def test_range(self):
        c = cellrange('A100:Z1000')
        self.assertEqual(c.start_col, 0)
        self.assertEqual(c.start_row, 99)
        self.assertEqual(c.end_col, 25)
        self.assertEqual(c.end_row, 999)

class test_cellrange_compare(unittest.TestCase):
    def test_comparison_wrongclass(self):
        c = cellrange('A1')
        self.assertFalse(c == 2)
        self.assertTrue(c != 2)
        
    def test_comparison_self(self):
        c = cellrange('A1')
        self.assertTrue(c == c)
        self.assertFalse(c != c)
        
    def test_comparison_identical(self):
        c = cellrange('A1')
        d = cellrange('A1')
        self.assertTrue(c == c)
        self.assertFalse(c != c)

    def test_comparison_different(self):
        c = cellrange('A1')
        d = cellrange('A2')
        self.assertFalse(c == d)
        self.assertTrue(c != d)

class test_cellrange_static_wrappers(unittest.TestCase):
    def test_name2posn_B2(self):
        c = cellrange()
        self.assertEqual(c.name2posn('B2'), (1,1))

    def test_posn2name_B2(self):
        c = cellrange()
        self.assertEqual(c.posn2name((1,1)), 'B2')

    def test_range2posn_with_A1_A99(self):
        c = cellrange()
        self.assertEqual(c.range2posn('A1:A99'), ((0,0), (0,98)))

###########################################################################
