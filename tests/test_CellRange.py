import unittest

from CellRange import CellRange

###########################################################################
class test_CellRange_constructor(unittest.TestCase):
    def test_arg_count(self):
        #0 args
        try:
            CellRange()
        except TypeError:
            self.fail("cell: raised TypeError unexpectedly (0 args)")

        #1 arg
        try:
            CellRange('A1')
        except TypeError:
            self.fail("cell: raised TypeError unexpectedly (1 arg)")

        #2 args
        try:
            CellRange(1,2)
        except TypeError:
            self.fail("cell: raised TypeError unexpectedly (2 args)")

        #3 args
        with self.assertRaises(TypeError):
            CellRange(1,2,3)

        #4 args
        try:
            CellRange(1,2,3,4)
        except TypeError:
            self.fail("cell: raised TypeError unexpectedly (2 args)")

        #>4 args
        with self.assertRaises(TypeError):
            CellRange(1,2,3)

    def test_invalid_arg_fails(self):
        with self.assertRaises(TypeError):
            CellRange(None)
        with self.assertRaises(TypeError):
            CellRange(False)
        with self.assertRaises(TypeError):
            CellRange(True)

    def test_invalid_two_arg_fails(self):
        with self.assertRaises(TypeError):
            CellRange(None,0)
            CellRange(0,None)

    def test_invalid_four_arg_fails(self):
        with self.assertRaises(TypeError):
            CellRange(None,0,0,0)
        with self.assertRaises(TypeError):
            CellRange(0,None,0,0)
        with self.assertRaises(TypeError):
            CellRange(0,0,None,0)
        with self.assertRaises(TypeError):
            CellRange(0,0,0,None)

###########################################################################
class test_cell_setget_posn(unittest.TestCase):
    def test_set_posn_empty(self):
        self.assertEqual(CellRange().posn(), ((0,0),(0,0)) )

    def test_set_posn_singleton_fails(self):
        with self.assertRaises(TypeError):
            CellRange(1)

    def test_set_posn_negative_fails(self):
        with self.assertRaises(TypeError):
            CellRange(-1,2)
        with self.assertRaises(TypeError):
            CellRange(1,-2)
        with self.assertRaises(TypeError):
            CellRange(1,2,-3,4)
        with self.assertRaises(TypeError):
            CellRange(1,2,3,-4)

    def test_set_posn_zero(self):
        self.assertEqual(CellRange().posn(), ((0,0),(0,0)) )
        self.assertEqual(CellRange(0,0).posn(), ((0,0),(0,0)) )
        self.assertEqual(CellRange(0,0,0,0).posn(), ((0,0),(0,0)) )

    def test_set_posn(self):
        self.assertEqual(CellRange(1,2).posn(), ((1,2),(1,2)) )
        self.assertEqual(CellRange(1,2,3,4).posn(), ((1,2),(3,4)) )

###########################################################################
class test_cell_set_name(unittest.TestCase):
    def test_set_name_empty(self):
        self.assertEqual(CellRange().posn(), ((0,0),(0,0)) )
        self.assertEqual(CellRange('').posn(), ((0,0),(0,0)) )
        self.assertEqual(CellRange(' ').posn(), ((0,0),(0,0)) )

    def test_set_name_dollars_ignored(self):
        self.assertEqual(CellRange('$').posn(), ((0,0),(0,0)) )
        self.assertEqual(CellRange(' $	').posn(), ((0,0),(0,0)) )

    def test_set_name_zero_fails(self):
        with self.assertRaises(TypeError):
            CellRange('0' )
        with self.assertRaises(TypeError):
            CellRange('00')
        with self.assertRaises(TypeError):
            CellRange('01')
        with self.assertRaises(TypeError):
            CellRange('02')
        with self.assertRaises(TypeError):
            CellRange('A0')
        with self.assertRaises(TypeError):
            CellRange('A' )
        with self.assertRaises(TypeError):
            CellRange('B' )

    def test_missing_row_numbers_fails(self):
        with self.assertRaises(TypeError):
            CellRange('A:Z1')
        with self.assertRaises(TypeError):
            CellRange('A1:Z')
        with self.assertRaises(TypeError):
            CellRange('A:Z')

    def test_set_name_A(self):
        self.assertEqual(CellRange('A1').posn(), ((0,0),(0,0)) )
        self.assertEqual(CellRange('A2').posn(), ((0,1),(0,1)) )

    def test_range2posn_with_A1_A99(self):
        self.assertEqual(CellRange('A1:A99').posn(), ((0,0),(0,98)))

    def test_range2posn_with_A1_C3(self):
        self.assertEqual(CellRange('A1:C3').posn(), ((0,0),(2,2)))

###########################################################################
class test_CellRange_get_name(unittest.TestCase):
    def test_CellRange_singleton_name_by_posn(self):
        self.assertEqual(CellRange(0,0).name(), 'A1')
        self.assertEqual(CellRange(0,1).name(), 'A2')
        self.assertEqual(CellRange(1,0).name(), 'B1')
        self.assertEqual(CellRange(1,1).name(), 'B2')

    def test_CellRange_singleton_name_by_name(self):
        self.assertEqual(CellRange('A1').name(), 'A1')
        self.assertEqual(CellRange('A2').name(), 'A2')
        self.assertEqual(CellRange('B1').name(), 'B1')
        self.assertEqual(CellRange('B2').name(), 'B2')

    def test_CellRange_range_pair_name_by_posn(self):
        self.assertEqual(CellRange(1,9, 25,9).name(), 'B10:Z10')
        self.assertEqual(CellRange(1,9, 26,10).name(), 'B10:AA11')
        self.assertEqual(CellRange(1,9, 27,11).name(), 'B10:AB12')

    def test_CellRange_range_pair_name_by_name(self):
        self.assertEqual(CellRange('B10:Z10').name(), 'B10:Z10')
        self.assertEqual(CellRange('B10:AA11').name(), 'B10:AA11')
        self.assertEqual(CellRange('B10:AB12').name(), 'B10:AB12')

###########################################################################
class test_CellRange_compare(unittest.TestCase):
    def test_comparison_wrongclass(self):
        c = CellRange('A1:Z12')
        self.assertFalse(c == 2)
        self.assertTrue(c != 2)
        
    def test_comparison_self(self):
        c = CellRange('A1:Z12')
        self.assertTrue(c == c)
        self.assertFalse(c != c)
        
    def test_comparison_identical(self):
        c = CellRange('A1:Z12')
        d = CellRange('A1')
        self.assertTrue(c == c)
        self.assertFalse(c != c)

    def test_comparison_different(self):
        c = CellRange('A1:Z12')
        d = CellRange('A2:Z12')
        self.assertFalse(c == d)
        self.assertTrue(c != d)

###########################################################################
