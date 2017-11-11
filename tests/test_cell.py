import unittest

from cell import cell

###########################################################################
class test_cell_constructor(unittest.TestCase):
    def test_arg_count(self):
        #0 args
        try:
            cell()
        except TypeError:
            self.fail("cell: raised TypeError unexpectedly (0 args)")

        #1 arg
        try:
            cell('A1')
        except TypeError:
            self.fail("cell: raised TypeError unexpectedly (1 arg)")

        #2 args
        try:
            cell(1,2)
        except TypeError:
            self.fail("cell: raised TypeError unexpectedly (2 args)")

        #>2 args
        with self.assertRaises(TypeError):
            cell(1,2,3)

    def test_invalid_arg_fails(self):
        with self.assertRaises(TypeError):
            cell(None)
        with self.assertRaises(TypeError):
            cell(False)
        with self.assertRaises(TypeError):
            cell(True)

###########################################################################
class test_cell_setget_posn(unittest.TestCase):
    def test_set_posn_empty(self):
        self.assertEqual(cell().posn(), (0,0))

    def test_set_posn_singleton_fails(self):
        with self.assertRaises(TypeError):
            cell(1)

    def test_set_posn_negative_fails(self):
        with self.assertRaises(TypeError):
            cell(-1,2)
        with self.assertRaises(TypeError):
            cell(1,-2)

    def test_set_posn_zero(self):
        self.assertEqual(cell(0,0).posn(), (0,0))

    def test_set_posn(self):
        self.assertEqual(cell(1,2).posn(), (1,2))

###########################################################################
class test_cell_set_name(unittest.TestCase):
    def test_set_name_empty(self):
        self.assertEqual(cell().posn(), (0,0))
        self.assertEqual(cell('').posn(), (0,0))
        self.assertEqual(cell(' ').posn(), (0,0))

    def test_set_name_dollars_ignored(self):
        self.assertEqual(cell('$').posn(), (0,0))
        self.assertEqual(cell(' $	').posn(), (0,0))

        with self.assertRaises(TypeError):
            cell('$0')
        with self.assertRaises(TypeError):
            cell('$A')
        with self.assertRaises(TypeError):
            cell('$A$0')

        self.assertEqual(cell('$A$1').posn(), (0,0))
        self.assertEqual(cell('$B$2').posn(), (1,1))

    def test_set_name_zero_fails(self):
        with self.assertRaises(TypeError):
            cell('0' )
        with self.assertRaises(TypeError):
            cell('00')
        with self.assertRaises(TypeError):
            cell('01')
        with self.assertRaises(TypeError):
            cell('02')
        with self.assertRaises(TypeError):
            cell('A0')
        with self.assertRaises(TypeError):
            cell('A' )
        with self.assertRaises(TypeError):
            cell('B' )

    def test_set_name_A(self):
        self.assertEqual(cell('A1').posn(), (0,0))
        self.assertEqual(cell('A2').posn(), (0,1))

    def test_set_name_B(self):
        self.assertEqual(cell('B1').posn(), (1,0))
        self.assertEqual(cell('B2').posn(), (1,1))

    def test_set_name_Z100(self):
        self.assertEqual(cell('Z100').posn(), (25,99))

    #boundaries 1 to 2 chars
    def test_set_name_Z1(self):
        self.assertEqual(cell('Z1').posn(), (25,0))

    def test_set_name_AA1(self):
        self.assertEqual(cell('AA1').posn(), (26,0))

    def test_set_name_AB1(self):
        self.assertEqual(cell('AB1').posn(), (27,0))

    #boundaries within 2 chars
    def test_set_name_AZ1(self):
        self.assertEqual(cell('AZ1').posn(), (51,0))

    def test_set_name_BA1(self):
        self.assertEqual(cell('BA1').posn(), (52,0))

    #boundaries 2 to 3 chars
    def test_set_name_ZZ1(self):
        self.assertEqual(cell('ZZ1' ).posn(), (701,0))

    def test_set_name_AAA1(self):
        self.assertEqual(cell('AAA1').posn(), (702,0))

###########################################################################
class test_cell_get_name(unittest.TestCase):
    def test_get_name_empty(self):
        self.assertEqual(cell().name(), 'A1')

    def test_get_name_in_A_to_Z(self):
        self.assertEqual(cell(0,0).name(), 'A1')
        self.assertEqual(cell(0,1).name(), 'A2')
        self.assertEqual(cell(1,0).name(), 'B1')
        self.assertEqual(cell(1,1).name(), 'B2')

    def test_get_name_crossing_Z_AA_AB(self):
        self.assertEqual(cell(25,0).name(), 'Z1')
        self.assertEqual(cell(26,0).name(), 'AA1')
        self.assertEqual(cell(27,0).name(), 'AB1')

    def test_get_name_crossing_AZ_AA_AB(self):
        self.assertEqual(cell(51,0).name(), 'AZ1')
        self.assertEqual(cell(52,0).name(), 'BA1')
        self.assertEqual(cell(53,0).name(), 'BB1')

    def test_get_name_crossing_ZZ_AAA_AAB(self):
        self.assertEqual(cell(701,0).name(), 'ZZ1')
        self.assertEqual(cell(702,0).name(), 'AAA1')
        self.assertEqual(cell(703,0).name(), 'AAB1')

###########################################################################
class test_cell_compare(unittest.TestCase):
    def test_comparison_wrongclass(self):
        c = cell('A1')
        self.assertFalse(c == 2)
        self.assertTrue(c != 2)
        
    def test_comparison_self(self):
        c = cell('A1')
        self.assertTrue(c == c)
        self.assertFalse(c != c)
        
    def test_comparison_identical(self):
        c = cell('A1')
        d = cell('A1')
        self.assertTrue(c == c)
        self.assertFalse(c != c)

    def test_comparison_different(self):
        c = cell('A1')
        d = cell('A2')
        self.assertFalse(c == d)
        self.assertTrue(c != d)

###########################################################################
