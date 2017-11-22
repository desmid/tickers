import unittest

from libreoffice.cell import Cell

###########################################################################
class test_Cell_constructor(unittest.TestCase):
    def test_arg_count(self):
        #0 args
        try:
            Cell()
        except TypeError:
            self.fail("Cell: raised TypeError unexpectedly (0 args)")

        #1 arg
        try:
            Cell('A1')
        except TypeError:
            self.fail("Cell: raised TypeError unexpectedly (1 arg)")

        #2 args
        try:
            Cell(1,2)
        except TypeError:
            self.fail("Cell: raised TypeError unexpectedly (2 args)")

        #>2 args
        with self.assertRaises(TypeError):
            Cell(1,2,3)

    def test_invalid_arg_fails(self):
        with self.assertRaises(TypeError):
            Cell(None)
        with self.assertRaises(TypeError):
            Cell(False)
        with self.assertRaises(TypeError):
            Cell(True)

###########################################################################
class test_Cell_setget_posn(unittest.TestCase):
    def test_set_posn_empty(self):
        self.assertEqual(Cell().posn(), (0,0))

    def test_set_posn_singleton_fails(self):
        with self.assertRaises(TypeError):
            Cell(1)

    def test_set_posn_negative_fails(self):
        with self.assertRaises(TypeError):
            Cell(-1,2)
        with self.assertRaises(TypeError):
            Cell(1,-2)

    def test_set_posn_zero(self):
        self.assertEqual(Cell(0,0).posn(), (0,0))

    def test_set_posn(self):
        self.assertEqual(Cell(1,2).posn(), (1,2))

###########################################################################
class test_Cell_set_name(unittest.TestCase):
    def test_set_name_empty(self):
        self.assertEqual(Cell().posn(), (0,0))
        self.assertEqual(Cell('').posn(), (0,0))
        self.assertEqual(Cell(' ').posn(), (0,0))

    def test_set_name_dollars_ignored(self):
        self.assertEqual(Cell('$').posn(), (0,0))
        self.assertEqual(Cell(' $	').posn(), (0,0))

        with self.assertRaises(TypeError):
            Cell('$0')
        with self.assertRaises(TypeError):
            Cell('$A$0')

        self.assertEqual(Cell('$A').posn(), (0,0))
        self.assertEqual(Cell('$1').posn(), (0,0))
        self.assertEqual(Cell('$A$1').posn(), (0,0))
        self.assertEqual(Cell('$B$2').posn(), (1,1))

    def test_set_name_colon_range_fails(self):
        with self.assertRaises(TypeError):
            Cell('A1:A2')

    def test_set_name_zero_fails(self):
        with self.assertRaises(TypeError):
            Cell('0' )
        with self.assertRaises(TypeError):
            Cell('00')

    def test_singleton_columns(self):
        with self.assertRaises(TypeError):
            Cell('A0')

        self.assertEqual(Cell('A').posn(), (0,0))
        self.assertEqual(Cell('B').posn(), (1,0))
        self.assertEqual(Cell('AAA1').posn(), (702,0))

    def test_singleton_rows(self):
        self.assertEqual(Cell('1').posn(), (0,0))
        self.assertEqual(Cell('01').posn(), (0,0))
        self.assertEqual(Cell('02').posn(), (0,1))

    def test_set_name_A(self):
        self.assertEqual(Cell('A1').posn(), (0,0))
        self.assertEqual(Cell('A2').posn(), (0,1))

    def test_set_name_B(self):
        self.assertEqual(Cell('B1').posn(), (1,0))
        self.assertEqual(Cell('B2').posn(), (1,1))

    def test_set_name_Z100(self):
        self.assertEqual(Cell('Z100').posn(), (25,99))

    #boundaries 1 to 2 chars
    def test_set_name_Z1(self):
        self.assertEqual(Cell('Z1').posn(), (25,0))

    def test_set_name_AA1(self):
        self.assertEqual(Cell('AA1').posn(), (26,0))

    def test_set_name_AB1(self):
        self.assertEqual(Cell('AB1').posn(), (27,0))

    #boundaries within 2 chars
    def test_set_name_AZ1(self):
        self.assertEqual(Cell('AZ1').posn(), (51,0))

    def test_set_name_BA1(self):
        self.assertEqual(Cell('BA1').posn(), (52,0))

    #boundaries 2 to 3 chars
    def test_set_name_ZZ1(self):
        self.assertEqual(Cell('ZZ1' ).posn(), (701,0))

    def test_set_name_AAA1(self):
        self.assertEqual(Cell('AAA1').posn(), (702,0))

###########################################################################
class test_Cell_get_name(unittest.TestCase):
    def test_get_name_empty(self):
        self.assertEqual(Cell().name(), 'A1')

    def test_get_name_in_A_to_Z(self):
        self.assertEqual(Cell(0,0).name(), 'A1')
        self.assertEqual(Cell(0,1).name(), 'A2')
        self.assertEqual(Cell(1,0).name(), 'B1')
        self.assertEqual(Cell(1,1).name(), 'B2')

    def test_get_name_crossing_Z_AA_AB(self):
        self.assertEqual(Cell(25,0).name(), 'Z1')
        self.assertEqual(Cell(26,0).name(), 'AA1')
        self.assertEqual(Cell(27,0).name(), 'AB1')

    def test_get_name_crossing_AZ_AA_AB(self):
        self.assertEqual(Cell(51,0).name(), 'AZ1')
        self.assertEqual(Cell(52,0).name(), 'BA1')
        self.assertEqual(Cell(53,0).name(), 'BB1')

    def test_get_name_crossing_ZZ_AAA_AAB(self):
        self.assertEqual(Cell(701,0).name(), 'ZZ1')
        self.assertEqual(Cell(702,0).name(), 'AAA1')
        self.assertEqual(Cell(703,0).name(), 'AAB1')

###########################################################################
class test_Cell_compare(unittest.TestCase):
    def test_comparison_wrongclass(self):
        c = Cell('A1')
        self.assertFalse(c == 2)
        self.assertTrue(c != 2)
        
    def test_comparison_self(self):
        c = Cell('A1')
        self.assertTrue(c == c)
        self.assertFalse(c != c)
        
    def test_comparison_identical(self):
        c = Cell('A1')
        d = Cell('A1')
        self.assertTrue(c == c)
        self.assertFalse(c != c)

    def test_comparison_different(self):
        c = Cell('A1')
        d = Cell('A2')
        self.assertFalse(c == d)
        self.assertTrue(c != d)

###########################################################################
class test_Cell_update(unittest.TestCase):
    def test_update_row_number(self):
        c = Cell('A')
        d = Cell('B100')
        self.assertEqual(c._internal(), (0, None))
        self.assertEqual(c.name(), 'A1')
        c.update_from(d)
        self.assertEqual(c._internal(), (0, 99))
        self.assertEqual(c.name(), 'A100')

    def test_update_column_name(self):
        c = Cell('1')
        d = Cell('B100')
        self.assertEqual(c._internal(), (None, 0))
        self.assertEqual(c.name(), 'A1')
        c.update_from(d)
        self.assertEqual(c._internal(), (1, 0))
        self.assertEqual(c.name(), 'B1')

    def test_update_column_name_and_row_number(self):
        c = Cell('')
        d = Cell('B100')
        self.assertEqual(c._internal(), (None, None))
        self.assertEqual(c.name(), 'A1')
        c.update_from(d)
        self.assertEqual(c._internal(), (1, 99))
        self.assertEqual(c.name(), 'B100')

    def test_update_no_force(self):
        c = Cell('A1')
        d = Cell('B100')
        self.assertEqual(c._internal(), (0, 0))
        self.assertEqual(c.name(), 'A1')
        c.update_from(d)
        self.assertEqual(c._internal(), (0, 0))
        self.assertEqual(c.name(), 'A1')

    def test_update_force(self):
        c = Cell('A1')
        d = Cell('B100')
        self.assertEqual(c.name(), 'A1')
        c.update_from(d, force=True)
        self.assertEqual(c.name(), 'B100')


###########################################################################
if __name__ == '__main__':
    unittest.main()

###########################################################################
