###########################################################################
import re

from Cell import Cell

class CellRange(object):
    """

    Represent a spreadsheet cell coordinate range. Create with a
    zero-based numeric quadruple (column,row, column,row) or a name
    string 'A1:B20', equivalently:
    
    CellRange(0,2,3,4)  #(column 0, row 2) to (column 3, row 4)
    CellRange('A3:D5')  #ditto, but as written on the spreadsheet

    Ignores $ signs. Raises TypeError if any numerical coordinates
    are <= 0 or if a name does not specify a complete column row.

    Methods:

    posn()  returns ((start_column, start_row), (end_column, end_row))

    Examples:

    posn()          =>  ((0,0), (0,0))
    posn('')        =>  ((0,0), (0,0))
    posn('A:Z')     =>  ((0,0), (25,0))
    posn('A1:A99')  =>  ((0,0), (0,98))
    posn('A1:C3')   =>  ((0,0), (2,2))
    """

    #probably should be able to construct from 1 or 2 Cell() obects as well

    def __init__(self, *args):
        if len(args) == 0:
            self._set_by_posn_two(0,0)
            return

        if len(args) == 1:
            self._set_by_name(*args)
            return

        if len(args) == 2:
            self._set_by_posn_two(*args)
            return

        if len(args) == 4:
            self._set_by_posn_four(*args)
            return

        raise TypeError(
            "CellRange() takes 0, 1, 2 or 4 arguments ({} supplied)".format(len(*args))
        )

    def posn(self): return ( (self.start.posn(), self.end.posn()) )

    def name(self):
        if self.start == self.end:
            return self.start.name()
        return self.start.name() + ':' + self.end.name()

    def _set_by_posn_four(self, start_col=0, start_row=0, end_col=0, end_row=0):
        if not isinstance(start_col, int) or not isinstance(start_row, int):
            raise TypeError("cell positions must be integers")
        if not isinstance(end_col, int) or not isinstance(end_row, int):
            raise TypeError("cell positions must be integers")
        if start_col < 0 or start_row < 0 or end_col < 0 or end_row < 0:
            raise TypeError("cell positions cannot be negative")

        self.start = Cell(start_col, start_row)
        self.end   = Cell(end_col, end_row)

    def _set_by_posn_two(self, start_col=0, start_row=0):
        if not isinstance(start_col, int) or not isinstance(start_row, int):
            raise TypeError("cell positions must be integers")
        if start_col < 0 or start_row < 0:
            raise TypeError("cell positions cannot be negative")

        self.start = Cell(start_col, start_row)
        self.end   = Cell(start_col, start_row)

    def _set_by_name(self, name):
        if not isinstance(name, str):
            raise TypeError("cell name range must be a string")
        self.start, self.end = self._name2posn(name)

    def _name2posn(self, name):
        try:
            c,r = name.split(':', 1)
        except:
            c,r = name, name
        return (Cell(c), Cell(r))

    def __eq__(self, other):
        if isinstance(other, self.__class__):
            return self.__dict__ == other.__dict__
        return False

    def __ne__(self, other):
        return not self.__eq__(other)

###########################################################################