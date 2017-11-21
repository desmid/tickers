###########################################################################
import logging
Logger = logging.getLogger('LoadPrices')
Logger.debug("Load: LibreOffice.Sheet")

###########################################################################
# http://www.openoffice.org/api/docs/common/ref/com/sun/star/sheet/CellFlags.html
# https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1sheet_1_1CellFlags.html
###########################################################################

import uno
from com.sun.star.sheet.CellFlags import VALUE, DATETIME, STRING, ANNOTATION, \
    FORMULA, HARDATTR, STYLES, OBJECTS, EDITATTR, FORMATTED

LO_CLEAR_FLAGS = (VALUE|STRING)

from LibreOffice import CellRange

###########################################################################
class DataColumn(object):
    """
    Represents a spreadsheet column range as a CellRange object and a list
    of data values.

      DataColumn(CellRange, list_of_values)

    Operators
      DataColumn[i]       sets or returns i'th data value.

    Methods
      cells()             returns CellRange object
      rows()              returns data list
      copy_empty(column)  returns an empty DataColumn of same dimension
                          using colname to contruct its CellRange.

    Raises
      IndexError
    """

    def __init__(self, cells, data):
        self.cellrange = cells
        self.vec = data

    def cells(self):
        return self.cellrange

    def rows(self):
        return self.vec

    def copy_empty(self, colname):
        newcells = DataSheet.asCellRange(colname, template=self.cellrange)
        return DataColumn(newcells, [''] * len(self.vec))

    def __setitem__(self, i, val):
        self.vec[i] = val

    def __getitem__(self, i):
        return self.vec[i]

    def __repr__(self):
        s = ','.join([str(c) for c in self.vec])
        return str(self.cellrange) + ' [' + s + ']'

###########################################################################
class DataFrame(object):
    """
    Represents a set of spreadsheet column ranges as a list of DataColumn
    objects of same dimension as a key DataColumn. The Datacolumns need not
    be adjacent or even in vertical register, but must be the same length
    as the key column.

      DataFrame(keycolumn, datacols)

    Methods
      keycol()   returns keycol DataColumn.
      columns()  returns DataColumn list.
      update(dict[key] = [val1, val2, ...])
                 iterates over self.keycol looking up keys in the supplied
                 dict; values are written to the corresponding DataColumns
                 to fill out the DataFrame.
    """

    def __init__(self, keycolumn, datacols):
        self.keycol = keycolumn
        self.frame = [ keycolumn.copy_empty(c) for c in datacols ]

    def keycol(self):
        return self.keycol

    def columns(self):
        return self.frame

    def update(self, datadict):
        for i,key in enumerate(self.keycol.rows()):
            for j,column in enumerate(self.frame):
                try:
                    column[i] = datadict[key][j]
                    Logger.debug("update: '%s'  (%d,%d)" % (key, i, j))
                except IndexError:
                    break

    def __repr__(self):
        s = ','.join([str(f) for f in self.frame])
        return '[' + s + ']'

###########################################################################
class DataSheet(object):
    """
    Represents a spreadsheet with basic column and block operations.

      DataSheet(libreoffice_sheet_object)

    Let 'colid' represent any type convertible to CellRange or containing a
    CellRange as its position (currently DataColumn).

    Methods
      asCellRange(string)  return a CellRange from a string.
      asCellRange(list)    return a list of CellRanges from list of strings.

      read_column(colid)        return DataColumn from spreadsheet.
      clear_column(colid)       clear spreadsheet column.
      write_column(DataColumn)  write spreadsheet column from DataColumn.

      clear_dataframe(DataFrame)  clear spreadsheet cells given by DataFrame.
      write_dataframe(DataFrame)  write spreadsheet cells from DataFrame.
    """

    def __init__(self, sheet):
        self.sheet = sheet

    @classmethod
    def asCellRange(cls, item, template=None):
        if isinstance(item, str):
            cr = CellRange(item)
            if template is not None:
                cr.update_from(template)
            return cr
        if isinstance(item, Cell):
            cr = CellRange(item)
            if template is not None:
                cr.update_from(template)
            return cr
        if isinstance(item, CellRange):
            cr = item
            if template is not None:
                cr.update_from(template)
            return cr
        if isinstance(item, list):
            return [cls.asCellRange(i, template) for i in item]
        raise TypeError("unexpected type '%s'" % str(item))

    def _get_cells(self, arg):
        if isinstance(arg, str):
            return self.asCellRange(arg)
        if isinstance(arg, CellRange):
            return arg
        if isinstance(arg, DataColumn):
            return arg.cells()
        raise TypeError("unexpected type '%s'" % str(arg))

    def _find_length(self, vec, end=''):
        i = len(vec)
        while i:
            if vec[i-1] != end:
                break
            i -= 1
        return i

    def read_column(self, column, truncate=False):
        cells = self._get_cells(column)

        ((start_col,start_row), (end_col,end_row)) = cells.posn()

        data = [
            self.sheet.getCellByPosition(start_col, row).getString()
            for row in range(start_row, end_row+1)
        ]

        if truncate:
            length = self._find_length(data)
            data = data[:length]
            if length > 0:
                end_row = start_row + length -1
            cells = CellRange(start_col,start_row, end_col,end_row)
        return DataColumn(cells, data)

    def clear_column(self, column):
        cells = self._get_cells(column)

        ((start_col,start_row), (_,end_row)) = cells.posn()

        for row in range(start_row, end_row+1):
            cell = self.sheet.getCellByPosition(start_col, row)
            cell.clearContents(LO_CLEAR_FLAGS)
            #Logger.debug("clear_column({},{})".format(start_col, row))

    def write_column(self, column):
        if not isinstance(column, DataColumn):
            raise TypeError("unexpected type '%s'" % str(column))

        cells = column.cells()

        ((start_col,start_row), (_,end_row)) = cells.posn()

        #Logger.debug('write_column: ' + str(cells))
        for i,row in enumerate(range(start_row, end_row+1)):

            cell = self.sheet.getCellByPosition(start_col, row)

            #Logger.debug('write_column:lookup: ' + str(i))
            value = column[i]

            if isinstance(value, float):
                #Logger.debug("write_row({},{})={}".format(start_col, i, value))
                cell.Value = value
                continue

            if isinstance(value, int):
                #Logger.debug("write_row({},{})={}".format(start_col, i, value))
                cell.Value = value
                continue

            if isinstance(value, bool):
                #Logger.debug("write_row({},{})={}".format(start_col, i, value))
                if value is True:
                    cell.Value = 1
                else:
                    cell.Value = 0
                continue

            if isinstance(value, str):
                #Logger.debug("write_row({},{})={}".format(start_col, i, value))
                cell.String = value
                continue

            #Logger.debug("write_row({},{})={}".format(start_col, i, value))
            cell.String = str(value)

    def clear_dataframe(self, frame):
        if not isinstance(frame, DataFrame):
            raise TypeError("unexpected type '%s'" % str(frame))
        for column in frame.columns():
            self.clear_column(column)

    def write_dataframe(self, frame):
        if not isinstance(frame, DataFrame):
            raise TypeError("unexpected type '%s'" % str(frame))
        for column in frame.columns():
            self.write_column(column)

###########################################################################
