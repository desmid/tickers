###########################################################################
import logging
Logger = logging.getLogger('LoadPrices')
Logger.debug("Load: libreoffice.datasheet")

###########################################################################
# http://www.openoffice.org/api/docs/common/ref/com/sun/star/sheet/CellFlags.html
# https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1sheet_1_1CellFlags.html
###########################################################################

import uno
from com.sun.star.sheet.CellFlags import VALUE, DATETIME, STRING, ANNOTATION, \
    FORMULA, HARDATTR, STYLES, OBJECTS, EDITATTR, FORMATTED

LO_CLEAR_FLAGS = (VALUE|STRING)

from libreoffice import CellRange

###########################################################################
def asCellRange(item, template=None):
    """
    asCellRange(string)     return a new CellRange from a string.
    asCellRange(Cell)       return a new CellRange from a Cell.
    asCellRange(CellRange)  return a new CellRange from a CellRange.
    asCellRange(list)       return a list of new CellRanges from a list.
    """
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
        return [asCellRange(i, template) for i in item]
    raise TypeError("unexpected type '%s'" % str(item))

###########################################################################
class DataColumn(object):
    """
    Represents a spreadsheet column range as a CellRange object and a list
    of data values.

      DataColumn(CellRange, list_of_values)

    Operators
      DataColumn[i]       sets or returns i'th data value.
      len(DataColumn)     returns number of data rows

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
        newcells = asCellRange(colname, template=self.cellrange)
        return DataColumn(newcells, [''] * len(self.vec))

    def __setitem__(self, i, val):
        self.vec[i] = val

    def __getitem__(self, i):
        return self.vec[i]

    def __len__(self):
        return len(self.vec)

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

      DataFrame(keycolumn, keylabels, datacols)

    Methods
      keycol()   returns keycol DataColumn.
      columns()  returns DataColumn list.
      update(dict[key] = [val1, val2, ...])
                 iterates over self.keycol looking up keys in the supplied
                 dict; values are written to the corresponding DataColumns
                 to fill out the DataFrame.
    """

    def __init__(self, keycolumn, keyvals, datacols):
        self.keycol = keycolumn
        self.keyvec = self.set_keymask(keyvals)
        self.cframe = [ keycolumn.copy_empty(c) for c in datacols ]

    def set_keymask(self, keyvals):
        vec = [True] * len(self.keycol)
        for i,key in enumerate(self.keycol.rows()):
            try:
                keyvals[key]
            except KeyError:
                vec[i] = False
        return vec

    def keycol(self):
        return self.keycol

    def columns(self):
        return self.cframe

    def has_data(self, i):
        #Logger.debug("has_data[%d]: %s" % (i, str(self.keyvec[i])))
        return self.keyvec[i]

    def update(self, datadict):
        for i,key in enumerate(self.keycol.rows()):
            for j,column in enumerate(self.cframe):
                try:
                    column[i] = datadict[key][j]
                    Logger.debug("update: '%s'  (%d,%d)" % (key, i, j))
                except IndexError:
                    break

    def __repr__(self):
        s = ','.join([str(f) for f in self.cframe])
        return '[' + s + ']'

###########################################################################
class DataSheet(object):
    """
    Represents a spreadsheet with basic column and block operations.

      DataSheet(libreoffice_sheet_object)

    Let 'colid' represent any type convertible to CellRange or containing a
    CellRange as its position (currently DataColumn).

    Methods
      read_column(colid)   return DataColumn from spreadsheet.

      clear_cell(col, row)           clear cell at numeric (col,row).
      clear_column(DataFrame, colid) clear column given by 'colid' using
                                     DataFrame to select cells.
      clear_dataframe(DataFrame)     clear cells given by DataFrame.

      write_cell(col, row, value)  write value to cell at numeric (col,row).
      write_column(DataFrame, DataColumn)     write column from DataColumn.
      write_dataframe(DataFrame)   write cells from DataFrame using
                                   DataFrame to select cells.
    """

    def __init__(self, sheet):
        self.sheet = sheet

    def _get_cells(self, arg):
        if isinstance(arg, str):
            return asCellRange(arg)
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

    def read_cell(self, col, row):
        return self.sheet.getCellByPosition(col, row).getString()

    def read_column(self, column, truncate=False):
        cells = self._get_cells(column)
        ((start_col,start_row), (end_col,end_row)) = cells.posn()
        data = [
            self.read_cell(start_col, row)
            for row in range(start_row, end_row+1)
        ]
        if truncate:
            length = self._find_length(data)
            data = data[:length]
            if length > 0:
                end_row = start_row + length -1
            cells = CellRange(start_col,start_row, end_col,end_row)
        return DataColumn(cells, data)

    def clear_cell(self, col, row):
        if not isinstance(col, int):
            raise TypeError("unexpected type '%s'" % str(col))
        if not isinstance(row, int):
            raise TypeError("unexpected type '%s'" % str(row))
        cell = self.sheet.getCellByPosition(col, row)
        cell.clearContents(LO_CLEAR_FLAGS)
        Logger.debug("clear_cell({},{})".format(col, row))

    def clear_column(self, frame, column):
        if not isinstance(frame, DataFrame):
            raise TypeError("unexpected type '%s'" % str(frame))
        cells = self._get_cells(column)
        #Logger.debug('clear_column: ' + str(cells))
        ((start_col,start_row), (_,end_row)) = cells.posn()
        for i,row in enumerate(range(start_row, end_row+1)):
            if frame.has_data(i):
                self.clear_cell(start_col, row)

    def clear_dataframe(self, frame):
        if not isinstance(frame, DataFrame):
            raise TypeError("unexpected type '%s'" % str(frame))
        for column in frame.columns():
            self.clear_column(frame, column)

    def write_cell(self, col, row, value):
        if not isinstance(col, int):
            raise TypeError("unexpected type '%s'" % str(col))
        if not isinstance(row, int):
            raise TypeError("unexpected type '%s'" % str(row))
        cell = self.sheet.getCellByPosition(col, row)
        #Logger.debug("write_row({},{})={}".format(start_col, i, value))
        if isinstance(value, float):
            cell.Value = value
        elif isinstance(value, int):
            cell.Value = value
        elif isinstance(value, bool):
            cell.Value = (0, 1)[value]
        elif isinstance(value, str):
            cell.String = value
        else:
            cell.String = str(value)

    def write_column(self, frame, column):
        if not isinstance(frame, DataFrame):
            raise TypeError("unexpected type '%s'" % str(frame))
        cells = self._get_cells(column)
        #Logger.debug('write_column: ' + str(cells))
        ((start_col,start_row), (_,end_row)) = cells.posn()
        for i,row in enumerate(range(start_row, end_row+1)):
            if not frame.has_data(i):
                continue
            self.write_cell(start_col, row, column[i])

    def write_dataframe(self, frame):
        if not isinstance(frame, DataFrame):
            raise TypeError("unexpected type '%s'" % str(frame))
        for column in frame.columns():
            self.write_column(frame, column)

###########################################################################
