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
    """
    def __init__(self, keycolumn, datacols):
        self.frame = [ keycolumn.copy_empty(c) for c in datacols ]

    def columns(self):
        return self.frame

    def __repr__(self):
        s = ','.join([str(f) for f in self.frame])
        return '[' + s + ']'

###########################################################################
class DataSheet(object):

    def __init__(self, sheet):
        self.sheet = sheet

    @classmethod
    def asCellRange(cls, item, template=None):
        if isinstance(item, str):
            cr = CellRange(item)
            if template is not None:
                cr.update_from(template)
            return cr
        if isinstance(item, list):
            return [cls.asCellRange(i, template) for i in item]
        raise TypeError("unexpected type '%s'" % str(item))

    def _check_arg_type(self, arg):
        if isinstance(arg, str):
            return self.asCellRange(arg)
        if isinstance(arg, CellRange):
            return arg
        if isinstance(arg, DataColumn):
            return arg.cells()
        raise TypeError("argument must be a string or CellRange or DataColumn")

    def _find_length(self, vec, end=''):
        i = len(vec)
        while i:
            if vec[i-1] != end:
                break
            i -= 1
        return i

    def clear_dataframe(self, dataframe):
        for column in dataframe.columns():
            self.clear_column(column)

    def write_dataframe(self, dataframe):
        for column in dataframe.columns():
            self.write_column(column)

    def read_column(self, column, truncate=False):

        cells = self._check_arg_type(column)

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
        cells = self._check_arg_type(column)

        ((start_col,start_row), (_,end_row)) = cells.posn()

        for row in range(start_row, end_row+1):
            cell = self.sheet.getCellByPosition(start_col, row)
            cell.clearContents(LO_CLEAR_FLAGS)
            #Logger.debug("clear_column({},{})".format(start_col, row))

    def write_column(self, column):
        cells = self._check_arg_type(column)

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

###########################################################################
