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

from LibreOffice import Cell, CellRange

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
        newcells = asCellRange(colname, template=self.cellrange)
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
def asCellRange(item, template=None):
    if isinstance(item, str):
        cr = CellRange(item)
        if template is not None:
            cr.update_from(template)
        return cr
    if isinstance(item, list):
        return [asCellRange(i, template) for i in item]
    raise TypeError("asCellRange() unexpected type '%s'" % str(item))

def check_arg_type(arg):
    if isinstance(arg, str):
        return asCellRange(arg)
    if isinstance(arg, CellRange):
        return arg
    if isinstance(arg, DataColumn):
        return arg.cells()
    raise TypeError("argument must be a string or CellRange or DataColumn")

def clear_dataframe(sheet, dataframe):
    for column in dataframe.columns():
        clear_column(sheet, column)

def write_dataframe(sheet, dataframe):
    for column in dataframe.columns():
        write_column(sheet, column)

def read_column(sheet, column, truncate=False):

    def find_length(vec, end=''):
        i = len(vec)
        while i:
            if vec[i-1] != end:
                break
            i -= 1
        return i

    cells = check_arg_type(column)

    ((start_col,start_row), (end_col,end_row)) = cells.posn()

    data = [
        sheet.getCellByPosition(start_col, row).getString()
        for row in range(start_row, end_row+1)
    ]

    if truncate:
        length = find_length(data)
        data = data[:length]
        if length > 0:
            end_row = start_row + length -1
        cells = CellRange(start_col,start_row, end_col,end_row)
    return DataColumn(cells, data)

def clear_column(sheet, column):
    cells = check_arg_type(column)

    ((start_col,start_row), (_,end_row)) = cells.posn()

    for row in range(start_row, end_row+1):
        cell = sheet.getCellByPosition(start_col, row)
        cell.clearContents(LO_CLEAR_FLAGS)
        #Logger.debug("clear_column({},{})".format(start_col, row))

def write_column(sheet, column):
    cells = check_arg_type(column)

    ((start_col,start_row), (_,end_row)) = cells.posn()

    #Logger.debug('write_column: ' + str(cells))
    for i,row in enumerate(range(start_row, end_row+1)):

        cell = sheet.getCellByPosition(start_col, row)
        
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
