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

def read_column(sheet, cells):
    ((start_col,start_row), (_,end_row)) = cells.posn()
    return [
        sheet.getCellByPosition(start_col, row).getString()
        for row in range(start_row, end_row+1)
    ]

def clear_column(sheet, keyrange, column):
    ((_,start_row), (_,end_row)) = keyrange.posn()
    (col,_) = column.posn()
    for row in range(start_row, end_row+1):
        cell = sheet.getCellByPosition(col, row)
        cell.clearContents(LO_CLEAR_FLAGS)
        #Logger.debug("clear_column({},{})".format(col, row))

def clear_column_list(sheet, keyrange, columns):
    for column in columns:
        clear_column(sheet, keyrange, column)

def write_row(sheet, row, datacols, data, key):
    for i,column in enumerate(datacols):
        (col,_) = column.posn()
        cell = sheet.getCellByPosition(col, row)
        try:
            datum, fmt = data[key][i], data.formats(i)
        except (KeyError, IndexError):
            continue
        if fmt == '%f':
            cell.Value = float(datum)
        elif fmt == '%s':
            cell.String = str(datum)
        else:
            cell.String = str(datum)
        #Logger.debug("write_row({},{})={}".format(col, row, datum))

def write_block(sheet, keyrange, datacols, data):
    if len(data) < 1: return
    ((key_col,start_row), (_,end_row)) = keyrange.posn()
    for row in range(start_row, end_row+1):
        key = sheet.getCellByPosition(key_col, row).getString()
        write_row(sheet, row, datacols, data, key)

###########################################################################
