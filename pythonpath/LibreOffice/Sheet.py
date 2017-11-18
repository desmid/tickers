###########################################################################
import logging
Logger = logging.getLogger('LoadPrices')
Logger.debug("Load: LibreOffice.Sheet")

###########################################################################
# http://www.openoffice.org/api/docs/common/ref/com/sun/star/sheet/CellFlags.html
# https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1sheet_1_1CellFlags.html
###########################################################################

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
        #Logger.debug("clear_column({},{})".format(col, row))
        cell = sheet.getCellByPosition(col, row)
        cell.clearContents(LO_CLEAR_FLAGS)

def clear_column_list(sheet, keyrange, columns):
    for column in columns:
        clear_column(sheet, keyrange, column)

def write_column(sheet, keyrange, column, datadict, sformat, i):
    if len(datadict) < 1: return
    ((key_col,start_row), (_, end_row)) = keyrange.posn()
    (col,_) = column.posn()
    for row in range(start_row, end_row+1):
        key = sheet.getCellByPosition(key_col, row).getString()
        try:
            data = datadict[key][i]
        except KeyError:
            continue
        cell = sheet.getCellByPosition(col, row)
        if sformat== '%f':
            cell.Value = float(data)
        elif sformat == '%s':
            cell.String = str(data)
        else:
            cell.String = str(data)
        #Logger.debug("write_column({},{},{})={}".format(col, row, key,
        #                 sheet.getCellByPosition(col, row).getString()
        #                                        ))

def write_column_list(sheet, keyrange, datacols, datadict, formats):
    if len(datadict) < 1: return
    for i,column in enumerate(datacols):
        write_column(sheet, keyrange, column, datadict, formats[i], i)

###########################################################################
