###########################################################################
import logging
Logger = logging.getLogger('LoadPrices')
Logger.debug("Load: LibreOffice.Sheet")

###########################################################################
# http://www.openoffice.org/api/docs/common/ref/com/sun/star/sheet/CellFlags.html
###########################################################################

def read_column(sheet, posn):
    ((keycol,start_row), (_, end_row)) = posn
    return [
        sheet.getCellByPosition(keycol, row).getString()
        for row in range(start_row, end_row+1)
    ]

def clear_columns(sheet, keyrange, datacols, flags=5):
    ((_,start_row), (_, end_row)) = keyrange
    for row in range(start_row, end_row+1):
        for (col,_) in datacols:
            cell = sheet.getCellByPosition(col, row)
            item = cell.getString()
            #Logger.debug("clear({},{})={}".format(col, row, item))
            cell.clearContents(flags)
        row += 1

def write_columns(sheet, keyrange, datacols, datadict, formats):
    if len(datadict) < 1: return
    ((keycol,start_row), (_, end_row)) = keyrange
    for row in range(start_row, end_row+1):
        key = sheet.getCellByPosition(keycol, row).getString()
        try:
            data = datadict[key]
        except KeyError:
            row += 1
            continue
        for i,(col,_) in enumerate(datacols):
            cell = sheet.getCellByPosition(col, row)
            if formats[i] == '%f':
                value = float(data[i])
                #Logger.debug("write({},{},{})={}".format(col, row, key, value))
                cell.Value = value
            elif formats[i] == '%s':
                value = str(data[i])
                #Logger.debug("write({},{},{})={}".format(col, row, key, value))
                cell.String = value
            else:
                value = str(data[i])
                #Logger.debug("write({},{},{})={}".format(col, row, key, value))
                cell.String = value
        row += 1

###########################################################################
