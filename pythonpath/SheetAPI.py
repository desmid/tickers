###########################################################################
import logging
Logger = logging.getLogger('LoadPrices')
Logger.debug("Load: SheetAPI")

###########################################################################
def read_column(sheet, posn):
    ((keycol,rowfirst), (_, rowlast)) = posn
    return [
        sheet.getCellByPosition(keycol, row).getString()
        for row in range(rowfirst, rowlast+1)
    ]

def clear_columns(sheet, keyrange, datacols, flags=5):
    """
    http://www.openoffice.org/api/docs/common/ref/com/sun/star/sheet/CellFlags.html
    """
    ((keycol,rowfirst), (_, rowlast)) = keyrange
    for row in range(rowfirst, rowlast+1):
        key = sheet.getCellByPosition(keycol, row).getString()
        for col in datacols:
            cell = sheet.getCellByPosition(col, row)
            item = cell.getString()
            #print("clear({},{})={}".format(col, row, item))
            cell.clearContents(flags)
        row += 1

def write_columns(sheet, keyrange, datacols, datadict):
    if len(datadict) < 1: return
    formats = datadict['%formats']
    ((keycol,rowfirst), (_, rowlast)) = keyrange
    for row in range(rowfirst, rowlast+1):
        key = sheet.getCellByPosition(keycol, row).getString()
        try:
            data = datadict[key]
        except KeyError:
            row += 1
            continue
        for i,col in enumerate(datacols):
            cell = sheet.getCellByPosition(col, row)
            if formats[i] == '%f':
                value = float(data[i])
                #print("write({},{},{})={}".format(col, row, key, value))
                cell.Value = value
            elif formats[i] == '%s':
                value = str(data[i])
                #print("write({},{},{})={}".format(col, row, key, value))
                cell.String = value
            else:
                value = str(data[i])
                #print("write({},{},{})={}".format(col, row, key, value))
                cell.String = value
        row += 1

###########################################################################