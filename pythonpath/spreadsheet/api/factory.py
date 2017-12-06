###########################################################################
import logging
Logger = logging.getLogger('LoadPrices')
Logger.debug("Load: spreadsheet.api.factory")

from spreadsheet.api.libreoffice import LibreOffice

###########################################################################
def Spreadsheet(api, **kwargs):
    if api == 'libreoffice':
        return LibreOffice(**kwargs)
    raise AttributeError("unknown spreadsheet type '%s'" % api)

###########################################################################
