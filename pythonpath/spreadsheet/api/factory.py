import logging
Logger = logging.getLogger('LoadPrices')
Logger.debug("Load: spreadsheet.api.factory")


def spreadsheet_api(name, **kwargs):
    if name == 'libreoffice':
        from . libreoffice import LibreOffice
        return LibreOffice(**kwargs)
    raise AttributeError("unknown spreadsheet type '%s'" % name)


class SpreadsheetAPI(object):

    def clear_cell(self, sheetname, col, row):
        """Virtual method: clear spreadsheet cell."""
        raise NotImplementedError()

    def read_cell_string(self, sheetname, col, row):
        """Virtual method: read value from spreadsheet cell as string."""
        raise NotImplementedError()

    def write_cell_numeric(self, sheetname, col, row, value):
        """Virtual method: write numeric value (float or integer) to
        spreadsheet cell."""
        raise NotImplementedError()

    def write_cell_boolean(self, sheetname, col, row, value):
        """Virtual method: write boolean value to spreadsheet cell."""
        raise NotImplementedError()

    def write_cell_string(self, sheetname, col, row, value):
        """Virtual method: write string to spreadsheet cell."""
        raise NotImplementedError()

    def show_box(self, text, title, value):
        """Virtual method: display message box popup."""
        raise NotImplementedError()

    def _as_boolean(self, value):
        if not isinstance(value, bool):
            raise TypeError("not a boolean '%s'" % str(value))
        return (0, 1)[value]
