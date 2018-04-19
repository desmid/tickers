import logging
Logger = logging.getLogger('LoadPrices')
Logger.debug("Load: spreadsheet.api.libreoffice")

###########################################################################
# http://www.openoffice.org/api/docs/common/ref/com/sun/star/sheet/CellFlags.html
# https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1sheet_1_1CellFlags.html
###########################################################################

from . factory import SpreadsheetAPI


class LibreOffice(SpreadsheetAPI):
    from com.sun.star.sheet.CellFlags import \
        VALUE, DATETIME, STRING, ANNOTATION, FORMULA, HARDATTR, \
        STYLES, OBJECTS, EDITATTR, FORMATTED

    Clear_Flags = (VALUE | STRING)

    from com.sun.star.awt.VclWindowPeerAttribute import \
        OK, OK_CANCEL, YES_NO, YES_NO_CANCEL, RETRY_CANCEL, \
        DEF_OK, DEF_CANCEL, DEF_RETRY, DEF_YES, DEF_NO

    MsgBox_Return = {
        OK:            1,
        OK_CANCEL:     2,
        YES_NO:        3,
        YES_NO_CANCEL: 4,
        RETRY_CANCEL:  5,
        DEF_OK:        6,
        DEF_CANCEL:    7,
        DEF_RETRY:     8,
        DEF_YES:       9,
        DEF_NO:        10,
    }

    def __init__(self, docroot=None):
        if docroot is None:
            raise AttributeError("missing xscriptcontext")
        self.docroot = docroot
        try:
            self.doc = docroot.getDocument()
        except Exception as e:
            raise AttributeError("could not open document: %s" % str(e))

    def _get_cell(self, sheetname, col, row):
        sheet = self.doc.getSheets().getByName(sheetname)
        return sheet.getCellByPosition(col, row)

    #@override
    def clear_cell(self, sheetname, col, row):
        cell = self._get_cell(sheetname, col, row)
        cell.clearContents(self.Clear_Flags)

    #@override
    def read_cell_string(self, sheetname, col, row):
        cell = self._get_cell(sheetname, col, row)
        return cell.getString()

    #@override
    def write_cell_numeric(self, sheetname, col, row, value):
        cell = self._get_cell(sheetname, col, row)
        cell.Value = value

    #@override
    def write_cell_boolean(self, sheetname, col, row, value):
        cell = self._get_cell(sheetname, col, row)
        cell.Value = self._as_boolean(value)

    #@override
    def write_cell_string(self, sheetname, col, row, value):
        cell = self._get_cell(sheetname, col, row)
        cell.String = value

    #@override
    def show_box(self, text, title, value=OK):
        parent = self.doc.CurrentController.Frame.ContainerWindow
        try:
            ctx = self.docroot.getComponentContext()
            val = self._show_newbox(parent, ctx, text, title, value)
        except Exception as e:
            Logger.debug(e)
            val = self._show_legacybox(parent, text, title, value)
        return val

    ########################################
    # LO - works with LO 5
    # OO - works with OO 4
    def _show_newbox(self, parent, ctx, text, title, value):
        tk = ctx.getServiceManager().createInstanceWithContext(
            "com.sun.star.awt.Toolkit", ctx
        )
        box = tk.createMessageBox(
            parent, 0, self.MsgBox_Return[value], title, text
        )
        return box.execute()

    ########################################
    # LO - works with LO 4 early versions
    #    - fails with LO 4.2.3.3 onwards
    # OO - works with OO Portable 3.2
    def _show_legacybox(self, parent, text, title, value):
        wd = WindowDescriptor()
        wd.Type = 1  # MODALTOP
        wd.WindowServiceName = "messbox"
        wd.ParentIndex = -1
        wd.Parent = parent
        wd.WindowAttributes = value
        tk = parent.getToolkit()
        box = tk.createWindow(wd)
        box.setMessageText(text)
        box.setCaptionText(title)
        return box.execute()
