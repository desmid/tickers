###########################################################################
import logging
Logger = logging.getLogger('LoadPrices')
Logger.debug("Load: libreoffice.controls")

###########################################################################
#Code heavily adapted from LemonFool:SimpleYahooPriceScrape.py
###########################################################################
from com.sun.star.awt.VclWindowPeerAttribute import \
    OK, OK_CANCEL, YES_NO, YES_NO_CANCEL, RETRY_CANCEL, DEF_OK, DEF_CANCEL, \
    DEF_RETRY, DEF_YES, DEF_NO

###########################################################################
class MessageBox(object):
    #popup return codes available to caller
    RETURN = {
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

    @classmethod
    def show(cls, XSCRIPTCONTEXT, text, title, value=OK):
        doc = XSCRIPTCONTEXT.getDocument()
        parent = doc.CurrentController.Frame.ContainerWindow
        try:
            con = XSCRIPTCONTEXT.getComponentContext()
            val = MessageBox._newbox(parent, con, text, title, value)
        except Exception as e:
            Logger.debug(str(e))
            val = MessageBox._legacybox(parent, text, title, value)
        return val

    ###########################################################################
    # LO:
    #   - works with LO 5
    # OO:
    #   - works with OO 4
    def _newbox(parent, con, text, title, value):
        tk = con.getServiceManager().createInstanceWithContext(
            "com.sun.star.awt.Toolkit", con
        )
        box = tk.createMessageBox(
            parent, 0, MessageBox.RETURN[value], title, text
        )
        return box.execute()

    ###########################################################################
    # LO:
    #   - works with LO 4 early versions
    #   - fails with LO 4.2.3.3 onwards
    # OO:
    #   - works with OO Portable 3.2
    def _legacybox(parent, text, title, value):
        wd = WindowDescriptor()
        wd.Type = 1  #MODALTOP
        wd.WindowServiceName = "messbox"
        wd.ParentIndex = -1
        wd.Parent = parent
        wd.WindowAttributes = value
        tk = parent.getToolkit()
        box = tk.createWindow(wd)
        box.setMessageText(text)
        box.setCaptionText(title)
        return box.execute()

###########################################################################
