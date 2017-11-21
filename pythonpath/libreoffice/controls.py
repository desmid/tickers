###########################################################################
import logging
Logger = logging.getLogger('LoadPrices')
Logger.debug("Load: libreoffice.controls")

###########################################################################
# spreadsheet controls
# adapted from LemonFool:SimpleYahooPriceScrape.py
###########################################################################
from com.sun.star.awt.VclWindowPeerAttribute import \
    OK, OK_CANCEL, YES_NO, YES_NO_CANCEL, RETRY_CANCEL, DEF_OK, DEF_CANCEL, \
    DEF_RETRY, DEF_YES, DEF_NO

# arbitrary mappings, for use by the caller:
#     OK = 1
#    YES = 2
#     NO = 3
# CANCEL = 4

# Message box test for OO or LO version
# Uses either messageBoxOO4 or messageBoxLO4. Pretty much pot luck which one
# works, depending on which version of OpenOffice or LibreOffice is used
def messageBox(XSCRIPTCONTEXT, msgText, msgTitle, msgButtons = OK):
    doc = XSCRIPTCONTEXT.getDocument()
    parentWin = doc.CurrentController.Frame.ContainerWindow
    try:
        msgRet = _messageBoxLO4(parentWin, msgText, msgTitle, msgButtons)
    except:
        ctx = XSCRIPTCONTEXT.getComponentContext()
        msgRet = _messageBoxOO4(parentWin, ctx, msgText, msgTitle, msgButtons)
    return msgRet

# Following works with OO4 and LO5 but not LO4 or OO Portable 3.2
# Needs numeric msgButtons 1,2,3 etc so translated by buttonDict
def _messageBoxOO4(parentWin, ctx, MsgText, MsgTitle, msgButtons):
    buttonDict = {OK: 1, OK_CANCEL: 2, YES_NO: 3, YES_NO_CANCEL: 4}
    msgButtons = buttonDict[msgButtons]
    toolkit = ctx.getServiceManager().createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)
    msgbox = toolkit.createMessageBox(parentWin, 0, msgButtons, MsgTitle, MsgText)
    return msgbox.execute()

# Show a message box with the UNO based toolkit.
# Works with LO4 and OOP3.2
# Does NOT work with LO 4.2.3.3
def _messageBoxLO4(parentWin, MsgText, MsgTitle, msgButtons):
    #describe window properties.
    aDescriptor = WindowDescriptor()
    aDescriptor.Type = 1                       #MODALTOP
    aDescriptor.WindowServiceName = "messbox"
    aDescriptor.ParentIndex = -1
    aDescriptor.Parent = parentWin
    aDescriptor.WindowAttributes = msgButtons
    tk = parentWin.getToolkit()
    msgbox = tk.createWindow(aDescriptor)
    msgbox.setMessageText(MsgText)
    msgbox.setCaptionText(MsgTitle)
    return msgbox.execute()

###########################################################################
