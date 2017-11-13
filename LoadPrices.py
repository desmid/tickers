###########################################################################
import sys
import os
import re

###########################################################################
# start logging
# https://docs.python.org/3/howto/logging.html#logging-advanced-tutorial
###########################################################################
import logging
import platform

LOGFILE = '/home/brown/TRADE/SOFTWARE/LibreOffice/out.log'
LOGFORMAT = '%(asctime)s %(levelname)5s [%(lineno)4s - %(funcName)-15s] %(message)s'
LOGLEVEL = logging.DEBUG

Logger = logging.getLogger('LoadPrices')
Logger.setLevel(LOGLEVEL)

tmp = logging.FileHandler(LOGFILE)
tmp.setFormatter(logging.Formatter(LOGFORMAT))
Logger.addHandler(tmp)
del tmp

Logger.debug("Start")
Logger.info(str(platform.uname()))
Logger.info("Python %s", sys.version)

###########################################################################
# LibreOffice/OpenOffice globals
###########################################################################
DOC = XSCRIPTCONTEXT.getDocument()

###########################################################################
# embedded pythonpath and imports
###########################################################################
import uno

PYTHONPATH = '/Scripts/python/pythonpath'

pythonpath = uno.fileUrlToSystemPath(DOC.URL) + PYTHONPATH
if pythonpath not in sys.path: sys.path.append(pythonpath)
del pythonpath

Logger.debug("New path: " + str(sys.path))

#embedded imports go here:
from Yahoo import Yahoo

###########################################################################
# the macros
###########################################################################
def get_yahoo_prices(*args):
    yahoo = Yahoo(DOC)
    yahoo.get('Sheet1', keys='A2:A200', datacols=['B', 'C'])
    messageBox("Processing finished", "Status")

def get_yahoo_fx(*args):
    yahoo = Yahoo(DOC)
    yahoo.get('Sheet1', keys='G2:G200', datacols=['H'])
    yahoo.get('Sheet1', keys='J2:J200', datacols=['I'])
    messageBox("Processing finished", "Status")

###########################################################################
# spreadsheet controls
# adapted from LemonFool:SimpleYahooPriceScrape.py
###########################################################################
from com.sun.star.awt.VclWindowPeerAttribute import OK, OK_CANCEL, YES_NO, \
    YES_NO_CANCEL, RETRY_CANCEL, DEF_OK, DEF_CANCEL, DEF_RETRY, DEF_YES, DEF_NO

# Our arbitrary mappings, for use by the caller:
#     OK = 1
#    YES = 2
#     NO = 3
# CANCEL = 4

# Message box test for OO or LO version
# Uses either messageBoxOO4 or messageBoxLO4. Pretty much pot luck which one
# works, depending on which version of OpenOffice or LibreOffice is used
def messageBox(msgText, msgTitle, msgButtons = OK):
    parentWin = DOC.CurrentController.Frame.ContainerWindow
    try:
        msgRet = messageBoxLO4(parentWin, msgText, msgTitle, msgButtons)
    except:
        ctx = XSCRIPTCONTEXT.getComponentContext()
        msgRet = messageBoxOO4(parentWin, ctx, msgText, msgTitle, msgButtons)
    return msgRet

# Following works with OO4 and LO5 but not LO4 or OO Portable 3.2
# Needs numeric msgButtons 1,2,3 etc so translated by buttonDict
def messageBoxOO4(parentWin, ctx, MsgText, MsgTitle, msgButtons):
    buttonDict = {OK: 1, OK_CANCEL: 2, YES_NO: 3, YES_NO_CANCEL: 4}
    msgButtons = buttonDict[msgButtons]
    toolkit = ctx.getServiceManager().createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)
    msgbox = toolkit.createMessageBox(parentWin, 0, msgButtons, MsgTitle, MsgText)
    msgRet =  msgbox.execute()
    return msgRet

# Show a message box with the UNO based toolkit.
# Works with LO4 and OOP3.2
# Does NOT work with LO 4.2.3.3
def messageBoxLO4(parentWin, MsgText, MsgTitle, msgButtons):
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
    msgRet = msgbox.execute()
    return msgRet

###########################################################################
# general utilities
# adapted from LemonFool:SimpleYahooPriceScrape.py
###########################################################################
import socket
try:
    import urllib.request
except:
    import urllib

WEB_TIMEOUT  = 10  #seconds
WEB_MAXTRIES = 5

def get_html(url, timeout=WEB_TIMEOUT, maxtries=WEB_MAXTRIES):

    if not isinstance(timeout, int): timeout = 1
    if timeout < 1: timeout = 1
    if not isinstance(maxtries, int): maxtries = 1
    if maxtries < 1: maxtries = 1

    hdr = {
        'User-Agent': 'Mozilla/5.0 AppleWebKit/537.11 (KHTML, like Gecko) Chrome/23.0.1271.64 Safari/537.11',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Charset': 'ISO-8859-1,utf-8;q=0.7,*;q=0.3',
        'Accept-Encoding': 'none',
        'Accept-Language': 'en-US,en;q=0.8',
        'Connection': 'keep-alive'
    }
    html = 'No Response'

    tries = 1
    while True:
        if tries > maxtries: break
        socket.setdefaulttimeout(timeout)
        try:
            if sys.version[0] == '3':
                req = urllib.request.Request(url=url, headers=hdr)
                response = urllib.request.urlopen(req)
            else:
                #data = urllib.urlencode(values)
                req = urllib2.Request(url, None, hdr)
                response = urllib2.urlopen(req)
            html = response.read()
            html = html.decode('1252', 'ignore')
            break
        except Exception as e:
            tries += 1
            timeout *= 2

    return html

###########################################################################
g_exportedScripts = get_yahoo_prices, get_yahoo_fx,
###########################################################################
