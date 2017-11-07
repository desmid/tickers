###########################################################################
# some code thanks to LemonFool software people and SimpleYahooPriceScrape
###########################################################################

###########################################################################
# macros
###########################################################################
import sys
import os
import re

def test_macro(XSCRIPTCONTEXT):
    doc = XSCRIPTCONTEXT.getDocument()
    sheets = doc.getSheets()
    sheet = sheets.getByName("Sheet1")
    tRange = sheet.getCellRangeByName("C4")
    tRange.String = "The Python version is %s.%s.%s" % sys.version_info[:3]
    tRange = sheet.getCellRangeByName("C5")
    tRange.String = sys.executable
    tRange = sheet.getCellRangeByName("C6")
    tRange.String = str(sys.path)
    tRange = sheet.getCellRangeByName("C7")
    tRange.String = str(__file__)
    tRange = sheet.getCellRangeByName("C8")
    tRange.String = str(os.path.dirname(os.path.realpath(__file__)))
    return None

def get_yahoo_prices(*args):
    get_yahoo_prices_body(XSCRIPTCONTEXT, sheetname='Sheet1', keyrow=1, keycol=0, datacols=[1,2])
    return None

###########################################################################
# macro guts
###########################################################################
def get_yahoo_prices_body(XSCRIPTCONTEXT, sheetname, keyrow, keycol, datacols):
    doc = XSCRIPTCONTEXT.getDocument()

    sheets = doc.getSheets()
    mySheet = sheets.getByName(sheetname)
    priceDict = {}

    symbols = sheet_read_symbols(mySheet, keyrow, keycol)
    sheet_clear_columns(mySheet, keyrow, keycol, datacols)

    url = 'https://query1.finance.yahoo.com/v7/finance/quote?symbols=' + ','.join(symbols)
    html = getHtml(url)
    priceDict = createPriceDict(html)

    sheet_write_columns(mySheet, keyrow, keycol, datacols, priceDict)

    messageBox(XSCRIPTCONTEXT, "Processing finished", "Status")

###########################################################################
# utilities
###########################################################################
import socket
try:
    import urllib.request
except:
    import urllib

webTimeOut = 10

def getHtml(myUrl):
    global webTimeOut
    
    try:
        if int(webTimeOut) < 1:
            webTimeOut = 1
        else:
            webTimeOut = int(webTimeOut)
    except:
        webTimeOut = 1

    socket.setdefaulttimeout(webTimeOut)    # set default timeout for web query

    hdr = {
        'User-Agent': 'Mozilla/5.0 AppleWebKit/537.11 (KHTML, like Gecko) Chrome/23.0.1271.64 Safari/537.11',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Charset': 'ISO-8859-1,utf-8;q=0.7,*;q=0.3',
        'Accept-Encoding': 'none',
        'Accept-Language': 'en-US,en;q=0.8',
        'Connection': 'keep-alive'
    }
    
    attempts = 1
    html = 'No Response'
        
    while True:
        try:
            if sys.version[0] == '3':
                req = urllib.request.Request(url=myUrl, headers=hdr)
                response = urllib.request.urlopen(req)
            else:
                #data = urllib.urlencode(values)
                req = urllib2.Request(myUrl, None, hdr)
                response = urllib2.urlopen(req)
            html = response.read()
            html = html.decode('1252', 'ignore')
        except Exception as e:
            attempts += 1
            webTimeOut *= 2
            socket.setdefaulttimeout(webTimeOut)
            if attempts > 5:
                break
        break

    return html

def createPriceDict(html):   
    priceDict = {}
    matchObj = re.search( r'.*?\[(.*?)\].*', html)
    data = ''
    if matchObj:
        data = matchObj.group(1)

    priceDict['%format'] = ['%f', '%s']
        
    while data:
        matchObj = re.search( r'.*?{(.*?)\}(.*)', data)
        if matchObj:
            symbolData = matchObj.group(1)
            symbol = ''
            price = ''
            currency = ''
            for symbolElement in symbolData.split(','):
                symbolElement = symbolElement.replace('"','')
                if 'symbol' in symbolElement:
                    symbol = (symbolElement[symbolElement.index(':')+1:])
                elif 'regularMarketPrice' in symbolElement:
                    price = (symbolElement[symbolElement.index(':')+1:])
                elif 'currency' in symbolElement:
                    currency = (symbolElement[symbolElement.index(':')+1:])
 
            priceDict[symbol] = [price, currency]
            data = matchObj.group(2)

    return(priceDict)

def sheet_read_symbols(sheet, keyrow=0, keycol=0):
    data = []
    p = re.compile(r'[A-Z0-9]+\.[A-Z]')
    for s in sheet_read_column(sheet, keyrow, keycol):
        if p.match(s):
            data.append(s)
    return data

def sheet_read_column(sheet, keyrow=0, keycol=0):
    data = []
    row = keyrow
    while True:
        key = sheet.getCellByPosition(keycol, row).getString()
        if len(key) == 0:
            break
        data.append(key)
        row += 1
    return data

def sheet_clear_columns(sheet, keyrow=0, keycol=0, datacols=[], flags=5):
    """
    http://www.openoffice.org/api/docs/common/ref/com/sun/star/sheet/CellFlags.html
    """
    row = keyrow
    while True:
        key = sheet.getCellByPosition(keycol, row).getString()
        if len(key) == 0:
            break
        for col in datacols:
            cell = sheet.getCellByPosition(col, row)
            item = cell.getString()
            #print("clear({},{})={}".format(col, row, item))
            cell.clearContents(flags)
        row += 1

def sheet_write_columns(sheet, keyrow=0, keycol=0, datacols=[], datadict={}):
    formats = datadict['%format']
    row = keyrow
    while True:
        key = sheet.getCellByPosition(keycol, row).getString()
        if len(key) == 0:
            break
        try:
            data = datadict[key]
        except KeyError:
            row += 1
            continue
        for i, col in enumerate(datacols):
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
                value = 'format?'
                #print("write({},{},{})={}".format(col, row, key, value))
                cell.String = value
        row += 1

####################################################################################
# controls
####################################################################################

# Message box test for OO or LO version
#
# Uses either messageBoxOO4 or messageBoxLO4. Pretty much pot luck which one
# works, depending on which version of OpenOffice or LibreOffice is used

from com.sun.star.awt.VclWindowPeerAttribute import OK, OK_CANCEL, YES_NO, \
    YES_NO_CANCEL, RETRY_CANCEL, DEF_OK, DEF_CANCEL, DEF_RETRY, DEF_YES, DEF_NO

#            OK: 4194304
#     OK_CANCEL: 8388608
#        YES_NO: 16777216
# YES_NO_CANCEL: 33554432
#
# return OK = 1
#       YES = 2
#        NO = 3
#    CANCEL = 0

def messageBox(XSCRIPTCONTEXT, msgText, msgTitle, msgButtons = OK):
    doc = XSCRIPTCONTEXT.getDocument()
    parentWin = doc.CurrentController.Frame.ContainerWindow
    try:
        msgRet = messageBoxLO4(parentWin, msgText, msgTitle, msgButtons)
    except:
        ctx = XSCRIPTCONTEXT.getComponentContext()
        msgRet = messageBoxOO4(parentWin, ctx, msgText, msgTitle, msgButtons)
    return msgRet

### following works with OO4 and LO5 but not LO4 or OO Portable 3.2
# Needs numeric msgButtons 1,2,3 etc so translated by buttonDict
def messageBoxOO4(parentWin, ctx, MsgText, MsgTitle, msgButtons):
    buttonDict = {OK: 1, OK_CANCEL: 2, YES_NO: 3, YES_NO_CANCEL: 0}
    msgButtons = buttonDict[msgButtons]
    toolkit = ctx.getServiceManager().createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)
    msgbox = toolkit.createMessageBox(parentWin, 0, msgButtons, MsgTitle, MsgText)
    msgRet =  msgbox.execute()
    return msgRet

# Show a message box with the UNO based toolkit. Works with LO4 and OOP3.2
# But.... does not work with LO 4.2.3.3
# msgButtons needs to be in format YES_NO_CANCEL or 33554432
def messageBoxLO4(parentWin, MsgText, MsgTitle, msgButtons):
    #describe window properties.
    aDescriptor = WindowDescriptor()
    aDescriptor.Type = 1                    # MODALTOP
    aDescriptor.WindowServiceName = "messbox"
    aDescriptor.ParentIndex = -1
    aDescriptor.Parent = parentWin
    aDescriptor.WindowAttributes = msgButtons   #MsgButton = OK
    tk = parentWin.getToolkit()
    msgbox = tk.createWindow(aDescriptor)
    msgbox.setMessageText(MsgText)
    msgbox.setCaptionText(MsgTitle)
    msgRet = msgbox.execute()
    return msgRet

###########################################################################
g_exportedScripts = test_macro, get_yahoo_prices,
###########################################################################
