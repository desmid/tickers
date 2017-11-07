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
    get_yahoo_prices_body(XSCRIPTCONTEXT,
                          'Sheet1', keys='A2:A200', datacols=['B', 'C'])
    return None

###########################################################################
# macro guts
###########################################################################
def get_yahoo_prices_body(XSCRIPTCONTEXT,
                          sheetname='Sheet1', keys='A2:A200', datacols=['B']):
    keyrange = range2posn(keys)
    datacols = [name2posn(i)[0] for i in datacols]

    doc = XSCRIPTCONTEXT.getDocument()

    sheets = doc.getSheets()
    mySheet = sheets.getByName(sheetname)
    priceDict = {}

    symbols = sheet_read_symbols(mySheet, keyrange)
    sheet_clear_columns(mySheet, keyrange, datacols)

    html = getHtml(make_yahoo_url(symbols))
    priceDict = createPriceDict(html)
    sheet_write_columns(mySheet, keyrange, datacols, priceDict)

    messageBox(XSCRIPTCONTEXT, "Processing finished", "Status")

def make_yahoo_url(symbols):
    URL = 'https://query1.finance.yahoo.com/v7/finance/quote?'
    return URL + 'symbols=' + ','.join(symbols)

###########################################################################
# spreadsheet utilities
###########################################################################
def name2posn(n=''):
    if isinstance(n, int):
        if n < 0: n = 0
        return (n,n)
    if not isinstance(n, str): return (0,0)
    n = re.sub('[^A-Z0-9:]', '', str(n).upper())
    #0-based arithmetic
    if n is '': n = '1'
    c,r,ordA = 0,0,ord('A')
    while len(n) > 0:
        v = ord(n[0])
        if v < ordA:
            r = int(n)
            break
        #print(n, v, c, r)
        c = 26*c + v-ordA+1
        #print(n, v, c, r)
        n = n[1:]
    if c > 0: c -= 1
    if r > 0: r -= 1
    return (c,r)

def range2posn(n=''):
    if isinstance(n, int): return (name2posn(n), name2posn(n))
    if not isinstance(n, str): return (name2posn(n), name2posn(n))
    try:
        c,r = n.split(':')
    except:
        c = n.split(':')[0]
        r = c
    if c == '': c = '0'
    if r == '': r = '0'
    return (name2posn(c), name2posn(r))

def sheet_read_symbols(sheet, posn):
    p = re.compile(r'[A-Z0-9]+\.[A-Z]')
    return [s for s in sheet_read_column(sheet, posn) if p.match(s)]

def sheet_read_column(sheet, posn):
    ((keycol,rowfirst), (_, rowlast)) = posn
    return [
        sheet.getCellByPosition(keycol, row).getString()
        for row in range(rowfirst, rowlast+1)
    ]

def sheet_clear_columns(sheet, keyrange, datacols, flags=5):
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

def sheet_write_columns(sheet, keyrange, datacols, datadict):
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
# controls
###########################################################################

# Message box test for OO or LO version
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
# general utilities
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

    #priceDict[ticker] = [regularMarketPrice, currency]
    priceDict['%formats'] = ['%f', '%s']
        
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
                    currency = re.sub('GBp', 'GBX', currency)
 
            priceDict[symbol] = [price, currency]
            data = matchObj.group(2)

    return(priceDict)

###########################################################################
g_exportedScripts = test_macro, get_yahoo_prices,
###########################################################################
