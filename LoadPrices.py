###########################################################################
import sys
import logging
import platform  #for logging system

###########################################################################
# logging
# https://docs.python.org/3/howto/logging.html#logging-advanced-tutorial
###########################################################################
LOGNAME = 'LoadPrices'
LOGFILE = 'out.log'
LOGFORMAT = '%(asctime)s %(levelname)5s [%(filename)-15s %(lineno)4s %(funcName)-25s] %(message)s'
LOGLEVEL = logging.DEBUG
LOGENABLE = True  #set to False to disable logfile

def createLogger(logname, logfile, logformat, loglevel, enabled):
    Logger = logging.getLogger(logname)
    Logger.setLevel(loglevel)
    if enabled:
        tmp = logging.FileHandler(logfile)
        tmp.setFormatter(logging.Formatter(logformat))
        Logger.addHandler(tmp)
    return Logger

Logger = createLogger(LOGNAME, LOGFILE, LOGFORMAT, LOGLEVEL, LOGENABLE)

Logger.info("Start")
Logger.info("Platform: " + ' '.join(platform.uname()))
Logger.info("Python: " + str(sys.version))

###########################################################################
# embedded zip archive pythonpath
PYTHONPATH = '/Scripts/python/pythonpath'

#insert embedded zip archive pythonpath
def prependPath(newpath):
    import uno
    doc = XSCRIPTCONTEXT.getDocument()
    pythonpath = uno.fileUrlToSystemPath(doc.URL) + newpath
    if pythonpath not in sys.path:
        sys.path.insert(0, pythonpath)

prependPath(PYTHONPATH)

Logger.info("Path: " + str(sys.path))

# import embedded
from sites import Yahoo
from spreadsheet.api.factory import Spreadsheet

###########################################################################
# the macros
###########################################################################
DOC = Spreadsheet('libreoffice', docroot=XSCRIPTCONTEXT)

def get_yahoo_stocks(*args):
    yahoo = Yahoo(DOC)
    try:
        yahoo.get_stocks('Sheet1', keyrange='A1:A200', datacols=['B', 'C'])
        DOC.showbox("Processing finished", "Status")
    except Warning as e:
        DOC.showbox(str(e), "Web lookup error")

def get_yahoo_fx(*args):
    yahoo = Yahoo(DOC)
    try:
        yahoo.get_fx('Sheet1', keyrange='E1:G200', datacols=['F'])
        DOC.showbox("Processing finished", "Status")
    except Warning as e:
        DOC.showbox(str(e), "Web lookup error")

def get_yahoo_indices(*args):
    yahoo = Yahoo(DOC)
    try:
        yahoo.get_indices('Sheet1', keyrange='H1:H200', datacols=['I', 'J'])
        DOC.showbox("Processing finished", "Status")
    except Warning as e:
        DOC.showbox(str(e), "Web lookup error")

g_exportedScripts = get_yahoo_stocks, get_yahoo_fx, get_yahoo_indices,
###########################################################################
