###########################################################################
import sys
import logging
import platform  #for logging system
import uno       #for pythonpath

###########################################################################
# globals
###########################################################################
# LO/OO spreadsheet
DOC = XSCRIPTCONTEXT.getDocument()

# logging
LOGNAME = 'LoadPrices'
LOGFILE = '/home/brown/TRADE/SOFTWARE/LibreOffice/out.log'
LOGFORMAT = '%(asctime)s %(levelname)5s [%(filename)-15s %(lineno)4s %(funcName)-25s] %(message)s'
LOGLEVEL = logging.DEBUG
LOGENABLE = True  #set to False to disable logging to logfile

# embedded pythonpath
PYTHONPATH = '/Scripts/python/pythonpath'

###########################################################################
# https://docs.python.org/3/howto/logging.html#logging-advanced-tutorial
def createLogger(logname, logfile, logformat, loglevel, enabled):
    Logger = logging.getLogger(logname)
    Logger.setLevel(loglevel)
    if enabled:
        tmp = logging.FileHandler(logfile)
        tmp.setFormatter(logging.Formatter(logformat))
        Logger.addHandler(tmp)
    return Logger

def prependPath(doc, newpath):
    pythonpath = uno.fileUrlToSystemPath(doc.URL) + newpath
    if pythonpath not in sys.path:
        sys.path.insert(0, pythonpath)

###########################################################################
Logger = createLogger(LOGNAME, LOGFILE, LOGFORMAT, LOGLEVEL, LOGENABLE)

prependPath(DOC, PYTHONPATH)

Logger.info("Start")
Logger.info(str(platform.uname()))
Logger.info("Python %s", sys.version)
Logger.info("New path: " + str(sys.path))

# import embedded
from sites import Yahoo
from libreoffice.controls import messageBox

###########################################################################
# the macros
###########################################################################
def get_yahoo_prices(*args):
    yahoo = Yahoo(DOC)
    yahoo.get('Sheet1', keyrange='A2:A200', datacols=['B', 'C'])
    messageBox(XSCRIPTCONTEXT, "Processing finished", "Status")

def get_yahoo_fx(*args):
    yahoo = Yahoo(DOC)
    yahoo.get('Sheet1', keyrange='G2:G200', datacols=['H'])
    yahoo.get('Sheet1', keyrange='J2:J200', datacols=['I'])
    messageBox(XSCRIPTCONTEXT, "Processing finished", "Status")

g_exportedScripts = get_yahoo_prices, get_yahoo_fx,
###########################################################################
