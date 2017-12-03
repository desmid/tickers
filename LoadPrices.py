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
LOGFILE = 'out.log'
LOGFORMAT = '%(asctime)s %(levelname)5s [%(filename)-15s %(lineno)4s %(funcName)-25s] %(message)s'
LOGLEVEL = logging.DEBUG
LOGENABLE = True  #set to False to disable logfile

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

#insert embedded zip archive pythonpath
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
from libreoffice.controls import MessageBox

###########################################################################
# the macros
###########################################################################
def get_yahoo_stocks(*args):
    yahoo = Yahoo(DOC)
    yahoo.get_stocks('Sheet1', keyrange='A1:A200', datacols=['B', 'C'])
    MessageBox.show(XSCRIPTCONTEXT, "Processing finished", "Status")

def get_yahoo_fx(*args):
    yahoo = Yahoo(DOC)
    yahoo.get_fx('Sheet1', keyrange='E1:G200', datacols=['F'])
    MessageBox.show(XSCRIPTCONTEXT, "Processing finished", "Status")

def get_yahoo_indices(*args):
    yahoo = Yahoo(DOC)
    yahoo.get_indices('Sheet1', keyrange='H1:H200', datacols=['I', 'J'])
    MessageBox.show(XSCRIPTCONTEXT, "Processing finished", "Status")

g_exportedScripts = get_yahoo_stocks, get_yahoo_fx, get_yahoo_indices,
###########################################################################
