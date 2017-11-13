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
LOGFORMAT = '%(asctime)s %(levelname)5s [%(filename)-15s %(lineno)4s %(funcName)-25s] %(message)s'
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
from LO_Controls import messageBox

###########################################################################
# the macros
###########################################################################
def get_yahoo_prices(*args):
    yahoo = Yahoo(DOC)
    yahoo.get('Sheet1', keys='A2:A200', datacols=['B2', 'C2'])
    messageBox(XSCRIPTCONTEXT, "Processing finished", "Status")

def get_yahoo_fx(*args):
    yahoo = Yahoo(DOC)
    yahoo.get('Sheet1', keys='G2:G200', datacols=['H2'])
    yahoo.get('Sheet1', keys='J2:J200', datacols=['I2'])
    messageBox(XSCRIPTCONTEXT, "Processing finished", "Status")

g_exportedScripts = get_yahoo_prices, get_yahoo_fx,
###########################################################################
