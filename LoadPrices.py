###########################################################################
# macros
###########################################################################
import sys
import os

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

# g_exportedScripts = test_macro,
###########################################################################
