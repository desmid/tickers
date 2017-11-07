# Reload macros without restarting LibreOffice
# http://hydrogeotools.blogspot.de/2014/03/libreoffice-and-python-macros.html
# files: reload_macros.py, my_macros.py

# When development is ready, move the macro1 function to a separate
# module, add XSCRIPTCONTEN.getDocument and expose to LibreOffice by
# g_exportedScripts:
#
# def macro1(*args):
#     thisdoc = XSCRIPTCONTEXT.getDocument()
#     #do stuff
#
# g_exportedScripts = macro1,
#
# Packaging into document:
#
# unzip xyz.ods -d newfolder
# cd newfolder
# mkdir -p Scripts/python
# cp /path/to/your_script.py Scripts/python
# vi META-INF/manifest.xml
#
# <manifest:file-entry manifest:full-path="Scripts/python/ReloadMacros.py" manifest:media-type=""/>
# <manifest:file-entry manifest:full-path="Scripts/python/LoadPrices.py" manifest:media-type=""/>
# <manifest:file-entry manifest:full-path="Scripts/python/" manifest:media-type="application/binary"/>
# <manifest:file-entry manifest:full-path="Scripts/" manifest:media-type="application/binary"/>
#
# ls | zip -r@ xyz_new.ods
# or
# zip -r ../xyz_new.ods .

import sys
#import os

sys.path.append('/home/brown/TRADE/SOFTWARE/LibreOffice/t/Scripts/python')

#logfile = os.path.dirname(__file__.replace("file://","").replace("%20"," ")) + "/output.txt"
logfile = '/home/brown/TRADE/SOFTWARE/LibreOffice/out.log'
sys.stdout = open(logfile, "a")

def reload_macros(*args):
    print('RLM: ' , XSCRIPTCONTEXT)

    import LoadPrices
    try:
        reload(LoadPrices)  #python 2
    except NameError:
        import imp          #python 3
        imp.reload(LoadPrices)

    from LoadPrices import test_macro
    
    test_macro(XSCRIPTCONTEXT)

g_exportedScripts = reload_macros,
