###########################################################################
import logging
Logger = logging.getLogger('LoadPrices')
Logger.debug("Load: Sites.Yahoo")

###########################################################################
import re
from LibreOffice import Cell, CellRange
from LibreOffice import Sheet
from Web import HttpAgent

###########################################################################
class Yahoo(object):

    URL_YAHOO = 'https://query1.finance.yahoo.com/v7/finance/quote?'

    CRE_EPIC    = re.compile(r'^([A-Z0-9]{2,4})\.?$')       # BP. => BP
    CRE_EPIC_EX = re.compile(r'^([A-Z0-9]{2,4}\.[A-Z]+)$')  # BP.L => BP.L
    CRE_INDEX   = re.compile(r'^(\^[A-Z0-9]+)$')            # ^FTSE => ^FTSE
    CRE_FXPAIR  = re.compile(r'^((?:[A-Z]{6}){1})(?:=X)?$') # EURGBP=X => EURGBP

    def __init__(self, doc=None):
        self.doc = doc
        self.web = HttpAgent()

    def get(self, sheetname='Sheet1', keys='A2:A200', datacols=['B']):
        sheet = self.doc.getSheets().getByName(sheetname)

        keyrange = CellRange(keys)
        datacols = [Cell(col) for col in datacols]

        Logger.debug('keyrange: ' + str(keyrange))
        Logger.debug('datacols: ' + str(datacols))

        keylist = Sheet.read_column(sheet, keyrange)

        sym2key = self.get_key_symbol_dict(keylist)

        #Logger.debug(str(sym2key))

        Sheet.clear_column_list(sheet, keyrange, datacols)

        url = self.build_url(sym2key.keys())

        Logger.debug(url)

        text = self.web.fetch(url)
        if not self.web.ok():
            raise KeyError("fetch URL {} FAILED".format(url))
            Logger.critical(str(self.web))

        dataDict = self.parse_json(text)

        #Logger.debug(dataDict)

        #repopulate datadict
        for s,k in sym2key.items():
            dataDict[k] = dataDict[s]

        Logger.debug(dataDict)

        Sheet.write_block(sheet, keyrange, datacols, dataDict, ['%f', '%s'])

    def get_key_symbol_dict(self, keylist):
        d = {}
        for key in keylist:
            m = self.CRE_EPIC.search(key)
            if m:
                Logger.debug('EPIC: ' + m.group(1))
                d[m.group(1)] = key
                continue
            m = self.CRE_EPIC_EX.search(key)
            if m:
                Logger.debug('EPIC_EX: ' + m.group(1))
                d[m.group(1)] = key
                continue
            m = self.CRE_INDEX.search(key)
            if m:
                Logger.debug('INDEX: ' + m.group(1))
                d[m.group(1)] = key
                continue
            m = self.CRE_FXPAIR.search(key)
            if m:
                Logger.debug('FXPAIR: ' + m.group(1))
                d[m.group(1) + '=X'] = key
                continue
        return d

    def build_url(self, symbols):
        return self.URL_YAHOO + 'symbols=' + ','.join(symbols)

    def parse_json(self, text):
        data = {}

        m = re.search(r'.*?\[(.*?)\].*', text)
        if not m: return data
        text = m.group(1)

        #data[ticker] = [regularMarketPrice, currency]
        while text:
            m = re.search(r'.*?{(.*?)\}(.*)', text)
            if not m: break

            symbolData = m.group(1)
            symbol, price, currency = '', '', ''

            for element in symbolData.split(','):
                element = element.replace('"','')
                try:
                    key,val = element.split(':', 1)
                except:
                    continue

                if 'symbol' == key:
                    symbol = val
                    continue

                if 'regularMarketPrice' == key:
                    price = val
                    continue

                if 'currency' == key:
                    #currency = val.replace('GBp', 'GBX')
                    currency = val
                    continue

            #Logger.debug("ypj: %s,%s,%s", symbol,price,currency)

            data[symbol] = [price, currency]
            text = m.group(2)

        return data

###########################################################################
