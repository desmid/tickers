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

        keycolumn = CellRange(keys)
        datacols = [Cell(c) for c in datacols]

        Logger.debug('keycolumn: ' + str(keycolumn))
        Logger.debug('datacols:  ' + str(datacols))

        keylist = Sheet.read_column(sheet, keycolumn)

        Sheet.clear_column_list(sheet, keycolumn, datacols)

        Logger.debug('keylist: ' + str(keylist))

        sym2key = self.get_key_symbol_dict(keylist)

        url = self.build_url(sym2key.keys())

        Logger.debug('url: ' + url)

        text = self.web.fetch(url)

        if not self.web.ok():
            raise KeyError("fetch URL {} FAILED".format(url))

        prices = priceDict(text)

        Logger.debug('prices: ' + str(prices))

        self.recover_prices4keys(sym2key, prices)

        Sheet.write_block(sheet, keycolumn, datacols, prices)

    def build_url(self, tickers):
        return self.URL_YAHOO + 'symbols=' + ','.join(tickers)

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

    def tickers(self):
        return self.sym2key.keys()

    def recover_prices4keys(self, sym2key, prices):
        for s,k in sym2key.items():
            prices[k] = prices[s]

###########################################################################
class priceDict(object):
    """
    class behaves like a dict of Yahoo ticker key to price information:

        data[ticker] = [regularMarketPrice, currency]

    Constructor initialises and parses input json string.

    object[key]  sets/returns value list for that key
    len(object)  returns size of dict
    formats() returns list of data formatting strings, ['%f', '%s']
    formats(i) returns i'th of data formatting string
    """
    formats = ['%f', '%s']

    def __init__(self, text=''):
        self.data = self._parse_json(text)

    def get_formats(self, i=None):
        if i is not None:
            return self.formats[i]
        return self.formats

    def __len__(self):
        return len(self.data)

    def __setitem__(self, key, item):
        self.data[key] = item

    def __getitem__(self, key):
        return self.data[key]

    def _parse_json(self, text):
        data = {}

        m = re.search(r'.*?\[(.*?)\].*', text)
        if not m: return data
        text = m.group(1)

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

    def __repr__(self):
        return str(self.data) + ', fmt=' + str(self.formats)

###########################################################################
