###########################################################################
import logging
Logger = logging.getLogger('LoadPrices')
Logger.debug("Load: Sites.Yahoo")

###########################################################################
import re
from LibreOffice import Sheet
from Web import HttpAgent

###########################################################################
class Yahoo(object):

    def __init__(self, doc=None):
        self.doc = doc
        self.web = HttpAgent()

    def get(self, sheetname='Sheet1', keys='A2:A200', datacols=['B']):
        sheet = self.doc.getSheets().getByName(sheetname)

        keycolumn = Sheet.asCellRange(keys)
        datacols  = Sheet.asCellRange(datacols)

        Logger.debug('keycolumn: ' + str(keycolumn))
        Logger.debug('datacols:  ' + str(datacols))

        keylist = Sheet.read_column(sheet, keycolumn)

        Sheet.clear_columns(sheet, keycolumn, datacols)

        Logger.debug('keylist: ' + str(keylist))

        prices = keyPriceDict(self, keylist)

        url = prices.make_url()

        Logger.debug('url: ' + url)

        text = self.web.fetch(url)

        if not self.web.ok():
            raise KeyError("fetch URL {} FAILED".format(url))

        prices.parse(text)

        Logger.debug('prices: ' + str(prices))

        Sheet.write_block(sheet, keycolumn, datacols, prices)

###########################################################################
class keyPriceDict(object):
    """
    Provides a read-only dict of spreadsheet cell value to Yahoo price
    information from a Yahoo generated JSON string:

      keyPriceDict(list_of_sheet_cell_values)
      keyPriceDict[cell_value] = [regularMarketPrice, currency]

    keyPriceDict[key]  returns value list for that key or a default
                       value if the key is non-whiespace and not matched
    len(priceDict)     returns size of contained price dict
    tickers()          returns list of extracted tickers
    parse(text)        parses a Yahoo JSON string and stores the result
    formats()          returns list of data formatting strings, ['%f', '%s']
    formats(i)         returns i'th of data formatting string
    """

    URL_BASE = 'https://query1.finance.yahoo.com/v7/finance/quote?'

    CRE_EPIC    = re.compile(r'^([A-Z0-9]{2,4})\.?$')       # BP. => BP
    CRE_EPIC_EX = re.compile(r'^([A-Z0-9]{2,4}\.[A-Z]+)$')  # BP.L => BP.L
    CRE_INDEX   = re.compile(r'^(\^[A-Z0-9]+)$')            # ^FTSE => ^FTSE
    CRE_FXPAIR  = re.compile(r'^((?:[A-Z]{6}){1})(?:=X)?$') # EURGBP=X => EURGBP

    def __init__(self, tickmaker, keylist):
        self.key2tick = self._keys_to_tickers(keylist)
        self.tick2price = None

    def formats(self, i=None):
        if i is None: return self.tick2price.formats()
        return self.tick2price.formats(i)

    def tickers(self):
        return self.key2tick.values()

    def make_url(self, tickers=None):
        if tickers is None:
            tickers = self.tickers()
        return self.URL_BASE + 'symbols=' + ','.join(tickers)

    def parse(self, text):
        self.tick2price = priceDict(text)

    def __repr__(self):
        return str(self.tick2price)

    def __len__(self):
        return len(self.tick2price)

    def __getitem__(self, key):
        try:
            ticker = self.key2tick[key]
            return self.tick2price[ticker]
        except KeyError:
            if key.strip() != '':
                return self.tick2price.defaults()
        raise KeyError

    def _keys_to_tickers(self, keylist):
        d = {}
        for key in keylist:
            m = self.CRE_EPIC.search(key)
            if m:
                d[key] = m.group(1)
                Logger.debug('EPIC: %s => %s' % (key, d[key]))
                continue

            m = self.CRE_EPIC_EX.search(key)
            if m:
                d[key] = m.group(1)
                Logger.debug('EPIC_EX: %s => %s' % (key, d[key]))
                continue

            m = self.CRE_INDEX.search(key)
            if m:
                d[key] = m.group(1)
                Logger.debug('INDEX: %s => %s' % (key, d[key]))
                continue

            m = self.CRE_FXPAIR.search(key)
            if m:
                d[key] = m.group(1) + '=X'
                Logger.debug('FXPAIR: %s => %s' % (key, d[key]))
                continue
        return d

###########################################################################
class priceDict(object):
    """
    Provides a read-only dict of Yahoo ticker to Yahoo price information
    from a Yahoo generated JSON string:

      priceDict(json_string)
      priceDict[ticker] = [regularMarketPrice, currency]

    Constructor initialises and parses input json string.

    priceDict[key]  returns value for that key as list
    len(priceDict)  returns size of dict
    names()         returns list of column names
    names(i)        returns i'th column name
    formats()       returns list of formatting strings, ['%f', '%s']
    formats(i)      returns i'th formatting string
    defaults        returns list of default values
    defaults(i)     returns i'th default value
    data()          returns whole internal dict
    """

    NAMES    = ['regularMarketPrice', 'currency']
    FORMATS  = ['%f', '%s']
    DEFAULTS = [0, 'n/a']

    def __init__(self, text=''):
        self._data = self._parse_json(text)

    def names(self, i=None):
        if i is None: return self.NAMES
        return self.NAMES[i]

    def formats(self, i=None):
        if i is None: return self.FORMATS
        return self.FORMATS[i]

    def defaults(self, i=None):
        if i is None: return self.DEFAULTS
        return self.DEFAULTS[i]

    def data(self):
        return self._data

    def __repr__(self):
        return str(self._data) + ', fmt=' + str(self.FORMATS)

    def __len__(self):
        return len(self._data)

    def __getitem__(self, key):
        return self._data[key]

    def _parse_json(self, text=''):
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
                    currency = val
                    continue

            #Logger.debug("ypj: %s,%s,%s", symbol,price,currency)
            data[symbol] = [price, currency]
            text = m.group(2)

        return data

###########################################################################
