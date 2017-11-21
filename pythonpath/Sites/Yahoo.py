###########################################################################
import logging
Logger = logging.getLogger('LoadPrices')
Logger.debug("Load: Sites.Yahoo")

###########################################################################
import re
from LibreOffice.Sheet import DataSheet, DataFrame
from Web import HttpAgent

###########################################################################
class Yahoo(object):
    """
    Spreadsheet driver for Yahoo queries.

    Methods
      get(sheetname, keyrange, datacols)

         reads keyrange columns in sheetname, extracts Yahoo tickers,
         assembles URL, fetches query, parses result, populates
         keycols of spreadsheet.

    Raises
      KeyError  if web fetch fails.

    """

    def __init__(self, doc=None):
        self.doc = doc
        self.web = HttpAgent()

    def get(self, sheetname='Sheet1', keyrange='A2:A200', datacols=['B']):
        sheet = DataSheet(self.doc.getSheets().getByName(sheetname))

        Logger.debug('keyrange:  ' + str(keyrange))
        Logger.debug('datacols:  ' + str(datacols))

        keydata = sheet.read_column(keyrange, truncate=True)
        Logger.debug('keydata: ' + str(keydata))

        dataframe = DataFrame(keydata, datacols)
        Logger.debug('dataframe: ' + str(dataframe))

        sheet.clear_dataframe(dataframe)

        keyticker = KeyTicker(keydata.rows())
        Logger.debug('tickers: ' + str(keyticker.tickers()))

        url = keyticker.url()
        Logger.debug('url: ' + url)

        text = self.web.fetch(url)

        if not self.web.ok():
            raise KeyError("fetch URL {} FAILED".format(url))

        pricedict = PriceDict(keyticker, text)
        Logger.debug('pricedict: ' + str(pricedict))

        dataframe.update(pricedict)
        Logger.debug('dataframe: ' + str(dataframe))

        sheet.write_dataframe(dataframe)

###########################################################################
class KeyTicker(object):
    """
    Provides a read-only dict of spreadsheet cell value to Yahoo tickers.

    Constructor
      KeyTicker(list_of_string)

    Operators
      KeyTicker[key]  returns ticker for that key or '' if unmatched.
      len(Key)        returns number of stored tickers.

    Methods
      tickers()     returns stored tickers as list.
      url(tickers)  returns composed URL using tickers (or stored tickers
                    if no argument).
    """

    URL_BASE = 'http://query1.finance.yahoo.com/v7/finance/quote?'

    #some patterns could be combined, but easier to test separately:

    #EPIC
    CRE_EPIC       = re.compile(r'^([A-Z0-9]{2,4})$')           # BP
    CRE_EPIC_DOT   = re.compile(r'^([A-Z0-9]{2,4})\.$')         # BP.
    CRE_EPIC_FLOOR = re.compile(r'^([A-Z0-9]{2,4}\.[A-Z]+)$')   # BP.L

    #index
    CRE_INDEX      = re.compile(r'^(\^[A-Z0-9]+)$')             # ^FTSE

    #currency pair
    CRE_FXPAIR_X   = re.compile(r'^([A-Z]{6}=X)$')              # EURGBP=X 
    CRE_FXPAIR_SEP = re.compile(r'^([A-Z]{3})[:/]([A-Z]{3})$')  # EUR:GBP
    CRE_FXPAIR_CH6 = re.compile(r'^([A-Z]{6})$')                # EURGBP

    def __init__(self, data=[]):
        self.key2tick = self._extract_tickers(data)

    def tickers(self):
        return self.key2tick.values()

    def url(self, tickers=[]):
        if len(tickers) < 1:
            tickers = self.tickers()
        return self.URL_BASE + 'symbols=' + ','.join(tickers)

    def __repr__(self):
        return str(self.key2tick)

    def __len__(self):
        return len(self.key2tick)

    def __getitem__(self, key):
        return self.key2tick[key]

    def _extract_tickers(self, data):
        d = {}
        for key in data:
            ticker = self._match_ticker(key)
            if ticker != '':
                d[key] = ticker
        return d

    def _match_ticker(self, text=''):
        m = self.CRE_EPIC.search(text)
        if m:
            ticker = m.group(1)
            Logger.debug('EPIC: %s => %s' % (text, ticker))
            return ticker

        m = self.CRE_EPIC_DOT.search(text)
        if m:
            ticker = m.group(1)
            Logger.debug('EPIC_DOT: %s => %s' % (text, ticker))
            return ticker

        m = self.CRE_EPIC_FLOOR.search(text)
        if m:
            ticker = m.group(1)
            Logger.debug('EPIC_FLOOR: %s => %s' % (text, ticker))
            return ticker

        m = self.CRE_INDEX.search(text)
        if m:
            ticker = m.group(1)
            Logger.debug('INDEX: %s => %s' % (text, ticker))
            return ticker

        m = self.CRE_FXPAIR_X.search(text)
        if m:
            ticker = m.group(1)
            Logger.debug('FXPAIR_X: %s => %s' % (text, ticker))
            return ticker

        m = self.CRE_FXPAIR_SEP.search(text)
        if m:
            ticker = m.group(1) + m.group(2) + '=X'
            Logger.debug('FXPAIR_SEP: %s => %s' % (text, ticker))
            return ticker

        m = self.CRE_FXPAIR_CH6.search(text)
        if m:
            ticker = m.group(1) + '=X'
            Logger.debug('FXPAIR_CH6: %s => %s' % (text, ticker))
            return ticker

        return ''

###########################################################################
class PriceDict(object):
    """
    Provides a read-only dict of spreadsheet cell value to Yahoo price
    information from a Yahoo generated JSON string.

    Constructor
      PriceDict(KeyTicker, text)

    Operators
      PriceDict[key]  returns price list for that key:
                        [regularMarketPrice, currency]
                      or a default list for an unmatched non-whitespace key.
      len(priceDict)  returns number of key,price pairs.

    Methods
      names()      returns list of column names.
      names(i)     returns i'th column name.
      formats()    returns list of data formatting strings, ['%f', '%s'].
      formats(i)   returns i'th data formatting string.
      defaults     returns list of default values.
      defaults(i)  returns i'th default value.

    Raises
      KeyError    if key lookup fails.
      IndexError  if names/formats/defaults index lookup fails.
    """

    NAMES    = ['regularMarketPrice', 'currency']
    FORMATS  = ['%f', '%s']
    DEFAULTS = [0, 'n/a']

    def __init__(self, key2ticker, text):
        self.key2ticker = key2ticker
        self.tick2price = self._parse_json(text)

    def names(self, i=None):
        if i is None: return self.NAMES
        return self.NAMES[i]

    def formats(self, i=None):
        if i is None: return self.FORMATS
        return self.FORMATS[i]

    def defaults(self, i=None):
        if i is None: return self.DEFAULTS
        return self.DEFAULTS[i]

    def __repr__(self):
        return str(self.tick2price) + ', fmt=' + str(self.FORMATS)

    def __len__(self):
        return len(self.tick2price)

    def __getitem__(self, key):
        try:
            ticker = self.key2ticker[key]
            return self.tick2price[ticker]
        except KeyError:
            if key.strip() != '':  #whitespace or empty
                return self.defaults()
        return ''

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
