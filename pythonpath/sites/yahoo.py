###########################################################################
import logging
Logger = logging.getLogger('LoadPrices')
Logger.debug("Load: Sites.Yahoo")

###########################################################################
import re
from libreoffice import DataSheet, DataFrame
from web import HttpAgent

###########################################################################
class Yahoo(object):
    """
    Spreadsheet driver for Yahoo queries.

    Methods
      get_stocks(sheetname, keyrange, datacols)
      get_fx(sheetname, keyrange, datacols)
      get_indices(sheetname, keyrange, datacols)

         read keyrange columns in sheetname, extract Yahoo tickers,
         assemble URL, fetch query, parse result, populate
         keycols of spreadsheet.

    Raises
      KeyError  if web fetch fails.

    """

    def __init__(self, doc=None):
        self.doc = doc
        self.web = HttpAgent()

    def get_stocks(self, *args, **kwargs):
        self._get(*args, mode='stock', **kwargs)

    def get_fx(self, *args, **kwargs):
        self._get(*args, mode='fx', **kwargs)

    def get_indices(self, *args, **kwargs):
        self._get(*args, mode='index', **kwargs)

    def _get(self, sheetname='Sheet1', keyrange='A2:A200', datacols=['B'],
            mode='stock'):

        sheet = DataSheet(self.doc.getSheets().getByName(sheetname))

        Logger.debug('keyrange: ' + str(keyrange))
        Logger.debug('datacols: ' + str(datacols))

        keydata = sheet.read_column(keyrange, truncate=True)
        Logger.debug('keydata: ' + str(keydata))

        if mode == 'stock':
            keyticker = KeyTickerStock(keydata.rows())
        elif mode == 'fx':
            keyticker = KeyTickerFX(keydata.rows())
        elif mode == 'index':
            keyticker = KeyTickerIndex(keydata.rows())
        else:
            raise AttributeError("unknown mode '%s'" % str(mode))

        Logger.debug('tickers: ' + str(keyticker.tickers()))

        dataframe = DataFrame(keydata, keyticker, datacols)
        Logger.debug('dataframe: ' + str(dataframe))

        sheet.clear_dataframe(dataframe)

        url = keyticker.url()
        Logger.debug('url: ' + url)

        text = self.web.fetch(url)

        if not self.web.ok():
            raise KeyError("fetch URL {} FAILED".format(url))

        pricedict = PriceDict(text, keyticker)
        Logger.debug('pricedict: ' + str(pricedict))

        dataframe.update(pricedict)
        Logger.debug('dataframe: ' + str(dataframe))

        sheet.write_dataframe(dataframe)

###########################################################################
class KeyTickerBase(object):
    """
    Base class provides a read-only dict of spreadsheet cell value to
    Yahoo tickers.

    Constructor
      KeyTicker(list_of_string)

    Operators
      KeyTicker[key]     returns ticker for that key or '' if unmatched.
      len(KeyTickerBase) returns number of stored tickers.

    Methods
      keys()        returns keys.
      items()       returns items.
      values()      returns values.
      tickers()     returns stored tickers as list.
      url(tickers)  returns composed URL using tickers (or stored tickers
                    if no argument).

    """

    URL_BASE = 'http://query1.finance.yahoo.com/v7/finance/quote?'

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

    def keys(self): return self.key2tick.keys()
    def items(self): return self.key2tick.items()
    def values(self): return self.key2tick.values()

    def _extract_tickers(self, data):
        d = {}
        for key in data:
            ticker = self.match_ticker(key)
            if ticker != '':
                d[key] = ticker
        return d

class KeyTickerStock(KeyTickerBase):
    """
    KeyTicker for stocks.
    """
    #patterns could be combined, but easier to test separately:
    CRE_EPIC       = re.compile(r'^([A-Z0-9]{2,4})$')          # BP
    CRE_EPIC_DOT   = re.compile(r'^([A-Z0-9]{2,4})\.$')        # BP.
    CRE_EPIC_FLOOR = re.compile(r'^([A-Z0-9]{2,4}\.[A-Z]+)$')  # BP.L

    def match_ticker(self, text=''):
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
        return ''

class KeyTickerFX(KeyTickerBase):
    """
    KeyTicker for FX currency pairs.
    """
    #patterns could be combined, but easier to test separately:
    CRE_FXPAIR_X   = re.compile(r'^([A-Z]{6}=X)$')              # EURGBP=X
    CRE_FXPAIR_SEP = re.compile(r'^([A-Z]{3})[:/]([A-Z]{3})$')  # EUR:GBP
    CRE_FXPAIR_CH6 = re.compile(r'^([A-Z]{6})$')                # EURGBP

    def match_ticker(self, text=''):
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

class KeyTickerIndex(KeyTickerBase):
    """
    KeyTicker for indices.
    """
    #patterns could be combined, but easier to test separately:
    CRE_INDEX_HAT  = re.compile(r'^(\^[A-Z][A-Z0-9]{2,})$')  # ^FTSE
    CRE_INDEX      = re.compile(r'^([A-Z][A-Z0-9]{2,})$')    # FTSE

    def match_ticker(self, text=''):
        m = self.CRE_INDEX_HAT.search(text)
        if m:
            ticker = m.group(1)
            Logger.debug('INDEX_HAT: %s => %s' % (text, ticker))
            return ticker

        m = self.CRE_INDEX.search(text)
        if m:
            ticker = '^' + m.group(1)
            Logger.debug('INDEX: %s => %s' % (text, ticker))
            return ticker
        return ''

###########################################################################
class PriceDict(object):
    """
    Provides a read-only dict of spreadsheet cell value to Yahoo price
    information from a Yahoo generated JSON string. The optional second
    argument, keydict, provides an indirection on lookups; keys will be
    looked up here first, then in the parsed price dict.

    Constructor
      PriceDict(text, keydict=None)

    Operators
      PriceDict[key]  returns price list for that key:
                        [regularMarketPrice, currency]
                      or a default list for an unmatched non-whitespace key.
      len(PriceDict)  returns number of key,price pairs.

    Methods
      names()      returns list of column names.
      names(i)     returns i'th column name, eg., ['price', 'currency'].

      formats()    returns list of data formatting strings, eg., ['%f', '%s'].
      formats(i)   returns i'th data formatting string.

      defaults     returns list of default values, eg., [0, 'n/a'].
      defaults(i)  returns i'th default value.

      data()       returns ticker to price list dict.

    Raises
      KeyError    if key lookup fails.
      IndexError  if names/formats/defaults index lookup fails.
    """

    NAMES    = ['regularMarketPrice', 'currency']
    FORMATS  = ['%f', '%s']
    DEFAULTS = [0, 'n/a']

    def __init__(self, text, key2ticker=None):
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

    def data(self):
        return self.tick2price

    def __repr__(self):
        return str(self.tick2price) + ', fmt=' + str(self.FORMATS)

    def __len__(self):
        return len(self.tick2price)

    def __getitem__(self, key):
        try:
            if self.key2ticker is None:  #direct key
                ticker = key
            else:  #indirect key
                ticker = self.key2ticker[key]
            return self.tick2price[ticker]
        except KeyError:
            if key.strip() != '':  #whitespace or empty
                return self.defaults()
        return ''

    def _parse_json(self, text=''):
        data = {}

        #extract inner dict key:value "result":[] which is an array
        #containing one dict
        m = re.search(r'"result":\[([^\]]*)\]', text)
        if not m: return data

        #strip double quotes throughout
        text = m.group(1).replace('"','')

        while text:
            #extract one ticker's dict from array
            m = re.search(r'{(.*?)}(.*)', text)
            if not m: break

            #split into key:value pairs
            keyvals = m.group(1).split(',')
            symbol, price, currency = '', '', ''

            for pair in keyvals:
                try:
                    key,val = pair.split(':', 1)
                except:
                    continue

                if key == 'symbol':
                    symbol = val
                    continue

                if key == 'regularMarketPrice':
                    price = val
                    continue

                if key == 'currency':
                    currency = val
                    continue

            #Logger.debug("ypj: %s,%s,%s", symbol,price,currency)
            data[symbol] = [price, currency]
            text = m.group(2)

        return data

###########################################################################
