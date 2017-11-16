import unittest

import context
from Sites import Yahoo

###########################################################################
# Yahoo queries/responses as of 2017-11-09
###########################################################################

###########################################################################
class test_cpd_simple(unittest.TestCase):

    # https://query1.finance.yahoo.com/v7/finance/quote?symbols=
    DATA_EMPTY_SYMBOL='{"quoteResponse":{"result":[],"error":null}}'

    # https://query1.finance.yahoo.com/v7/finance/quote?symbols=GARBAGE
    DATA_UNKNOWN_SYMBOL = '{"quoteResponse":{"result":[],"error":null}}'

    # https://query1.finance.yahoo.com/v7/finance/quote?symbols=AAT.L
    DATA_KNOWN_SYMBOL_WRONG_EXCHANGE = '{"quoteResponse":{"result":[{"language":"en-US","quoteType":"EQUITY","tradeable":false,"symbol":"AAT.L"}],"error":null}}'

    def setUp(self): self.o = Yahoo()
    def setTeardown(self): self.o = None

    def test_empty(self):
        r = self.o.parse_json('')
        self.assertEqual(r, {})

    def test_garbage(self):
        r = self.o.parse_json('garbage')
        self.assertEqual(r, {})

    def test_empty_symbol(self):
        r = self.o.parse_json(self.DATA_EMPTY_SYMBOL)
        self.assertEqual(r, {})

    def test_unknown_symbol(self):
        r = self.o.parse_json(self.DATA_UNKNOWN_SYMBOL)
        self.assertEqual(r, {})

    def test_known_symbol_wrong_exchange(self):
        r = self.o.parse_json(self.DATA_KNOWN_SYMBOL_WRONG_EXCHANGE)
        self.assertEqual(r, {'AAT.L': ['', '']} )

###########################################################################
class test_cpd_share(unittest.TestCase):

    # https://query1.finance.yahoo.com/v7/finance/quote?symbols=GSK.L
    DATA_ONE_SHARE = '{"quoteResponse":{"result":[{"language":"en-US","quoteType":"EQUITY","quoteSourceName":"Delayed Quote","currency":"GBp","regularMarketChangePercent":0.5582434,"regularMarketPreviousClose":1343.5,"bid":1350.5,"ask":1351.0,"bidSize":1754,"askSize":4179,"messageBoardId":"finmb_275442","fullExchangeName":"LSE","longName":"GlaxoSmithKline plc","financialCurrency":"GBP","averageDailyVolume3Month":8922796,"averageDailyVolume10Day":10966586,"fiftyTwoWeekLowChange":15.0,"fiftyTwoWeekLowChangePercent":0.011227545,"fiftyTwoWeekHighChange":-373.5,"fiftyTwoWeekHighChangePercent":-0.21658452,"fiftyTwoWeekLow":1336.0,"fiftyTwoWeekHigh":1724.5,"earningsTimestamp":1508914800,"earningsTimestampStart":1517828340,"earningsTimestampEnd":1518177600,"trailingAnnualDividendRate":0.8,"trailingPE":34.55243,"trailingAnnualDividendYield":5.9545966E-4,"marketState":"REGULAR","epsTrailingTwelveMonths":39.1,"epsForward":110.48,"sharesOutstanding":4918680064,"bookValue":0.166,"fiftyDayAverage":1462.6,"fiftyDayAverageChange":-111.599976,"fiftyDayAverageChangePercent":-0.07630246,"twoHundredDayAverage":1558.3348,"twoHundredDayAverageChange":-207.33484,"twoHundredDayAverageChangePercent":-0.13304897,"marketCap":66451369984,"market":"gb_market","shortName":"GLAXOSMITHKLINE PLC ORD 25P","exchangeDataDelayedBy":20,"priceHint":2,"exchange":"LSE","forwardPE":0.12228458,"priceToBook":8138.5547,"sourceInterval":15,"exchangeTimezoneName":"Europe/London","exchangeTimezoneShortName":"GMT","gmtOffSetMilliseconds":0,"tradeable":false,"regularMarketPrice":1351.0,"regularMarketTime":1510241147,"regularMarketChange":7.5,"regularMarketOpen":1349.5,"regularMarketDayHigh":1365.0,"regularMarketDayLow":1345.15,"regularMarketVolume":7346447,"symbol":"GSK.L"}],"error":null}}'

    # https://query1.finance.yahoo.com/v7/finance/quote?symbols=GSK.L,VOD.L
    DATA_TWO_SHARES = '{"quoteResponse":{"result":[{"language":"en-US","quoteType":"EQUITY","quoteSourceName":"Delayed Quote","currency":"GBp","gmtOffSetMilliseconds":0,"fiftyTwoWeekHighChangePercent":-0.3305776,"fiftyTwoWeekLow":178.4,"fiftyTwoWeekHigh":267.32,"fullExchangeName":"LSE","longName":"Barclays PLC","financialCurrency":"GBP","averageDailyVolume3Month":41093009,"forwardPE":0.08242746,"fiftyDayAverage":189.30714,"fiftyDayAverageChange":-10.357147,"fiftyDayAverageChangePercent":-0.05471081,"twoHundredDayAverage":200.38643,"averageDailyVolume10Day":66348256,"fiftyTwoWeekLowChange":0.55000305,"regularMarketPrice":178.95,"regularMarketTime":1510241642,"regularMarketChange":-0.5,"regularMarketOpen":179.45,"regularMarketDayHigh":180.55,"regularMarketDayLow":178.4,"regularMarketVolume":20943903,"regularMarketChangePercent":-0.27862915,"earningsTimestampStart":1469599200,"earningsTimestampEnd":1470045600,"trailingAnnualDividendRate":0.03,"trailingAnnualDividendYield":1.6717748E-4,"regularMarketPreviousClose":179.45,"bid":178.95,"ask":179.0,"bidSize":4233,"askSize":3500,"messageBoardId":"finmb_323899","earningsTimestamp":1508983200,"market":"gb_market","shortName":"BARCLAYS PLC ORD 25P","fiftyTwoWeekLowChangePercent":0.003082977,"fiftyTwoWeekHighChange":-88.37001,"twoHundredDayAverageChange":-21.436432,"twoHundredDayAverageChangePercent":-0.106975466,"marketCap":30500952064,"tradeable":false,"epsTrailingTwelveMonths":-3.2,"epsForward":21.71,"exchange":"LSE","sharesOutstanding":17044400128,"bookValue":3.745,"marketState":"REGULAR","exchangeDataDelayedBy":20,"priceHint":2,"priceToBook":47.78371,"sourceInterval":15,"exchangeTimezoneName":"Europe/London","exchangeTimezoneShortName":"GMT","symbol":"BARC.L"},{"language":"en-US","quoteType":"EQUITY","quoteSourceName":"Delayed Quote","currency":"GBp","gmtOffSetMilliseconds":0,"fiftyTwoWeekHighChangePercent":-0.05536553,"fiftyTwoWeekLow":186.5,"fiftyTwoWeekHigh":233.9,"fullExchangeName":"LSE","longName":"Vodafone Group Plc","financialCurrency":"EUR","averageDailyVolume3Month":53942259,"forwardPE":24.549997,"fiftyDayAverage":213.77715,"fiftyDayAverageChange":7.1728516,"fiftyDayAverageChangePercent":0.03355294,"twoHundredDayAverage":217.56653,"averageDailyVolume10Day":59550585,"fiftyTwoWeekLowChange":34.449997,"regularMarketPrice":220.95,"regularMarketTime":1510241696,"regularMarketChange":5.050003,"regularMarketOpen":216.45,"regularMarketDayHigh":221.2,"regularMarketDayLow":215.95,"regularMarketVolume":52058796,"regularMarketChangePercent":2.3390474,"trailingAnnualDividendRate":0.172,"trailingAnnualDividendYield":7.9666515E-4,"regularMarketPreviousClose":215.9,"bid":220.95,"ask":221.0,"bidSize":0,"askSize":0,"messageBoardId":"finmb_324490","earningsTimestamp":1510648200,"market":"gb_market","shortName":"VODAFONE GROUP PLC ORD USD0.20 ","fiftyTwoWeekLowChangePercent":0.18471849,"fiftyTwoWeekHighChange":-12.949997,"twoHundredDayAverageChange":3.3834686,"twoHundredDayAverageChangePercent":0.015551421,"marketCap":59194269696,"tradeable":false,"epsTrailingTwelveMonths":-26.2,"epsForward":0.09,"exchange":"LSE","sharesOutstanding":26790799360,"bookValue":3.156,"marketState":"REGULAR","exchangeDataDelayedBy":20,"priceHint":2,"priceToBook":70.00951,"sourceInterval":15,"exchangeTimezoneName":"Europe/London","exchangeTimezoneShortName":"GMT","symbol":"VOD.L"}],"error":null}}'

    def setUp(self): self.o = Yahoo()
    def setTeardown(self): self.o = None

    def test_one_share(self):
        r = self.o.parse_json(self.DATA_ONE_SHARE)
        self.assertEqual(r, { 'GSK.L': ['1351.0', 'GBp'] })

    def test_two_shares(self):
        r = self.o.parse_json(self.DATA_TWO_SHARES)
        self.assertEqual(r, {
            'BARC.L': ['178.95', 'GBp'],
            'VOD.L':  ['220.95', 'GBp'],
        })

###########################################################################
class test_cpd_index(unittest.TestCase):

    # https://query1.finance.yahoo.com/v7/finance/quote?symbols=^FTSE
    DATA_ONE_INDEX = '{"quoteResponse":{"result":[{"language":"en-US","quoteType":"INDEX","currency":"GBP","market":"gb_market","exchangeDataDelayedBy":15,"exchange":"FGI","askSize":0,"messageBoardId":"finmb_INDEXFTSE","fullExchangeName":"FTSE Index","averageDailyVolume3Month":737611713,"averageDailyVolume10Day":799295187,"fiftyTwoWeekLowChange":813.8696,"fiftyTwoWeekLowChangePercent":0.12186048,"fiftyTwoWeekHighChange":-106.430176,"fiftyTwoWeekHighChangePercent":-0.014005814,"fiftyTwoWeekLow":6678.7,"fiftyTwoWeekHigh":7599.0,"fiftyDayAverage":7467.536,"fiftyDayAverageChange":25.033691,"fiftyDayAverageChangePercent":0.0033523361,"twoHundredDayAverage":7422.032,"twoHundredDayAverageChange":70.5376,"twoHundredDayAverageChangePercent":0.009503812,"sourceInterval":15,"exchangeTimezoneName":"Europe/London","exchangeTimezoneShortName":"GMT","gmtOffSetMilliseconds":0,"priceHint":2,"regularMarketChangePercent":-0.4933834,"regularMarketPreviousClose":7529.72,"bid":0.0,"ask":0.0,"bidSize":0,"shortName":"FTSE 100","regularMarketPrice":7492.57,"regularMarketTime":1510241538,"regularMarketChange":-37.15039,"regularMarketOpen":7529.72,"regularMarketDayHigh":7532.2,"regularMarketDayLow":7476.89,"regularMarketVolume":0,"tradeable":false,"marketState":"REGULAR","symbol":"^FTSE"}],"error":null}}'

    def setUp(self): self.o = Yahoo()
    def setTeardown(self): self.o = None

    def test_one_index(self):
        r = self.o.parse_json(self.DATA_ONE_INDEX)
        self.assertEqual(r, { '^FTSE': ['7492.57', 'GBP'] })  #NOTE: key is ^FTSE

###########################################################################
class test_cpd_currency_fx(unittest.TestCase):

    # https://query1.finance.yahoo.com/v7/finance/quote?symbols=GBPEUR=X
    DATA_ONE_FX = '{"quoteResponse":{"result":[{"quoteType":"CURRENCY","currency":"EUR","regularMarketPrice":1.1277249,"regularMarketTime":1510243440,"regularMarketChange":-0.0033086538,"regularMarketOpen":1.1310632,"regularMarketDayHigh":1.1292582,"regularMarketDayLow":1.1296135,"regularMarketPreviousClose":1.1310335,"regularMarketChangePercent":-0.29253367,"fiftyTwoWeekRange":"1.1258918 - 1.1592116","exchange":"CCY","messageBoardId":"finmb_GBPEUR_X","priceHint":5,"fiftyTwoWeekLow":1.1258918,"fiftyTwoWeekHigh":1.1592116,"shortName":"GBP/EUR","exchangeTimezoneShortName":"GMT","gmtOffSetMilliseconds":0,"sourceInterval":15,"exchangeTimezoneName":"Europe/London","tradeable":false,"marketState":"REGULAR","fullExchangeName":"CCY","market":"ccy_market","exchangeDataDelayedBy":0,"symbol":"GBPEUR=X"}],"error":null}}'

    def setUp(self): self.o = Yahoo()
    def setTeardown(self): self.o = None

    def test_one_fx(self):
        r = self.o.parse_json(self.DATA_ONE_FX)
        self.assertEqual(r, { 'GBPEUR=X': ['1.1277249', 'EUR'] })  #NOTE: key is GBPEUR=X

###########################################################################
class test_cpd_miscellaneous(unittest.TestCase):
    
    # https://query1.finance.yahoo.com/v7/finance/quote?symbols=AAL.L,AAL.L
    DATA_SAME_SHARE_REPEATED = '{"quoteResponse":{"result":[{"language":"en-US","quoteType":"EQUITY","currency":"GBp","sharesOutstanding":1399869952,"tradeable":false,"bookValue":16.229,"fiftyDayAverage":1415.2428,"fiftyTwoWeekHighChange":-50.5,"fiftyTwoWeekHighChangePercent":-0.032909743,"fiftyTwoWeekLow":950.1,"twoHundredDayAverageChangePercent":0.21708904,"marketCap":20774070272,"trailingAnnualDividendRate":0.48,"trailingPE":5.059666,"regularMarketChangePercent":-1.2312812,"bid":1440.0,"fiftyDayAverageChange":68.7572,"fiftyDayAverageChangePercent":0.048583325,"twoHundredDayAverage":1219.3027,"twoHundredDayAverageChange":264.69727,"priceHint":2,"regularMarketPreviousClose":1502.5,"ask":1500.0,"bidSize":450,"askSize":650,"messageBoardId":"finmb_409115","epsTrailingTwelveMonths":293.3,"epsForward":1.91,"sourceInterval":15,"exchangeTimezoneName":"Europe/London","exchangeTimezoneShortName":"GMT","gmtOffSetMilliseconds":0,"marketState":"POSTPOST","trailingAnnualDividendYield":3.1946754E-4,"regularMarketPrice":1484.0,"regularMarketTime":1510247699,"regularMarketChange":-18.5,"regularMarketOpen":1498.5,"regularMarketDayHigh":1500.5,"regularMarketDayLow":1457.0,"regularMarketVolume":3333147,"exchange":"LSE","fullExchangeName":"LSE","longName":"Anglo American plc","financialCurrency":"USD","averageDailyVolume3Month":5541003,"averageDailyVolume10Day":5140918,"fiftyTwoWeekHigh":1534.5,"earningsTimestamp":1501120800,"exchangeDataDelayedBy":20,"shortName":"ANGLO AMERICAN PLC ORD USD0.549","forwardPE":7.769634,"priceToBook":91.441246,"market":"gb_market","fiftyTwoWeekLowChange":533.9,"fiftyTwoWeekLowChangePercent":0.5619409,"symbol":"AAL.L"},{"language":"en-US","quoteType":"EQUITY","currency":"GBp","sharesOutstanding":1399869952,"tradeable":false,"bookValue":16.229,"fiftyDayAverage":1415.2428,"fiftyTwoWeekHighChange":-50.5,"fiftyTwoWeekHighChangePercent":-0.032909743,"fiftyTwoWeekLow":950.1,"twoHundredDayAverageChangePercent":0.21708904,"marketCap":20774070272,"trailingAnnualDividendRate":0.48,"trailingPE":5.059666,"regularMarketChangePercent":-1.2312812,"bid":1440.0,"fiftyDayAverageChange":68.7572,"fiftyDayAverageChangePercent":0.048583325,"twoHundredDayAverage":1219.3027,"twoHundredDayAverageChange":264.69727,"priceHint":2,"regularMarketPreviousClose":1502.5,"ask":1500.0,"bidSize":450,"askSize":650,"messageBoardId":"finmb_409115","epsTrailingTwelveMonths":293.3,"epsForward":1.91,"sourceInterval":15,"exchangeTimezoneName":"Europe/London","exchangeTimezoneShortName":"GMT","gmtOffSetMilliseconds":0,"marketState":"POSTPOST","trailingAnnualDividendYield":3.1946754E-4,"regularMarketPrice":1484.0,"regularMarketTime":1510247699,"regularMarketChange":-18.5,"regularMarketOpen":1498.5,"regularMarketDayHigh":1500.5,"regularMarketDayLow":1457.0,"regularMarketVolume":3333147,"exchange":"LSE","fullExchangeName":"LSE","longName":"Anglo American plc","financialCurrency":"USD","averageDailyVolume3Month":5541003,"averageDailyVolume10Day":5140918,"fiftyTwoWeekHigh":1534.5,"earningsTimestamp":1501120800,"exchangeDataDelayedBy":20,"shortName":"ANGLO AMERICAN PLC ORD USD0.549","forwardPE":7.769634,"priceToBook":91.441246,"market":"gb_market","fiftyTwoWeekLowChange":533.9,"fiftyTwoWeekLowChangePercent":0.5619409,"symbol":"AAL.L"}],"error":null}}'

    # https://query1.finance.yahoo.com/v7/finance/quote?symbols=AAL.L,GBPUSD=X
    DATA_MIXED='{"quoteResponse":{"result":[{"language":"en-US","quoteType":"CURRENCY","currency":"USD","regularMarketPreviousClose":1.3116475,"bid":1.315824,"ask":1.3157376,"regularMarketChangePercent":0.31842265,"fiftyTwoWeekLow":1.1995153,"fiftyTwoWeekHigh":1.3615816,"regularMarketPrice":1.315824,"regularMarketChange":0.0041764975,"regularMarketOpen":1.3116819,"regularMarketDayHigh":1.316517,"regularMarketDayLow":1.3086778,"fiftyDayAverage":1.3245234,"twoHundredDayAverage":1.3030384,"shortName":"GBP/USD","bidSize":0,"askSize":0,"messageBoardId":"finmb_GBP_X","fiftyTwoWeekLowChange":0.025539994,"fiftyTwoWeekLowChangePercent":0.021291928,"priceHint":5,"fiftyTwoWeekHighChange":-0.07369,"fiftyTwoWeekHighChangePercent":-0.054120883,"gmtOffSetMilliseconds":0,"exchange":"CCY","twoHundredDayAverageChangePercent":-0.005722838,"tradeable":false,"regularMarketTime":1510257480,"regularMarketVolume":0,"market":"ccy_market","fiftyDayAverageChange":0.004991472,"fiftyDayAverageChangePercent":0.003768504,"twoHundredDayAverageChange":-0.0074570775,"fullExchangeName":"CCY","averageDailyVolume3Month":0,"averageDailyVolume10Day":0,"marketState":"REGULAR","sourceInterval":15,"exchangeTimezoneName":"Europe/London","exchangeTimezoneShortName":"GMT","exchangeDataDelayedBy":0,"symbol":"GBPUSD=X"},{"language":"en-US","quoteType":"EQUITY","currency":"GBp","epsTrailingTwelveMonths":293.3,"epsForward":1.91,"regularMarketPreviousClose":1502.5,"bid":1440.0,"ask":1500.0,"bidSize":450,"askSize":650,"messageBoardId":"finmb_409115","regularMarketChangePercent":-1.2312812,"fiftyTwoWeekLowChange":533.9,"fiftyTwoWeekLowChangePercent":0.5619409,"priceHint":2,"fiftyTwoWeekHighChange":-50.5,"fiftyTwoWeekHighChangePercent":-0.032909743,"fiftyTwoWeekLow":950.1,"fiftyTwoWeekHigh":1534.5,"gmtOffSetMilliseconds":0,"trailingAnnualDividendYield":3.1946754E-4,"earningsTimestamp":1501120800,"trailingAnnualDividendRate":0.48,"trailingPE":5.059666,"exchange":"LSE","sharesOutstanding":1399869952,"bookValue":16.229,"twoHundredDayAverageChangePercent":0.21525227,"tradeable":false,"regularMarketPrice":1484.0,"regularMarketTime":1510247699,"regularMarketChange":-18.5,"regularMarketOpen":1498.5,"regularMarketDayHigh":1500.5,"regularMarketDayLow":1457.0,"regularMarketVolume":3333147,"market":"gb_market","fiftyDayAverage":1417.0695,"fiftyDayAverageChange":66.93054,"fiftyDayAverageChangePercent":0.04723166,"twoHundredDayAverage":1221.1456,"twoHundredDayAverageChange":262.85437,"marketCap":20774070272,"shortName":"ANGLO AMERICAN PLC ORD USD0.549","fullExchangeName":"LSE","longName":"Anglo American plc","financialCurrency":"USD","averageDailyVolume3Month":5587442,"averageDailyVolume10Day":5140918,"marketState":"POSTPOST","forwardPE":7.769634,"priceToBook":91.441246,"sourceInterval":15,"exchangeTimezoneName":"Europe/London","exchangeTimezoneShortName":"GMT","exchangeDataDelayedBy":20,"symbol":"AAL.L"}],"error":null}}'

    def setUp(self): self.o = Yahoo()
    def setTeardown(self): self.o = None

    def test_same_share_repeated(self):
        r = self.o.parse_json(self.DATA_SAME_SHARE_REPEATED)
        self.assertEqual(r, {'AAL.L': ['1484.0', 'GBp'] })

    def test_mixed_query(self):
        r = self.o.parse_json(self.DATA_MIXED)
        self.assertEqual(r, {'AAL.L':    ['1484.0',   'GBp'],
                             'GBPUSD=X': ['1.315824', 'USD'],
                         })

###########################################################################
class test_cpd_known_anomalies(unittest.TestCase):

    # Known anomaly:
    #
    # https://www.lemonfool.co.uk/viewtopic.php?f=27&t=8183&start=60#p93146 (vrdriver)
    #
    # https://query1.finance.yahoo.com/v7/finance/quote?symbols=42TE.L
    DATA_ANOMALY_42TE = '{"quoteResponse":{"result":[{"language":"en-US","quoteType":"BOND","currency":"GBP","regularMarketPreviousClose":129.5,"bid":0.0,"ask":0.0,"bidSize":1000,"askSize":3000,"messageBoardId":"finmb_3433747","tradeable":false,"fiftyTwoWeekHighChangePercent":0.0,"fiftyTwoWeekLow":129.5,"fiftyTwoWeekHigh":129.5,"fullExchangeName":"LSE","longName":"Co-operative Group Limited","fiftyTwoWeekLowChange":0.0,"priceHint":2,"market":"gb_market","marketState":"POSTPOST","earningsTimestamp":1474642740,"exchange":"LSE","sourceInterval":15,"exchangeTimezoneName":"Europe/London","exchangeTimezoneShortName":"GMT","gmtOffSetMilliseconds":0,"exchangeDataDelayedBy":20,"regularMarketChangePercent":0.0,"shortName":"CO-OPERATIVE GROUP LIMITED 11% ","fiftyTwoWeekLowChangePercent":0.0,"fiftyTwoWeekHighChange":0.0,"regularMarketPrice":129.5,"regularMarketTime":1510245000,"regularMarketChange":0.0,"regularMarketOpen":125.25,"regularMarketDayHigh":129.5,"regularMarketDayLow":129.5,"regularMarketVolume":7099,"symbol":"42TE.L"}],"error":null}}'

    def setUp(self): self.o = Yahoo()
    def setTeardown(self): self.o = None

    # def test_anomaly_42TE(self):
    #     r = self.o.parse_json(self.DATA_ANOMALY_42TE)
    #     self.assertEqual(r, { '42TE.L': ['129.5', 'GBp'] })  #currency reported in GBP not GBp

###########################################################################
if __name__ == '__main__':
    unittest.main()

###########################################################################
