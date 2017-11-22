import unittest

import context
from yahoo_json import *  #yahoo json strings

from sites.yahoo import PriceDict

###########################################################################
class test_cpd_simple(unittest.TestCase):
    def test_empty(self):
        o = PriceDict('')
        self.assertEqual(o.data(), {})

    def test_empty_symbol(self):
        o = PriceDict(DATA_EMPTY_SYMBOL)
        self.assertEqual(o.data(), {})

    def test_unknown_symbol(self):
        o = PriceDict(DATA_UNKNOWN_SYMBOL)
        self.assertEqual(o.data(), {})

    def test_known_symbol_wrong_exchange(self):
        o = PriceDict(DATA_KNOWN_SYMBOL_WRONG_EXCHANGE)
        self.assertEqual(o.data(), {'AAT.L': ['', '']} )

###########################################################################
class test_cpd_share(unittest.TestCase):
    def test_one_share(self):
        o = PriceDict(DATA_ONE_SHARE)
        self.assertEqual(o.data(), { 'GSK.L': ['1351.0', 'GBp'] })

    def test_two_shares(self):
        o = PriceDict(DATA_TWO_SHARES)
        self.assertEqual(o.data(), {
            'BARC.L': ['178.95', 'GBp'],
            'VOD.L':  ['220.95', 'GBp'],
        })

###########################################################################
class test_cpd_index(unittest.TestCase):
    def test_one_index(self):
        o = PriceDict(DATA_ONE_INDEX)  #NOTE: key is ^FTSE
        self.assertEqual(o.data(), { '^FTSE': ['7492.57', 'GBP'] })

###########################################################################
class test_cpd_currency_fx(unittest.TestCase):
    def test_one_fx(self):
        o = PriceDict(DATA_ONE_FX)  #NOTE: key is GBPEUR=X
        self.assertEqual(o.data(), { 'GBPEUR=X': ['1.1277249', 'EUR'] })

###########################################################################
class test_cpd_miscellaneous(unittest.TestCase):
    def test_same_share_repeated(self):
        o = PriceDict(DATA_SAME_SHARE_REPEATED)
        self.assertEqual(o.data(), {'AAL.L': ['1484.0', 'GBp'] })

    def test_mixed_query(self):
        o = PriceDict(DATA_MIXED)
        self.assertEqual(o.data(), {'AAL.L':    ['1484.0',   'GBp'],
                             'GBPUSD=X': ['1.315824', 'USD'],
                         })

###########################################################################
# class test_cpd_known_anomalies(unittest.TestCase):
#     def test_anomaly_42TE(self):
#         o = PriceDict(DATA_ANOMALY_42TE)  #currency in GBP not GBp
#         self.assertEqual(o.data(), { '42TE.L': ['129.5', 'GBp'] })

###########################################################################
class test_cpd_accessors(unittest.TestCase):
    def test_names(self):
        o = PriceDict('')
        self.assertEqual(o.names(), ['regularMarketPrice', 'currency'])

    def test_formats(self):
        o = PriceDict('')
        self.assertEqual(o.formats(), ['%f','%s'])

    def test_defaults(self):
        o = PriceDict('')
        self.assertEqual(o.defaults(), [0, 'n/a'])

###########################################################################
class test_cpd_dict_getitem_direct(unittest.TestCase):
    def test_lookup_empty_returns_empty(self):
        o = PriceDict('')
        self.assertEqual(o[''],   '')     #empty ok
        self.assertEqual(o[' 	 '], '')  #whitespace ok

    def test_lookup_not_ticker_returns_defaults(self):
        o = PriceDict('')
        self.assertEqual(o['NOT_A_KEY'], o.defaults())

    def test_lookup_one(self):
        o = PriceDict(DATA_ONE_SHARE)
        self.assertEqual(o['GSK.L'], ['1351.0', 'GBp'])

    def test_indexing(self):
        o = PriceDict(DATA_TWO_SHARES)
        self.assertEqual(o['BARC.L'], ['178.95', 'GBp'])
        self.assertEqual(o['VOD.L'],  ['220.95', 'GBp'])

###########################################################################
class test_cpd_dict_getitem_direct_and_indirect(unittest.TestCase):
    def test_lookup_one(self):
        DIRECT_INDIRECT_LOOKUP = {
            'GSK.L'   : 'GSK.L',
            'indirect': 'GSK.L',
        }
        o = PriceDict(DATA_ONE_SHARE, DIRECT_INDIRECT_LOOKUP)
        self.assertEqual(o['GSK.L'],    ['1351.0', 'GBp'])
        self.assertEqual(o['indirect'], ['1351.0', 'GBp'])

###########################################################################
if __name__ == '__main__':
    unittest.main()

###########################################################################
