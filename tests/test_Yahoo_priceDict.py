import unittest

import context
from yahoo_json import *  #yahoo json strings

from Sites.Yahoo import priceDict

###########################################################################
class test_cpd_misc(unittest.TestCase):
    def test_formats(self):
        o = priceDict('')
        self.assertEqual(o.get_formats(), ['%f','%s'])

###########################################################################
class test_cpd_lookups(unittest.TestCase):
    def test_lookup_empty(self):
        o = priceDict()
        with self.assertRaises(KeyError):
            o['NOT_A_KEY']

    def test_lookup_one(self):
        o = priceDict(DATA_ONE_SHARE)
        self.assertEqual(o['GSK.L'], ['1351.0', 'GBp'])

    def test_indexing(self):
        o = priceDict(DATA_TWO_SHARES)
        self.assertEqual(o['BARC.L'], ['178.95', 'GBp'])
        self.assertEqual(o['VOD.L'], ['220.95', 'GBp'])

###########################################################################
class test_cpd_simple(unittest.TestCase):

    def test_empty(self):
        o = priceDict('')
        self.assertEqual(o.get_data(), {})

    def test_garbage(self):
        o = priceDict('garbage')
        self.assertEqual(o.get_data(), {})

    def test_empty_symbol(self):
        o = priceDict(DATA_EMPTY_SYMBOL)
        self.assertEqual(o.get_data(), {})

    def test_unknown_symbol(self):
        o = priceDict(DATA_UNKNOWN_SYMBOL)
        self.assertEqual(o.get_data(), {})

    def test_known_symbol_wrong_exchange(self):
        o = priceDict(DATA_KNOWN_SYMBOL_WRONG_EXCHANGE)
        self.assertEqual(o.get_data(), {'AAT.L': ['', '']} )

###########################################################################
class test_cpd_share(unittest.TestCase):

    def test_one_share(self):
        o = priceDict(DATA_ONE_SHARE)
        self.assertEqual(o.get_data(), { 'GSK.L': ['1351.0', 'GBp'] })

    def test_two_shares(self):
        o = priceDict(DATA_TWO_SHARES)
        self.assertEqual(o.get_data(), {
            'BARC.L': ['178.95', 'GBp'],
            'VOD.L':  ['220.95', 'GBp'],
        })

###########################################################################
class test_cpd_index(unittest.TestCase):

    def test_one_index(self):
        o = priceDict(DATA_ONE_INDEX)
        self.assertEqual(o.get_data(), { '^FTSE': ['7492.57', 'GBP'] })  #NOTE: key is ^FTSE

###########################################################################
class test_cpd_currency_fx(unittest.TestCase):

    def test_one_fx(self):
        o = priceDict(DATA_ONE_FX)
        self.assertEqual(o.get_data(), { 'GBPEUR=X': ['1.1277249', 'EUR'] })  #NOTE: key is GBPEUR=X

###########################################################################
class test_cpd_miscellaneous(unittest.TestCase):
    
    def test_same_share_repeated(self):
        o = priceDict(DATA_SAME_SHARE_REPEATED)
        self.assertEqual(o.get_data(), {'AAL.L': ['1484.0', 'GBp'] })

    def test_mixed_query(self):
        o = priceDict(DATA_MIXED)
        self.assertEqual(o.get_data(), {'AAL.L':    ['1484.0',   'GBp'],
                             'GBPUSD=X': ['1.315824', 'USD'],
                         })

###########################################################################
class test_cpd_known_anomalies(unittest.TestCase):
    pass

    # def test_anomaly_42TE(self):
    #     o = priceDict(DATA_ANOMALY_42TE)
    #     self.assertEqual(o.get_data(), { '42TE.L': ['129.5', 'GBp'] })  #currency reported in GBP not GBp

###########################################################################
if __name__ == '__main__':
    unittest.main()

###########################################################################
