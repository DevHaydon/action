import sys
import os
import unittest
from unittest.mock import patch

sys.path.insert(0, os.path.abspath('3_trading_floor'))
import market

class MarketIntegrationTest(unittest.TestCase):
    def test_get_share_price_no_key(self):
        with patch.object(market, 'polygon_api_key', None), \
             patch('market.read_market', return_value=None):
            price = market.get_share_price('AAPL')
            self.assertEqual(price, 0.0)

    def test_get_share_price_api_error(self):
        with patch.object(market, 'polygon_api_key', 'key'), \
             patch.object(market, 'get_share_price_polygon', side_effect=RuntimeError('fail')) as func, \
             patch('market.read_market', return_value=None), \
             patch('market.log_exception') as log_exc:
            price = market.get_share_price('AAPL', retries=1)
            self.assertEqual(price, 0.0)
            self.assertEqual(func.call_count, 2)
            self.assertTrue(log_exc.called)

if __name__ == '__main__':
    unittest.main()
