import sys
import os
import unittest
from unittest.mock import patch

sys.path.insert(0, os.path.abspath('3_trading_floor'))
import market

class MarketIntegrationTest(unittest.TestCase):
    def test_get_share_price_no_key(self):
        with patch.object(market, 'polygon_api_key', None), \
             patch('random.randint', return_value=42):
            price = market.get_share_price('AAPL')
            self.assertEqual(price, 42.0)

    def test_get_share_price_api_error(self):
        with patch.object(market, 'polygon_api_key', 'key'), \
             patch.object(market, 'get_share_price_polygon', side_effect=RuntimeError('fail')), \
             patch('random.randint', return_value=77):
            price = market.get_share_price('AAPL')
            self.assertEqual(price, 77.0)

if __name__ == '__main__':
    unittest.main()
