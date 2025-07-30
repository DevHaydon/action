import sys
import os
import unittest
from unittest.mock import patch

# Make trading floor modules importable
sys.path.insert(0, os.path.abspath('3_trading_floor'))

import accounts


class AccountOperationsTest(unittest.TestCase):
    def setUp(self):
        # In-memory store instead of sqlite
        self.store = {}

        def fake_write(name, data):
            self.store[name.lower()] = data

        def fake_read(name):
            return self.store.get(name.lower())

        patches = [
            patch('accounts.write_account', side_effect=fake_write),
            patch('accounts.read_account', side_effect=fake_read),
            patch('accounts.write_log'),
            patch('accounts.get_share_price', return_value=100.0),
        ]
        for p in patches:
            p.start()
        self.addCleanup(lambda: [p.stop() for p in patches])

        self.account = accounts.Account.get('Alice')

    def test_buy_and_sell_shares(self):
        # Buy 10 shares at $100 each plus spread
        self.account.buy_shares('AAPL', 10, 'init')
        expected_balance_after_buy = accounts.INITIAL_BALANCE - 10 * 100 * (1 + accounts.SPREAD)
        self.assertAlmostEqual(self.account.balance, expected_balance_after_buy)
        self.assertEqual(self.account.holdings, {'AAPL': 10})

        # Sell 5 of them
        self.account.sell_shares('AAPL', 5, 'take profit')
        expected_balance_after_sell = expected_balance_after_buy + 5 * 100 * (1 - accounts.SPREAD)
        self.assertAlmostEqual(self.account.balance, expected_balance_after_sell)
        self.assertEqual(self.account.holdings, {'AAPL': 5})
        self.assertEqual(len(self.account.transactions), 2)

    def test_profit_loss_calculation(self):
        self.account.buy_shares('AAPL', 10, 'init')
        self.account.sell_shares('AAPL', 5, 'take profit')
        value = self.account.calculate_portfolio_value()
        pnl = self.account.calculate_profit_loss(value)
        # Expected small negative due to spread costs
        self.assertAlmostEqual(pnl, -3.0, places=1)

    def test_max_order_size_enforced(self):
        with patch('accounts.MAX_ORDER_SIZE', 5):
            with self.assertRaises(ValueError):
                self.account.buy_shares('AAPL', 6, 'too big')

    def test_daily_trade_limit_enforced(self):
        with patch('accounts.DAILY_TRADE_LIMIT', 2), patch('accounts.MAX_ORDER_SIZE', 100):
            self.account.buy_shares('AAPL', 1, 't1')
            self.account.sell_shares('AAPL', 1, 't2')
            with self.assertRaises(ValueError):
                self.account.buy_shares('AAPL', 1, 't3')


if __name__ == '__main__':
    unittest.main()
