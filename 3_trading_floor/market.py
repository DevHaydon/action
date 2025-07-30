from polygon import RESTClient
from dotenv import load_dotenv
import os
from datetime import datetime
from typing import Dict
from database import write_market, read_market
from functools import lru_cache
from logger import log_exception
import time

load_dotenv(override=True)

polygon_api_key = os.getenv("POLYGON_API_KEY")
polygon_plan = os.getenv("POLYGON_PLAN")

is_paid_polygon = polygon_plan == "paid"
is_realtime_polygon = polygon_plan == "realtime"

# simple in-memory cache of latest successful prices
price_cache: Dict[str, float] = {}

def is_market_open() -> bool:
    client = RESTClient(polygon_api_key)
    market_status = client.get_market_status()
    return market_status.market == "open"

def get_all_share_prices_polygon_eod() -> dict[str, float]:
    client = RESTClient(polygon_api_key)

    probe = client.get_previous_close_agg("SPY")[0]
    last_close = datetime.fromtimestamp(probe.timestamp/1000).date()

    results = client.get_grouped_daily_aggs(last_close, adjusted=True, include_otc=False)
    return {result.ticker: result.close for result in results}

@lru_cache(maxsize=2)
def get_market_for_prior_date(today):
    market_data = read_market(today)
    if not market_data:
        market_data = get_all_share_prices_polygon_eod()
        write_market(today, market_data)
    return market_data

def get_share_price_polygon_eod(symbol) -> float:
    today = datetime.now().date().strftime("%Y-%m-%d")
    market_data = get_market_for_prior_date(today)
    return market_data.get(symbol, 0.0)

def get_share_price_polygon_min(symbol) -> float:
    client = RESTClient(polygon_api_key)
    result = client.get_snapshot_ticker("stocks", symbol)
    return result.min.close

def get_share_price_polygon(symbol) -> float:
    if is_paid_polygon:
        return get_share_price_polygon_min(symbol)
    else:
        return get_share_price_polygon_eod(symbol)


def _get_cached_price(symbol: str) -> float:
    """Return the last known price for ``symbol`` from memory or today's DB."""
    if symbol in price_cache:
        return price_cache[symbol]

    today = datetime.now().date().strftime("%Y-%m-%d")
    market_data = read_market(today)
    if market_data:
        price = market_data.get(symbol)
        if price is not None:
            price_cache[symbol] = price
            return price
    return 0.0

def get_share_price(symbol, retries: int = 2) -> float:
    """Return the latest share price for ``symbol``.

    Attempts to fetch the price from Polygon up to ``retries`` + 1 times
    before falling back to cached data. Any exceptions raised by the API are
    logged for monitoring.
    """
    if polygon_api_key:
        for attempt in range(retries + 1):
            try:
                price = get_share_price_polygon(symbol)
                price_cache[symbol] = price
                return price
            except Exception as e:
                log_exception("market", e, "Polygon API error")
                if attempt < retries:
                    time.sleep(0.1)
    return _get_cached_price(symbol)
