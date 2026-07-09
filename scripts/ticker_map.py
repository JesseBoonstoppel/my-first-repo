"""Maps the portfolio sheet's "TICKER:MARKET" identifiers to Yahoo Finance symbols.

Yahoo's ticker for a given company often does not match the plain ticker used
on the exchange itself (e.g. ING Groep trades as "INGA" on Euronext Amsterdam,
not "ING"), so exact overrides take priority over the generic suffix rules.
"""

# Exact overrides for tickers where the exchange code doesn't map cleanly.
OVERRIDES = {
    "AD:AMS": "AD.AS",
    "ING:AMS": "INGA.AS",
    "SHELL:AMS": "SHELL.AS",
    "HTGC:NYSE": "HTGC",
}

# Fallback: exchange code -> Yahoo suffix, for tickers with no override.
MARKET_SUFFIX = {
    "AMS": ".AS",
    "NYSE": "",
    "NASDAQ": "",
    "LON": ".L",
    "PAR": ".PA",
    "FRA": ".DE",
    "AEX": ".AS",
}


def to_yahoo_symbol(ticker: str, market: str | None = None) -> str:
    """Convert a "TICKER:MARKET" sheet identifier to a Yahoo Finance symbol."""
    if ticker in OVERRIDES:
        return OVERRIDES[ticker]

    base, _, exchange = ticker.partition(":")
    exchange = exchange or market or ""
    suffix = MARKET_SUFFIX.get(exchange.upper(), "")
    return f"{base}{suffix}"
