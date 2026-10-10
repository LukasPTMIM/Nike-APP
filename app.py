"""NIKE / any-ticker interactive valuation app — Team 5, Phase III (B), CBS International Business School.
ONE-FILE VERSION: everything (data loading, valuation maths, offline sample data, user interface) is in this file.
How to run (Mac Terminal or Windows Command Prompt):
    pip install streamlit yfinance pandas numpy plotly openpyxl
    streamlit run app.py
Features (as required by the brief):
  * Yahoo Finance data via yfinance for any ticker (NKE, AAPL, MSFT, ADS.DE, ...)
  * Residual income valuation: Revenue x growth -> Net income (x margin) -> Dividends (x payout)
    -> Book value roll-forward -> RI = NI - ke x beginning BV -> value per share
  * Editable assumptions: revenue growth, net margin, payout ratio, cost of equity (CAPM), terminal growth,
    forecast period (3-10 years), optional year-by-year path
  * Trading comparables: EV/Revenue, EV/EBITDA, P/E (TTM & forward), P/Book with implied share prices
  * Football field, market price comparison, forecast tables, charts, sensitivity heatmaps, CSV/Excel download
  * "Part A Excel model (NKE)" preset uses the supplied Team 5 assumptions (workbook not independently verified)
  * Offline fallback sample (Oct 9, 2026) if Yahoo Finance cannot be reached
Not investment advice.
"""
from __future__ import annotations
import io
import json
import math
from dataclasses import dataclass
import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
# =============================================================================
# OFFLINE SAMPLE DATA (used only if Yahoo Finance is unreachable)
# =============================================================================
SAMPLE = json.loads(r"""
{
"as_of": "Unverified sample supplied with original code; claimed date: Oct 9, 2026",
"risk_free": 0.0524,
"companies": {
  "NKE": {
   "ticker": "NKE",
   "name": "NIKE, Inc.",
   "currency": "USD",
   "price_currency": "USD",
   "price": 34.71,
   "shares": 1485308829,
   "market_cap": 51560000000,
   "ev": 54310000000,
   "revenue_ttm": 45890000000,
   "ebitda": 4890000000,
   "eps_ttm": 2.09,
   "eps_fwd": 1.35,
   "bvps": 10.27,
   "beta": 1.1,
   "ev_revenue": 1.18,
   "ev_ebitda": 11.11,
   "pe_ttm": 16.64,
   "pe_fwd": 25.71,
   "pb": 3.38,
   "low52": null,
   "high52": null,
   "target_low": 23.0,
   "target_mean": 47.56,
   "target_high": 75.0,
   "sector": "Consumer Cyclical",
   "industry": "Footwear & Accessories",
   "summary": "NIKE, Inc. designs, develops, markets and sells athletic footwear, apparel, equipment and accessories under the NIKE, Jordan and Converse brands, through wholesale partners and NIKE Direct (owned stores and digital).",
   "hist": {
    "Revenue": {
     "2024": 51362,
     "2025": 46309,
     "2026": 46398
    },
    "Net income": {
     "2024": 5700,
     "2025": 3219,
     "2026": 3108
    },
    "Book value of equity": {
     "2024": 14430,
     "2025": 13213,
     "2026": 14865
    },
    "Dividends paid": {
     "2024": 2169,
     "2025": 2300,
     "2026": 2407
    }
   }
  },
  "ADS.DE": {
   "ticker": "ADS.DE",
   "name": "adidas AG",
   "currency": "EUR",
   "price_currency": "EUR",
   "price": 146.75,
   "market_cap": 25570000000,
   "ev": 30390000000,
   "revenue_ttm": 26040000000,
   "ebitda": 2660000000,
   "eps_ttm": 7.77,
   "ev_revenue": 1.17,
   "ev_ebitda": 9.2,
   "pe_ttm": 18.89,
   "pe_fwd": 14.17,
   "pb": 4.03,
   "beta": 1.17
  },
  "DECK": {
   "ticker": "DECK",
   "name": "Deckers Outdoor Corp.",
   "currency": "USD",
   "price_currency": "USD",
   "price": 82.89,
   "market_cap": 11290000000,
   "ev": 10160000000,
   "revenue_ttm": 5530000000,
   "ebitda": 1330000000,
   "eps_ttm": 7.09,
   "ev_revenue": 1.84,
   "ev_ebitda": 7.65,
   "pe_ttm": 11.69,
   "pe_fwd": 10.88,
   "pb": 4.92,
   "beta": 1.12
  },
  "ONON": {
   "ticker": "ONON",
   "name": "On Holding AG",
   "currency": "USD",
   "price_currency": "USD",
   "price": 35.09,
   "market_cap": 11730000000,
   "ev": 10890000000,
   "revenue_ttm": 3990000000,
   "ebitda": 606930000,
   "eps_ttm": 1.47,
   "ev_revenue": 2.73,
   "ev_ebitda": 17.94,
   "pe_ttm": 23.95,
   "pe_fwd": 19.03,
   "pb": 4.96,
   "beta": 2.08
  },
  "LULU": {
   "ticker": "LULU",
   "name": "lululemon athletica inc.",
   "currency": "USD",
   "price_currency": "USD",
   "price": 94.5,
   "market_cap": 10460000000,
   "ev": 11210000000,
   "revenue_ttm": 11090000000,
   "ebitda": 2380000000,
   "eps_ttm": 12.22,
   "ev_revenue": 1.01,
   "ev_ebitda": 4.7,
   "pe_ttm": 7.74,
   "pe_fwd": 11.34,
   "pb": 2.2,
   "beta": 0.92
  },
  "AS": {
   "ticker": "AS",
   "name": "Amer Sports, Inc.",
   "currency": "USD",
   "price_currency": "USD",
   "price": 27.06,
   "market_cap": 15740000000,
   "ev": 16130000000,
   "revenue_ttm": 7440000000,
   "ebitda": 1210000000,
   "eps_ttm": 0.96,
   "ev_revenue": 2.17,
   "ev_ebitda": 13.29,
   "pe_ttm": 28.3,
   "pe_fwd": 18.77,
   "pb": 2.3,
   "beta": 2.03
  },
  "7936.T": {
   "ticker": "7936.T",
   "name": "ASICS Corporation",
   "currency": "JPY",
   "price_currency": "JPY",
   "price": 4141,
   "market_cap": 2830000000000,
   "ev": 2790000000000,
   "revenue_ttm": 942600000000,
   "ebitda": 212450000000,
   "eps_ttm": 178.63,
   "ev_revenue": 2.96,
   "ev_ebitda": 13.15,
   "pe_ttm": 23.18,
   "pe_fwd": 21.17,
   "pb": 7.86,
   "beta": 1.0
  },
  "UAA": {
   "ticker": "UAA",
   "name": "Under Armour, Inc.",
   "currency": "USD",
   "price_currency": "USD",
   "price": 4.93,
   "market_cap": 2120000000,
   "ev": 3100000000,
   "revenue_ttm": 4930000000,
   "ebitda": 173040000,
   "eps_ttm": -1.15,
   "ev_revenue": 0.62,
   "ev_ebitda": 17.8,
   "pe_ttm": null,
   "pe_fwd": 62.32,
   "pb": 1.47,
   "beta": 1.62
  },
  "PUM.DE": {
   "ticker": "PUM.DE",
   "name": "PUMA SE",
   "currency": "EUR",
   "price_currency": "EUR",
   "price": 22.01,
   "market_cap": 3240000000,
   "ev": 5520000000,
   "revenue_ttm": 6990000000,
   "ebitda": -17000000,
   "eps_ttm": -3.02,
   "ev_revenue": 0.79,
   "ev_ebitda": null,
   "pe_ttm": null,
   "pe_fwd": null,
   "pb": 1.78,
   "beta": 1.03
  }
}
}
""")
# =============================================================================
# VALUATION MATHS (residual income, comps, sensitivity)
# =============================================================================
# ---------------------------------------------------------------------------
# Residual income
# ---------------------------------------------------------------------------
@dataclass
class RIResult:
    table: pd.DataFrame          # one row per forecast year
    book_value_0: float
    pv_ri: float
    terminal_value: float
    pv_terminal: float
    equity_value: float
    value_per_share: float
def _as_list(x, n):
    if isinstance(x, (list, tuple, np.ndarray, pd.Series)):
        vals = list(x)
        if len(vals) < n:                       # extend with the last value
            vals = vals + [vals[-1]] * (n - len(vals))
        return [float(v) for v in vals[:n]]
    return [float(x)] * n
def residual_income(revenue_0: float, book_value_0: float, shares: float,
                    growth, net_margin, payout: float, cost_of_equity: float,
                    terminal_growth: float, years: int, first_year: int | None = None) -> RIResult:
    """Simplified residual income valuation (as specified in the Phase III brief).
    Revenue_t   = Revenue_(t-1) x (1 + growth_t)
    NetIncome_t = Revenue_t x net margin_t
    Dividends_t = NetIncome_t x payout
    BV_t        = BV_(t-1) + NetIncome_t - Dividends_t
    RI_t        = NetIncome_t - ke x BV_(t-1)
    Value       = BV_0 + sum PV(RI_t) + PV(terminal value)
    TV          = RI_N x (1 + g) / (ke - g), discounted with the year-N factor
    """
    if cost_of_equity <= terminal_growth:
        raise ValueError("Cost of equity must be higher than terminal growth.")
    numeric = [revenue_0, book_value_0, shares, payout, cost_of_equity, terminal_growth]
    if not all(np.isfinite(v) for v in numeric):
        raise ValueError("Please enter finite numbers for all assumptions.")
    if years < 1 or revenue_0 < 0:
        raise ValueError("Forecast years must be positive and revenue cannot be negative.")
    if shares <= 0:
        raise ValueError("Share count must be positive.")
    g_path = _as_list(growth, years)
    m_path = _as_list(net_margin, years)
    if not all(np.isfinite(v) for v in g_path + m_path):
        raise ValueError("Fill in every growth and margin cell with a number.")
    rows = []
    rev, bv = revenue_0, book_value_0
    for t in range(1, years + 1):
        rev = rev * (1 + g_path[t - 1])
        ni = rev * m_path[t - 1]
        div = ni * payout
        bv_beg = bv
        charge = cost_of_equity * bv_beg
        ri = ni - charge
        bv = bv_beg + ni - div
        dfac = 1 / (1 + cost_of_equity) ** t
        rows.append({
            "Year": (first_year + t - 1) if first_year else f"Year {t}",
            "Revenue growth": g_path[t - 1], "Revenue": rev, "Net margin": m_path[t - 1],
            "Net income": ni, "Dividends": div, "Beginning BV": bv_beg, "Ending BV": bv,
            "ROE": ni / bv_beg if bv_beg else np.nan, "Equity charge": charge,
            "Residual income": ri, "Discount factor": dfac, "PV of RI": ri * dfac,
        })
    table = pd.DataFrame(rows)
    ri_n = table["Residual income"].iloc[-1]
    tv = ri_n * (1 + terminal_growth) / (cost_of_equity - terminal_growth)
    pv_tv = tv * table["Discount factor"].iloc[-1]
    pv_ri = table["PV of RI"].sum()
    equity = book_value_0 + pv_ri + pv_tv
    return RIResult(table, book_value_0, pv_ri, tv, pv_tv, equity, equity / shares)
def sensitivity_grid(base_kwargs: dict, ke_values, g_values) -> pd.DataFrame:
    """Value per share for each cost of equity (rows) x terminal growth (columns)."""
    out = pd.DataFrame(index=[f"{k:.2%}" for k in ke_values], columns=[f"{g:.2%}" for g in g_values], dtype=float)
    for k in ke_values:
        for g in g_values:
            kw = dict(base_kwargs, cost_of_equity=k, terminal_growth=g)
            try:
                out.loc[f"{k:.2%}", f"{g:.2%}"] = residual_income(**kw).value_per_share
            except ValueError:
                out.loc[f"{k:.2%}", f"{g:.2%}"] = np.nan
    return out
def capm(risk_free: float, beta: float, erp: float) -> float:
    return risk_free + beta * erp
# ---------------------------------------------------------------------------
# Trading comparables
# ---------------------------------------------------------------------------
MULTIPLES = ["EV/Revenue", "EV/EBITDA", "P/E (TTM)", "P/E (Fwd)", "P/B"]
def _ratio(num, den):
    try:
        if num is None or den is None or den == 0 or (isinstance(den, float) and math.isnan(den)):
            return np.nan
        return num / den
    except TypeError:
        return np.nan
def company_multiples(c: dict) -> dict:
    """Multiples for one company. Uses Yahoo's own ratios when present (provider-supplied ratios),
    otherwise computes them from the raw fields."""
    ev, mcap, price = c.get("ev"), c.get("market_cap"), c.get("price")
    out = {
        "EV/Revenue": c.get("ev_revenue") or _ratio(ev, c.get("revenue_ttm")),
        "EV/EBITDA": c.get("ev_ebitda") or _ratio(ev, c.get("ebitda")),
        "P/E (TTM)": c.get("pe_ttm") or _ratio(price, c.get("eps_ttm")),
        "P/E (Fwd)": c.get("pe_fwd") or _ratio(price, c.get("eps_fwd")),
        "P/B": c.get("pb") or _ratio(price, c.get("bvps")),
    }
    # Negative or absurd multiples are not meaningful for valuation
    limits = {"EV/Revenue": 30, "EV/EBITDA": 60, "P/E (TTM)": 80, "P/E (Fwd)": 80, "P/B": 40}
    for k, v in out.items():
        if v is None or not np.isfinite(v) or v <= 0 or v > limits[k]:
            out[k] = np.nan
    return out
def comps_table(peers: list[dict]) -> pd.DataFrame:
    rows = []
    for p in peers:
        m = company_multiples(p)
        rows.append({"Ticker": p["ticker"], "Company": p.get("name", p["ticker"]),
                     "Currency": p.get("price_currency", ""), "Source": p.get("source", ""), "Price": p.get("price"),
                     "Market cap (m)": (p.get("market_cap") or np.nan) / 1e6 if p.get("market_cap") else np.nan,
                     **m})
    return pd.DataFrame(rows)
def comps_stats(table: pd.DataFrame) -> pd.DataFrame:
    s = table[MULTIPLES]
    return pd.DataFrame({"25th pct": s.quantile(0.25), "Median": s.median(), "Mean": s.mean(),
                         "75th pct": s.quantile(0.75), "N": s.count()}).T
def implied_prices(target: dict, stats: pd.DataFrame) -> pd.DataFrame:
    """Apply peer multiples (25th / median / 75th) to the target's own metrics."""
    shares = target.get("shares")
    ev, mcap = target.get("ev"), target.get("market_cap")
    net_debt = (ev - mcap) if (ev is not None and mcap is not None) else 0.0
    metric = {"EV/Revenue": target.get("revenue_ttm"), "EV/EBITDA": target.get("ebitda"),
              "P/E (TTM)": target.get("eps_ttm"), "P/E (Fwd)": target.get("eps_fwd"), "P/B": target.get("bvps")}
    rows = []
    for mult in MULTIPLES:
        base = metric[mult]
        if base is None or not np.isfinite(base) or base <= 0 or stats.loc["N", mult] == 0:
            continue
        vals = {}
        for q in ["25th pct", "Median", "75th pct"]:
            m = stats.loc[q, mult]
            if mult.startswith("EV"):
                vals[q] = (m * base - net_debt) / shares if shares else np.nan
            else:
                vals[q] = m * base
        rows.append({"Method": mult, "Target metric": base, "Low": vals["25th pct"],
                     "Median": vals["Median"], "High": vals["75th pct"]})
    return pd.DataFrame(rows)
# =============================================================================
# DATA LAYER (Yahoo Finance via yfinance, with offline fallback)
# =============================================================================
def _clean(x):
    try:
        if x is None:
            return None
        x = float(x)
        return None if math.isnan(x) or math.isinf(x) else x
    except (TypeError, ValueError):
        return None
def _row(df: pd.DataFrame | None, names: list[str]) -> pd.Series | None:
    if df is None or df.empty:
        return None
    for n in names:
        if n in df.index:
            s = pd.to_numeric(df.loc[n], errors="coerce").dropna()
            if not s.empty:
                return s.sort_index()
    return None
def _from_yahoo(ticker: str) -> dict:
    import yfinance as yf  # imported here so the offline mode works without network
    t = yf.Ticker(ticker)
    info = t.info or {}
    price = _clean(info.get("currentPrice")) or _clean(info.get("regularMarketPrice"))
    if price is None:
        raise ValueError(f"No market data returned for '{ticker}'.")
    hist = pd.DataFrame()
    try:
        inc, bs, cf = t.income_stmt, t.balance_sheet, t.cashflow
        rev = _row(inc, ["Total Revenue", "Operating Revenue"])
        ni = _row(inc, ["Net Income", "Net Income Common Stockholders"])
        eq = _row(bs, ["Stockholders Equity", "Common Stock Equity", "Total Equity Gross Minority Interest"])
        div = _row(cf, ["Cash Dividends Paid", "Common Stock Dividend Paid"])
        cols = {"Revenue": rev, "Net income": ni, "Book value of equity": eq,
                "Dividends paid": div.abs() if div is not None else None}
        hist = pd.DataFrame({k: v for k, v in cols.items() if v is not None}) / 1e6
        hist.index = [d.year if hasattr(d, "year") else d for d in hist.index]
        hist = hist.dropna(how="all").sort_index()
    except Exception:  # statements are optional; the comps still work without them
        hist = pd.DataFrame()
    return {
        "ticker": ticker.upper(),
        "name": info.get("longName") or info.get("shortName") or ticker.upper(),
        "currency": info.get("financialCurrency") or info.get("currency") or "",
        "price_currency": info.get("currency") or "",
        "price": price,
        "shares": _clean(info.get("sharesOutstanding")) or _clean(info.get("impliedSharesOutstanding")),
        "market_cap": _clean(info.get("marketCap")),
        "ev": _clean(info.get("enterpriseValue")),
        "revenue_ttm": _clean(info.get("totalRevenue")),
        "ebitda": _clean(info.get("ebitda")),
        "eps_ttm": _clean(info.get("trailingEps")),
        "eps_fwd": _clean(info.get("forwardEps")),
        "bvps": _clean(info.get("bookValue")),
        "beta": _clean(info.get("beta")),
        "ev_revenue": _clean(info.get("enterpriseToRevenue")),
        "ev_ebitda": _clean(info.get("enterpriseToEbitda")),
        "pe_ttm": _clean(info.get("trailingPE")),
        "pe_fwd": _clean(info.get("forwardPE")),
        "pb": _clean(info.get("priceToBook")),
        "low52": _clean(info.get("fiftyTwoWeekLow")),
        "high52": _clean(info.get("fiftyTwoWeekHigh")),
        "target_low": _clean(info.get("targetLowPrice")),
        "target_mean": _clean(info.get("targetMeanPrice")),
        "target_high": _clean(info.get("targetHighPrice")),
        "sector": info.get("sector", ""),
        "industry": info.get("industry", ""),
        "summary": info.get("longBusinessSummary", ""),
        "hist": hist,
        "source": "Yahoo Finance (live)",
    }
def _from_sample(ticker: str) -> dict:
    d = dict(SAMPLE["companies"][ticker.upper()])
    hist = pd.DataFrame(d.pop("hist", {}))
    if not hist.empty:
        hist.index = [int(i) for i in hist.index]
    d["hist"] = hist.sort_index()
    if not d.get("shares") and d.get("market_cap") and d.get("price"):
        d["shares"] = d["market_cap"] / d["price"]          # derive share count for peers
    if not d.get("bvps") and d.get("pb") and d.get("price"):
        d["bvps"] = d["price"] / d["pb"]                    # derive book value per share
    d["source"] = f"Offline sample ({SAMPLE['as_of']})"
    return d
def get_company(ticker: str, allow_sample: bool = True) -> dict:
    """Live Yahoo data; falls back to the bundled sample if Yahoo is unreachable."""
    ticker = ticker.strip().upper()
    try:
        return _from_yahoo(ticker)
    except Exception as exc:
        if allow_sample and ticker in SAMPLE["companies"]:
            d = _from_sample(ticker)
            d["warning"] = f"Yahoo Finance unavailable ({type(exc).__name__}); showing offline sample data."
            return d
        raise
def get_risk_free() -> tuple[float, str]:
    """10-year US Treasury yield from Yahoo (^TNX), else the sample value."""
    try:
        import yfinance as yf
        px = yf.Ticker("^TNX").history(period="5d")["Close"].dropna()
        if not px.empty:
            return float(px.iloc[-1]) / 100, "Yahoo Finance ^TNX (10-year Treasury)"
    except Exception:
        pass
    return SAMPLE["risk_free"], f"Offline sample ({SAMPLE['as_of']})"
def sample_tickers() -> list[str]:
    return sorted(SAMPLE["companies"].keys())
# =============================================================================
# STREAMLIT USER INTERFACE
# =============================================================================
st.set_page_config(page_title="Equity Valuation App", page_icon=":material/query_stats:", layout="wide")
ACCENT = "#1F5AA6"
GREY = "#9AA0A6"
GREEN = "#2E7D32"
RED = "#C62828"
# Assumptions from the Team 5 Excel model (Phase III A, Base case) — used by the "Part A" preset.
PART_A = {
    "growth": [-0.0159, 0.047, 0.047, 0.0326, 0.0326],
    "margin": [0.073210, 0.086265, 0.087784, 0.090880, 0.093909],
    "payout": 0.6232, "rf": 0.0444, "beta": 1.11, "erp": 0.0418, "tg": 0.0215,
    "years": 5, "shares_m": 1479.8, "revenue_0": 46398.0, "bv_0": 14865.0, "first_year": 2027,
}
# ---------------------------------------------------------------------------
# Data loading (cached for one hour so the app stays fast)
# ---------------------------------------------------------------------------
@st.cache_data(ttl=3600, show_spinner=False)
def load_company(ticker: str, demo: bool = False) -> dict:
    return _from_sample(ticker) if demo else get_company(ticker)
@st.cache_data(ttl=3600, show_spinner=False)
def load_rf() -> tuple[float, str]:
    return get_risk_free()
def fmt_money(x, cur="", dec=2):
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return "n/a"
    sym = {"USD": "$", "EUR": "€", "JPY": "¥", "GBP": "£", "CHF": "CHF "}.get(cur, "")
    return f"{sym}{x:,.{dec}f}" if sym else f"{x:,.{dec}f} {cur}".strip()
def fmt_big(x, cur=""):
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return "n/a"
    for div, suf in [(1e12, "tn"), (1e9, "bn"), (1e6, "m")]:
        if abs(x) >= div:
            return fmt_money(x / div, cur, 1) + suf
    return fmt_money(x, cur, 0)
# ---------------------------------------------------------------------------
# Sidebar — company and assumptions
# ---------------------------------------------------------------------------
st.sidebar.title("Valuation inputs")
st.sidebar.subheader("1 · Company")
demo = st.sidebar.toggle("Demo mode (no internet required)", value=True)
if st.sidebar.button("Refresh data"):
    st.cache_data.clear()
    st.rerun()
ticker = st.sidebar.text_input("Stock ticker (Yahoo Finance)", "NKE",
                               help="Examples: NKE, AAPL, MSFT, ADS.DE (adidas, Frankfurt), 7936.T (ASICS, Tokyo)").strip().upper()
peer_str = st.sidebar.text_input("Comparable companies (comma-separated)", "ADS.DE, DECK, ONON, LULU, AS, 7936.T",
                                 help="Peers used for EV/Revenue, EV/EBITDA, P/E and P/Book.")
with st.spinner(f"Loading {ticker} from Yahoo Finance…"):
    try:
        co = load_company(ticker, demo)
    except Exception as exc:
        st.error(f"Could not load **{ticker}**: {exc}. Check the ticker on finance.yahoo.com and your internet connection.")
        st.info(f"Offline sample tickers available: {', '.join(sample_tickers())}")
        st.stop()
cur = co.get("currency", "")
hist = co.get("hist", pd.DataFrame())
shares_m_data = (co.get("shares") or 0) / 1e6
rf_live, rf_src = (SAMPLE["risk_free"], "Unverified demo assumption") if demo else load_rf()
# Defaults derived from the latest financials
rev0_data = float(hist["Revenue"].dropna().iloc[-1]) if "Revenue" in hist and hist["Revenue"].notna().any() else (co.get("revenue_ttm") or 0) / 1e6
if "Book value of equity" in hist and hist["Book value of equity"].notna().any():
    bv0_data = float(hist["Book value of equity"].dropna().iloc[-1])
else:
    bv0_data = (co.get("bvps") or 0) * shares_m_data
last_year = int(hist.index[-1]) if len(hist.index) else None
def _default_growth():
    if "Revenue" in hist and hist["Revenue"].notna().sum() >= 2:
        r = hist["Revenue"].dropna()
        n = min(3, len(r) - 1)
        cagr = (r.iloc[-1] / r.iloc[-1 - n]) ** (1 / n) - 1
        return float(np.clip(cagr, -0.05, 0.15))
    return 0.03
def _default_margin():
    if {"Revenue", "Net income"} <= set(hist.columns):
        m = (hist["Net income"] / hist["Revenue"]).dropna()
        if not m.empty and m.iloc[-1] > 0:
            return float(np.clip(m.iloc[-1], 0.0, 0.4))
    return 0.05
def _default_payout():
    if {"Dividends paid", "Net income"} <= set(hist.columns):
        p = (hist["Dividends paid"] / hist["Net income"]).dropna()
        if not p.empty and np.isfinite(p.iloc[-1]):
            return float(np.clip(p.iloc[-1], 0.0, 1.0))
    return 0.0
auto = {"growth": _default_growth(), "margin": _default_margin(), "payout": _default_payout(),
        "rf": rf_live, "beta": co.get("beta") or 1.0, "erp": 0.0418, "tg": 0.025, "years": 5}
st.sidebar.subheader("2 · Assumptions")
presets = ["From latest financials", "Bear case", "Bull case"]
if ticker == "NKE":
    presets.append("Part A Excel model (NKE)")
preset = st.sidebar.selectbox("Starting point", presets,
                              help="Sets the default values below. You can still change every input.")
d = dict(auto)
if preset == "Bear case":
    d.update(growth=auto["growth"] - 0.02, margin=max(auto["margin"] - 0.015, 0.0), erp=auto["erp"] + 0.005, tg=auto["tg"] - 0.005)
elif preset == "Bull case":
    d.update(growth=auto["growth"] + 0.02, margin=auto["margin"] + 0.015, erp=auto["erp"] - 0.005, tg=auto["tg"] + 0.005)
elif preset.startswith("Part A"):
    d.update(growth=PART_A["growth"][0], margin=PART_A["margin"][0], payout=PART_A["payout"], rf=PART_A["rf"],
             beta=PART_A["beta"], erp=PART_A["erp"], tg=PART_A["tg"], years=PART_A["years"])
k = f"{ticker}|{preset}|{demo}"   # widget keys: defaults reset when the ticker or preset changes
years = st.sidebar.slider("Forecast period (years)", 3, 10, int(d["years"]), key=f"yrs|{k}")
growth = st.sidebar.number_input("Revenue growth (% per year)", -30.0, 50.0, round(d["growth"] * 100, 2), 0.25, key=f"g|{k}") / 100
margin = st.sidebar.number_input("Net income margin (%)", -20.0, 60.0, round(d["margin"] * 100, 2), 0.25, key=f"m|{k}") / 100
payout = st.sidebar.number_input("Dividend payout ratio (%)", 0.0, 150.0, round(d["payout"] * 100, 2), 1.0, key=f"p|{k}") / 100
with st.sidebar.expander("Cost of equity (CAPM)", expanded=True):
    rf = st.number_input("Risk-free rate (%)", 0.0, 15.0, round(float(np.clip(d["rf"] * 100, 0, 15)), 2), 0.05, key=f"rf|{k}",
                         help=f"Default: {rf_src}") / 100
    beta = st.number_input("Beta", 0.0, 4.0, round(float(np.clip(d["beta"], 0, 4)), 2), 0.05, key=f"b|{k}")
    erp = st.number_input("Equity risk premium (%)", 0.0, 15.0, round(d["erp"] * 100, 2), 0.05, key=f"erp|{k}") / 100
    ke_capm = capm(rf, beta, erp)
    override = st.checkbox("Override cost of equity", key=f"ov|{k}")
    ke = st.number_input("Cost of equity (%)", 0.0, 30.0, round(float(np.clip(ke_capm * 100, 0, 30)), 3), 0.05, key=f"ke|{k}") / 100 if override else ke_capm
    st.caption(f"Cost of equity = {rf:.2%} + {beta:.2f} × {erp:.2%} = **{ke_capm:.2%}**")
tg = st.sidebar.number_input("Terminal growth (%)", -2.0, 6.0, round(d["tg"] * 100, 2), 0.05, key=f"tg|{k}") / 100
# Base-year data (overridable)
with st.sidebar.expander("Base-year data", expanded=preset.startswith("Part A")):
    pa = preset.startswith("Part A")
    rev0 = st.number_input(f"Revenue, base year ({cur} m)", value=float(PART_A["revenue_0"] if pa else round(rev0_data, 1)), key=f"r0|{k}")
    bv0 = st.number_input(f"Book value of equity, base year ({cur} m)", value=float(PART_A["bv_0"] if pa else round(bv0_data, 1)), key=f"bv0|{k}")
    shares_m = st.number_input("Shares outstanding (m)", value=float(PART_A["shares_m"] if pa else round(shares_m_data, 1)), key=f"sh|{k}")
# Optional year-by-year path
custom = st.sidebar.toggle("Edit growth and margin year by year", value=preset.startswith("Part A"), key=f"yy|{k}")
first_year = (last_year + 1) if last_year else None
if preset.startswith("Part A"):
    first_year = PART_A["first_year"]
if custom:
    if preset.startswith("Part A"):
        base_g = (PART_A["growth"] + [PART_A["growth"][-1]] * 10)[:years]
        base_m = (PART_A["margin"] + [PART_A["margin"][-1]] * 10)[:years]
    else:
        base_g, base_m = [growth] * years, [margin] * years
    path = pd.DataFrame({"Year": [f"{first_year + i}" if first_year else f"Year {i + 1}" for i in range(years)],
                         "Revenue growth %": np.round(np.array(base_g) * 100, 3),
                         "Net margin %": np.array(base_m) * 100})
    st.sidebar.caption("Edit the cells below:")
    path = st.sidebar.data_editor(path, hide_index=True, disabled=["Year"], key=f"path|{k}|{years}")
    g_path = (path["Revenue growth %"].astype(float) / 100).tolist()
    m_path = (path["Net margin %"].astype(float) / 100).tolist()
else:
    g_path, m_path = growth, margin
# ---------------------------------------------------------------------------
# Valuation
# ---------------------------------------------------------------------------
ri_kwargs = dict(revenue_0=rev0, book_value_0=bv0, shares=shares_m, growth=g_path, net_margin=m_path,
                 payout=payout, cost_of_equity=ke, terminal_growth=tg, years=years, first_year=first_year)
try:
    ri = residual_income(**ri_kwargs)
except ValueError as exc:
    st.error(str(exc))
    st.stop()
price = co.get("price")
same_currency = bool(cur and cur == co.get("price_currency"))
upside = ri.value_per_share / price - 1 if price and same_currency else np.nan
# Peers
peer_list = list(dict.fromkeys(p.strip().upper() for p in peer_str.split(",") if p.strip() and p.strip().upper() != ticker))
peers, failed = [], []
with st.spinner("Loading comparable companies…"):
    for p in peer_list:
        try:
            peers.append(load_company(p, demo))
        except Exception:
            failed.append(p)
ctab = comps_table(peers) if peers else pd.DataFrame()
cstats = comps_stats(ctab) if not ctab.empty else pd.DataFrame()
implied = implied_prices(co, cstats) if not cstats.empty and same_currency else pd.DataFrame()
comps_mid = float(implied["Median"].median()) if not implied.empty else np.nan
# ---------------------------------------------------------------------------
# Header
# ---------------------------------------------------------------------------
st.title(f"{co['name']} ({ticker})")
src_line = f"Data: {co['source']} · Reporting currency: {cur or 'n/a'} · Risk-free default: {rf_src}"
st.caption(src_line)
if "Offline sample" in co["source"]:
    st.warning("DEMO DATA: supplied with the original code and not independently verified. Do not treat these values as current market data.")
if any("Offline sample" in p.get("source", "") for p in peers):
    st.info("Some or all peer figures use unverified demo data. See the Source column.")
if co.get("warning"):
    st.warning(co["warning"])
if co.get("price_currency") and cur and co["price_currency"] != cur:
    st.warning(f"Share price is quoted in {co['price_currency']} but statements are in {cur}. "
               "Per-share values from the RI model are in the reporting currency. Market comparisons and implied comps are disabled until currencies are aligned.")
if failed:
    st.info(f"Could not load these peers: {', '.join(failed)}")
c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Market price", fmt_money(price, co.get("price_currency", cur)), border=True)
c2.metric("Residual income value", fmt_money(ri.value_per_share, cur), f"{upside:+.1%} vs market" if price and same_currency else None, border=True)
c3.metric("Trading comps (median)", fmt_money(comps_mid, cur),
          f"{comps_mid / price - 1:+.1%} vs market" if price and same_currency and np.isfinite(comps_mid) else None, border=True)
c4.metric("Cost of equity", f"{ke:.2%}", f"terminal g {tg:.2%}", delta_color="off", border=True)
c5.metric("Market cap", fmt_big(co.get("market_cap"), co.get("price_currency", cur)), border=True)
tabs = st.tabs(["Residual income", "Trading comps", "Football field", "Sensitivity", "Company data", "Method"])
# ---------------------------------------------------------------------------
# Tab 1 — Residual income
# ---------------------------------------------------------------------------
with tabs[0]:
    left, right = st.columns([3, 2])
    with left:
        st.subheader("Forecast")
        t = ri.table.set_index("Year")
        show = pd.DataFrame(index=t.index)
        pct_cols = ["Revenue growth", "Net margin", "ROE"]
        for col in ["Revenue growth", "Revenue", "Net margin", "Net income", "Dividends", "Beginning BV", "Ending BV",
                    "ROE", "Equity charge", "Residual income", "Discount factor", "PV of RI"]:
            if col in pct_cols:
                show[col] = t[col].map(lambda v: f"{v:.2%}")
            elif col == "Discount factor":
                show[col] = t[col].map(lambda v: f"{v:.4f}")
            else:
                show[col] = t[col].map(lambda v: f"{v:,.0f}")
        st.dataframe(show.T, height=460)
        st.caption(f"{cur} millions except ratios. RI = Net income − Cost of equity × Beginning book value.")
    with right:
        st.subheader("Value build-up")
        wf = go.Figure(go.Waterfall(
            orientation="v", measure=["absolute", "relative", "relative", "total"],
            x=["Book value today", "PV of RI (forecast)", "PV of terminal value", "Equity value"],
            y=[ri.book_value_0, ri.pv_ri, ri.pv_terminal, 0],
            text=[f"{v:,.0f}" for v in [ri.book_value_0, ri.pv_ri, ri.pv_terminal, ri.equity_value]],
            textposition="outside", connector={"line": {"color": GREY, "dash": "dot"}},
            increasing={"marker": {"color": ACCENT}}, totals={"marker": {"color": "#37474F"}}))
        wf.update_layout(height=330, margin=dict(l=10, r=10, t=10, b=10), yaxis_title=f"{cur} m", showlegend=False)
        st.plotly_chart(wf)
        st.markdown(
            f"**{fmt_big(ri.equity_value * 1e6, cur)}** equity value ÷ **{shares_m:,.1f}m** shares = "
            f"**{fmt_money(ri.value_per_share, cur)} per share**  \n"
            + (f"Terminal value is **{ri.pv_terminal / ri.equity_value:.0%}** of the total." if ri.equity_value else "Terminal share: n/a (zero equity value)."))
    st.subheader("Forecast charts")
    a, b = st.columns(2)
    with a:
        fig = go.Figure()
        fig.add_bar(x=t.index.astype(str), y=t["Revenue"], name="Revenue", marker_color="#C5CAE9")
        fig.add_bar(x=t.index.astype(str), y=t["Net income"], name="Net income", marker_color=ACCENT)
        fig.update_layout(barmode="group", height=320, title="Revenue and net income", yaxis_title=f"{cur} m",
                          margin=dict(l=10, r=10, t=40, b=10), legend=dict(orientation="h", y=-0.2))
        st.plotly_chart(fig)
    with b:
        fig = go.Figure()
        fig.add_bar(x=t.index.astype(str), y=t["Residual income"], name="Residual income", marker_color=ACCENT)
        fig.add_scatter(x=t.index.astype(str), y=t["PV of RI"], name="PV of RI", mode="lines+markers", line=dict(color=GREEN))
        fig.add_scatter(x=t.index.astype(str), y=t["ROE"] * 100, name="ROE (%)", mode="lines+markers",
                        line=dict(color=GREY, dash="dot"), yaxis="y2")
        fig.update_layout(height=320, title="Residual income and ROE", yaxis_title=f"{cur} m",
                          yaxis2=dict(overlaying="y", side="right", title="ROE %", showgrid=False),
                          margin=dict(l=10, r=10, t=40, b=10), legend=dict(orientation="h", y=-0.2))
        st.plotly_chart(fig)
    # Downloads
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        pd.DataFrame({"Assumption": ["Ticker", "Revenue base (m)", "Book value base (m)", "Shares (m)", "Payout",
                                     "Cost of equity", "Terminal growth", "Years", "Value per share", "Market price"],
                      "Value": [ticker, rev0, bv0, shares_m, payout, ke, tg, years, ri.value_per_share, price]}
                     ).to_excel(xw, sheet_name="Assumptions", index=False)
        ri.table.to_excel(xw, sheet_name="RI forecast", index=False)
        if not ctab.empty:
            ctab.to_excel(xw, sheet_name="Comps", index=False)
            implied.to_excel(xw, sheet_name="Implied prices", index=False)
    d1, d2 = st.columns(2)
    d1.download_button("Download forecast (CSV)", ri.table.to_csv(index=False), f"{ticker}_ri_forecast.csv", "text/csv")
    d2.download_button("Download full valuation (Excel)", buf.getvalue(), f"{ticker}_valuation.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
# ---------------------------------------------------------------------------
# Tab 2 — Trading comps
# ---------------------------------------------------------------------------
with tabs[1]:
    if ctab.empty:
        st.info("Add comparable tickers in the sidebar to see trading multiples.")
    else:
        st.subheader("Peer multiples")
        own = company_multiples(co)
        show = ctab.copy()
        show.loc[len(show)] = {"Ticker": ticker, "Company": f"{co['name']} (target)", "Currency": co.get("price_currency", ""),
                               "Price": price, "Market cap (m)": (co.get("market_cap") or np.nan) / 1e6, **own}
        cfg = {m: st.column_config.NumberColumn(m, format="%.1fx") for m in MULTIPLES}
        cfg.update({"Price": st.column_config.NumberColumn(format="%.2f"),
                    "Market cap (m)": st.column_config.NumberColumn(format="%,.0f")})
        st.dataframe(show, hide_index=True, column_config=cfg)
        st.caption("Multiples from Yahoo Finance where available, otherwise computed; verify currency consistency for foreign listings. "
                   "Negative or extreme multiples are excluded as not meaningful.")
        s1, s2 = st.columns([3, 2])
        with s1:
            st.markdown("**Peer statistics**")
            st.dataframe(cstats.style.format("{:.2f}"))
        with s2:
            fig = go.Figure()
            fig.add_bar(x=MULTIPLES, y=[cstats.loc["Median", m] for m in MULTIPLES], name="Peer median", marker_color=GREY)
            fig.add_bar(x=MULTIPLES, y=[own[m] for m in MULTIPLES], name=ticker, marker_color=ACCENT)
            fig.update_layout(barmode="group", height=300, margin=dict(l=10, r=10, t=30, b=10),
                              title=f"{ticker} vs peer median", legend=dict(orientation="h", y=-0.25))
            st.plotly_chart(fig)
        st.subheader("Implied share price from peer multiples")
        if implied.empty:
            st.info("Not enough positive metrics to apply the multiples.")
        else:
            st.dataframe(implied, hide_index=True, column_config={
                "Target metric": st.column_config.NumberColumn(format="%,.2f"),
                "Low": st.column_config.NumberColumn("Low (25th pct)", format="%.2f"),
                "Median": st.column_config.NumberColumn(format="%.2f"),
                "High": st.column_config.NumberColumn("High (75th pct)", format="%.2f")})
            st.caption("EV multiples: (multiple × metric − net debt) ÷ shares, with net debt = Yahoo EV − market cap. "
                       "P/E and P/B: multiple × EPS or book value per share.")
# ---------------------------------------------------------------------------
# Tab 3 — Football field
# ---------------------------------------------------------------------------
with tabs[2]:
    st.subheader("Valuation range vs market price")
    bars = []
    try:
        lo = residual_income(**dict(ri_kwargs, cost_of_equity=ke + 0.01, terminal_growth=tg - 0.005)).value_per_share
        hi_ke = max(ke - 0.01, tg + 0.01)
        hi = residual_income(**dict(ri_kwargs, cost_of_equity=hi_ke, terminal_growth=tg + 0.005)).value_per_share
        bars.append(("Residual income (ke ±1pp, g ±0.5pp)", lo, hi, ri.value_per_share))
    except ValueError:
        pass
    for _, r in implied.iterrows():
        bars.append((f"Comps: {r['Method']} (peer 25th–75th)", r["Low"], r["High"], r["Median"]))
    if same_currency and co.get("low52") and co.get("high52"):
        bars.append(("52-week trading range", co["low52"], co["high52"], None))
    if same_currency and co.get("target_low") and co.get("target_high"):
        bars.append(("Analyst price targets", co["target_low"], co["target_high"], co.get("target_mean")))
    if bars:
        ff = go.Figure()
        names = [b[0] for b in bars][::-1]
        for nm, lo_, hi_, mid in bars[::-1]:
            lo_, hi_ = min(lo_, hi_), max(lo_, hi_)
            color = ACCENT if nm.startswith("Residual") else ("#78909C" if nm.startswith("Comps") else "#B0BEC5")
            ff.add_bar(y=[nm], x=[hi_ - lo_], base=[lo_], orientation="h", marker_color=color, showlegend=False,
                       text=[f"{lo_:,.2f} – {hi_:,.2f}"], textposition="inside", insidetextanchor="middle",
                       hovertemplate=f"{nm}<br>Low %{{base:,.2f}}<br>High {hi_:,.2f}<extra></extra>")
            if mid is not None and np.isfinite(mid):
                ff.add_scatter(y=[nm], x=[mid], mode="markers", marker=dict(symbol="line-ns-open", size=22, color="black", line=dict(width=3)),
                               showlegend=False, hovertemplate=f"Mid-point {mid:,.2f}<extra></extra>")
        if price and same_currency:
            ff.add_vline(x=price, line_dash="dash", line_color=RED,
                         annotation_text=f"Market {fmt_money(price, co.get('price_currency', cur))}", annotation_position="top")
        ff.update_layout(height=120 + 60 * len(bars), margin=dict(l=10, r=30, t=40, b=10), xaxis_title=f"Value per share ({cur})",
                         yaxis=dict(categoryorder="array", categoryarray=names))
        st.plotly_chart(ff)
        st.caption("Bars show each method's range; the black tick is the mid-point (base case or peer median). "
                   "Red dashed line = current market price.")
# ---------------------------------------------------------------------------
# Tab 4 — Sensitivity
# ---------------------------------------------------------------------------
with tabs[3]:
    a, b = st.columns(2)
    with a:
        st.subheader("Cost of equity × terminal growth")
        kes = [ke + x for x in (-0.02, -0.01, -0.005, 0, 0.005, 0.01, 0.02)]
        gs = [tg + x for x in (-0.01, -0.005, 0, 0.005, 0.01)]
        grid = sensitivity_grid(ri_kwargs, kes, gs)
        hm = go.Figure(go.Heatmap(z=grid.values, x=grid.columns, y=grid.index, colorscale="RdBu", zmid=price if same_currency else None,
                                  text=np.vectorize(lambda v: "" if np.isnan(v) else f"{v:,.2f}")(grid.values.astype(float)),
                                  texttemplate="%{text}", colorbar=dict(title="Value")))
        hm.update_layout(height=380, xaxis_title="Terminal growth", yaxis_title="Cost of equity",
                         margin=dict(l=10, r=10, t=10, b=10))
        st.plotly_chart(hm)
    with b:
        st.subheader("Revenue growth × net margin")
        gsteps = [-0.02, -0.01, 0, 0.01, 0.02]
        msteps = [-0.02, -0.01, 0, 0.01, 0.02]
        base_g = np.array(g_path if isinstance(g_path, list) else [g_path] * years)
        base_m = np.array(m_path if isinstance(m_path, list) else [m_path] * years)
        z = [[residual_income(**dict(ri_kwargs, growth=list(base_g + dg), net_margin=list(base_m + dm))).value_per_share
              for dg in gsteps] for dm in msteps]
        hm2 = go.Figure(go.Heatmap(z=z, x=[f"{x:+.0%}" for x in gsteps], y=[f"{x:+.0%}" for x in msteps],
                                   colorscale="RdBu", zmid=price if same_currency else None, texttemplate="%{z:,.2f}", colorbar=dict(title="Value")))
        hm2.update_layout(height=380, xaxis_title="Change in revenue growth (pp)", yaxis_title="Change in net margin (pp)",
                          margin=dict(l=10, r=10, t=10, b=10))
        st.plotly_chart(hm2)
    st.caption("Blue cells are above the current market price, red cells below (colour centred on the market price).")
# ---------------------------------------------------------------------------
# Tab 5 — Company data
# ---------------------------------------------------------------------------
with tabs[4]:
    st.subheader("Historical financials")
    if hist.empty:
        st.info("Yahoo returned no annual statements for this ticker.")
    else:
        h = hist.copy()
        if {"Net income", "Revenue"} <= set(h.columns):
            h["Net margin"] = h["Net income"] / h["Revenue"]
            h["Revenue growth"] = h["Revenue"].pct_change()
        if {"Net income", "Book value of equity"} <= set(h.columns):
            h["ROE (on ending BV)"] = h["Net income"] / h["Book value of equity"]
        if {"Dividends paid", "Net income"} <= set(h.columns):
            h["Payout ratio"] = h["Dividends paid"] / h["Net income"]
        fmt = {c: "{:,.0f}" for c in ["Revenue", "Net income", "Book value of equity", "Dividends paid"] if c in h}
        fmt.update({c: "{:.1%}" for c in ["Net margin", "Revenue growth", "ROE (on ending BV)", "Payout ratio"] if c in h})
        table = pd.DataFrame({c: h[c].map(lambda v, f=fmt[c]: "" if pd.isna(v) else f.format(v))
                              for c in h.columns if c in fmt})
        st.dataframe(table.T)
        st.caption(f"Fiscal years, {cur} millions. Source: {co['source']}.")
        if "Revenue" in h:
            fig = go.Figure()
            fig.add_bar(x=h.index.astype(str), y=h["Revenue"], name="Revenue", marker_color="#C5CAE9")
            if "Net income" in h:
                fig.add_bar(x=h.index.astype(str), y=h["Net income"], name="Net income", marker_color=ACCENT)
            fig.update_layout(barmode="group", height=300, margin=dict(l=10, r=10, t=30, b=10), title="History",
                              legend=dict(orientation="h", y=-0.2))
            st.plotly_chart(fig)
    if co.get("summary"):
        with st.expander("Business description"):
            st.write(co["summary"])
# ---------------------------------------------------------------------------
# Tab 6 — Method
# ---------------------------------------------------------------------------
with tabs[5]:
    st.subheader("How the residual income value is calculated")
    st.latex(r"\text{Revenue}_t = \text{Revenue}_{t-1}(1+g_t)\qquad \text{NI}_t = \text{Revenue}_t \times m_t\qquad \text{Div}_t = \text{NI}_t \times \text{payout}")
    st.latex(r"BV_t = BV_{t-1} + \text{NI}_t - \text{Div}_t \qquad RI_t = \text{NI}_t - k_e \times BV_{t-1}")
    st.latex(r"V_0 = BV_0 + \sum_{t=1}^{N}\frac{RI_t}{(1+k_e)^t} + \frac{RI_N(1+g)}{(k_e-g)(1+k_e)^N}")
    st.markdown(
        "- **Cost of equity** from CAPM: risk-free rate + beta × equity risk premium.\n"
        "- **Defaults** come from the latest annual statements on Yahoo Finance: revenue growth = up to 3-year CAGR, "
        "net margin and payout = latest year. Change any of them in the sidebar.\n"
        "- **Part A preset (NKE)** loads the supplied Team 5 Base case path. The original workbook was not provided; "
        "agreement with that workbook has not been verified.\n"
        "- **Trading comps** apply the peer 25th percentile, median and 75th percentile multiples to the target's own "
        "revenue, EBITDA, EPS and book value per share.\n"
        "- Yahoo data can be incomplete or delayed. Always sanity-check inputs against the company's filings.")
    st.caption("Built for CBS International Business School, Phase III (B). Not investment advice.")