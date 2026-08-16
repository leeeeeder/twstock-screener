#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Taiwan stock screener (value / dividend focus).

Fetches open data published by TWSE (Taiwan Stock Exchange) and TPEx (Taipei
Exchange), merges it into one record per common stock, and writes:

  output/twstock_<date>.csv   full dataset, UTF-8 BOM so Excel opens it cleanly
  output/twstock_<date>.html  self-contained interactive screener (no CDN)

Only the Python standard library is required.

Usage:
    python twstock_screener.py                # normal run (uses today's cache)
    python twstock_screener.py --years 10     # longer dividend history window
    python twstock_screener.py --no-cache     # force re-download
    python twstock_screener.py --skip-history # fast run, no dividend history
"""

from __future__ import annotations

import argparse
import csv
import datetime as dt
import http.client
import json
import os
import ssl
import sys
import time
import urllib.error
import urllib.parse
import urllib.request

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CACHE_DIR = os.path.join(BASE_DIR, "cache")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")

USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) twstock-screener/1.0"

# TWSE serves a certificate without a Subject Key Identifier extension, which
# Python 3.13 rejects under its default strict X.509 profile. Drop only that
# flag; chain and hostname verification stay on.
SSL_CTX = ssl.create_default_context()
SSL_CTX.verify_flags &= ~ssl.VERIFY_X509_STRICT

# --- endpoints ---------------------------------------------------------------

TWSE_VALUATION = "https://openapi.twse.com.tw/v1/exchangeReport/BWIBBU_ALL"
TWSE_QUOTES = "https://openapi.twse.com.tw/v1/exchangeReport/STOCK_DAY_ALL"
TWSE_EXRIGHT = ("https://www.twse.com.tw/rwd/zh/exRight/TWT49U"
                "?startDate={y}0101&endDate={y}1231&response=json")
TWSE_INCOME = [
    "https://openapi.twse.com.tw/v1/opendata/t187ap06_L_ci",    # general
    "https://openapi.twse.com.tw/v1/opendata/t187ap06_L_mim",   # cross-industry
    "https://openapi.twse.com.tw/v1/opendata/t187ap06_L_basi",  # banking
    "https://openapi.twse.com.tw/v1/opendata/t187ap06_L_fh",    # financial holding
    "https://openapi.twse.com.tw/v1/opendata/t187ap06_L_ins",   # insurance
    "https://openapi.twse.com.tw/v1/opendata/t187ap06_L_bd",    # securities / futures
]

# Industry name as plain text, keyed by company code -- no code table to maintain.
# These also carry 營業利益, which the t187ap06_* income statements do not expose
# under a single consistent name.
TWSE_INDUSTRY = "https://openapi.twse.com.tw/v1/opendata/t187ap14_L"
TPEX_INDUSTRY = "https://www.tpex.org.tw/openapi/v1/mopsfin_t187ap14_O"

# The datasets above skip financial holdings and foreign (-KY) issuers. Company
# basic info covers everyone but gives the industry as a numeric code, so the
# code -> name table is derived from the companies that appear in both.
TWSE_BASIC = "https://openapi.twse.com.tw/v1/opendata/t187ap03_L"
TPEX_BASIC = "https://www.tpex.org.tw/openapi/v1/mopsfin_t187ap03_O"

# TPEx publishes only the current day's ex-dividend results, and no multi-year
# equivalent of TWSE's TWT49U exists: mopsfin_t187ap39_O stopped at ROC 110 and
# every dated query on the old and new sites ignores the date. So each run folds
# today's rows into a history file kept in the repo, and the record grows from
# the day tracking started rather than reaching backwards.
TPEX_EXRIGHT_TODAY = "https://www.tpex.org.tw/openapi/v1/tpex_exright_daily"
HISTORY_PATH = os.path.join(BASE_DIR, "data", "tpex_exright_history.json")

TPEX_VALUATION = "https://www.tpex.org.tw/openapi/v1/tpex_mainboard_peratio_analysis"
TPEX_QUOTES = "https://www.tpex.org.tw/openapi/v1/tpex_mainboard_daily_close_quotes"
# The "A" variants carry Chinese field names and stay current; the plain
# mopsfin_t187ap06_O_* endpoints use English keys with null 年度/季別.
TPEX_INCOME = [
    "https://www.tpex.org.tw/openapi/v1/mopsfin_t187ap06_O_ciA",
    "https://www.tpex.org.tw/openapi/v1/mopsfin_t187ap06_O_mimA",
    "https://www.tpex.org.tw/openapi/v1/mopsfin_t187ap06_O_basiA",
    "https://www.tpex.org.tw/openapi/v1/mopsfin_t187ap06_O_fhA",
    "https://www.tpex.org.tw/openapi/v1/mopsfin_t187ap06_O_insA",
    "https://www.tpex.org.tw/openapi/v1/mopsfin_t187ap06_O_bdA",
]

# Monthly revenue. The exchanges publish the year-on-year change themselves, so
# it is read rather than recomputed -- no second opinion on the denominator.
TWSE_REVENUE = "https://openapi.twse.com.tw/v1/opendata/t187ap05_L"
TPEX_REVENUE = "https://www.tpex.org.tw/openapi/v1/mopsfin_t187ap05_O"

# Balance sheets, for debt ratio and ROE. Only the _ci (general industry) files
# are substantial -- 967 and 857 rows. The other five carry a handful of rows
# each, sometimes without 資產總計/權益總計 at all and with the company code under
# an English key, so every field here is read defensively. Financial holdings,
# banks and insurers therefore end up without these ratios, the same gap they
# already have for 毛利率.
TWSE_BALANCE = [
    f"https://openapi.twse.com.tw/v1/opendata/t187ap07_L_{v}"
    for v in ("ci", "mim", "basi", "fh", "ins", "bd")
]
TPEX_BALANCE = [
    f"https://www.tpex.org.tw/openapi/v1/mopsfin_t187ap07_O_{v}"
    for v in ("ci", "mim", "basi", "fh", "ins", "bd")
]


# --- small helpers -----------------------------------------------------------

def log(msg: str) -> None:
    print(msg, flush=True)


def to_float(value):
    """Parse a number out of the exchanges' string fields; '' / '-' / 'N/A' -> None."""
    if value is None:
        return None
    text = str(value).strip().replace(",", "").replace("%", "")
    if text in ("", "-", "--", "N/A", "n/a", "無", "不適用"):
        return None
    try:
        return float(text)
    except ValueError:
        return None


def pick(row: dict, *needles):
    """Return the first value whose key contains every needle (field names vary)."""
    for key, value in row.items():
        if all(n in key for n in needles):
            return value
    return None


def row_code(row: dict) -> str:
    """Company code, whether the dataset names the column in Chinese or English."""
    value = row.get("公司代號") or row.get("SecuritiesCompanyCode") or pick(row, "公司代號")
    return str(value or "").strip()


def is_common_stock(code: str) -> bool:
    """Keep ordinary shares only: exactly four digits, not an ETF (00xx)."""
    return len(code) == 4 and code.isdigit() and not code.startswith("00")


def roc_to_iso(roc: str):
    """'1150814' -> '2026-08-14'. Returns None if unparseable."""
    digits = "".join(ch for ch in str(roc) if ch.isdigit())
    if len(digits) != 7:
        return None
    return f"{int(digits[:3]) + 1911:04d}-{digits[3:5]}-{digits[5:7]}"


# Minimum gap between requests to the same host. TWSE throttles per IP: a run
# that fired ~10 of its endpoints back to back had every single one come back
# as non-JSON while TPEx, a different host, was untouched. Pacing costs about
# fifteen seconds per run and removes the whole failure mode.
MIN_HOST_INTERVAL = 1.5
_last_request: dict[str, float] = {}


def throttle(url: str) -> None:
    host = urllib.parse.urlsplit(url).netloc
    wait = MIN_HOST_INTERVAL - (time.monotonic() - _last_request.get(host, 0.0))
    if wait > 0:
        time.sleep(wait)
    _last_request[host] = time.monotonic()


def fetch_json(url: str, cache_key: str, use_cache: bool = True, retries: int = 4):
    """GET JSON with an on-disk cache. Returns None if the source is unavailable."""
    path = os.path.join(CACHE_DIR, cache_key + ".json")
    if use_cache and os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as fh:
                return json.load(fh)
        except (OSError, json.JSONDecodeError):
            pass  # corrupt cache, refetch

    request = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    for attempt in range(1, retries + 1):
        throttle(url)
        try:
            with urllib.request.urlopen(request, timeout=90, context=SSL_CTX) as resp:
                payload = json.loads(resp.read().decode("utf-8"))
            os.makedirs(os.path.dirname(path), exist_ok=True)
            with open(path, "w", encoding="utf-8") as fh:
                json.dump(payload, fh, ensure_ascii=False)
            return payload
        # OSError covers URLError, TimeoutError and reset connections;
        # HTTPException covers IncompleteRead, which a truncated multi-megabyte
        # response raises and which used to escape and kill the whole run.
        except (OSError, http.client.HTTPException, json.JSONDecodeError) as exc:
            if attempt == retries:
                log(f"  ! 取得失敗（已略過）: {url}\n    {type(exc).__name__}: {exc}")
                return None
            # Back off hard: if this is throttling, short retries make it worse.
            time.sleep(5 * attempt)
    return None


# --- source loaders ----------------------------------------------------------

def load_valuation(stocks: dict, use_cache: bool) -> str:
    """TWSE + TPEx daily P/E, P/B and dividend yield. Returns the data date."""
    data_date = ""

    log("→ 上市 本益比／殖利率／股價淨值比 …")
    rows = fetch_json(TWSE_VALUATION, "twse_valuation", use_cache) or []
    for row in rows:
        code = str(row.get("Code", "")).strip()
        if not is_common_stock(code):
            continue
        data_date = data_date or roc_to_iso(row.get("Date")) or ""
        stocks.setdefault(code, new_stock(code, row.get("Name", "").strip(), "上市"))
        rec = stocks[code]
        rec["per"] = to_float(row.get("PEratio"))
        rec["pbr"] = to_float(row.get("PBratio"))
        rec["yield_pct"] = to_float(row.get("DividendYield"))
    log(f"  上市 {len(rows)} 筆")

    log("→ 上櫃 本益比／殖利率／股價淨值比 …")
    rows = fetch_json(TPEX_VALUATION, "tpex_valuation", use_cache) or []
    for row in rows:
        code = str(row.get("SecuritiesCompanyCode", "")).strip()
        if not is_common_stock(code):
            continue
        data_date = data_date or roc_to_iso(row.get("Date")) or ""
        stocks.setdefault(code, new_stock(code, row.get("CompanyName", "").strip(), "上櫃"))
        rec = stocks[code]
        rec["per"] = to_float(row.get("PriceEarningRatio"))
        rec["pbr"] = to_float(row.get("PriceBookRatio"))
        rec["yield_pct"] = to_float(row.get("YieldRatio"))
        rec["cash_div"] = to_float(row.get("DividendPerShare"))
    log(f"  上櫃 {len(rows)} 筆")
    return data_date


def load_quotes(stocks: dict, use_cache: bool) -> None:
    """TWSE + TPEx daily closing prices and turnover."""
    log("→ 上市 每日收盤行情 …")
    rows = fetch_json(TWSE_QUOTES, "twse_quotes", use_cache) or []
    for row in rows:
        code = str(row.get("Code", "")).strip()
        if code not in stocks:
            continue
        rec = stocks[code]
        rec["close"] = to_float(row.get("ClosingPrice"))
        rec["change"] = to_float(row.get("Change"))
        rec["high"] = to_float(row.get("HighestPrice"))
        rec["low"] = to_float(row.get("LowestPrice"))
        shares = to_float(row.get("TradeVolume"))
        rec["volume_lots"] = round(shares / 1000) if shares is not None else None
        rec["turnover"] = to_float(row.get("TradeValue"))

    log("→ 上櫃 每日收盤行情 …")
    rows = fetch_json(TPEX_QUOTES, "tpex_quotes", use_cache) or []
    for row in rows:
        code = str(row.get("SecuritiesCompanyCode", "")).strip()
        if code not in stocks:
            continue
        rec = stocks[code]
        rec["close"] = to_float(row.get("Close"))
        rec["change"] = to_float(row.get("Change"))
        rec["high"] = to_float(row.get("High"))
        rec["low"] = to_float(row.get("Low"))
        shares = to_float(row.get("TradingShares"))
        rec["volume_lots"] = round(shares / 1000) if shares is not None else None
        rec["turnover"] = to_float(row.get("TransactionAmount"))

    # Derived fields the exchanges do not publish directly.
    for rec in stocks.values():
        close, change = rec["close"], rec["change"]
        if close is not None and change is not None:
            prev = close - change
            rec["change_pct"] = round(change / prev * 100, 2) if prev else None
        # TWSE publishes yield but not cash dividend per share; back it out.
        if rec["cash_div"] is None and close and rec["yield_pct"]:
            rec["cash_div"] = round(close * rec["yield_pct"] / 100, 2)


def load_revenue(stocks: dict, use_cache: bool) -> None:
    """Monthly revenue and its year-on-year change, as published."""
    log("→ 月營收年增率 …")
    matched = 0
    for key, url in [("twse_revenue", TWSE_REVENUE), ("tpex_revenue", TPEX_REVENUE)]:
        for row in fetch_json(url, key, use_cache) or []:
            code = row_code(row)
            if code not in stocks:
                continue
            rec = stocks[code]
            # Published at full float precision (1.5379365538929763); round it.
            yoy = to_float(pick(row, "營業收入", "去年同月增減"))
            cum = to_float(pick(row, "累計營業收入", "前期比較增減"))
            rec["rev_yoy"] = round(yoy, 2) if yoy is not None else None
            rec["rev_cum_yoy"] = round(cum, 2) if cum is not None else None
            month = str(row.get("資料年月") or "").strip()
            rec["rev_month"] = f"{month[:3]}/{month[3:]}" if len(month) == 5 else month
            matched += 1
    log(f"  對應到 {matched} 檔")


def load_balance(stocks: dict, use_cache: bool) -> None:
    """Balance sheet totals, for the debt ratio and ROE."""
    log("→ 資產負債表（負債比／ROE）…")
    matched = 0
    sources = [("twse_balance", TWSE_BALANCE), ("tpex_balance", TPEX_BALANCE)]
    for prefix, urls in sources:
        for url in urls:
            key = f"{prefix}_{url.rsplit('_', 1)[-1]}"
            for row in fetch_json(url, key, use_cache) or []:
                code = row_code(row)
                if code not in stocks:
                    continue
                rec = stocks[code]
                # Several of these files omit the totals entirely; keep whatever
                # is there and let compute_ratios() skip the incomplete ones.
                assets = to_float(row.get("資產總計"))
                debt = to_float(row.get("負債總計"))
                equity = to_float(row.get("權益總計"))
                book = to_float(row.get("每股參考淨值"))
                if assets is not None:
                    rec["assets"] = assets
                if debt is not None:
                    rec["debt"] = debt
                if equity is not None:
                    rec["equity"] = equity
                if book is not None:
                    rec["book_value"] = book
                matched += 1
    log(f"  對應到 {matched} 檔")


def compute_ratios(stocks: dict) -> None:
    """
    Debt ratio and cumulative ROE.

    ROE uses 本期淨利 from the income statement, which is a year-to-date figure
    (115Q2 means the first half only). It is deliberately NOT annualised: for a
    cyclical company doubling a half-year profit is badly misleading. The column
    is labelled 累計 so the period is read off the 財報期別 column.
    """
    for rec in stocks.values():
        assets, debt = rec["assets"], rec["debt"]
        if assets and debt is not None:
            rec["debt_ratio"] = round(debt / assets * 100, 1)
        equity, profit = rec["equity"], rec["net_income"]
        if equity and profit is not None:
            rec["roe"] = round(profit / equity * 100, 1)


def load_income(stocks: dict, use_cache: bool) -> None:
    """Latest cumulative income statement: EPS, revenue, gross and net margin."""
    log("→ 綜合損益表（EPS／毛利率／淨利率）…")
    sources = [("twse_income", TWSE_INCOME), ("tpex_income", TPEX_INCOME)]
    loaded = 0
    for prefix, urls in sources:
        for url in urls:
            key = f"{prefix}_{url.rsplit('_', 1)[-1]}"
            rows = fetch_json(url, key, use_cache) or []
            for row in rows:
                code = str(pick(row, "公司代號") or "").strip()
                if code not in stocks:
                    continue
                rec = stocks[code]
                year = str(pick(row, "年度") or "").strip()
                quarter = str(pick(row, "季別") or "").strip()
                rec["eps"] = to_float(pick(row, "基本每股盈餘"))
                rec["eps_period"] = f"{year}Q{quarter}" if year and quarter else ""
                revenue = to_float(pick(row, "營業收入"))
                gross = to_float(pick(row, "營業毛利"))
                net = to_float(pick(row, "本期淨利"))
                rec["revenue"] = revenue
                rec["net_income"] = net      # kept raw so compute_ratios() can do ROE
                if revenue and gross is not None:
                    rec["gross_margin"] = round(gross / revenue * 100, 2)
                if revenue and net is not None:
                    rec["net_margin"] = round(net / revenue * 100, 2)
                loaded += 1
    log(f"  對應到 {loaded} 筆財報")


def load_industry(stocks: dict, use_cache: bool) -> None:
    """Industry name plus operating margin, in two passes (named, then by code)."""
    log("→ 產業別 …")
    matched = 0
    for key, url in [("twse_industry", TWSE_INDUSTRY), ("tpex_industry", TPEX_INDUSTRY)]:
        for row in fetch_json(url, key, use_cache) or []:
            code = row_code(row)
            if code not in stocks:
                continue
            rec = stocks[code]
            rec["industry"] = str(row.get("產業別") or "").strip()
            revenue = to_float(row.get("營業收入"))
            operating = to_float(row.get("營業利益"))
            if revenue and operating is not None:
                rec["op_margin"] = round(operating / revenue * 100, 2)
            matched += 1

    # Second pass: numeric industry codes for everyone, including the companies
    # the first pass missed (financial holdings, insurers, -KY issuers).
    code_of: dict[str, str] = {}
    for key, url in [("twse_basic", TWSE_BASIC), ("tpex_basic", TPEX_BASIC)]:
        for row in fetch_json(url, key, use_cache) or []:
            company = row_code(row)
            industry_id = str(row.get("產業別")
                              or row.get("SecuritiesIndustryCode") or "").strip()
            if company and industry_id:
                code_of[company] = industry_id

    # Learn code -> name from companies the first pass already named, then apply
    # that name to every company sharing the code. The published labels are not
    # self-consistent -- code 17 comes through as 金融業 for 39 companies and
    # 金融保險業 for 2 -- and an industry that splits into two names would
    # silently halve any filter on it. The code is the real classification, the
    # name only a label for it, so the majority label becomes canonical.
    votes: dict[str, dict[str, int]] = {}
    for company, industry_id in code_of.items():
        rec = stocks.get(company)
        if rec and rec["industry"]:
            tally = votes.setdefault(industry_id, {})
            tally[rec["industry"]] = tally.get(rec["industry"], 0) + 1
    names = {industry_id: max(tally.items(), key=lambda kv: kv[1])[0]
             for industry_id, tally in votes.items()}

    filled = renamed = 0
    for company, industry_id in code_of.items():
        rec = stocks.get(company)
        canonical = names.get(industry_id)
        if not rec or not canonical:
            continue
        if not rec["industry"]:
            rec["industry"] = canonical
            filled += 1
        elif rec["industry"] != canonical:
            rec["industry"] = canonical
            renamed += 1

    unknown = sum(1 for r in stocks.values() if not r["industry"])
    log(f"  直接對應 {matched} 檔，用產業代碼補上 {filled} 檔，"
        f"統一名稱 {renamed} 檔，仍無分類 {unknown} 檔")


def compute_payout(stocks: dict) -> None:
    """
    Payout ratio = cash dividend / trailing-twelve-month EPS.

    Both exchanges define 本益比 as 收盤價 ÷ 最近四季每股稅後盈餘, so TTM EPS is
    recoverable as close / P/E -- a better denominator than the cumulative EPS on
    the income statement, which covers only part of the year. Loss-making firms
    have a blank P/E, so they simply get no payout ratio.
    """
    for rec in stocks.values():
        per, close, div = rec["per"], rec["close"], rec["cash_div"]
        if not per or per <= 0 or not close or div is None:
            continue
        eps_ttm = close / per
        if eps_ttm <= 0:
            continue
        rec["eps_ttm"] = round(eps_ttm, 2)
        rec["payout"] = round(div / eps_ttm * 100, 1)


def accumulate_tpex_dividends(stocks: dict, use_cache: bool) -> dict:
    """
    Fold today's OTC ex-dividend rows into the repo's history file.

    Returns the history. Streaks are only reported for calendar years the file
    covers from start to finish -- a few months of records must never be dressed
    up as a dividend record, which would be worse than the blank it replaces.
    """
    log("→ 上櫃除權息累積 …")
    today = dt.date.today().isoformat()

    history = {"tracking_since": today, "events": {}}
    if os.path.exists(HISTORY_PATH):
        try:
            with open(HISTORY_PATH, "r", encoding="utf-8") as fh:
                loaded = json.load(fh)
            if isinstance(loaded.get("events"), dict):
                history = loaded
                history.setdefault("tracking_since", today)
        except (OSError, json.JSONDecodeError) as exc:
            log(f"  ! 歷史檔讀取失敗，改為重新開始: {exc}")

    rows = fetch_json(TPEX_EXRIGHT_TODAY, "tpex_exright_today", use_cache) or []
    added = 0
    for row in rows:
        code = str(row.get("SecuritiesCompanyCode", "")).strip()
        if not is_common_stock(code):
            continue
        iso = roc_to_iso(row.get("Date"))
        cash = to_float(row.get("CashDividend"))
        if not iso or not cash:          # 權-only events pay no cash
            continue
        # Keyed by code+date so re-running the same day cannot double count.
        key = f"{code}@{iso}"
        if key not in history["events"]:
            history["events"][key] = round(cash, 4)
            added += 1

    os.makedirs(os.path.dirname(HISTORY_PATH), exist_ok=True)
    with open(HISTORY_PATH, "w", encoding="utf-8") as fh:
        json.dump(history, fh, ensure_ascii=False, indent=1, sort_keys=True)

    # Only whole calendar years that started on or after tracking began can be
    # trusted; anything earlier is a partial year we happened to catch the tail of.
    since = history["tracking_since"]
    first_full = int(since[:4]) + (0 if since[5:] == "01-01" else 1)
    last_full = dt.date.today().year - 1
    complete = list(range(first_full, last_full + 1))

    paid: dict[str, set[int]] = {}
    for key in history["events"]:
        code, _, iso = key.partition("@")
        paid.setdefault(code, set()).add(int(iso[:4]))

    counted = 0
    for code, rec in stocks.items():
        if rec["market"] != "上櫃":
            continue
        rec["div_tracking"] = since
        if not complete:
            continue
        years_paid = {y for y in paid.get(code, set()) if y in complete}
        rec["div_years"] = len(years_paid)
        streak = 0
        cursor = last_full
        while cursor in years_paid:
            streak += 1
            cursor -= 1
        rec["div_streak"] = streak
        counted += 1

    log(f"  新增 {added} 筆，累計 {len(history['events'])} 筆，"
        f"自 {since} 起追蹤")
    if complete:
        log(f"  可用的完整年度 {complete[0]}–{complete[-1]}，已為 {counted} 檔上櫃計算")
    else:
        log("  尚無完整年度，上櫃連續配息年數維持空白（僅標示追蹤起始日）")
    return history


def load_dividend_history(stocks: dict, years: int, use_cache: bool) -> dict:
    """
    Count cash-dividend years from the TWSE ex-dividend result table (TWT49U).

    This table covers listed (上市) stocks only. OTC history is accumulated
    separately by accumulate_tpex_dividends() because no multi-year TPEx source
    exists.
    """
    this_year = dt.date.today().year
    year_range = list(range(this_year - years + 1, this_year + 1))
    log(f"→ 上市除權息紀錄 {year_range[0]}–{year_range[-1]} （{len(year_range)} 年）…")

    paid: dict[str, set[int]] = {}
    for index, year in enumerate(year_range):
        # Past years are final; only the current year needs a daily refresh.
        cache_key = f"exright/twse_{year}"
        cacheable = use_cache or year < this_year
        cached = cacheable and os.path.exists(
            os.path.join(CACHE_DIR, cache_key + ".json"))
        payload = fetch_json(TWSE_EXRIGHT.format(y=year), cache_key, cacheable)
        if not payload or payload.get("stat") != "OK":
            log(f"  ! {year} 年無資料")
            continue
        fields = payload.get("fields", [])
        try:
            code_at = fields.index("股票代號")
            kind_at = fields.index("權/息")
        except ValueError:
            code_at, kind_at = 1, 6
        count = 0
        for row in payload.get("data", []):
            code = str(row[code_at]).strip()
            # "息" = cash, "權息" = cash + stock, "權" = stock only (not a payout).
            if code in stocks and "息" in str(row[kind_at]):
                paid.setdefault(code, set()).add(year)
                count += 1
        log(f"  {year}: {count} 檔配息")
        # TWSE throttles roughly 3 requests / 5 seconds; no need to wait on cache hits.
        if not cached and index < len(year_range) - 1:
            time.sleep(3)

    for code, rec in stocks.items():
        if rec["market"] != "上市":
            continue
        got = paid.get(code, set())
        rec["div_years"] = len(got)
        rec["last_div_year"] = max(got) if got else None
        # Streak counted backwards from the most recent year; a company that has
        # not yet gone ex-dividend this year still keeps last year's streak.
        streak, cursor = 0, this_year
        if this_year not in got:
            cursor = this_year - 1
        while cursor in got:
            streak += 1
            cursor -= 1
        rec["div_streak"] = streak

    return {"from": year_range[0], "to": year_range[-1], "span": len(year_range)}


def new_stock(code: str, name: str, market: str) -> dict:
    return {
        "code": code, "name": name, "market": market, "industry": "",
        "close": None, "change": None, "change_pct": None,
        "high": None, "low": None, "volume_lots": None, "turnover": None,
        "per": None, "pbr": None, "yield_pct": None, "cash_div": None,
        "eps": None, "eps_period": "", "eps_ttm": None, "payout": None,
        "revenue": None, "net_income": None,
        "gross_margin": None, "net_margin": None, "op_margin": None,
        "rev_yoy": None, "rev_cum_yoy": None, "rev_month": "",
        "assets": None, "debt": None, "equity": None, "book_value": None,
        "debt_ratio": None, "roe": None,
        "div_years": None, "div_streak": None, "last_div_year": None,
        "div_tracking": "",
    }


# --- output ------------------------------------------------------------------

CSV_COLUMNS = [
    ("code", "代號"), ("name", "名稱"), ("market", "市場"), ("industry", "產業別"),
    ("close", "收盤價"), ("change", "漲跌"), ("change_pct", "漲跌幅%"),
    ("volume_lots", "成交量(張)"), ("turnover", "成交金額(元)"),
    ("per", "本益比"), ("pbr", "股價淨值比"), ("yield_pct", "殖利率%"),
    ("cash_div", "現金股利(元)"), ("payout", "配息率%"), ("eps_ttm", "EPS(近四季估)"),
    ("eps", "EPS(累計)"), ("eps_period", "財報期別"),
    ("rev_yoy", "月營收年增率%"), ("rev_cum_yoy", "累計營收年增率%"),
    ("rev_month", "營收月份"),
    ("gross_margin", "毛利率%"), ("op_margin", "營益率%"), ("net_margin", "淨利率%"),
    ("roe", "ROE(累計)%"), ("debt_ratio", "負債比%"), ("book_value", "每股淨值"),
    ("div_streak", "連續配息年數"), ("div_years", "配息年數"),
    ("last_div_year", "最近配息年"), ("div_tracking", "配息追蹤起始"),
]


def write_csv(path: str, records: list) -> None:
    with open(path, "w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.writer(fh)
        writer.writerow([label for _, label in CSV_COLUMNS])
        for rec in records:
            writer.writerow([rec.get(key) if rec.get(key) is not None else ""
                             for key, _ in CSV_COLUMNS])


# Fields the page actually uses. Everything else stays in the CSV only: with
# ~2000 rows, each field name is repeated 2000 times in the payload, so both the
# column list and the row-as-array encoding below exist to keep the page small.
PAGE_FIELDS = [
    "code", "name", "market", "industry",
    "close", "change_pct", "volume_lots",
    "per", "pbr", "yield_pct", "cash_div", "payout", "eps_ttm",
    "eps", "eps_period", "rev_yoy", "rev_cum_yoy", "rev_month",
    "gross_margin", "op_margin", "net_margin", "roe", "debt_ratio", "book_value",
    "div_streak", "div_years", "last_div_year", "div_tracking",
]


def render_html(records: list, meta: dict, template_path: str) -> str:
    with open(template_path, "r", encoding="utf-8") as fh:
        template = fh.read()
    payload = json.dumps(
        {
            "meta": meta,
            "fields": PAGE_FIELDS,
            "rows": [[rec.get(f) for f in PAGE_FIELDS] for rec in records],
        },
        ensure_ascii=False, separators=(",", ":"))
    return template.replace("/*__DATA__*/null", payload)


def write_html(path: str, html: str) -> None:
    with open(path, "w", encoding="utf-8") as fh:
        fh.write(html)


def write_artifact_html(path: str, html: str) -> None:
    """
    Strip the document wrapper for publishing as a Claude Artifact.

    Artifacts wrap the uploaded file in their own <!doctype>/<head>/<body>, so
    the file must contain page content only -- <title> and <style> stay (the
    title names the artifact), the document tags around them go.

    The charset declaration is kept even though it normally lives in the head we
    are dropping: without it a host that serves the page without a charset falls
    back to Latin-1 and every Chinese character turns to mojibake. If the wrapper
    already declares UTF-8 its declaration wins and this one is simply ignored.
    """
    head_start = html.index("<title>")
    head_end = html.index("</head>")
    body_start = html.index("<body>") + len("<body>")
    body_end = html.index("</body>")
    head = html[head_start:head_end].strip()
    body = html[body_start:body_end].strip()
    with open(path, "w", encoding="utf-8") as fh:
        fh.write('<meta charset="utf-8">\n' + head + "\n\n" + body + "\n")


# --- main --------------------------------------------------------------------

def main() -> int:
    parser = argparse.ArgumentParser(description="台股選股資料抓取工具")
    parser.add_argument("--years", type=int, default=8,
                        help="除權息歷史回溯年數（預設 8）")
    parser.add_argument("--no-cache", action="store_true",
                        help="忽略當日快取，重新下載")
    parser.add_argument("--skip-history", action="store_true",
                        help="跳過除權息歷史（快很多，但沒有連續配息年數）")
    parser.add_argument("--allow-partial", action="store_true",
                        help="即使有一邊市場完全抓不到也照樣輸出（預設會中止）")
    args = parser.parse_args()

    today = dt.date.today().strftime("%Y%m%d")
    global CACHE_DIR
    CACHE_DIR = os.path.join(BASE_DIR, "cache", today)
    os.makedirs(CACHE_DIR, exist_ok=True)
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    use_cache = not args.no_cache

    stocks: dict[str, dict] = {}
    data_date = load_valuation(stocks, use_cache)
    if not stocks:
        log("\n沒有取得任何資料，請檢查網路連線後重試。")
        return 1

    # A whole market missing means an exchange blocked or broke, not that the
    # market is empty. Writing that file anyway looks successful while silently
    # dropping half the universe, so stop unless the caller opted in.
    listed_count = sum(1 for r in stocks.values() if r["market"] == "上市")
    otc_count = len(stocks) - listed_count
    if (not listed_count or not otc_count) and not args.allow_partial:
        missing = "上市" if not listed_count else "上櫃"
        log(f"\n中止：完全沒有取得{missing}資料"
            f"（上市 {listed_count} / 上櫃 {otc_count}）。")
        log("這通常代表該交易所擋住了這台機器，而不是真的沒有股票。")
        log("確定要輸出不完整的結果，請加上 --allow-partial。")
        return 2

    load_quotes(stocks, use_cache)

    # Quotes come from different endpoints than the valuation counts checked
    # above, so they can fail on their own and leave every price blank while the
    # page still looks populated.
    priced = sum(1 for r in stocks.values() if r["close"] is not None)
    if priced < len(stocks) * 0.8 and not args.allow_partial:
        log(f"\n中止：只有 {priced}/{len(stocks)} 檔取得收盤價，行情資料不完整。")
        log("確定要輸出不完整的結果，請加上 --allow-partial。")
        return 2

    load_industry(stocks, use_cache)
    load_income(stocks, use_cache)
    load_revenue(stocks, use_cache)
    load_balance(stocks, use_cache)
    compute_payout(stocks)
    compute_ratios(stocks)

    history = {"from": None, "to": None, "span": 0}
    if not args.skip_history:
        history = load_dividend_history(stocks, args.years, use_cache)
    # Runs even with --skip-history: missing a day loses that day's records
    # permanently, since TPEx only ever serves the current day.
    accumulate_tpex_dividends(stocks, use_cache)

    records = sorted(stocks.values(), key=lambda r: r["code"])
    listed = sum(1 for r in records if r["market"] == "上市")

    meta = {
        "data_date": data_date,
        "generated_at": dt.datetime.now().strftime("%Y-%m-%d %H:%M"),
        "total": len(records),
        "listed": listed,
        "otc": len(records) - listed,
        "history_from": history["from"],
        "history_to": history["to"],
        "history_span": history["span"],
    }

    csv_path = os.path.join(OUTPUT_DIR, f"twstock_{today}.csv")
    html_path = os.path.join(OUTPUT_DIR, f"twstock_{today}.html")
    # Deliberately not date-stamped: republishing the same file path updates the
    # existing artifact in place, so a shared link keeps working after a refresh.
    artifact_path = os.path.join(OUTPUT_DIR, "artifact.html")
    # The GitHub Pages entry point. This is the COMPLETE document -- artifact.html
    # is a body fragment that only works inside the Artifact wrapper.
    index_path = os.path.join(OUTPUT_DIR, "index.html")
    html = render_html(records, meta, os.path.join(BASE_DIR, "template.html"))
    write_csv(csv_path, records)
    write_html(html_path, html)
    write_html(index_path, html)
    write_artifact_html(artifact_path, html)

    log("\n完成。")
    log(f"  資料日期  {data_date or '（未取得）'}")
    log(f"  普通股    {len(records)} 檔（上市 {listed} / 上櫃 {len(records) - listed}）")
    log(f"  CSV       {csv_path}")
    log(f"  選股介面  {html_path}")
    log(f"  發布版    {artifact_path}")
    log("\n用瀏覽器打開選股介面那個 .html 就能開始篩選。")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        log("\n已中止。")
        sys.exit(130)
