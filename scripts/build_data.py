#!/usr/bin/env python3
"""
US active equity ETF tracker - daily build (GitHub Actions).

Holdings sources (verified by probe 2026-11-03)
  ark       ARK daily CSV (assets.ark-funds.com)                 full, daily
  jpm       JPMorgan daily holdings xlsx                         full, daily
  capgroup  Capital Group daily-holdings download (xlsx)         full, daily
  zacks     Zacks ETF holdings page (BAI, DYNF, CHAT, FBCG, TCHP) full list; update cadence checked via as-of / content changes

Each run stores a snapshot per ETF in data/hist/<T>.json (only when the portfolio changed),
then compares consecutive snapshots:
  new  / exit           position opened / closed
  buy  / sell           share count changed beyond the fund-wide creation/redemption factor
Fund-wide flows move every position by the same ratio; that common ratio (median) is removed first,
so only manager decisions remain.
"""
import csv
import datetime as dt
import hashlib
import io
import json
import re
import statistics
import time
from pathlib import Path

import requests

ROOT = Path(__file__).resolve().parent.parent
DATA = ROOT / "data"
HIST = DATA / "hist"
CFG = json.loads((ROOT / "config" / "etfs.json").read_text(encoding="utf-8"))
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
H = {"User-Agent": UA, "Accept": "*/*", "Accept-Language": "en-US,en;q=0.9"}
TODAY = dt.date.today()
DIAG = []
TRADE_MIN = 0.02            # flow-adjusted share change needed to call it a buy/sell (2%)
KEEP_SNAPS = 60


def log(*a):
    print(*a, flush=True)


def diag(src, msg):
    DIAG.append({"src": src, "msg": str(msg)[:400]})
    log("DIAG", src, msg)


def get(url, timeout=60):
    for i in range(3):
        try:
            r = requests.get(url, headers=H, timeout=timeout)
            if r.status_code == 200:
                return r
        except requests.RequestException:
            pass
        time.sleep(2 * (i + 1))
    return None


def num(x):
    if x is None:
        return None
    if isinstance(x, (int, float)):
        return float(x)
    s = str(x).replace("$", "").replace(",", "").replace("%", "").strip()
    if s in ("", "-", "--", "N/A"):
        return None
    neg = s.startswith("(") and s.endswith(")")
    try:
        v = float(s.strip("()"))
        return -v if neg else v
    except ValueError:
        return None


def to_date(s):
    if isinstance(s, (dt.date, dt.datetime)):
        return (s.date() if isinstance(s, dt.datetime) else s).isoformat()
    s = str(s or "").strip()
    for fmt in ("%m/%d/%Y", "%m/%d/%y", "%Y-%m-%d", "%b %d, %Y", "%B %d, %Y", "%d-%b-%Y"):
        try:
            return dt.datetime.strptime(s, fmt).date().isoformat()
        except ValueError:
            pass
    return ""


def row(key, name, shares, weight, value=None, typ=""):
    return {"key": key, "name": name, "shares": shares, "weight": weight, "value": value, "type": typ}


# ================================================================== parsers
def parse_ark(text):
    rows, asof = [], ""
    for r in csv.DictReader(io.StringIO(text)):
        name = (r.get("company") or "").strip()
        if not name or r.get("date", "").lower().startswith("the "):
            continue
        asof = asof or to_date(r.get("date"))
        tk = (r.get("ticker") or "").strip().upper()
        rows.append(row(tk or (r.get("cusip") or name).upper(), name, num(r.get("shares")),
                        num(r.get("weight (%)")), num(r.get("market value ($)"))))
    return rows, asof


HEAD = {
    "ticker": ("ticker", "symbol"),
    "name": ("security description", "description", "security name", "name", "holding", "issuer"),
    "shares": ("shares/par", "shares", "quantity", "par value", "nominal"),
    "value": ("market value", "market val", "value (usd)", "value"),
    "weight": ("% of net assets", "% of net", "% of fund", "weight", "% of market value", "percent", "%"),
    "type": ("security type", "asset type", "asset class", "type"),
}


def parse_table_rows(table):
    """Generic header-detecting parser for issuer spreadsheets."""
    asof = ""
    hdr_i, cols = None, {}
    for i, r in enumerate(table[:40]):
        cells = [str(c or "").strip() for c in r]
        low = [c.lower() for c in cells]
        if not asof:
            for c in cells:
                m = re.search(r"(\d{1,2}/\d{1,2}/\d{2,4})", c)
                if m and ("as of" in c.lower() or len(c) < 12):
                    asof = to_date(m.group(1))
            for c in r:
                if isinstance(c, (dt.date, dt.datetime)):
                    asof = asof or to_date(c)
        hits = sum(any(k == x or (len(k) > 3 and k in x) for x in low) for ks in (HEAD["ticker"], HEAD["shares"], HEAD["weight"]) for k in ks[:3])
        if hits >= 2:
            hdr_i = i
            for field, keys in HEAD.items():
                for k in keys:
                    j = next((n for n, x in enumerate(low) if x == k), None)
                    if j is None and len(k) > 3:
                        j = next((n for n, x in enumerate(low) if k in x), None)
                    if j is not None and j not in cols.values():
                        cols[field] = j
                        break
            break
    if hdr_i is None:
        return [], asof
    out = []
    for r in table[hdr_i + 1:]:
        cell = lambda f: (r[cols[f]] if f in cols and cols[f] < len(r) else None)
        name = str(cell("name") or "").strip()
        tk = str(cell("ticker") or "").strip().upper()
        if not name and not tk:
            continue
        out.append(row(tk or name.upper(), name or tk, num(cell("shares")), num(cell("weight")), num(cell("value")), str(cell("type") or "")))
    return out, asof


def parse_xlsx(blob):
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(blob), read_only=True, data_only=True)
    best = ([], "")
    for ws in wb.worksheets:
        table = [list(r) for r in ws.iter_rows(values_only=True)]
        res = parse_table_rows(table)
        if len(res[0]) > len(best[0]):
            best = res
    return best


def parse_zacks(html):
    m = re.search(r"etf_holdings\.formatted_data\s*=\s*(\[.*?\]);", html, re.S)
    if not m:
        return [], ""
    data = json.loads(m.group(1))
    rows = []
    for r in data:
        name = re.sub(r"<[^>]+>", "", r[0])
        t = re.search(r'title="([^"]+)"', r[0])
        name = t.group(1) if t else name.replace("..", "").strip()
        tk = re.search(r'rel="([^"]+)"', r[1])
        tk = (tk.group(1) if tk else re.sub(r"<[^>]+>", "", r[1])).strip().upper()
        rows.append(row(tk or name.upper(), name, num(r[2]), num(r[3])))
    asof = ""
    m2 = re.search(r"[Hh]oldings?[^<]{0,80}?as of\s*([0-9]{1,2}/[0-9]{1,2}/[0-9]{2,4})", html) or \
         re.search(r"as of\s*([0-9]{1,2}/[0-9]{1,2}/[0-9]{4})", html)
    if m2:
        asof = to_date(m2.group(1))
    return rows, asof


# ================================================================== fetchers
def fetch_zacks(t):
    r = get(f"https://www.zacks.com/funds/etf/{t}/holding")
    return (*parse_zacks(r.text), "zacks") if r else ([], "", "zacks")


def fetch_holdings(e):
    """Issuer file first; Zacks when the issuer file is missing, too small, or stale (> 6 days)."""
    rows, asof, src = fetch_issuer(e)
    stale = bool(asof) and (TODAY - dt.date.fromisoformat(asof)).days > 6
    if src != "zacks" and (len(rows) < 5 or stale):
        zr, za, zs = fetch_zacks(e["ticker"])
        if len(zr) >= 5:
            diag(e["ticker"], f"issuer file {'stale ' + asof if stale else 'unavailable'} -> Zacks used")
            return zr, za, zs
    return rows, asof, src


def fetch_issuer(e):
    t = e["ticker"]
    src = e["issuer"] if (e["issuer"] in ("ark", "capgroup") or (e["issuer"] == "jpm" and e.get("cusip"))) else "zacks"
    try:
        if src == "ark":
            r = get(f"https://assets.ark-funds.com/fund-documents/funds-etf-csv/{e['csv']}")
            return (*parse_ark(r.content.decode("utf-8", "replace")), src) if r else ([], "", src)
        if src == "jpm":
            r = get(f"https://am.jpmorgan.com/FundsMarketingHandler/excel?type=dailyETFHoldings&cusip={e['cusip']}&country=us&role=adv&fundType=N_ETF&locale=en-US&isUnderlyingHolding=false&isProxyHolding=false")
            return (*parse_xlsx(r.content), src) if r else ([], "", src)
        if src == "capgroup":
            r = get(f"https://www.capitalgroup.com/api/investments/investment-service/v1/etfs/{t.lower()}/download/daily-holdings?audience=individual&redirect=true")
            return (*parse_xlsx(r.content), src) if r else ([], "", src)
        return fetch_zacks(t)
    except Exception as ex:
        diag(t, f"holdings error: {ex}")
        return [], "", src


EQUITY_RX = re.compile(r"COMMON|STOCK|EQUITY(?! LINK)|ADR|REIT|SHARES|ORDINARY", re.I)
CASHLIKE_RX = re.compile(r"CASH|TREAS|MONEY|LIQUID|\bFUND\b|\bUSD\b|BLK CSH|MMF|COLLATERAL|DERIV|FUTURE", re.I)
DROP_RX = re.compile(r"LINKED|\bELN\b|NOTE|CASH|MONEY MARKET|FUTURE|OPTION|SWAP|FORWARD|TREASURY|T-BILL|DERIVATIVE|REPO", re.I)


def clean(rows, e):
    out = []
    for r in rows:
        tag = f"{r['type']} {r['name']} {r['key']}"
        if DROP_RX.search(tag) and not (r["type"] and EQUITY_RX.search(r["type"]) and not DROP_RX.search(r["type"])):
            continue                           # ELNs (JEPI), cash, futures etc. are not stock picks
        if r["shares"] is None and r["weight"] is None:
            continue
        if r["key"] in ("N/A", "NA", "--", "-") and CASHLIKE_RX.search(r["name"] + " " + r["key"]):
            continue
        if r["key"] in ("N/A", "NA", "--", "-"):
            r["key"] = r["name"].upper()[:40]          # keep foreign/unlisted names as their own key
        out.append(r)
    # merge duplicate keys
    m = {}
    for r in out:
        if r["key"] in m:
            a = m[r["key"]]
            a["shares"] = (a["shares"] or 0) + (r["shares"] or 0)
            a["weight"] = (a["weight"] or 0) + (r["weight"] or 0)
            a["value"] = (a["value"] or 0) + (r["value"] or 0) if (a["value"] is not None or r["value"] is not None) else None
        else:
            m[r["key"]] = dict(r)
    rows = list(m.values())
    tw = sum(r["weight"] or 0 for r in rows)
    if 0.3 < tw < 1.5:                                  # file gives fractions (0.06 = 6%)
        for r in rows:
            r["weight"] = (r["weight"] or 0) * 100 if r["weight"] is not None else None
    # weights as % of the stock sleeve when a file gives none
    if all(r["weight"] is None for r in rows):
        tot = sum(r["value"] or 0 for r in rows) or 1
        for r in rows:
            r["weight"] = (r["value"] or 0) / tot * 100
    return sorted(rows, key=lambda r: -(r["weight"] or 0))


# ================================================================== history + diff
def snap_hash(rows):
    return hashlib.sha1(json.dumps([[r["key"], r["shares"]] for r in rows], sort_keys=True).encode()).hexdigest()[:12]


def shares_ok(rows):
    sh = [r["shares"] for r in rows if r["shares"]]
    return len(sh) >= max(3, len(rows) // 2) and statistics.median(sh) >= 1000


def diff_weight(prev, now):
    """Fallback when share counts are unusable: new/exit + large weight moves (price moves also count -> labelled)."""
    P = {r["key"]: r for r in prev}
    N = {r["key"]: r for r in now}
    out = []
    for k, n in N.items():
        p = P.get(k)
        if not p:
            out.append({"key": k, "name": n["name"], "kind": "new", "shares": n["shares"], "prev_shares": 0, "chg": None, "weight": n["weight"], "prev_weight": 0})
        elif (p["weight"] or 0) > 0.2 and n["weight"] is not None:
            rel = n["weight"] / p["weight"] - 1
            if abs(rel) >= 0.25 and abs(n["weight"] - p["weight"]) >= 0.3:
                out.append({"key": k, "name": n["name"], "kind": "buy" if rel > 0 else "sell", "shares": n["shares"], "prev_shares": p["shares"],
                            "chg": rel, "weight": n["weight"], "prev_weight": p["weight"]})
    for k, p in P.items():
        if k not in N:
            out.append({"key": k, "name": p["name"], "kind": "exit", "shares": 0, "prev_shares": p["shares"], "chg": -1.0, "weight": 0, "prev_weight": p["weight"]})
    order = {"new": 0, "exit": 1, "buy": 2, "sell": 3}
    out.sort(key=lambda d: (order[d["kind"]], -max(d["weight"] or 0, d["prev_weight"] or 0)))
    return out, 1.0


def diff(prev, now):
    if not (shares_ok(prev) and shares_ok(now)):
        d, f = diff_weight(prev, now)
        for x in d:
            x["basis"] = "weight"
        return d, f
    P = {r["key"]: r for r in prev}
    N = {r["key"]: r for r in now}
    common = [k for k in N if k in P and (P[k]["shares"] or 0) > 0 and (N[k]["shares"] or 0) > 0]
    ratios = [N[k]["shares"] / P[k]["shares"] for k in common]
    flow = statistics.median(ratios) if ratios else 1.0
    out = []
    for k, n in N.items():
        p = P.get(k)
        if not p:
            out.append({"key": k, "name": n["name"], "kind": "new", "shares": n["shares"], "prev_shares": 0, "chg": None,
                        "weight": n["weight"], "prev_weight": 0})
        elif (p["shares"] or 0) > 0 and n["shares"] is not None:
            adj = n["shares"] / (p["shares"] * flow) - 1
            if abs(adj) >= TRADE_MIN:
                out.append({"key": k, "name": n["name"], "kind": "buy" if adj > 0 else "sell", "shares": n["shares"],
                            "prev_shares": p["shares"], "chg": adj, "weight": n["weight"], "prev_weight": p["weight"]})
    for k, p in P.items():
        if k not in N:
            out.append({"key": k, "name": p["name"], "kind": "exit", "shares": 0, "prev_shares": p["shares"], "chg": -1.0,
                        "weight": 0, "prev_weight": p["weight"]})
    order = {"new": 0, "exit": 1, "buy": 2, "sell": 3}
    out.sort(key=lambda d: (order[d["kind"]], -max(d["weight"] or 0, d["prev_weight"] or 0)))
    return out, flow


def update_history(t, rows, asof):
    HIST.mkdir(parents=True, exist_ok=True)
    p = HIST / f"{t}.json"
    hist = json.loads(p.read_text(encoding="utf-8")) if p.exists() else []
    h = snap_hash(rows)
    date = asof or TODAY.isoformat()
    changed = not hist or (hist[-1]["hash"] != h and date >= hist[-1]["asof"])
    if changed:
        if hist and hist[-1]["asof"] == date:
            hist[-1] = {"asof": date, "hash": h, "rows": rows}      # same day re-published: replace
        else:
            hist.append({"asof": date, "hash": h, "rows": rows})
        hist = hist[-KEEP_SNAPS:]
        p.write_text(json.dumps(hist, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    return hist


# ================================================================== market data (Yahoo chart, price returns)
def market(t):
    r = get(f"https://query1.finance.yahoo.com/v8/finance/chart/{t}?range=1y&interval=1d")
    if not r:
        return {}
    try:
        res = r.json()["chart"]["result"][0]
        ts = res["timestamp"]
        q = res["indicators"]["quote"][0]
        px = [(dt.date.fromtimestamp(a), c, v) for a, c, v in zip(ts, q["close"], q["volume"]) if c]
        last_d, last = px[-1][0], px[-1][1]

        def ret(days):
            tgt = last_d - dt.timedelta(days=days)
            base = next((c for d0, c, _ in px if d0 >= tgt), None)
            return last / base - 1 if base else None
        ytd_base = next((c for d0, c, _ in px if d0.year == last_d.year), None)
        vols = [v * c for d0, c, v in px if v and d0 >= last_d - dt.timedelta(days=90)]
        return {"price": last, "price_date": last_d.isoformat(), "r1m": ret(30), "r3m": ret(91), "r1y": ret(365),
                "ytd": last / ytd_base - 1 if ytd_base else None,
                "adv_usd": sum(vols) / len(vols) if vols else None}
    except Exception as ex:
        diag(t, f"market error: {ex}")
        return {}


def aum(t):
    r = get(f"https://stockanalysis.com/etf/{t.lower()}/")
    if not r:
        return None
    m = re.search(r"Assets[^$]{0,200}?\$([\d\.]+)\s*([BMK])", r.text)
    if not m:
        return None
    return float(m.group(1)) * {"B": 1e9, "M": 1e6, "K": 1e3}[m.group(2)]


# ================================================================== main
def main():
    out = []
    for e in CFG["etfs"]:
        t = e["ticker"]
        raw, asof, src = fetch_holdings(e)
        rows = clean(raw, e)
        log(t, src, "rows", len(raw), "->", len(rows), "asof", asof)
        if len(rows) < 5:
            diag(t, f"few holdings parsed ({len(rows)}) from {src}")
        if len(rows) > CFG.get("max_holdings", 300):
            diag(t, f"{len(rows)} holdings (> {CFG.get('max_holdings', 300)}): consider removing from the list")
        hist = update_history(t, rows, asof) if rows else (json.loads((HIST / f"{t}.json").read_text()) if (HIST / f"{t}.json").exists() else [])
        trades = []
        for a, b in zip(hist, hist[1:]):
            d, flow = diff(a["rows"], b["rows"])
            if d:
                trades.append({"date": b["asof"], "prev": a["asof"], "flow": flow, "items": d})
        trades = trades[::-1][:30]
        latest = hist[-1] if hist else {"asof": "", "rows": []}
        m = market(t)
        m["aum"] = aum(t)
        stale = bool(latest["asof"]) and (TODAY - dt.date.fromisoformat(latest["asof"])).days > 6
        if stale:
            diag(t, f"holdings date {latest['asof']} is old (source not updated)")
        out.append({**e, "source": src, "asof": latest["asof"], "asof_from_file": bool(asof), "n": len(latest["rows"]),
                    "stale": stale, "basis": "shares" if shares_ok(latest["rows"]) else "weight",
                    "holdings": latest["rows"], "trades": trades, "snapshots": len(hist), "mkt": m})
        time.sleep(1)
    res = {"status": "ok", "generated_at": dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds"),
           "trade_min": TRADE_MIN, "etfs": out, "diag": DIAG}
    DATA.mkdir(parents=True, exist_ok=True)
    (DATA / "etfs.json").write_text(json.dumps(res, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    log("wrote data/etfs.json", len(out), "etfs", len(DIAG), "diag")


if __name__ == "__main__":
    main()
