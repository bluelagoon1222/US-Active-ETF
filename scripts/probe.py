#!/usr/bin/env python3
"""
Probe: can GitHub Actions download each issuer's DAILY holdings file?
Writes probe/result.json with status, size, and a short sample per ETF and per URL pattern.
No API keys needed.
"""
import datetime as dt
import json
import re
from pathlib import Path

import requests

ROOT = Path(__file__).resolve().parent.parent
CFG = json.loads((ROOT / "config" / "etfs.json").read_text(encoding="utf-8"))
OUT = ROOT / "probe" / "result.json"
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/124.0 Safari/537.36")
H = {"User-Agent": UA, "Accept": "*/*", "Accept-Language": "en-US,en;q=0.9"}
R = {"run_at_utc": dt.datetime.utcnow().isoformat(timespec="seconds"), "checks": []}


def rec(name, ok, **kw):
    kw.update({"name": name, "ok": ok})
    R["checks"].append(kw)
    print(("OK   " if ok else "FAIL ") + name, {k: (str(v)[:100]) for k, v in kw.items() if k not in ("name", "ok", "sample")})


def get(url, **kw):
    try:
        r = requests.get(url, headers=H, timeout=40, **kw)
        ctype = r.headers.get("content-type", "")
        body = r.content
        txt = body[:4000].decode("utf-8", "replace") if ("text" in ctype or "json" in ctype or "csv" in ctype or body[:1] in (b"{", b"[", b"<")) else f"<binary {len(body)} bytes, {ctype}>"
        return r.status_code, ctype, len(body), txt, r.url
    except Exception as e:
        return None, "", 0, f"error {e}", url


def try_urls(name, urls, must=None):
    for u in urls:
        st, ct, n, txt, final = get(u)
        ok = st == 200 and n > 300 and (must is None or re.search(must, txt, re.I))
        rec(name, bool(ok), url=u, final=final, status=st, ctype=ct, bytes=n, sample=txt[:1500])
        if ok:
            return u
    return None


# ------------------------------------------------------------------ iShares (BAI, DYNF): screener -> product page -> CSV
def probe_ishares():
    scr = "https://www.ishares.com/us/product-screener/product-screener-v3.1.jsn?dcrPath=/templatedata/config/product-screener-v3/data/en/us-ishares/ishares-product-screener-backend-config&siteEntryPassthrough=true"
    st, ct, n, txt, _ = get(scr)
    pages = {}
    if st == 200:
        try:
            js = requests.get(scr, headers=H, timeout=60).json()
            for pid, d in js.items():
                tk = (d.get("localExchangeTicker") or "").upper()
                if tk in ("BAI", "DYNF"):
                    pages[tk] = d.get("productPageUrl") or pid
        except Exception as e:
            txt = f"json error {e}"
    rec("ishares_screener", bool(pages), status=st, bytes=n, pages=pages, sample=txt[:400])
    known = {"BAI": "/us/products/339081/ishares-a-i-innovation-and-tech-active-etf"}
    for tk in ("BAI", "DYNF"):
        page = pages.get(tk) or known.get(tk)
        if not page:
            continue
        page = page if page.startswith("http") else "https://www.ishares.com" + page
        try_urls(f"ishares_{tk}", [page + f"/1467271812596.ajax?fileType=csv&fileName={tk}_holdings&dataType=fund"], must=r"Ticker|Name")


# ------------------------------------------------------------------ ARK
def probe_ark():
    for e in [x for x in CFG["etfs"] if x["issuer"] == "ark"]:
        try_urls(f"ark_{e['ticker']}", [
            f"https://assets.ark-funds.com/fund-documents/funds-etf-csv/{e['csv']}",
            f"https://ark-funds.com/wp-content/uploads/funds-etf-csv/{e['csv']}",
        ], must=r"ticker|company")


# ------------------------------------------------------------------ JPMorgan JEPI
def probe_jpm():
    c = "46641Q332"
    try_urls("jpm_JEPI", [
        f"https://am.jpmorgan.com/FundsMarketingHandler/excel?type=dailyETFHoldings&cusip={c}&country=us&role=adv&fundType=N_ETF&locale=en-US&isUnderlyingHolding=false&isProxyHolding=false",
        f"https://am.jpmorgan.com/FundsMarketingHandler/excel?type=dailyETFHoldings&cusip={c}&country=us&role=adv&locale=en-US",
        f"https://am.jpmorgan.com/FundsMarketingHandler/product-data?cusip={c}&country=us&role=adv&language=en&userLoggedIn=false&version=8.24.1_1",
    ])


# ------------------------------------------------------------------ pages that should link to a holdings file
def probe_pages():
    pages = {
        "CHAT": ["https://www.roundhillinvestments.com/etf/chat/"],
        "CGGR": ["https://www.capitalgroup.com/individual/investments/fund/cggr", "https://www.capitalgroup.com/advisor/investments/etf/cggr"],
        "CGDV": ["https://www.capitalgroup.com/individual/investments/fund/cgdv"],
        "FBCG": ["https://www.fidelity.com/etfs/fbcg", "https://fundresearch.fidelity.com/mutual-funds/summary/316092311"],
        "TCHP": ["https://www.troweprice.com/personal-investing/tools/fund-research/TCHP", "https://www.troweprice.com/financial-intermediary/us/en/investments/etfs/blue-chip-growth-etf.html"],
    }
    for tk, urls in pages.items():
        for u in urls:
            st, ct, n, txt, final = get(u)
            full = ""
            try:
                full = requests.get(u, headers=H, timeout=40).text if st == 200 else ""
            except Exception:
                pass
            links = sorted(set(re.findall(r'(https?://[^"\'\s>]+?(?:holding|Holding|HOLDING)[^"\'\s>]*)', full)))[:12]
            files = sorted(set(re.findall(r'(https?://[^"\'\s>]+?\.(?:csv|xlsx|xls)(?:\?[^"\'\s>]*)?)', full)))[:12]
            rec(f"page_{tk}", st == 200, url=u, final=final, status=st, bytes=len(full), holding_links=links, files=files,
                sample=re.sub(r"\s+", " ", full[full.lower().find("holding"):][:600]) if "holding" in full.lower() else "")


# ------------------------------------------------------------------ universal fallback: stockanalysis.com holdings page
def probe_fallback():
    for e in CFG["etfs"]:
        u = f"https://stockanalysis.com/etf/{e['ticker'].lower()}/holdings/"
        st, ct, n, txt, final = get(u)
        try:
            full = requests.get(u, headers=H, timeout=40).text if st == 200 else ""
        except Exception:
            full = ""
        rows = len(re.findall(r"<tr", full))
        asof = re.search(r"[Aa]s of ([A-Z][a-z]{2,8} \d{1,2}, \d{4})", full)
        rec(f"fallback_{e['ticker']}", st == 200 and rows > 5, url=u, status=st, rows=rows, asof=asof.group(1) if asof else "")


if __name__ == "__main__":
    for fn in (probe_ishares, probe_ark, probe_jpm, probe_pages, probe_fallback):
        try:
            fn()
        except Exception as e:
            rec(fn.__name__, False, error=str(e)[:200])
    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(json.dumps(R, ensure_ascii=False, indent=1), encoding="utf-8")
    print("wrote", OUT)
