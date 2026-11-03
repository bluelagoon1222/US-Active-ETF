#!/usr/bin/env python3
"""Probe round 2: fill the gaps (iShares passthrough, Roundhill, Capital Group, Fidelity, T. Rowe, Zacks) + parse samples."""
import datetime as dt, io, json, re
from pathlib import Path
import requests
import subprocess, sys
try:
    import openpyxl  # noqa
except ImportError:
    subprocess.run([sys.executable, "-m", "pip", "install", "-q", "openpyxl"], check=False)

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "probe" / "result.json"
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
H = {"User-Agent": UA, "Accept": "*/*", "Accept-Language": "en-US,en;q=0.9"}
R = {"run_at_utc": dt.datetime.utcnow().isoformat(timespec="seconds"), "round": 2, "checks": []}


def rec(name, ok, **kw):
    kw.update({"name": name, "ok": ok}); R["checks"].append(kw)
    print(("OK   " if ok else "FAIL ") + name, {k: str(v)[:90] for k, v in kw.items() if k not in ("name", "ok", "sample")})


def fetch(url, timeout=40):
    try:
        r = requests.get(url, headers=H, timeout=timeout)
        return r
    except Exception as e:
        class X: status_code = None; content = b""; text = f"error {e}"; headers = {}; url = url
        return X()


def text_probe(name, url, must):
    r = fetch(url)
    t = r.text if r.content else ""
    ok = r.status_code == 200 and bool(re.search(must, t[:20000], re.I)) and not t.lstrip().lower().startswith("<!doctype")
    rec(name, ok, url=url, status=r.status_code, ctype=r.headers.get("content-type", ""), bytes=len(r.content), sample=t[:1200])
    return ok, t


# 1) iShares with site-entry passthrough
for tk, page in (("BAI", "/us/products/339081/ishares-a-i-innovation-and-tech-active-etf"),
                 ("DYNF", "/us/products/307283/ishares-u-s-equity-factor-rotation-active-etf")):
    text_probe(f"ishares_{tk}", f"https://www.ishares.com{page}/1467271812596.ajax?fileType=csv&fileName={tk}_holdings&dataType=fund&siteEntryPassthrough=true", r"Ticker")

# 2) ARK CSV parse sample
r = fetch("https://assets.ark-funds.com/fund-documents/funds-etf-csv/ARK_INNOVATION_ETF_ARKK_HOLDINGS.csv")
rec("ark_ARKK_csv", r.status_code == 200, status=r.status_code, sample=r.content[:900].decode("utf-8", "replace"))

# 3) JPMorgan JEPI xlsx parse sample
r = fetch("https://am.jpmorgan.com/FundsMarketingHandler/excel?type=dailyETFHoldings&cusip=46641Q332&country=us&role=adv&fundType=N_ETF&locale=en-US&isUnderlyingHolding=false&isProxyHolding=false")
try:
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(r.content), read_only=True, data_only=True)
    ws = wb.worksheets[0]
    rows = [[("" if c is None else str(c)) for c in row] for row in ws.iter_rows(max_row=14, values_only=True)]
    rec("jpm_JEPI_xlsx", True, sheets=wb.sheetnames, sample=json.dumps(rows)[:1800])
except Exception as e:
    rec("jpm_JEPI_xlsx", False, error=str(e)[:200])

# 4) Roundhill CHAT: what the page loads
r = fetch("https://www.roundhillinvestments.com/etf/chat/")
t = r.text or ""
eps = sorted(set(re.findall(r'["\'](/[^"\']*(?:holding|Holding|data)[^"\']*\.(?:json|csv|xlsx|txt))["\']', t)))[:15]
srcs = sorted(set(re.findall(r'(?:src|href)=["\']([^"\']+\.(?:js|json|csv|xlsx))["\']', t)))[:25]
tds = len(re.findall(r"<td", t))
i = t.lower().find("top holdings"); i = i if i >= 0 else t.lower().find("holdings")
rec("roundhill_page", r.status_code == 200, tds=tds, endpoints=eps, srcs=srcs, sample=re.sub(r"\s+", " ", t[i:i + 2500]))
for u in ("https://www.roundhillinvestments.com/assets/data/CHAT_holdings.csv",
          "https://www.roundhillinvestments.com/assets/data/holdings/CHAT.csv",
          "https://www.roundhillinvestments.com/etf/chat/holdings.csv"):
    text_probe("roundhill_try", u, r"CHAT|Ticker|Name")

# 5) Capital Group
for u in ("https://www.capitalgroup.com/individual/investments/fund/cggr.html",
          "https://www.capitalgroup.com/advisor/investments/fund/cggr",
          "https://www.capitalgroup.com/api/investments/investment-service/v1/etfs/cggr/download/daily-holdings?audience=advisor&redirect=true",
          "https://www.capitalgroup.com/api/investments/investment-service/v1/etfs/cggr/download/daily-holdings?audience=individual&redirect=true"):
    r = fetch(u)
    t = r.text if (r.content and b"\x00" not in r.content[:200]) else f"<binary {len(r.content)} {r.headers.get('content-type','')}>"
    links = sorted(set(re.findall(r'(https?://[^"\'\s>]*(?:holding|Holding)[^"\'\s>]*)', t)))[:8]
    rec("capgroup", r.status_code == 200, url=u, status=r.status_code, ctype=r.headers.get("content-type", ""), bytes=len(r.content), links=links, sample=t[:600])

# 6) Fidelity FBCG
for u in ("https://digital.fidelity.com/prgw/digital/research/quote/dashboard/summary?symbol=FBCG",
          "https://fundresearch.fidelity.com/mutual-funds/composition/316092311",
          "https://www.actionsxchangerepository.fidelity.com/ShowDocument/ComplianceEnvelope/DailyHoldings_FBCG.csv"):
    r = fetch(u)
    t = r.text or ""
    rec("fidelity", r.status_code == 200, url=u, status=r.status_code, bytes=len(r.content), sample=re.sub(r"\s+", " ", t[:700]))

# 7) T. Rowe Price TCHP: look for holdings data in the ETF page
r = fetch("https://www.troweprice.com/financial-intermediary/us/en/investments/etfs/blue-chip-growth-etf.html")
t = r.text or ""
keys = sorted(set(re.findall(r'"([A-Za-z\-]*[Hh]olding[A-Za-z\-]*)"', t)))[:40]
apis = sorted(set(re.findall(r'(/[A-Za-z0-9_\-/\.]*(?:api|services|bin)[A-Za-z0-9_\-/\.]*)', t)))[:30]
rec("troweprice_page", r.status_code == 200, keys=keys, apis=apis)

# 8) Zacks full holdings (universal candidate)
for tk in ("CGGR", "CHAT", "FBCG", "TCHP", "CGDV"):
    r = fetch(f"https://www.zacks.com/funds/etf/{tk}/holding")
    t = r.text or ""
    m = re.search(r"etf_holdings\.formatted_data\s*=\s*(\[.*?\]);", t, re.S)
    n = len(json.loads(m.group(1))) if m else 0
    rec(f"zacks_{tk}", n > 25, status=r.status_code, rows=n, sample=(m.group(1)[:500] if m else t[:300]))

OUT.parent.mkdir(parents=True, exist_ok=True)
OUT.write_text(json.dumps(R, ensure_ascii=False, indent=1), encoding="utf-8")
print("wrote", OUT)
