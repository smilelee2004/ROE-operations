#!/usr/bin/env python3
# /// script
# requires-python = ">=3.10"
# dependencies = [
#     "openpyxl>=3.1",
# ]
# ///
"""
fetch_sec_xbrl.py
=================

Read a MonitorList.xlsx of US tickers (sheet "美股", column A = Ticker ID),
look each one up on the SEC, and save its XBRL data under <outdir>/<TICKER>/.

SEC JSON endpoints used (all free, no API key, but require a contact
User-Agent per https://www.sec.gov/os/accessing-edgar-data):

  1. https://www.sec.gov/files/company_tickers.json
     Ticker -> CIK mapping for common US-listed issuers (~10k).

  2. https://www.sec.gov/files/company_tickers_exchange.json
     Broader mapping — includes many ADRs (GSK, TM, NVO, ...). Used as
     a fallback when a ticker is missing from the first list.

  3. https://data.sec.gov/submissions/CIK{cik10}.json
     List of every filing by this issuer (accession, form, date, primary doc).
     We save this so you can pick the latest 10-K / 20-F / 10-Q later.

  4. https://data.sec.gov/api/xbrl/companyfacts/CIK{cik10}.json
     The issuer's full XBRL fact book aggregated across every filing
     (us-gaap + dei + entity-specific extensions). This is "XBRL data"
     in its most useful JSON form — one file, all history.

  5. (optional, --with-instance) the primary document of the latest
     10-K or 20-F: https://www.sec.gov/Archives/edgar/data/{cik}/{adsh_nodashes}/{primary_document}
     For post-2020 filings this is inline-XBRL HTML (.htm) that carries
     every XBRL tag inline. Useful when you want the raw source filing.

Output layout (default)::

    D:\work\SourceCode\ROE-operations\財報\
      RTX/
        submissions.json
        companyfacts.json
        latest_10K.htm          (if --with-instance, optional)
        meta.json               (what we fetched, when, filing accession)
      DOW/
        ...
      _fetch_log.txt            (one-line-per-ticker audit log)

Rate limit: 10 requests/second per SEC policy. The script sleeps 110ms
between calls and retries 429/5xx with exponential backoff.

Unsponsored ADRs (CABGY, CLPBY, HNNMY, RHHBY, ...) do not file with
the SEC, so they won't appear in either ticker list. The script logs
them as NOT_REGISTERED and continues.

Usage::

    pip install openpyxl
    python fetch_sec_xbrl.py \
        --list MonitorList.xlsx \
        --email "you@example.com"
    # -> writes to D:\work\SourceCode\ROE-operations\財報\ by default

    # Override the output folder if you want somewhere else:
    python fetch_sec_xbrl.py --list MonitorList.xlsx \
        --outdir D:\other\path\財報 --email you@example.com

    # Also grab each issuer's latest 10-K / 20-F primary document:
    python fetch_sec_xbrl.py --list MonitorList.xlsx \
        --email you@example.com --with-instance

    # Only (re)fetch a subset:
    python fetch_sec_xbrl.py --list MonitorList.xlsx \
        --email you@example.com --tickers TM,NFLX,ABNB

Re-running is safe: by default existing files are skipped; pass --force
to overwrite.
"""
from __future__ import annotations

import argparse
import gzip
import io
import json
import re
import sys
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable, Optional
import urllib.error
import urllib.request

# ------------------------------------------------------------------------
# SEC endpoints & policy
# ------------------------------------------------------------------------
SEC_TICKERS_URL          = "https://www.sec.gov/files/company_tickers.json"
SEC_TICKERS_EXCHANGE_URL = "https://www.sec.gov/files/company_tickers_exchange.json"
SEC_SUBMISSIONS_URL      = "https://data.sec.gov/submissions/CIK{cik10}.json"
SEC_COMPANYFACTS_URL     = "https://data.sec.gov/api/xbrl/companyfacts/CIK{cik10}.json"
SEC_ARCHIVE_URL          = "https://www.sec.gov/Archives/edgar/data/{cik}/{adsh_no_dashes}/{doc}"

# SEC asks for ≤10 req/sec from a single client. 110ms is safely under.
MIN_DELAY_SEC = 0.11

# Retry policy for 429 / 5xx
MAX_RETRIES   = 5
BACKOFF_BASE  = 1.5  # 1.5, 2.25, 3.38, 5.06, 7.59 seconds


# ------------------------------------------------------------------------
# HTTP with throttling + retry (stdlib only — avoids requests dependency)
# ------------------------------------------------------------------------
class SecClient:
    def __init__(self, user_agent: str):
        if not user_agent or "@" not in user_agent:
            raise ValueError(
                "SEC requires a descriptive User-Agent with contact email, "
                'e.g. "Smile Home smilelee@example.com"'
            )
        self.ua = user_agent
        self._last_call = 0.0

    def _sleep_if_needed(self) -> None:
        delta = time.monotonic() - self._last_call
        if delta < MIN_DELAY_SEC:
            time.sleep(MIN_DELAY_SEC - delta)

    def _request(self, url: str) -> bytes:
        """GET with throttling + retry. Returns decoded body bytes."""
        last_err: Optional[Exception] = None
        for attempt in range(MAX_RETRIES):
            self._sleep_if_needed()
            req = urllib.request.Request(
                url,
                headers={
                    "User-Agent":      self.ua,
                    "Accept":          "application/json, text/html, */*",
                    "Accept-Encoding": "gzip, deflate",
                    "Host":            url.split("/")[2],
                },
            )
            try:
                with urllib.request.urlopen(req, timeout=30) as resp:
                    self._last_call = time.monotonic()
                    raw = resp.read()
                    if resp.headers.get("Content-Encoding") == "gzip":
                        raw = gzip.decompress(raw)
                    return raw
            except urllib.error.HTTPError as e:
                self._last_call = time.monotonic()
                last_err = e
                # 404 is usually terminal (not a real ticker or no filings).
                if e.code == 404:
                    raise
                if e.code in (429, 500, 502, 503, 504):
                    wait = BACKOFF_BASE ** attempt
                    print(f"    HTTP {e.code} on {url}; retry in {wait:.1f}s "
                          f"({attempt+1}/{MAX_RETRIES})", file=sys.stderr)
                    time.sleep(wait)
                    continue
                raise
            except urllib.error.URLError as e:
                last_err = e
                wait = BACKOFF_BASE ** attempt
                print(f"    network error {e.reason}; retry in {wait:.1f}s "
                      f"({attempt+1}/{MAX_RETRIES})", file=sys.stderr)
                time.sleep(wait)
                continue
        raise RuntimeError(f"exhausted retries for {url}: {last_err}")

    def get_json(self, url: str) -> dict:
        return json.loads(self._request(url).decode("utf-8"))

    def get_bytes(self, url: str) -> bytes:
        return self._request(url)


# ------------------------------------------------------------------------
# Ticker → CIK resolution
# ------------------------------------------------------------------------
@dataclass
class TickerHit:
    cik:      int
    name:     str
    ticker:   str
    exchange: str = ""
    source:   str = ""          # "company_tickers" | "exchange"


def build_ticker_index(client: SecClient) -> dict[str, TickerHit]:
    """
    Download both SEC ticker lists, return {UPPERCASE_TICKER: TickerHit}.
    The 'exchange' list wins over the primary list when they overlap
    because it carries exchange info.
    """
    index: dict[str, TickerHit] = {}

    # Primary list
    print("  loading company_tickers.json ...", file=sys.stderr)
    primary = client.get_json(SEC_TICKERS_URL)
    # format: {"0": {"cik_str": 320193, "ticker": "AAPL", "title": "Apple Inc."}, ...}
    for row in primary.values():
        t = str(row["ticker"]).upper()
        index[t] = TickerHit(
            cik=int(row["cik_str"]),
            name=row["title"],
            ticker=t,
            source="company_tickers",
        )

    # Exchange list (broader; includes many ADRs)
    print("  loading company_tickers_exchange.json ...", file=sys.stderr)
    ex = client.get_json(SEC_TICKERS_EXCHANGE_URL)
    # format: {"fields": ["cik","name","ticker","exchange"], "data": [[...],...]}
    cols = ex["fields"]
    ci = cols.index("cik")
    ni = cols.index("name")
    ti = cols.index("ticker")
    xi = cols.index("exchange")
    for row in ex["data"]:
        t = str(row[ti] or "").upper()
        if not t:
            continue
        index[t] = TickerHit(
            cik=int(row[ci]),
            name=str(row[ni]),
            ticker=t,
            exchange=str(row[xi] or ""),
            source="exchange",
        )
    return index


# ------------------------------------------------------------------------
# Monitor list reader
# ------------------------------------------------------------------------
def read_monitor_list(path: Path) -> list[str]:
    """
    Read tickers from MonitorList.xlsx. Accepts:
      - Sheet name '美股' (primary), or the first sheet as fallback.
      - Any column whose header is 'Ticker ID' / 'Ticker' / 'Symbol';
        defaults to column A.
    Returns a de-duplicated uppercase list, preserving first-seen order.
    """
    try:
        from openpyxl import load_workbook
    except ImportError:
        sys.exit("openpyxl is required: pip install openpyxl")

    wb = load_workbook(path, data_only=True, read_only=True)
    ws = wb["美股"] if "美股" in wb.sheetnames else wb[wb.sheetnames[0]]

    # Try to detect a header row
    first = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    header_labels = {"ticker id", "ticker", "symbol", "代號", "股票代號"}
    col_idx = 1
    start_row = 1
    if first and any(str(v).strip().lower() in header_labels for v in first if v):
        for c, v in enumerate(first, start=1):
            if v and str(v).strip().lower() in header_labels:
                col_idx = c
                break
        start_row = 2

    out, seen = [], set()
    for r in range(start_row, ws.max_row + 1):
        v = ws.cell(r, col_idx).value
        if v is None:
            continue
        t = str(v).strip().upper()
        if not t or t in seen:
            continue
        seen.add(t)
        out.append(t)
    return out


# ------------------------------------------------------------------------
# Per-ticker fetch
# ------------------------------------------------------------------------
@dataclass
class FetchResult:
    ticker:       str
    status:       str            # OK / NOT_REGISTERED / HTTP_ERROR / SKIPPED
    cik:          Optional[int] = None
    name:         str = ""
    form_latest:  str = ""       # latest 10-K / 20-F form type
    accession:    str = ""
    period_end:   str = ""
    messages:     list[str] = field(default_factory=list)

    def log_line(self) -> str:
        bits = [self.ticker.ljust(6), self.status.ljust(16)]
        if self.cik:
            bits.append(f"CIK={self.cik:<10d}")
        if self.form_latest:
            bits.append(f"{self.form_latest:<5s} {self.period_end}")
        if self.name:
            bits.append(self.name[:60])
        if self.messages:
            bits.append("| " + " ; ".join(self.messages))
        return " ".join(bits)


def _cik10(cik: int) -> str:
    return f"{cik:010d}"


def _save_json(path: Path, obj: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(obj, ensure_ascii=False, indent=2), encoding="utf-8")


def _find_latest_annual(submissions: dict) -> Optional[dict]:
    """Return the most recent 10-K or 20-F filing record from submissions.recent."""
    recent = submissions.get("filings", {}).get("recent", {})
    forms  = recent.get("form", [])
    if not forms:
        return None
    adsh   = recent.get("accessionNumber", [])
    prim   = recent.get("primaryDocument", [])
    fdate  = recent.get("filingDate", [])
    period = recent.get("reportDate", [])
    for i, f in enumerate(forms):
        if f in ("10-K", "20-F", "40-F"):
            return {
                "form":             f,
                "accessionNumber":  adsh[i]   if i < len(adsh)   else "",
                "primaryDocument":  prim[i]   if i < len(prim)   else "",
                "filingDate":       fdate[i]  if i < len(fdate)  else "",
                "reportDate":       period[i] if i < len(period) else "",
            }
    return None


def fetch_one(
    client:        SecClient,
    ticker:        str,
    hit:           Optional[TickerHit],
    outdir:        Path,
    with_instance: bool,
    force:         bool,
) -> FetchResult:
    fr = FetchResult(ticker=ticker, status="OK")

    if hit is None:
        fr.status = "NOT_REGISTERED"
        fr.messages.append(
            "not in SEC company_tickers (unsponsored ADR / OTC / non-US issuer)"
        )
        return fr

    fr.cik  = hit.cik
    fr.name = hit.name
    tdir = outdir / ticker
    tdir.mkdir(parents=True, exist_ok=True)
    cik10 = _cik10(hit.cik)

    # 1) submissions.json ---------------------------------------------------
    subs_path = tdir / "submissions.json"
    if subs_path.exists() and not force:
        submissions = json.loads(subs_path.read_text(encoding="utf-8"))
        fr.messages.append("submissions.json cached")
    else:
        try:
            submissions = client.get_json(SEC_SUBMISSIONS_URL.format(cik10=cik10))
            _save_json(subs_path, submissions)
        except urllib.error.HTTPError as e:
            fr.status = "HTTP_ERROR"
            fr.messages.append(f"submissions: HTTP {e.code}")
            return fr

    # 2) companyfacts.json (the XBRL fact book) -----------------------------
    cf_path = tdir / "companyfacts.json"
    if cf_path.exists() and not force:
        fr.messages.append("companyfacts.json cached")
    else:
        try:
            cf = client.get_json(SEC_COMPANYFACTS_URL.format(cik10=cik10))
            _save_json(cf_path, cf)
        except urllib.error.HTTPError as e:
            # Some issuers (recent IPOs, foreign filers with no XBRL yet) 404 here.
            fr.messages.append(f"companyfacts: HTTP {e.code} (issuer may not file XBRL yet)")

    # 3) Latest 10-K / 20-F metadata + optional instance --------------------
    latest = _find_latest_annual(submissions)
    if latest:
        fr.form_latest = latest["form"]
        fr.accession   = latest["accessionNumber"]
        fr.period_end  = latest["reportDate"]
        if with_instance and latest["primaryDocument"]:
            adsh_nd = latest["accessionNumber"].replace("-", "")
            suffix  = Path(latest["primaryDocument"]).suffix or ".htm"
            doc_path = tdir / f"latest_{latest['form'].replace('-', '')}{suffix}"
            if doc_path.exists() and not force:
                fr.messages.append(f"{doc_path.name} cached")
            else:
                url = SEC_ARCHIVE_URL.format(
                    cik=hit.cik,
                    adsh_no_dashes=adsh_nd,
                    doc=latest["primaryDocument"],
                )
                try:
                    data = client.get_bytes(url)
                    doc_path.write_bytes(data)
                    fr.messages.append(f"saved {doc_path.name} ({len(data):,} B)")
                except urllib.error.HTTPError as e:
                    fr.messages.append(f"instance: HTTP {e.code}")
    else:
        fr.messages.append("no 10-K / 20-F in recent filings")

    # 4) meta.json — audit breadcrumb --------------------------------------
    meta = {
        "ticker":            ticker,
        "cik":               hit.cik,
        "name":              hit.name,
        "exchange":          hit.exchange,
        "resolver":          hit.source,
        "fetched_utc":       time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
        "latest_annual":     latest or {},
        "files": {
            "submissions":   "submissions.json",
            "companyfacts":  "companyfacts.json" if cf_path.exists() else None,
        },
    }
    _save_json(tdir / "meta.json", meta)
    return fr


# ------------------------------------------------------------------------
# main
# ------------------------------------------------------------------------
def main() -> int:
    ap = argparse.ArgumentParser(
        description=__doc__.split("\n\n")[0],
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    ap.add_argument("--list",   type=Path, required=True,
                    help="Path to MonitorList.xlsx")
    ap.add_argument("--outdir", type=Path,
                    default=Path(r"D:\work\SourceCode\ROE-operations\財報"),
                    help=r"Root folder to save per-ticker data "
                         r"(default: D:\work\SourceCode\ROE-operations\財報)")
    ap.add_argument("--email",  type=str, required=True,
                    help='Contact email for SEC User-Agent (required by SEC). '
                         'Example: "Smile Home smilelee@example.com"')
    ap.add_argument("--tickers", type=str, default="",
                    help="Comma-separated subset to fetch (default: all in the list)")
    ap.add_argument("--with-instance", action="store_true",
                    help="Also download the latest 10-K/20-F primary document "
                         "(inline XBRL .htm)")
    ap.add_argument("--force", action="store_true",
                    help="Overwrite existing files (default: skip if present)")
    args = ap.parse_args()

    if not args.list.exists():
        sys.exit(f"--list file not found: {args.list}")

    tickers_all = read_monitor_list(args.list)
    if args.tickers:
        wanted = {t.strip().upper() for t in args.tickers.split(",") if t.strip()}
        tickers = [t for t in tickers_all if t in wanted]
        missing = wanted - set(tickers)
        if missing:
            print(f"Warning: {sorted(missing)} not in MonitorList", file=sys.stderr)
    else:
        tickers = tickers_all

    print(f"Processing {len(tickers)} tickers from {args.list.name} "
          f"-> {args.outdir}/", flush=True)

    args.outdir.mkdir(parents=True, exist_ok=True)
    user_agent = args.email if " " in args.email else f"MonitorList-Fetcher {args.email}"
    client = SecClient(user_agent)

    # Build ticker index once
    print("Building SEC ticker index ...", flush=True)
    index = build_ticker_index(client)
    print(f"  {len(index):,} tickers indexed", flush=True)

    # Fetch each ticker
    results: list[FetchResult] = []
    for i, t in enumerate(tickers, start=1):
        hit = index.get(t)
        print(f"[{i:>3}/{len(tickers)}] {t:<6s}  "
              f"{'CIK='+str(hit.cik) if hit else 'not in SEC lists'}",
              flush=True)
        try:
            fr = fetch_one(client, t, hit, args.outdir,
                           with_instance=args.with_instance, force=args.force)
        except Exception as e:
            fr = FetchResult(ticker=t, status="HTTP_ERROR",
                             cik=(hit.cik if hit else None),
                             name=(hit.name if hit else ""))
            fr.messages.append(f"{type(e).__name__}: {e}")
        results.append(fr)
        print("      " + fr.log_line(), flush=True)

    # Write audit log
    log_path = args.outdir / "_fetch_log.txt"
    with log_path.open("w", encoding="utf-8") as f:
        f.write(f"# fetch_sec_xbrl.py run at {time.strftime('%Y-%m-%d %H:%M:%S %Z')}\n")
        f.write(f"# source: {args.list}\n")
        f.write(f"# count : {len(results)}\n")
        for r in results:
            f.write(r.log_line() + "\n")

    # Summary
    n_ok       = sum(1 for r in results if r.status == "OK")
    n_notreg   = sum(1 for r in results if r.status == "NOT_REGISTERED")
    n_err      = sum(1 for r in results if r.status == "HTTP_ERROR")
    print("")
    print(f"Done: {n_ok} OK, {n_notreg} NOT_REGISTERED, {n_err} HTTP_ERROR")
    print(f"Log : {log_path}")
    return 0 if n_err == 0 else 2


if __name__ == "__main__":
    sys.exit(main())
