#!/usr/bin/env python3
"""
Sinokor/Heung-A schedule API fetcher
=====================================
Pulls vessel rotation data directly from the backend that powers
e-heunga.com's public "Schedule (Vessel)" search, and writes it out in
the same column shape generate_vessel_schedule.py expects from a SKED.xls
export. This is what lets the whole pipeline run without anyone manually
forwarding a SKED file: a scheduled job runs this, then feeds its output
straight into `generate_vessel_schedule.py --full`.

ENDPOINT
--------
    POST https://ebizapi.sinokor.co.kr/Schedule/Vessel
    body: {"compcd": "HALK", "nacd": "", "token": "", "vsl": "<code>", "month": "YYYYMM"}

Found by inspecting e-heunga.com's own network calls while using its
vessel-schedule search -- it is NOT a documented/published integration,
takes no auth token (the site itself sends an empty one), and could
change or disappear without notice. If this script starts failing,
that's the first thing to check -- fall back to a manual SKED.xls in
the meantime.

USAGE
-----
    # fetch every vessel code currently in sked_current.xlsx
    python fetch_sinokor_schedule.py --out input/auto-fetch.xlsx

    # fetch specific vessels only
    python fetch_sinokor_schedule.py --vessels KMBK,SWSR --out input/auto-fetch.xlsx

To track a vessel this script has never seen before (not yet in
sked_current.xlsx), add its code to input/vessel_codes.json, e.g.:
    ["KMBK", "SWSR", "NEWCODE"]

REQUIRES: pandas, openpyxl, requests
    pip install pandas openpyxl requests
"""
import argparse
import json
import sys
import time
from datetime import datetime
from pathlib import Path

import pandas as pd
import requests

SCRIPT_DIR = Path(__file__).resolve().parent
SKED_CURRENT_PATH = SCRIPT_DIR / "sked_current.xlsx"
VESSEL_CODES_PATH = SCRIPT_DIR / "input" / "vessel_codes.json"

API_URL = "https://ebizapi.sinokor.co.kr/Schedule/Vessel"
COMPANY_CODE = "HALK"  # Heung-A Line Korea, as sent by e-heunga.com itself

SKED_COLUMNS = [
    "Service", "Vessel", "Vessel Name", "Op.Liner", "Vyg", "Bound", "Vyg Bound",
    "Wharf", "POL", "POD", "ETA Date", "ETB Date", "ETD Date",
]


def known_vessel_codes():
    """Vessel codes to track: every code in the current canonical dataset,
    plus anything listed in input/vessel_codes.json (for a vessel that
    hasn't shown up in a SKED file yet but we already know the code for)."""
    codes = set()
    if SKED_CURRENT_PATH.exists():
        df = pd.read_excel(SKED_CURRENT_PATH)
        codes.update(str(c) for c in df["Vessel"].dropna().unique().tolist())
    if VESSEL_CODES_PATH.exists():
        codes.update(json.loads(VESSEL_CODES_PATH.read_text(encoding="utf-8")))
    return sorted(codes)


def month_strings(months_forward=2, months_back=0):
    """['202609', '202610', ...] from months_back before today through
    months_forward after, inclusive of the current month."""
    now = datetime.now()
    y, m = now.year, now.month
    m -= months_back
    while m <= 0:
        m += 12
        y -= 1
    out = []
    for _ in range(months_back + months_forward + 1):
        out.append(f"{y}{m:02d}")
        m += 1
        if m > 12:
            m = 1
            y += 1
    return out


def fetch_vessel_month(session, vsl, month):
    resp = session.post(API_URL, json={
        "compcd": COMPANY_CODE, "nacd": "", "token": "", "vsl": vsl, "month": month,
    }, timeout=20)
    resp.raise_for_status()
    return resp.json()


def split_vyg(vyg):
    """'2610S' -> ('2610', 'S'); '' / None -> (None, None)."""
    if not vyg:
        return None, None
    return vyg[:-1], vyg[-1]


def build_rows(records):
    rows = []
    for r in records:
        vyg, bound = split_vyg(r.get("VYG"))
        rows.append({
            "Service": r.get("SVC"),
            "Vessel": r.get("VSL"),
            "Vessel Name": r.get("VSLNM"),
            "Op.Liner": "",  # not exposed by this endpoint
            "Vyg": vyg,
            "Bound": bound,
            "Vyg Bound": r.get("VYG"),
            "Wharf": r.get("WHARF"),
            "POL": r.get("PORT"),
            "POD": None,  # filled in below, once sorted chronologically
            "ETA Date": r.get("ETA"),
            "ETB Date": r.get("ETB"),
            "ETD Date": r.get("ETD"),
        })
    return rows


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--out", required=True, help="Output .xlsx path, SKED-format")
    ap.add_argument("--vessels", help="Comma-separated vessel codes (default: every code in sked_current.xlsx + input/vessel_codes.json)")
    ap.add_argument("--months-forward", type=int, default=2)
    ap.add_argument("--months-back", type=int, default=0)
    ap.add_argument("--sleep", type=float, default=0.3, help="Seconds between API calls")
    ap.add_argument("--min-vessels-ratio", type=float, default=0.6,
                    help="Abort (nonzero exit) if fewer than this fraction of the requested vessels "
                         "returned any data -- a scheduled job should fail loudly on a bad fetch "
                         "rather than push a half-empty dashboard live (default: 0.6)")
    args = ap.parse_args()

    vessels = [v.strip() for v in args.vessels.split(",") if v.strip()] if args.vessels else known_vessel_codes()
    if not vessels:
        sys.exit("[error] no vessel codes to fetch -- pass --vessels or make sure sked_current.xlsx exists")

    months = month_strings(args.months_forward, args.months_back)
    print(f"[1/3] Fetching {len(vessels)} vessel(s) x {len(months)} month(s) ({', '.join(months)}) from {API_URL} ...")

    session = requests.Session()
    session.headers.update({"Content-Type": "application/json", "Accept": "application/json"})

    per_vessel_records = {}
    failures = []
    for vsl in vessels:
        records = []
        for month in months:
            try:
                records.extend(fetch_vessel_month(session, vsl, month))
            except Exception as e:
                failures.append((vsl, month, str(e)))
            time.sleep(args.sleep)
        per_vessel_records[vsl] = records
        print(f"       {vsl}: {len(records)} call(s)")

    if failures:
        print(f"[warn] {len(failures)} request(s) failed:")
        for vsl, month, err in failures:
            print(f"       {vsl} {month}: {err}")

    vessels_with_data = [v for v, recs in per_vessel_records.items() if recs]
    ratio = len(vessels_with_data) / len(vessels)
    if ratio < args.min_vessels_ratio:
        sys.exit(f"[error] only {len(vessels_with_data)}/{len(vessels)} vessels returned data "
                 f"({ratio:.0%}, below --min-vessels-ratio {args.min_vessels_ratio:.0%}) -- "
                 f"treating this as a bad fetch (API issue?) and refusing to write output")

    print("[2/3] Deriving POD (next port) per vessel and assembling rows ...")
    all_rows = []
    for vsl, records in per_vessel_records.items():
        if not records:
            continue
        seen = set()
        uniq = []
        for r in records:
            key = (r.get("SVC"), r.get("VSL"), r.get("VYG"), r.get("PORT"), r.get("ETA"))
            if key in seen:
                continue
            seen.add(key)
            uniq.append(r)
        uniq.sort(key=lambda r: r.get("ETA") or "")
        rows = build_rows(uniq)
        # POD = the next *different* port visited -- a vessel can appear
        # twice at the same call under two service codes (e.g. KHS1 and
        # KTS1 sharing a leg), so naively taking row i+1 sometimes points
        # right back at the port we're already at.
        for i, row in enumerate(rows):
            pod = None
            for j in range(i + 1, len(rows)):
                if rows[j]["POL"] != row["POL"]:
                    pod = rows[j]["POL"]
                    break
            row["POD"] = pod
        all_rows.extend(rows)

    if not all_rows:
        sys.exit("[error] no schedule data returned for any vessel -- API may be down, or every vessel code is stale")

    df = pd.DataFrame(all_rows, columns=SKED_COLUMNS)
    for c in ["ETA Date", "ETB Date", "ETD Date"]:
        df[c] = pd.to_datetime(df[c], errors="coerce")
    df = df.dropna(subset=["ETA Date"]).sort_values(["Vessel Name", "ETA Date"])

    out_path = Path(args.out)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(out_path, index=False)
    print(f"[3/3] Wrote {len(df)} row(s), {df['Vessel Name'].nunique()} vessel(s) -> {out_path}")


if __name__ == "__main__":
    main()
