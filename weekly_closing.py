"""
weekly_closing.py - "THAI Weekly closing time" calendar for THBKK / THLCH.

One row per week (Sun-Sat), one card per Thai port call on the day of its
ETB: ETB / ETD, DRY and RF closing, and a red/blue tag when the call moved
since the baseline snapshot. A "Changes since ..." table sits above the
calendar. Styles are inline (Calibri / MsoNormal) so the page pastes into
an Outlook mail intact.

Closing rules (same as the earlier weekly closing sheets):
    THBKK  DRY = ETB - 12h   RF = ETB - 6h
    THLCH  DRY = ETB - 24h   RF = ETB - 12h

Usage:
    python weekly_closing.py                         # baseline = 24h ago
    python weekly_closing.py --since 2026-10-21T15:30   # last issue's time
"""
import argparse
import html
import json
from datetime import datetime, timedelta

import pandas as pd

import generate_vessel_schedule as g

CUT_HOURS = {"THBKK": (12, 6), "THLCH": (24, 12)}
WEEKS = 5
NAVY, LINE, HEAD_BG, GREY = "#203764", "#9BB7D4", "#DCE6F1", "#595959"
RED, BLUE, GREEN = "#C00000", "#2E5A94", "#2E7D32"
PORT_STYLE = {  # card background, left border / tag colour
    "THBKK": ("#EEF2F8", NAVY),
    "THLCH": ("#E6F3F2", "#13817D"),
}
FONT = "font-family:Calibri,sans-serif"
DOW = ["SUN", "MON", "TUE", "WED", "THU", "FRI", "SAT"]

esc = html.escape


def span(text, size, color=NAVY, bold=False, extra=""):
    t = f"<b>{text}</b>" if bold else text
    return f'<span style="font-size:{size}pt;{FONT};color:{color}{extra}">{t}</span>'


def d_hm(t):
    return f"{t.day}/{t:%H%M}"


def delta_text(h):
    return f"{h:+g}h".replace("+-", "-") if abs(h) < 48 else f"{h / 24:+.1f}d"


def cutoffs(port, etb):
    dry, rf = CUT_HOURS[port]
    return etb - timedelta(hours=dry), etb - timedelta(hours=rf)


def th_calls(df):
    d = df[df["POL"].isin(g.TH_PORTS)].dropna(subset=["Vessel Name", "Vyg Bound", "POL"])
    d = d.sort_values("ETA Date").drop_duplicates(["Vessel Name", "Vyg Bound", "POL"])
    d = d.assign(ETB=d["ETB Date"].fillna(d["ETA Date"]), ETD=d["ETD Date"].fillna(d["ETA Date"]))
    return d


def change_tag(ch):
    if ch is None:
        return ""
    if ch["kind"] == "new":
        return span("NEW", 8.5, GREEN, True)
    color = RED if ch["kind"] == "delayed" else BLUE
    return span(f'{ch["kind"].upper()} {delta_text(ch["deltaHours"])}', 8.5, color, True)


def card(r, ch):
    port = r["POL"]
    bg, accent = PORT_STYLE[port]
    dry, rf = cutoffs(port, r["ETB"])
    tag = (f'<span style="font-size:8pt;{FONT};color:white;background:{accent};'
           f'padding:0 3px;mso-highlight:{accent}"><b>{port[2:]}</b></span>')
    key = " ".join(str(x) for x in (r["Service"], r["Vessel"], r["Vyg Bound"],
                                    r["Vessel Name"], port[2:])).lower()
    lines = [
        f'{span(esc(str(r["Vessel"]) + " " + str(r["Vyg Bound"])), 10.5, bold=True)} {tag} {span(esc(str(r["Service"])), 8, GREY)}',
        span(esc(r["Vessel Name"]), 8, GREY),
        span(f'ETB {r["ETB"]:%H%M} &middot; ETD {d_hm(r["ETD"])}', 9.5),
    ]
    if ch is not None:
        lines.append(change_tag(ch))
    lines.append(span(f"DRY {d_hm(dry)} &middot; RF {d_hm(rf)}", 9.5, RED))
    return (f'<div class=v data-k="{esc(key)}" style="background:{bg};border-left:solid {accent} 3.0pt;'
            f'padding:3pt 5pt;margin:0 0 4pt;white-space:nowrap">' + "<br>".join(lines) + "</div>")


def changes_table(rows, since_label):
    th = lambda t: (f'<td style="background:{HEAD_BG};border:solid {LINE} 1.0pt;padding:2pt 6pt">'
                    f'<p class=MsoNormal style="margin:0">{span(t, 9.5, bold=True)}</p></td>')
    td = lambda inner: (f'<td style="border:solid {LINE} 1.0pt;padding:2pt 6pt;white-space:nowrap">'
                        f'<p class=MsoNormal style="margin:0">{inner}</p></td>')

    def pair(old, new, color=NAVY):
        o = d_hm(old) if old is not None else None
        n = d_hm(new) if new is not None else None
        if n is None:
            return span(f'<s>{o}</s>', 9.5, "#7F7F7F")
        if o is None or o == n:
            return span(n, 9.5, color, True)
        return span(f'<span style="color:#7F7F7F"><s>{o}</s></span> &rarr; <b>{n}</b>', 9.5, color)

    title = (f'<p class=MsoNormal style="margin:0 0 3px">'
             f'{span(f"Changes since the last update {since_label}", 12, bold=True)}</p>')
    if not rows:
        return title + f'<p class=MsoNormal style="margin:0 0 12px">{span("No ETB / ETD changes.", 9.5, GREY)}</p>'
    out = [title, '<table border=0 cellspacing=0 cellpadding=0 style="border-collapse:collapse;margin:0 0 12px"><tr>',
           "".join(th(t) for t in ("PORT", "VESSEL / VOY", "CHANGE", "ETB", "ETD", "DRY", "RF")), "</tr>"]
    for r in rows:
        ts = lambda k: pd.Timestamp(r[k]) if r.get(k) else None
        o_etb, n_etb, o_etd, n_etd = ts("oldEtb"), ts("newEtb"), ts("oldEtd"), ts("newEtd")
        if r["kind"] == "removed":
            kind = span("REMOVED", 9.5, "#7F7F7F", True)
            n_etb = n_etd = None
        else:
            kind = change_tag(r).replace("8.5pt", "9.5pt")
        o_cut = cutoffs(r["port"], o_etb) if o_etb is not None else (None, None)
        n_cut = cutoffs(r["port"], n_etb) if n_etb is not None else (None, None)
        out.append("<tr>" + "".join([
            td(span(r["port"][2:], 9.5, bold=True)),
            td(span(f'<b>{esc(r["code"])} {esc(r["vygBound"])}</b> {esc(r["vessel"])}', 9.5)),
            td(kind),
            td(pair(o_etb, n_etb)), td(pair(o_etd, n_etd)),
            td(pair(o_cut[0], n_cut[0], RED)), td(pair(o_cut[1], n_cut[1], RED)),
        ]) + "</tr>")
    out.append("</table>")
    return "".join(out)


def calendar(calls, changes, today, start):
    cell = (f'border:solid {LINE} 1.0pt;padding:3pt 4pt;vertical-align:top')
    out = [f'<table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0 style="border-collapse:collapse;width:100%;min-width:1100px">',
           f'<tr><td width=52 style="{cell};background:{HEAD_BG}"></td>']
    out += [f'<td width=150 style="{cell};background:{HEAD_BG};text-align:center">{span(d, 10, bold=True)}</td>' for d in DOW]
    out.append("</tr>")
    by_day = {}
    for _, r in calls.iterrows():
        by_day.setdefault(r["ETB"].date(), []).append(r)
    for w in range(WEEKS):
        wk_start = start + timedelta(days=7 * w)
        wk_no = (wk_start + timedelta(days=1)).isocalendar()[1]
        out.append(f'<tr><td style="{cell};background:{NAVY};text-align:center;vertical-align:middle">'
                   f'{span("WK", 10, "white", True)}<br>{span(wk_no, 12, "white", True)}</td>')
        for i in range(7):
            day = wk_start + timedelta(days=i)
            is_today, past = day == today, day < today
            bg = "#FFF6D5" if is_today else ("#F7F7F7" if past else "white")
            if is_today:
                head = (f'<span style="font-size:11pt;{FONT};color:white;background:{RED};padding:0 2px"><b>{day.day}</b></span> '
                        + span(f'{day:%b} <b>TODAY</b>', 8.5, RED))
            elif day.day == 1 or (w == 0 and i == 0):
                head = span(f"<b>{day.day}</b> {day:%b}", 11) if day.day != 1 else span(f'<b>1</b> <span style="font-size:8.5pt;color:{GREY}">{day:%b}</span>', 11)
            else:
                head = span(day.day, 11, bold=True)
            cards = sorted(by_day.get(day, []), key=lambda r: r["ETB"])
            body = "".join(card(r, changes.get((r["Vessel Name"], r["Vyg Bound"], r["POL"]))) for r in cards)
            out.append(f'<td style="{cell};background:{bg};min-width:150px"><div style="margin:0 0 3pt">{head}</div>{body}</td>')
        out.append("</tr>")
    out.append("</table>")
    return "".join(out)


PAGE_CSS = """
body{margin:0;background:#f4f6fa;font-family:Calibri,'Segoe UI',Arial,sans-serif;color:#203764}
header{background:#203764;color:#fff;padding:12px 20px;display:flex;align-items:center;gap:14px}
header .logo{width:52px;height:52px;border-radius:50%;background:#fff;flex:none;object-fit:contain}
header h1{margin:0;font-size:19px;letter-spacing:.3px}
header p{margin:3px 0 0;font-size:12.5px;opacity:.85}
.bar{position:sticky;top:0;z-index:5;background:#f4f6fa;padding:10px 20px;display:flex;flex-wrap:wrap;gap:8px;align-items:center;border-bottom:1px solid #dde3ec}
#q{flex:1 1 280px;max-width:420px;font-size:14px;padding:8px 12px;border:1px solid #9BB7D4;border-radius:8px;outline:none}
#q:focus{border-color:#203764;box-shadow:0 0 0 3px rgba(32,55,100,.12)}
.chip{font-size:12.5px;padding:5px 11px;border:1px solid #9BB7D4;border-radius:14px;background:#fff;color:#203764;cursor:pointer}
.chip.on{background:#203764;color:#fff;border-color:#203764}
#cnt{font-size:12.5px;color:#667085;margin-left:4px}
main{padding:14px 20px 30px;overflow-x:auto}
main table{background:#fff}
.v.hit{outline:2px solid #F2C200}
.v.off{display:none}
.foot{font-size:12px;color:#667085;padding:0 20px 24px}
@media (max-width:640px){header,.bar,main,.foot{padding-left:12px;padding-right:12px}}
"""

PAGE_JS = """
const q=document.getElementById('q'),cnt=document.getElementById('cnt');let port='';
function run(){const t=q.value.trim().toLowerCase().split(/\\s+/).filter(Boolean);let n=0,tot=0;
 document.querySelectorAll('.v').forEach(v=>{tot++;const k=v.dataset.k,
  ok=t.every(w=>k.includes(w))&&(!port||k.split(' ').includes(port));
  v.classList.toggle('off',!ok);v.classList.toggle('hit',ok&&t.length>0);if(ok)n++;});
 cnt.textContent=(t.length||port)?n+' / '+tot+' calls':tot+' calls';}
q.addEventListener('input',run);
document.querySelectorAll('.chip[data-p]').forEach(c=>c.onclick=()=>{port=c.dataset.p;
 document.querySelectorAll('.chip[data-p]').forEach(x=>x.classList.toggle('on',x===c));run();});
document.getElementById('clr').onclick=()=>{q.value='';run();q.focus();};
document.addEventListener('keydown',ev=>{if(ev.key==='/'&&document.activeElement!==q){ev.preventDefault();q.focus();}});
run();
"""


FOOT_RULES = ("Times D/HHMM; a card sits on the day of its ETB. DRY = dry closing, RF = reefer closing. "
              "BKK: DRY ETB-12h / RF ETB-6h &middot; LCH: DRY ETB-24h / RF ETB-12h.")


def build(df_new, now, since=None):
    """Changes table + calendar as one inline-styled HTML fragment (the
    dashboard's Weekly closing view embeds it; main() wraps it as a page).
    Returns (fragment, info) with info = calls / rows / since_label / services."""
    df_new = g.only_th_vessels(df_new)
    since = since or now - timedelta(hours=24)
    snap = g.snapshot_before((now - since).total_seconds() / 3600, now)
    if snap:
        rows = g.build_call_changes(g.load_sked(snap[1]), df_new, now)
        since_label = f"{snap[0]:%d/%m %H:%M}"
    else:
        rows, since_label = [], "(no earlier snapshot)"
    today = now.date()
    start = today - timedelta(days=(today.weekday() + 1) % 7)
    end = start + timedelta(days=7 * WEEKS)
    # Only calls that land in the calendar's weeks
    in_window = lambda iso: iso is not None and start <= datetime.fromisoformat(iso).date() < end
    rows = [r for r in rows if in_window(r["newEtb"]) or in_window(r["oldEtb"])]
    changes = {(r["vessel"], r["vygBound"], r["port"]): r for r in rows if r["kind"] != "removed"}
    calls = th_calls(df_new)
    calls = calls[(calls["ETB"].dt.date >= start) & (calls["ETB"].dt.date < end)]
    services = ", ".join(dict.fromkeys(calls.sort_values("ETB")["Service"].astype(str)))
    fragment = f"<div class=chg>{changes_table(rows, since_label)}</div>{calendar(calls, changes, today, start)}"
    iso = lambda t: t.isoformat()
    data = {
        "start": start.isoformat(), "weeks": WEEKS, "sinceLabel": since_label, "cutHours": CUT_HOURS,
        "changes": rows,
        "calls": [{
            "code": str(r["Vessel"]), "vyg": str(r["Vyg Bound"]), "name": str(r["Vessel Name"]),
            "service": str(r["Service"]), "port": r["POL"], "etb": iso(r["ETB"]), "etd": iso(r["ETD"]),
            "change": (lambda c: c and {"kind": c["kind"], "deltaHours": c.get("deltaHours")})(
                changes.get((r["Vessel Name"], r["Vyg Bound"], r["POL"]))),
        } for _, r in calls.sort_values("ETB").iterrows()],
    }
    return fragment, {"calls": len(calls), "rows": len(rows), "since_label": since_label,
                      "services": services, "data": data}


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--since", help="Baseline time (Bangkok), e.g. the last issue: 2026-10-21T15:30 (default: 24h ago)")
    ap.add_argument("--out", help="Output HTML (default: Desktop/THAI Weekly closing time <date>.html)")
    args = ap.parse_args()

    meta = json.loads(g.SKED_META_PATH.read_text(encoding="utf-8"))
    updated = now = datetime.fromisoformat(meta["current"]["updated"])
    since = datetime.fromisoformat(args.since) if args.since else None
    fragment, info = build(g.load_sked(g.SKED_CURRENT_PATH), now, since)

    out = args.out or str(g.SCRIPT_DIR.parent / f"THAI Weekly closing time {now:%Y-%m-%d}.html")
    logo = g._data_uri(g.LOGO_PATH, "logo hidden")
    page = (
        '<!doctype html><html lang=en><head><meta charset="utf-8">'
        '<meta name=viewport content="width=device-width,initial-scale=1">'
        f"<title>Weekly Closing Time</title><style>{PAGE_CSS}</style></head><body>"
        f'<header>{f"<img class=logo src={chr(34)}{logo}{chr(34)} alt=HAL>" if logo else ""}<div>'
        "<h1>HEUNG A LINE (THAILAND) &middot; Weekly Schedule &amp; Closing Time</h1>"
        f"<p>Update {updated:%d/%m/%Y %H:%M} (Thailand time) &middot; services: {esc(info['services'])}</p></div></header>"
        '<div class=bar><input id=q type=search placeholder="Search vessel, voyage, name or service... (press /)" autocomplete=off>'
        '<button class=chip id=clr>Clear</button><button class="chip on" data-p="">All ports</button>'
        '<button class=chip data-p="bkk">BKK</button><button class=chip data-p="lch">LCH</button><span id=cnt></span></div>'
        f"<main>{fragment}</main>"
        f"<p class=foot>{FOOT_RULES} Red / blue tag = ETB change since {info['since_label']}.<br>"
        f"Source: Heung-A e-Service (e-heunga.com) vessel schedule, {updated:%d %b %H:%M}.</p>"
        f"<script>{PAGE_JS}</script></body></html>"
    )
    with open(out, "w", encoding="utf-8") as f:
        f.write(page)
    print(f"{info['calls']} calls, {info['rows']} change rows since {info['since_label']} -> {out}")


if __name__ == "__main__":
    main()
