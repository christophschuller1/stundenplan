# -*- coding: utf-8 -*-
"""
CIS Stundenplan Builder (06:00 Europe/Vienna)
- Login via HTTP Basic (Playwright http_credentials)
- Öffnet Semesterpläne und lädt XLSX für "1. Semester … IKTF(ü)"
- Fallback: direkter Download-Link (dms.php?id=848)
- Parsed KW-Sheets, baut HTML + ICS
"""
import os, re, math, time
from pathlib import Path
import datetime as dt

import pandas as pd
from ics import Calendar, Event
from pytz import timezone
from jinja2 import Template
from playwright.sync_api import sync_playwright

TZ = timezone("Europe/Vienna")
BASE = Path(__file__).parent.resolve()
PUBLIC = BASE / "public"
PUBLIC.mkdir(exist_ok=True)

CIS_LOGIN_URL = "https://cis.miles.ac.at/cis/"
SEMESTERPLAENE_URL = "https://cis.miles.ac.at/cis/index.php"
DOWNLOAD_XLSX_TO = BASE / "latest.xlsx"

# <<<< Falls sich die ID ändert, hier anpassen >>>>
FALLBACK_XLSX_URL = "https://cis.miles.ac.at/cms/dms.php?id=848"


# --------------------- Playwright: Login + Download --------------------- #

def _pick_frames_by_url(page):
    menu_fr, content_fr = None, None
    for fr in page.frames:
        u = (fr.url or "").lower()
        if "/cis/menu.php" in u:
            menu_fr = fr
        if "/cms/" in u:
            content_fr = fr
    return menu_fr, content_fr


def _click_semesterplaene_anywhere(page):
    """Klicke Semesterpläne (href enthält 'semester' oder Text 'Semesterpläne') – im Menü-Frame oder auf der Seite."""
    menu_fr, _ = _pick_frames_by_url(page)

    # im Menü-Frame versuchen
    if menu_fr:
        links = menu_fr.locator("a")
        for i in range(links.count()):
            el = links.nth(i)
            href = (el.get_attribute("href") or "").lower()
            if "semester" in href:
                try:
                    el.click(timeout=2500)
                    return True
                except Exception:
                    pass
        for pat in [r"Semesterpl[aä]ne", r"LV[- ]?Plan"]:
            try:
                menu_fr.get_by_role("link", name=re.compile(pat, re.I)).click(timeout=2500)
                return True
            except Exception:
                pass

    # auf der Hauptseite versuchen
    links = page.locator("a")
    for i in range(links.count()):
        el = links.nth(i)
        href = (el.get_attribute("href") or "").lower()
        if "semester" in href:
            try:
                el.click(timeout=2500)
                return True
            except Exception:
                pass
    for pat in [r"Semesterpl[aä]ne", r"LV[- ]?Plan"]:
        try:
            page.get_by_role("link", name=re.compile(pat, re.I)).click(timeout=2500)
            return True
        except Exception:
            pass
    return False


def _find_xlsx_link(ctx):
    """
    Suche einen Link für '1. Semester … IKTF(ü)' ODER beliebige dms.php?id=…-XLSX-Links.
    ctx = content-frame oder page
    """
    # 1) Textlink "1. Semester … IKTF(ü)"
    sem_text_re = re.compile(r"^\s*1\.\s*Semester.*IKTF(\u00fc|ue)?\s*$", re.I)
    try:
        t = ctx.get_by_role("link", name=sem_text_re).first
        if t.count() > 0:
            # Excel-Icon im Umfeld
            row = t.locator("xpath=..")
            xlsx = row.locator("a[href$='.xlsx'], a[href$='.xls']")
            if xlsx.count() == 0:
                row2 = row.locator("xpath=..")
                xlsx = row2.locator("a[href$='.xlsx'], a[href$='.xls'], a:has(img[alt*='Excel'])")
            if xlsx.count() > 0:
                return xlsx.first
            # evtl. ist der Textlink selbst die Datei
            return t
    except Exception:
        pass

    # 2) Generisch: dms.php?id=… Links
    try:
        cand = ctx.locator("a[href*='dms.php?id=']")
        if cand.count() > 0:
            # Bevorzuge solche, deren sichtbarer Text '1. Semester' und 'IKTF' enthält
            for i in range(min(200, cand.count())):
                el = cand.nth(i)
                txt = (el.inner_text() or "").strip()
                if re.search(r"1\.\s*Semester", txt, re.I) and re.search(r"IKTF", txt, re.I):
                    return el
            return cand.first
    except Exception:
        pass

    # 3) Jede .xlsx/.xls Datei
    try:
        cand = ctx.locator("a[href$='.xlsx'], a[href$='.xls']")
        if cand.count() > 0:
            return cand.first
    except Exception:
        pass

    return None


def login_and_download_xlsx():
    user = os.environ["CIS_USER"]
    pw = os.environ["CIS_PASS"]

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(
            accept_downloads=True,
            http_credentials={"username": user, "password": pw}
        )
        page = context.new_page()

        # Login
        page.goto(CIS_LOGIN_URL, wait_until="domcontentloaded")
        page.goto(SEMESTERPLAENE_URL, wait_until="domcontentloaded")
        time.sleep(0.6)

        # Versuche, auf "Semesterpläne" zu navigieren (wenn es eine Übersichtsseite gibt)
        _click_semesterplaene_anywhere(page)
        page.wait_for_load_state("domcontentloaded")
        time.sleep(0.6)

        # Content-Kontext (Frame oder Page)
        _, content_fr = _pick_frames_by_url(page)
        ctx = content_fr if content_fr else page

        # Versuche Link automatisch zu finden
        target = _find_xlsx_link(ctx)

        # Fallback: direkter Link (ID)
        if not target:
            print("[DEBUG] Kein Link gefunden – nutze Fallback dms.php?id=848")
            with page.expect_download() as dl:
                # direkter GET auf die Datei-URL
                page.goto(FALLBACK_XLSX_URL, wait_until="domcontentloaded")
            dl.value.save_as(str(DOWNLOAD_XLSX_TO))
        else:
            print("[DEBUG] Klicke gefundenen Download-Link …")
            with page.expect_download() as dl:
                target.click()
            dl.value.save_as(str(DOWNLOAD_XLSX_TO))

        context.close()
        browser.close()


# --------------------- XLSX -> Events --------------------- #

def list_kw_sheets(xls: pd.ExcelFile):
    kws = []
    for s in xls.sheet_names:
        if re.fullmatch(r"\d{2}", str(s)) or re.fullmatch(r"\d{2}", str(s).zfill(2)):
            kws.append(s)
        elif re.fullmatch(r"\d{2}", str(s)[-2:]):
            kws.append(s)
    def keyf(k):
        try: return int(str(k)[-2:])
        except: return 999
    return sorted(kws, key=keyf)


def try_parse_time(cell) -> dt.time | None:
    if cell is None or (isinstance(cell, float) and math.isnan(cell)): return None
    if isinstance(cell, dt.time): return cell
    if isinstance(cell, dt.datetime): return dt.time(cell.hour, cell.minute)
    s = str(cell).strip()
    m = re.match(r"^(\d{1,2}):(\d{2})$", s)
    if m:
        hh, mm = int(m.group(1)), int(m.group(2))
        if 0 <= hh < 24 and 0 <= mm < 60: return dt.time(hh, mm)
    return None


def extract_dates_from_header(df):
    weekday_re = re.compile(r"^(Montag|Dienstag|Mittwoch|Donnerstag|Freitag|Samstag|Sonntag)$", re.I)
    date_re = re.compile(r"(\d{1,2})[.\s](\d{1,2})[.\s](\d{4})")
    col_date = {}
    for r in range(0, min(30, len(df))):
        for c in range(0, df.shape[1]):
            val = str(df.iat[r, c]).strip()
            if weekday_re.match(val):
                for look in range(1, 5):
                    rr = r + look
                    if rr >= len(df): break
                    s2 = str(df.iat[rr, c]).strip()
                    m = date_re.search(s2)
                    if m:
                        d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
                        try: col_date[c] = dt.date(y, mo, d)
                        except: pass
                        break
    return col_date


def parse_xlsx_to_events(xlsx: Path):
    xls = pd.ExcelFile(xlsx)
    events = []
    for sheet in list_kw_sheets(xls):
        df = pd.read_excel(xls, sheet_name=sheet, header=None)
        if df.empty: continue

        # 1) Zeitspalte
        time_col = None
        for c in range(min(5, df.shape[1])):
            got = sum(1 for r in range(min(200, len(df))) if try_parse_time(df.iat[r, c]))
            if got > 5:
                time_col = c; break
        if time_col is None: continue

        # 2) Datumsspalten (Tage)
        col_dates = extract_dates_from_header(df)
        day_cols = sorted(col_dates.keys())
        if not day_cols: continue

        # 3) Start der Zeitraster
        start_row = None
        for r in range(len(df)):
            if try_parse_time(df.iat[r, time_col]):
                cnt = sum(1 for k in range(r, min(r+10, len(df))) if try_parse_time(df.iat[k, time_col]))
                if cnt >= 3: start_row = r; break
        if start_row is None: continue

        # 4) Slots -> Events
        r = start_row
        while r < len(df):
            t = try_parse_time(df.iat[r, time_col])
            if not t:
                r += 1
                continue
            start_time = t
            for c in day_cols:
                cell = str(df.iat[r, c]).strip()
                if not cell or cell.lower() == "nan":
                    continue
                rr = r + 1
                while rr < len(df):
                    t2 = try_parse_time(df.iat[rr, time_col])
                    if not t2: break
                    if str(df.iat[rr, c]).strip() != cell: break
                    rr += 1
                end_time = try_parse_time(df.iat[rr-1, time_col])
                end_dt = dt.datetime.combine(col_dates[c], end_time) + dt.timedelta(minutes=5)
                start_dt = dt.datetime.combine(col_dates[c], start_time)
                title, lecturer, room = cell, "", ""
                parts = [p.strip() for p in re.split(r"\s*\|\s*|\n", cell) if p.strip()]
                if len(parts)>=1: title = parts[0]
                if len(parts)>=2: lecturer = parts[1]
                if len(parts)>=3: room = parts[2]
                events.append({
                    "title": title, "lecturer": lecturer, "room": room,
                    "start": TZ.localize(start_dt), "end": TZ.localize(end_dt)
                })
            r += 1

    # Duplikate raus
    events = [e for e in events if e["end"] > e["start"]]
    uniq = {(e["title"], e["lecturer"], e["room"], e["start"], e["end"]): e for e in events}
    return list(uniq.values())


# --------------------- Exporte --------------------- #

def build_ics(events):
    cal = Calendar()
    for e in events:
        ev = Event()
        ev.name = e["title"]; ev.begin = e["start"]; ev.end = e["end"]
        desc = []
        if e["lecturer"]: desc.append(f"Dozent: {e['lecturer']}")
        if e["room"]: desc.append(f"Raum: {e['room']}")
        ev.description = "\n".join(desc) if desc else ""
        ev.location = e["room"] or ""
        cal.events.add(ev)
    (PUBLIC / "stundenplan.ics").write_text(str(cal), encoding="utf-8")


def build_html(events):
    tmpl = Template("""
<!doctype html>
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Stundenplan</title>
<style>
  body{font-family:system-ui,-apple-system,Segoe UI,Roboto,Arial,sans-serif;margin:16px}
  .day{margin:12px 0;padding:12px;border:1px solid #ddd;border-radius:12px}
  .ev{padding:8px 10px;border-radius:10px;border:1px solid #eee;margin:8px 0}
  .t{font-weight:600}
  .sub{color:#555;font-size:0.9rem}
  .hdr{display:flex;justify-content:space-between;align-items:center;margin-bottom:8px}
  .badge{font-size:0.8rem;background:#f3f3f3;border-radius:999px;padding:2px 8px}
</style>
<div class="hdr">
  <h1>Stundenplan</h1>
  <a class="badge" href="stundenplan.ics">Kalender abonnieren (ICS)</a>
</div>
<div>Automatisch aktualisiert täglich 06:00 (Europe/Vienna).</div>
{% for day, items in days %}
  <div class="day">
    <div class="t">{{ day.strftime("%A, %d.%m.%Y") }}</div>
    {% for e in items %}
      <div class="ev">
        <div class="t">{{ e.title }}</div>
        <div class="sub">{{ e.start.strftime("%H:%M") }}–{{ e.end.strftime("%H:%M") }}{% if e.room %} | {{ e.room }}{% endif %}{% if e.lecturer %} | {{ e.lecturer }}{% endif %}</div>
      </div>
    {% endfor %}
  </div>
{% endfor %}
""")
    by_day = {}
    for e in events:
        d = e["start"].date()
        by_day.setdefault(d, []).append(e)
    days = []
    for d, items in by_day.items():
        items.sort(key=lambda x: x["start"])
        days.append((d, items))
    days.sort(key=lambda x: x[0])
    html = tmpl.render(days=days)
    (PUBLIC / "index.html").write_text(html, encoding="utf-8")


def main():
    login_and_download_xlsx()
    events = parse_xlsx_to_events(DOWNLOAD_XLSX_TO)
    now = TZ.localize(dt.datetime.now())
    events = [e for e in events if (e["end"] >= now - dt.timedelta(days=7) and e["start"] <= now + dt.timedelta(days=120))]
    build_ics(events)
    build_html(events)
    print(f"OK: {len(events)} Termine exportiert.")


if __name__ == "__main__":
    main()
