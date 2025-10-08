# -*- coding: utf-8 -*-
"""
Daily 06:00 builder:
- Logs into CIS (HTTP Basic Auth popup) using Playwright http_credentials
- Navigates to Semesterpläne and downloads the XLSX (first .xlsx link on the page by default)
- Parses the workbook (sheets named by KW like '41', '42', ...)
- Builds a mobile-first index.html and an ICS feed
"""
import os, re, math
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


def login_and_download_xlsx():
    import re, time
    user = os.environ["CIS_USER"]
    pw = os.environ["CIS_PASS"]
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(
            accept_downloads=True,
            http_credentials={"username": user, "password": pw}
        )
        page = context.new_page()

        # 1) Start + Login (HTTP Basic)
        page.goto(CIS_LOGIN_URL, wait_until="domcontentloaded")
        page.goto(SEMESTERPLAENE_URL, wait_until="domcontentloaded")

        # 2) Ins Menü "Semesterpläne" – Fallback "LV-Plan"
        clicked = False
        try:
            page.get_by_role("link", name=re.compile(r"Semesterpl[aä]ne", re.I)).click(timeout=4000)
            page.wait_for_load_state("domcontentloaded"); clicked = True
        except Exception:
            pass
        if not clicked:
            try:
                page.get_by_role("link", name=re.compile(r"LV[- ]?Plan", re.I)).click(timeout=4000)
                page.wait_for_load_state("domcontentloaded"); clicked = True
            except Exception:
                pass

        # 3) Bereich „Semesterpläne Archiv“ – Dropdowns + Button
        try:
            ctx = page  # falls Frames nötig wären, hier ersetzen

            # Studiengang-Select finden
            prog_sel = None
            for selcss in [
                "select:has(option:has-text('Studiengang auswählen'))",
                "select[name*='studiengang']",
                "select"
            ]:
                loc = ctx.locator(selcss)
                if loc.count() > 0:
                    prog_sel = loc.first
                    break

            # Semester-Select finden
            sem_sel = None
            for selcss in [
                "select:has(option:has-text('Studiensemester auswählen'))",
                "select[name*='semester']",
                "select"
            ]:
                loc = ctx.locator(selcss)
                if loc.count() > 0:
                    sem_sel = loc.nth(1) if (prog_sel and loc.count() > 1) else loc.first
                    break

            import re as _re
            def select_by_text(select_loc, pattern):
                """Suche Option per Text (Regex) und wähle dann per value aus."""
                if not select_loc:
                    return False
                opts = select_loc.locator("option")
                n = opts.count()
                for i in range(n):
                    txt = (opts.nth(i).inner_text() or "").strip()
                    if _re.search(pattern, txt, _re.I):
                        val = opts.nth(i).get_attribute("value") or ""
                        if val:
                            select_loc.select_option(value=val)
                            return True
                return False

            ok_prog = select_by_text(prog_sel, r"(IKTF|IKT|Mil-IKTF)")
            ok_sem = select_by_text(sem_sel, r"^\s*1\s*$")

            # Button "Semesterplan laden" klicken
            clicked_btn = False
            for sel in [
                "button:has-text('Semesterplan laden')",
                "input[type='submit'][value*='Semesterplan']",
                "input[value*='Semesterplan']",
                "button:has-text('laden')"
            ]:
                try:
                    ctx.locator(sel).first.click(timeout=3000)
                    clicked_btn = True
                    break
                except Exception:
                    continue

            if not clicked_btn:
                print("[DEBUG] Konnte Button 'Semesterplan laden' nicht klicken – fahre fort.")

            page.wait_for_load_state("domcontentloaded")
            time.sleep(1)
        except Exception as e:
            print(f"[DEBUG] Auswahl Semesterpläne Archiv fehlgeschlagen: {e}")

        # 4) DEBUG: Links/Buttons listen
        links = page.locator("a")
        print(f"[DEBUG] Links nach Archiv-Laden: {links.count()}")
        for i in range(min(150, links.count())):
            try:
                el = links.nth(i)
                href = (el.get_attribute("href") or "").strip()
                txt = (el.inner_text() or "").strip().replace("\n", " ")
                if "xlsx" in href.lower() or "xls" in href.lower() or "excel" in txt.lower():
                    print(f"[DEBUG] Kandidat {i:03d}: text='{txt}' href='{href}'")
            except Exception:
                pass

        # 5) Excel-/Export-Link oder -Button suchen und klicken
        candidates = [
            "a[href$='.xlsx']",
            "a[href$='.xls']",
            "a:has-text('Excel')",
            "a:has-text('XLSX')",
            "button:has-text('Excel')",
            "button:has-text('Export')",
            "input[value*='Excel']",
            "input[value*='Export']",
        ]
        target = None
        for sel in candidates:
            loc = page.locator(sel)
            if loc.count() > 0:
                target = loc.first
                print(f"[DEBUG] Treffer: {sel}")
                break

        if not target:
            for i in range(links.count()):
                el = links.nth(i)
                href = (el.get_attribute("href") or "").lower()
                txt = (el.inner_text() or "").lower()
                if ".xlsx" in href or ".xls" in href or "excel" in txt or "export" in txt:
                    target = el
                    print(f"[DEBUG] Fallback-Link gewählt: text='{txt}' href='{href}'")
                    break

        if not target:
            raise RuntimeError("Kein Excel-/Export-Link gefunden. Selektor anpassen.")

        with page.expect_download() as dl:
            target.click()
        download = dl.value
        download.save_as(str(DOWNLOAD_XLSX_TO))

        context.close()
        browser.close()


def list_kw_sheets(xls: pd.ExcelFile):
    kws = []
    for s in xls.sheet_names:
        if re.fullmatch(r"\d{2}", str(s)) or re.fullmatch(r"\d{2}", str(s).zfill(2)):
            kws.append(s)
        elif re.fullmatch(r"\d{2}", str(s)[-2:]):
            kws.append(s)
    def keyf(k):
        try:
            return int(str(k)[-2:])
        except:
            return 999
    return sorted(kws, key=keyf)


def try_parse_time(cell) -> dt.time | None:
    if cell is None or (isinstance(cell, float) and math.isnan(cell)):
        return None
    if isinstance(cell, dt.time):
        return cell
    if isinstance(cell, dt.datetime):
        return dt.time(cell.hour, cell.minute)
    s = str(cell).strip()
    m = re.match(r"^(\d{1,2}):(\d{2})$", s)
    if m:
        hh, mm = int(m.group(1)), int(m.group(2))
        if 0 <= hh < 24 and 0 <= mm < 60:
            return dt.time(hh, mm)
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
                        try:
                            col_date[c] = dt.date(y, mo, d)
                        except:
                            pass
                        break
    return col_date


def parse_xlsx_to_events(xlsx: Path):
    xls = pd.ExcelFile(xlsx)
    events = []
    for sheet in list_kw_sheets(xls):
        df = pd.read_excel(xls, sheet_name=sheet, header=None)
        if df.empty:
            continue

        time_col = None
        for c in range(min(5, df.shape[1])):
            got = 0
            for r in range(min(200, len(df))):
                if try_parse_time(df.iat[r, c]):
                    got += 1
            if got > 5:
                time_col = c
                break
        if time_col is None:
            continue

        col_dates = extract_dates_from_header(df)
        day_cols = sorted(col_dates.keys())
        if not day_cols:
            continue

        start_row = None
        for r in range(len(df)):
            if try_parse_time(df.iat[r, time_col]):
                cnt = 0
                for k in range(r, min(r + 10, len(df))):
                    if try_parse_time(df.iat[k, time_col]):
                        cnt += 1
                if cnt >= 3:
                    start_row = r
                    break
        if start_row is None:
            continue

        r = start_row
        while r < len(df):
            t = try_parse_time(df.iat[r, time_col])
            if not t:
                r += 1
                continue
            start_time = t
            for c in day_cols:
                cell = str(df.iat[r, c]).strip()
                if not cell or cell.lower() in ("nan",):
                    continue
                rr = r + 1
                while rr < len(df):
                    t2 = try_parse_time(df.iat[rr, time_col])
                    if not t2:
                        break
                    cell2 = str(df.iat[rr, c]).strip()
                    if cell2 != cell:
                        break
                    rr += 1
                end_time = try_parse_time(df.iat[rr - 1, time_col])
                end_dt = dt.datetime.combine(col_dates[c], end_time) + dt.timedelta(minutes=5)
                start_dt = dt.datetime.combine(col_dates[c], start_time)
                title, lecturer, room = cell, "", ""
                parts = [p.strip() for p in re.split(r"\s*\|\s*|\n", cell) if p.strip()]
                if len(parts) >= 1:
                    title = parts[0]
                if len(parts) >= 2:
                    lecturer = parts[1]
                if len(parts) >= 3:
                    room = parts[2]
                events.append({
                    "title": title,
                    "lecturer": lecturer,
                    "room": room,
                    "start": TZ.localize(start_dt),
                    "end": TZ.localize(end_dt),
                })
            r += 1

    events = [e for e in events if e["end"] > e["start"]]
    uniq = {}
    for e in events:
        key = (e["title"], e["lecturer"], e["room"], e["start"], e["end"])
        uniq[key] = e
    return list(uniq.values())


def build_ics(events):
    cal = Calendar()
    for e in events:
        ev = Event()
        ev.name = e["title"]
        ev.begin = e["start"]
        ev.end = e["end"]
        desc = []
        if e["lecturer"]:
            desc.append(f"Dozent: {e['lecturer']}")
        if e["room"]:
            desc.append(f"Raum: {e['room']}")
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
    start_cut = now - dt.timedelta(days=7)
    end_cut = now + dt.timedelta(days=120)
    events = [e for e in events if (e["end"] >= start_cut and e["start"] <= end_cut)]
    build_ics(events)
    build_html(events)
    print(f"OK: {len(events)} Termine exportiert.")


if __name__ == "__main__":
    main()
