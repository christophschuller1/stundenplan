# -*- coding: utf-8 -*-
"""
Daily 06:00 builder:
- Logs into CIS (HTTP Basic Auth) using Playwright http_credentials
- Opens 'Semesterpläne' list and downloads the XLSX for '1. Semester … IKTF'
- Parses workbook (KW sheets) and builds HTML + ICS
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


def _pick_frames_by_url(page):
    """Erkenne Menü/Content-Frames robust über URL-Muster (Namen sind teils leer)."""
    menu_fr, content_fr = None, None
    for fr in page.frames:
        u = (fr.url or "").lower()
        if "/cis/menu.php" in u:
            menu_fr = fr
        if "/cms/" in u or "lvplan" in u or "stpl_" in u or "semester" in u:
            content_fr = fr
    return menu_fr, content_fr


def login_and_download_xlsx():
    import time
    user = os.environ["CIS_USER"]
    pw = os.environ["CIS_PASS"]

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(
            accept_downloads=True,
            http_credentials={"username": user, "password": pw}
        )
        page = context.new_page()

        # 1) Start + Login
        page.goto(CIS_LOGIN_URL, wait_until="domcontentloaded")
        page.goto(SEMESTERPLAENE_URL, wait_until="domcontentloaded")

        time.sleep(0.8)
        print(f"[DEBUG] Frames gefunden: {len(page.frames)}")
        for i, fr in enumerate(page.frames):
            print(f"[DEBUG] FRAME {i}: name='{fr.name}' url='{fr.url}'")

        # 2) Menü-/Content-Frames anhand URL wählen
        menu_fr, content_fr = _pick_frames_by_url(page)

        # 3) Im Menü „Semesterpläne“ klicken (Fallback LV-Plan)
        def _click_in_menu():
            if not menu_fr:
                return False
            for pat in [r"Semesterpl[aä]ne", r"LV[- ]?Plan"]:
                try:
                    menu_fr.get_by_role("link", name=re.compile(pat, re.I)).click(timeout=3000)
                    return True
                except Exception:
                    pass
            # Fallback: alle Links im Menü scannen
            try:
                links = menu_fr.locator("a")
                for i in range(min(150, links.count())):
                    el = links.nth(i)
                    txt = (el.inner_text() or "").strip()
                    if re.search(r"Semesterpl[aä]ne|LV[- ]?Plan", txt, re.I):
                        el.click()
                        return True
            except Exception:
                pass
            return False

        if not _click_in_menu():
            # Kein Menü? Versuche Top-Seite
            for pat in [r"Semesterpl[aä]ne", r"LV[- ]?Plan"]:
                try:
                    page.get_by_role("link", name=re.compile(pat, re.I)).click(timeout=3000)
                    break
                except Exception:
                    pass

        # Nach Klick Frames neu wählen
        time.sleep(0.8)
        menu_fr, content_fr = _pick_frames_by_url(page)
        ctx = content_fr if content_fr else page
        print(f"[DEBUG] Nutze Kontext: {'content-frame' if content_fr else 'page'} url='{getattr(ctx,'url','n/a')}'")

        # 4) Zieltext: 1. Semester … IKTF  (Umlaut tolerant)
        sem_text_re = re.compile(r"^\s*1\.\s*Semester.*(IKTF|IKT)\s*(\u00fc|ue)?\s*$", re.I)

        # 5) Versuch A: Excel-Icon im gleichen Eintrag
        download_done = False
        try:
            text_link = ctx.get_by_role("link", name=sem_text_re).first
            if text_link.count() > 0:
                row = text_link.locator("xpath=..")
                xlsx = row.locator("a[href$='.xlsx'], a[href$='.xls']")
                if xlsx.count() == 0:
                    row2 = row.locator("xpath=..")
                    xlsx = row2.locator("a[href$='.xlsx'], a[href$='.xls'], a:has(img[alt*='Excel'])")
                if xlsx.count() > 0:
                    print("[DEBUG] Klicke Excel-Icon …")
                    with page.expect_download() as dl:
                        xlsx.first.click()
                    dl.value.save_as(str(DOWNLOAD_XLSX_TO))
                    download_done = True
        except Exception as e:
            print(f"[DEBUG] Icon-Download fehlgeschlagen: {e}")

        # 6) Versuch B: Textlink direkt (Page-Download-Watcher!)
        if not download_done:
            try:
                link = ctx.get_by_role("link", name=sem_text_re).first
                if link.count() == 0:
                    raise RuntimeError("Textlink nicht gefunden.")
                print("[DEBUG] Klicke Textlink …")
                with page.expect_download() as dl:
                    link.click()
                dl.value.save_as(str(DOWNLOAD_XLSX_TO))
                download_done = True
            except Exception as e:
                print(f"[DEBUG] Textlink-Download fehlgeschlagen: {e}")

        # 7) Letzter Fallback: heuristisch Links durchsuchen
        if not download_done:
            links = ctx.locator("a")
            n = links.count()
            print(f"[DEBUG] Fallback: scanne {n} Links …")
            for i in range(min(400, n)):
                try:
                    el = links.nth(i)
                    href = (el.get_attribute("href") or "").lower()
                    txt = (el.inner_text() or "").strip()
                    if (".xlsx" in href or ".xls" in href) and re.search(r"1\.\s*Semester.*(ik|iktf)", txt, re.I):
                        print(f"[DEBUG] Fallback-Link: text='{txt}' href='{href}'")
                        with page.expect_download() as dl:
                            el.click()
                        dl.value.save_as(str(DOWNLOAD_XLSX_TO))
                        download_done = True
                        break
                except Exception:
                    pass

        if not download_done:
            raise RuntimeError("Kein Download für '1. Semester … IKTF' gefunden.")

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
                if not cell or cell.lower() == "nan":
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
                if len(parts) >= 1: title = parts[0]
                if len(parts) >= 2: lecturer = parts[1]
                if len(parts) >= 3: room = parts[2]
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
    start_cut = now - dt.timedelta(days=7)
    end_cut = now + dt.timedelta(days=120)
    events = [e for e in events if (e["end"] >= start_cut and e["start"] <= end_cut)]
    build_ics(events)
    build_html(events)
    print(f"OK: {len(events)} Termine exportiert.")


if __name__ == "__main__":
    main()
