# -*- coding: utf-8 -*-
"""
Daily 06:00 builder:
- Loggt am CIS via HTTP Basic (Playwright http_credentials) ein
- Öffnet die Seite „Semesterpläne“ (Link mit href enthält 'semester' – im Frame oder direkt auf der Seite)
- Lädt die XLSX für „1. Semester … IKTF(ü)“
- Parsed die KW-Sheets, baut HTML + ICS
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


# --------------------- Playwright: Login + Download --------------------- #

def _pick_frames_by_url(page):
    """Erkenne Menü/Content-Frames robust über URL (Frames sind unbenannt)."""
    menu_fr, content_fr = None, None
    for fr in page.frames:
        u = (fr.url or "").lower()
        if "/cis/menu.php" in u:           # linke Navi
            menu_fr = fr
        if "/cms/" in u:                   # Inhaltsframe
            content_fr = fr
    return menu_fr, content_fr


def _click_semesterplaene_anywhere(page):
    """
    Klicke einen Link zur Semesterpläne-Seite – egal ob im Menü-Frame oder direkt auf der Seite.
    Kriterium: href enthält 'semester' (robuster als reiner Text). Fallback: Text enthält 'Semesterpläne'.
    """
    menu_fr, content_fr = _pick_frames_by_url(page)

    # 1) Im Menü-Frame
    if menu_fr:
        links = menu_fr.locator("a")
        n = links.count()
        print(f"[DEBUG] Menü-Links: {n}")
        for i in range(min(60, n)):
            try:
                el = links.nth(i)
                href = (el.get_attribute("href") or "").strip()
                txt = (el.inner_text() or "").strip()
                print(f"[DEBUG] MENU[{i:02d}] txt='{txt}' href='{href}'")
            except Exception:
                pass
        # a) nach href
        for i in range(n):
            el = links.nth(i)
            href = (el.get_attribute("href") or "").lower()
            if "semester" in href:
                try:
                    el.click(timeout=3000)
                    print(f"[DEBUG] Menü: Klick auf href='{href}'")
                    return True
                except Exception:
                    pass
        # b) nach Text
        for pat in [r"Semesterpl[aä]ne", r"LV[- ]?Plan"]:
            try:
                menu_fr.get_by_role("link", name=re.compile(pat, re.I)).click(timeout=3000)
                print(f"[DEBUG] Menü: Klick via Text '{pat}'")
                return True
            except Exception:
                pass

    # 2) Auf der Hauptseite (kein oder leeres Menü)
    links = page.locator("a")
    n = links.count()
    for i in range(min(120, n)):
        el = links.nth(i)
        href = (el.get_attribute("href") or "").lower()
        if "semester" in href:
            try:
                el.click(timeout=3000)
                print(f"[DEBUG] Page: Klick auf href='{href}'")
                return True
            except Exception:
                pass
    for pat in [r"Semesterpl[aä]ne", r"LV[- ]?Plan"]:
        try:
            page.get_by_role("link", name=re.compile(pat, re.I)).click(timeout=3000)
            print(f"[DEBUG] Page: Klick via Text '{pat}'")
            return True
        except Exception:
            pass

    return False


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

        time.sleep(0.8)
        print(f"[DEBUG] Frames gefunden: {len(page.frames)}")
        for i, fr in enumerate(page.frames):
            print(f"[DEBUG] FRAME {i}: name='{fr.name}' url='{fr.url}'")

        # Auf Semesterpläne wechseln
        if not _click_semesterplaene_anywhere(page):
            print("[DEBUG] Konnte 'Semesterpläne' nicht klicken – versuche weiter auf der Seite.")
        page.wait_for_load_state("domcontentloaded")
        time.sleep(0.8)

        # Content-Kontext bestimmen
        menu_fr, content_fr = _pick_frames_by_url(page)
        ctx = content_fr if content_fr else page
        print(f"[DEBUG] Nutze Kontext: {'content-frame' if content_fr else 'page'} url='{getattr(ctx, 'url', 'n/a')}'")

        # Grobe Plausibilitätsprüfung
        all_links = ctx.locator("a")
        n_links = all_links.count()
        print(f"[DEBUG] Links auf Inhaltsseite: {n_links}")
        if n_links < 15:
            print("[DEBUG] WARN: Wenig Links – vermutlich noch nicht auf der Liste. (Siehe MENU[..])")

        # Ziel: „1. Semester … IKTF(ü)“
        sem_text_re = re.compile(r"^\s*1\.\s*Semester.*IKTF(\u00fc|ue)?\s*$", re.I)

        download_done = False

        # A) Excel-Icon neben dem Textlink
        try:
            tlink = ctx.get_by_role("link", name=sem_text_re).first
            if tlink.count() > 0:
                row = tlink.locator("xpath=..")
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

        # B) Textlink direkt (falls er selbst Download auslöst)
        if not download_done:
            try:
                link = ctx.get_by_role("link", name=sem_text_re).first
                if link.count() > 0:
                    print("[DEBUG] Klicke Textlink …")
                    with page.expect_download() as dl:
                        link.click()
                    dl.value.save_as(str(DOWNLOAD_XLSX_TO))
                    download_done = True
            except Exception as e:
                print(f"[DEBUG] Textlink-Download fehlgeschlagen: {e}")

        # C) Heuristik: finde Zeile mit „1. Semester“ & „IKTF“, dann .xlsx-Icon daneben
        if not download_done:
            print("[DEBUG] Heuristik: suche Zeile mit '1. Semester' + 'IKTF' …")
            for i in range(min(400, n_links)):
                try:
                    el = all_links.nth(i)
                    txt = (el.inner_text() or "").strip()
                    if re.search(r"1\.\s*Semester", txt, re.I) and re.search(r"IKTF", txt, re.I):
                        row = el.locator("xpath=..")
                        xlsx = row.locator("a[href$='.xlsx'], a[href$='.xls']")
                        if xlsx.count() == 0:
                            row2 = row.locator("xpath=..")
                            xlsx = row2.locator("a[href$='.xlsx'], a[href$='.xls']")
                        if xlsx.count() > 0:
                            print(f"[DEBUG] Fallback: klicke Icon neben '{txt}' …")
                            with page.expect_download() as dl:
                                xlsx.first.click()
                            dl.value.save_as(str(DOWNLOAD_XLSX_TO))
                            download_done = True
                            break
                except Exception:
                    pass

        if not download_done:
            raise RuntimeError("Kein Download für '1. Semester … IKTF' gefunden.")

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

        # 1) Zeitspalte finden
        time_col = None
        for c in range(min(5, df.shape[1])):
            got = sum(1 for r in range(min(200, len(df))) if try_parse_time(df.iat[r, c]))
            if got > 5:
                time_col = c; break
        if time_col is None: continue

        # 2) Datumsspalten (Tage) finden
        col_dates = extract_dates_from_header(df)
        day_cols = sorted(col_dates.keys())
        if not day_cols: continue

        # 3) Start der Zeitraster finden
        start_row = None
        for r in range(len(df)):
            if try_parse_time(df.iat[r, time_col]):
                cnt = sum(1 for k in range(r, min(r+10, len(df))) if try_parse_time(df.iat[k, time_col]))
                if cnt >= 3: start_row = r; break
        if start_row is None: continue

        # 4) Slots zu Events aggregieren
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

    # Duplikate entfernen
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
