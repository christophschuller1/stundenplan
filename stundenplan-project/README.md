# Stundenplan Auto-Update (Web + ICS)

**Was es macht**
- Lädt täglich um **06:00 (Europe/Vienna)** den aktuellen Stundenplan (Excel) vom CIS.
- Baut daraus eine **mobilfreundliche Website** (`index.html`) und einen **Kalender-Feed** (`stundenplan.ics`).
- Veröffentlicht beides über **GitHub Pages** (ein Link für alle).
- Keine Logins für die Klasse nötig. Nur einmaliges Einrichten.

**Voraussetzung**
- GitHub-Account: `christophschuller1`
- Repo **öffentlich**
- In GitHub → Settings → Pages: Source = `gh-pages` Branch
- Secrets setzen (Settings → Secrets → Actions):
  - `CIS_USER` = CIS-Benutzername
  - `CIS_PASS` = CIS-Passwort

**Ergebnis-Links (nach dem ersten Lauf)**
- Web: `https://christophschuller1.github.io/stundenplan-auto/`
- Kalender (ICS): `https://christophschuller1.github.io/stundenplan-auto/stundenplan.ics`

> Hinweis: Kalender-Apps aktualisieren abonniertes ICS in Intervallen (3–24h). Die **Webansicht** ist immer sofort aktuell.

## Einrichten (Einmalig, ~10 Minuten)

1) Neues **öffentliches** Repository `stundenplan-auto` bei GitHub anlegen.
2) Dieses Projekt in das Repo hochladen (Dateien in der Wurzel belassen).
3) In **Settings → Pages**: Branch `gh-pages` wählen (falls noch nicht vorhanden, entsteht beim ersten Workflow-Run).
4) In **Settings → Secrets → Actions** zwei Secrets anlegen:
   - `CIS_USER`
   - `CIS_PASS`
5) Workflow starten (Actions → "Build Stundenplan (06:00 Vienna)" → "Run workflow") oder auf 06:00 warten.

## Anpassung Download-Link (falls nötig)
Im Script `fetch_and_build.py` gibt es den Block `find_xlsx_link(page)`. Sollte der Excel-Link auf der CIS-Seite anders heißen/liegen, kann dort der Selektor/Text angepasst werden.
