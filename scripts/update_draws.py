"""Add new Eurojackpot draws to data.json.

Runs in GitHub Actions. Retries a slow source and leaves data unchanged if it stays unreachable. Reads the yearly results archive on euro-jackpot.net,
checks every overlapping draw against data.json, and appends only draws newer
than the latest one already stored. Nothing is written if the source disagrees
with existing data or cannot be read.
"""
import datetime as dt
import json
import re
import sys
import time
import urllib.request
from pathlib import Path

DATA = Path(__file__).resolve().parent.parent / "data.json"
URL = "https://www.euro-jackpot.net/results-archive-{year}"
ROW = re.compile(r'href="/results/(\d{2})-(\d{2})-(\d{4})".*?<ul[^>]*class="[^"]*balls[^"]*"[^>]*>(.*?)</ul>', re.S)
BALL = re.compile(r'<li[^>]*class="[^"]*\bball\b[^"]*"[^>]*>\s*<span>\s*(\d{1,2})\s*</span>')
EURO = re.compile(r'<li[^>]*class="[^"]*\beuro\b[^"]*"[^>]*>\s*<span>\s*(\d{1,2})\s*</span>')


HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/129.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-GB,en;q=0.9",
}


def fetch(year: int) -> str:
    """Fetch one archive page, retrying slow or refused requests."""
    last = None
    for attempt in range(1, 5):
        try:
            req = urllib.request.Request(URL.format(year=year), headers=HEADERS)
            with urllib.request.urlopen(req, timeout=60) as r:
                return r.read().decode("utf-8", "replace")
        except Exception as e:
            last = e
            print(f"Forsøg {attempt} for {year} mislykkedes: {e}")
            time.sleep(15 * attempt)
    raise RuntimeError(last)


def parse(html: str) -> dict:
    out = {}
    for d, m, y, ul in ROW.findall(html):
        main = sorted(int(x) for x in BALL.findall(ul))
        star = sorted(int(x) for x in EURO.findall(ul))
        out[f"{y}-{m}-{d}"] = (main, star)
    return out


def valid(date: str, main: list, star: list) -> bool:
    wd = dt.date.fromisoformat(date).weekday()  # Mon=0 … Tue=1, Fri=4
    max_star = 8 if date <= "2014-10-03" else 10 if date <= "2022-03-18" else 12
    return (wd in (1, 4) and len(main) == 5 and len(set(main)) == 5 and all(1 <= n <= 50 for n in main)
            and len(star) == 2 and len(set(star)) == 2 and all(1 <= n <= max_star for n in star))


def main() -> int:
    doc = json.loads(DATA.read_text(encoding="utf-8"))
    have = {r[0]: (r[1:6], r[6:8]) for r in doc["draws"]}
    latest = max(have)
    today = dt.date.today()

    found = {}
    for year in range(int(latest[:4]), today.year + 1):
        try:
            found.update(parse(fetch(year)))
        except Exception as e:  # source unreachable: keep data as is, try again next run
            print(f"::warning::Kunne ikke hente {year} ({e}). Data er uændret; der prøves igen ved næste kørsel.")
            return 0
    if not found:
        print("Ingen trækninger fundet på kildesiden. Sidens opbygning kan være ændret.")
        return 1

    mismatch = [d for d, (m, s) in found.items() if d in have and (sorted(have[d][0]) != m or sorted(have[d][1]) != s)]
    overlap = sum(1 for d in found if d in have)
    if mismatch:
        print("Kilden afviger fra eksisterende data for: " + ", ".join(sorted(mismatch)) + ". Intet skrevet.")
        return 1

    new = sorted(d for d in found if d > latest)
    bad = [d for d in new if not valid(d, *found[d])]
    new = [d for d in new if d not in bad]
    for d in bad:
        print(f"Springer over {d}: ugyldige tal {found[d]}")

    doc["updated"] = dt.datetime.now(dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    for d in new:
        m, s = found[d]
        doc["draws"].append([d, *m, *s])
    doc["draws"].sort(key=lambda r: r[0], reverse=True)
    DATA.write_text(json.dumps(doc, separators=(",", ":")), encoding="utf-8")
    print(f"Kontrolleret {overlap} eksisterende trækninger mod kilden. Nye: " + (", ".join(new) if new else "ingen"))
    return 0


if __name__ == "__main__":
    sys.exit(main())
