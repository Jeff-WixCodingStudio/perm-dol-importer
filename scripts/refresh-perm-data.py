import json
import re
import zipfile
from collections import Counter, defaultdict
from datetime import datetime, timedelta
from html import unescape
from pathlib import Path
from tempfile import NamedTemporaryFile
from urllib.request import Request, urlopen
from xml.etree.ElementTree import iterparse

PERFORMANCE_PAGE = "https://www.dol.gov/agencies/eta/foreign-labor/performance"
OUTPUT_PATH = Path(__file__).resolve().parents[1] / "public" / "perm-dashboard.json"
NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
HEADERS = {
    "User-Agent": "PERM-Dashboard-Importer/1.0 (public data dashboard)",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Referer": "https://www.dol.gov/",
}


def fetch_bytes(url):
    with urlopen(Request(url, headers=HEADERS), timeout=180) as response:
        return response.read()


def newest_workbook_url(html):
    urls = [unescape(match.group(1)) for match in re.finditer(r'href=["\']([^"\']+\.xlsx(?:\?[^"\']*)?)["\']', html, re.I)]
    urls = [url if url.startswith("http") else "https://www.dol.gov" + url for url in urls]
    urls = [url for url in urls if "PERM" in url.upper() and "record_layout" not in url.lower()]
    return sorted(set(urls), key=lambda url: int(re.search(r"FY(20\d{2})", url, re.I).group(1)) if re.search(r"FY(20\d{2})", url, re.I) else 0)[-1]


def shared_strings(archive):
    values = []
    with archive.open("xl/sharedStrings.xml") as stream:
        for event, element in iterparse(stream, events=("end",)):
            if element.tag == f"{NS}si":
                values.append("".join(text.text or "" for text in element.iter(f"{NS}t")))
                element.clear()
    return values


def column_index(reference):
    letters = re.match(r"[A-Z]+", reference).group(0)
    result = 0
    for letter in letters:
        result = result * 26 + ord(letter) - 64
    return result - 1


def cell_value(cell, strings):
    value = cell.find(f"{NS}v")
    raw = "" if value is None else value.text or ""
    if cell.get("t") == "s":
        return strings[int(raw)] if raw else ""
    if cell.get("t") == "inlineStr":
        return "".join(text.text or "" for text in cell.iter(f"{NS}t"))
    return raw


def normalize(value):
    return re.sub(r"[^A-Z0-9]", "_", str(value or "").upper())


def parse_date(value):
    if not value:
        return None
    try:
        if re.fullmatch(r"\d+(\.\d+)?", str(value)):
            return datetime(1899, 12, 30) + timedelta(days=float(value))
        return datetime.fromisoformat(str(value).replace("Z", "+00:00")).replace(tzinfo=None)
    except ValueError:
        for pattern in ("%m/%d/%Y", "%m/%d/%y", "%Y-%m-%d"):
            try:
                return datetime.strptime(str(value), pattern)
            except ValueError:
                pass
    return None


def empty_state():
    return {"daily": Counter(), "weekly": Counter(), "monthly": Counter(), "weekday": Counter(), "letters": Counter(), "month_letters": defaultdict(Counter), "latest": None}


def add_record(state, date, status, employer):
    state["latest"] = date if state["latest"] is None or date > state["latest"] else state["latest"]
    state["daily"][date.date().isoformat()] += 1
    monday = date.date() - timedelta(days=date.weekday())
    state["weekly"][monday.isoformat()] += 1
    state["monthly"][date.strftime("%Y-%m")] += 1
    state["weekday"][date.weekday()] += 1
    letter = employer.strip()[:1].upper() if employer.strip()[:1].isalpha() else "X"
    state["letters"][letter] += 1
    state["month_letters"][date.strftime("%Y-%m")][letter] += 1


def build_series(state):
    latest = state["latest"] or datetime(1970, 1, 1)
    daily_dates = [latest - timedelta(days=29 - index) for index in range(30)]
    month_dates = [datetime(latest.year, latest.month, 1)]
    for _ in range(11):
        current = month_dates[0]
        month_dates.insert(0, datetime(current.year - (1 if current.month == 1 else 0), 12 if current.month == 1 else current.month - 1, 1))
    latest_month = latest.strftime("%Y-%m")
    weekly_starts = [latest.date() - timedelta(days=latest.weekday() + 21 - index * 7) for index in range(4)]
    return {
        "daily": [state["daily"][date.date().isoformat()] for date in daily_dates],
        "dailyLabels": [date.strftime("%b %-d") for date in daily_dates],
        "weekly": [state["weekday"][day] for day in range(7)],
        "weeklyLabels": ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"],
        "weeklyVolume": [state["weekly"][date.isoformat()] for date in weekly_starts],
        "weeklyVolumeLabels": [date.strftime("%b %-d") for date in weekly_starts],
        "monthlyVolume": [state["monthly"][date.strftime("%Y-%m")] for date in month_dates],
        "monthlyVolumeLabels": [date.strftime("%b %Y") for date in month_dates],
        "activity": [state["letters"][chr(65 + index)] for index in range(26)],
        "activeLetters": [state["month_letters"][latest_month][chr(65 + index)] for index in range(26)],
        "detailedAsOf": latest.isoformat() + "Z",
    }


def parse_workbook(path):
    processed = empty_state()
    certified = empty_state()
    with zipfile.ZipFile(path) as archive:
        strings = shared_strings(archive)
        with archive.open("xl/worksheets/sheet1.xml") as stream:
            header = None
            for event, row in iterparse(stream, events=("end",)):
                if row.tag != f"{NS}row":
                    continue
                values = {column_index(cell.get("r")): cell_value(cell, strings) for cell in row.findall(f"{NS}c")}
                if header is None:
                    candidate = {normalize(value): index for index, value in values.items()}
                    decision_column = next((candidate.get(name) for name in ("DECISION_DATE", "DATE_OF_DECISION", "CASE_DECISION_DATE", "DETERMINATION_DATE") if candidate.get(name) is not None), None)
                    if decision_column is not None:
                        header = (decision_column, candidate.get("CASE_STATUS", candidate.get("STATUS")), candidate.get("EMP_BUSINESS_NAME", candidate.get("EMPLOYER_NAME")))
                elif header:
                    date = parse_date(values.get(header[0]))
                    if date:
                        status = str(values.get(header[1], "")).upper()
                        employer = str(values.get(header[2], ""))
                        add_record(processed, date, status, employer)
                        if "CERTIFIED" in status or "CERTIFICATION" in status:
                            add_record(certified, date, status, employer)
                row.clear()
    if processed["latest"] is None:
        raise RuntimeError("The DOL workbook did not have decision rows with a usable DECISION_DATE column.")
    return processed, certified


html = fetch_bytes(PERFORMANCE_PAGE).decode("utf-8", errors="replace")
workbook_url = newest_workbook_url(html)
if not workbook_url:
    raise RuntimeError("No PERM disclosure workbook was found on the DOL performance page.")
with NamedTemporaryFile(suffix=".xlsx") as temporary:
    temporary.write(fetch_bytes(workbook_url))
    temporary.flush()
    processed, certified = parse_workbook(temporary.name)

OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
OUTPUT_PATH.write_text(json.dumps({
    "generatedAt": datetime.utcnow().isoformat() + "Z",
    "sourceUrl": workbook_url,
    "sourceRelease": workbook_url.rsplit("/", 1)[-1],
    "chartSeries": {"Certified": build_series(certified), "Processed": build_series(processed)},
}) + "\n")
print(json.dumps({"sourceUrl": workbook_url, "latestDecision": processed["latest"].isoformat()}))
