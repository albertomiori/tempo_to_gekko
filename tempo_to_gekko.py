#!/usr/bin/env python3
"""Convert a Jira Tempo "Worklogs" .xls export into a plain-text list of
entries ready to copy into Gekko (which has no import, no date field and
no task field -- just hours + free-text description per line).

Usage:
    python3 tempo_to_gekko.py <input.xls> [output.txt]

Output format (one line per entry, hours padded to 5 chars):
    <hours>  <description>
with a "# <date>" header and a "Total: Xh" line per day, and a final
"Grand total: Xh" line.

Two kinds of worklogs:
  - "Issue" worklogs (everything except WORK_DESCRIPTION_ISSUES): rows
    are aggregated per (day, issue). Text is
    "Working on Jira - <KEY> - <summary>"
    ("Working on SN Ticket - ..." when the Issue Type is SN_ISSUE_TYPE; just "<summary>"
    for PLAIN_ISSUES, e.g. the Bootcamp).
  - "Meeting" worklogs (WORK_DESCRIPTION_ISSUES, e.g. EHEALTH-9360):
    every row is its own entry and the text is the Work Description
    ("Working on issue ..." / blank becomes "Daily Meeting").

Merging rules (a Gekko line must be at least 1h), applied per day:
  - Any entry >= 1h is its own line.
  - Short (<1h) meeting entries are merged together if they reach 1h.
    If not, they are absorbed by the shortest >=1h meeting entry.
  - Short issue entries are merged together if they reach 1h.
  - Anything still left over is folded into that SAME day's largest
    line, so a day's total never changes. If the day has no other line
    at all, it goes in a "MANUAL REVIEW" section.
  - A merged line that contains a meeting shows each part's duration,
    e.g. "Daily Meeting (30m), KEY - summary (7h 30m)". Merged
    lines made only of issues, or containing a PLAIN_ISSUES entry, show
    no durations.
"""
import sys
import datetime
from collections import defaultdict

import xlrd

MIN_HOURS = 1.0
WORK_DESCRIPTION_ISSUES = {"EHEALTH-9360"}
PLAIN_ISSUES = {"EHEALTH-17019"}        # text is just the issue summary
SN_ISSUE_TYPE = "SN Request"            # issue type -> "Working on SN Ticket - ..."
JIRA_PREFIX = "Working on Jira - "
SN_PREFIX = "Working on SN Ticket - "
DAILY_MEETING = "Daily Meeting"
STRIP_PREFIXES = ("DG Sante - ",)
TEXT_REWRITES = {
    "Weekly PLANING Meeting": "Weekly PLANNING Meeting",
    "AI4DEV Introduction": "AI4DEV Introduction Meeting",
}


def load_worklogs(path=None, data=None):
    """-> list of dicts: idx, date, issue, hours, summary, wdesc
    Reads from a file path, or from raw bytes (used by the browser version)."""
    wb = xlrd.open_workbook(path) if data is None else xlrd.open_workbook(file_contents=data)
    sheet = wb.sheet_by_name("Worklogs")
    header = [sheet.cell_value(0, c) for c in range(sheet.ncols)]
    col = {name: i for i, name in enumerate(header)}

    rows = []
    for r in range(1, sheet.nrows):
        issue = sheet.cell_value(r, col["Issue Key"])
        if not issue:
            continue
        date_tuple = xlrd.xldate_as_tuple(sheet.cell_value(r, col["Work date"]), wb.datemode)
        rows.append({
            "idx": len(rows),
            "date": datetime.date(*date_tuple[:3]),
            "issue": issue,
            "hours": float(sheet.cell_value(r, col["Hours"])),
            "summary": str(sheet.cell_value(r, col["Issue summary"])).strip(),
            "type": str(sheet.cell_value(r, col["Issue Type"])).strip(),
            "wdesc": str(sheet.cell_value(r, col["Work Description"]) or "").strip(),
        })
    return rows


def meeting_text(row):
    text = row["wdesc"]
    if not text:
        return row["summary"]
    if text.startswith("Working on issue"):
        return DAILY_MEETING
    for prefix in STRIP_PREFIXES:
        if text.startswith(prefix):
            text = text[len(prefix):]
    return TEXT_REWRITES.get(text, text)


def issue_part(issue, hours, summary, issue_type):
    if issue in PLAIN_ISSUES:
        return {"kind": "plain", "hours": hours, "bare": summary, "prefix": "", "idx": -1}
    prefix = SN_PREFIX if issue_type == SN_ISSUE_TYPE else JIRA_PREFIX
    return {"kind": "issue", "hours": hours, "bare": f"{issue} - {summary}",
            "prefix": prefix, "idx": -1}


def day_parts(day_rows, meeting_issues=WORK_DESCRIPTION_ISSUES):
    """-> (issue_parts sorted by key, meeting_parts in export order)"""
    issues = {}
    meetings = []
    for row in day_rows:
        if row["issue"] in meeting_issues:
            meetings.append({"kind": "meeting", "hours": row["hours"],
                             "bare": meeting_text(row), "prefix": "", "idx": row["idx"]})
        else:
            key = row["issue"]
            if key in issues:
                issues[key]["hours"] += row["hours"]
            else:
                issues[key] = issue_part(key, row["hours"], row["summary"], row["type"])
    return [issues[k] for k in sorted(issues)], meetings


def format_duration(hours):
    h, m = divmod(round(hours * 60), 60)
    return " ".join(p for p in (f"{h}h" if h else "", f"{m}m" if m else "") if p)


def compose(parts):
    """parts -> (total_hours, text)"""
    total = sum(p["hours"] for p in parts)
    kinds = {p["kind"] for p in parts}
    if len(parts) == 1:
        p = parts[0]
        return total, p["prefix"] + p["bare"]
    if "meeting" not in kinds and "plain" not in kinds:
        # issues only: prefix on the first one, no durations
        return total, ", ".join((p["prefix"] if i == 0 else "") + p["bare"]
                                for i, p in enumerate(parts))
    order = {"plain": 0, "meeting": 1, "issue": 2}
    ordered = sorted(parts, key=lambda p: (order[p["kind"]],
                                           p["bare"] != DAILY_MEETING))
    if "plain" in kinds:
        return total, ", ".join(p["bare"] for p in ordered)
    return total, ", ".join(f"{p['bare']} ({format_duration(p['hours'])})" for p in ordered)


def build_lines(rows, meeting_issues=WORK_DESCRIPTION_ISSUES):
    """-> (lines, manual_review); each is a list of (date, hours, text)"""
    by_day = defaultdict(list)
    for row in rows:
        by_day[row["date"]].append(row)

    lines = []          # (date, category, order, hours, text)
    manual_review = []  # (date, hours, text)

    for work_date in sorted(by_day):
        issues, meetings = day_parts(by_day[work_date], meeting_issues)
        big_i = [p for p in issues if p["hours"] >= MIN_HOURS]
        small_i = [p for p in issues if p["hours"] < MIN_HOURS]
        big_m = [[p] for p in meetings if p["hours"] >= MIN_HOURS]
        small_m = [p for p in meetings if p["hours"] < MIN_HOURS]

        # groups: lists of parts that become one line
        groups = [[p] for p in big_i] + big_m
        small_m_total = sum(p["hours"] for p in small_m)
        small_i_total = sum(p["hours"] for p in small_i)

        if small_m and small_m_total >= MIN_HOURS:
            groups.append(small_m)
            small_m = []
        elif small_m and big_m:
            # absorbed by the shortest >=1h meeting (the latest one on a tie)
            target = min(reversed(big_m), key=lambda g: g[0]["hours"])
            target.extend(small_m)
            small_m = []
        if small_i and small_i_total >= MIN_HOURS:
            groups.append(small_i)
            small_i = []

        leftover = small_m + small_i
        if leftover:
            leftover_total = sum(p["hours"] for p in leftover)
            if leftover_total >= MIN_HOURS:
                groups.append(leftover)
            elif groups:
                # fold into this SAME day's largest line (first on a tie)
                target = max(groups, key=lambda g: sum(p["hours"] for p in g))
                target.extend(leftover)
            else:
                _, text = compose(leftover)
                manual_review.append((work_date, leftover_total, text))

        for group in groups:
            hours, text = compose(group)
            if any(p["kind"] == "meeting" for p in group):
                category = 1
                order = min(p["idx"] for p in group if p["kind"] == "meeting")
            elif len(group) == 1:
                category, order = 0, 0
            else:
                category, order = 2, 0
            lines.append((work_date, category, order, hours, text))

    lines.sort(key=lambda x: (x[0], x[1], x[2]))  # stable: keeps key order
    return [(d, h, t) for d, _, _, h, t in lines], manual_review


def format_hours(h):
    # trim trailing .0 noise while keeping e.g. 1.5, 0.75
    return f"{h:.2f}".rstrip("0").rstrip(".") if "." in f"{h:.2f}" else f"{h:.2f}"


def render_output(lines, manual_review, rows):
    """-> (text, messages): the Gekko text and the summary/warning lines"""
    out = []
    current_day = None
    day_total = 0.0
    for work_date, hours, text in lines:
        if work_date != current_day:
            if current_day is not None:
                out.append(f"Total: {format_hours(day_total)}h\n")
            out.append(f"\n# {work_date.isoformat()}\n")
            current_day = work_date
            day_total = 0.0
        out.append(f"{format_hours(hours):<5}{text}\n")
        day_total += hours
    if current_day is not None:
        out.append(f"Total: {format_hours(day_total)}h\n")

    if manual_review:
        out.append("\n\n==== MANUAL REVIEW (still under 1h, no same-day entry to merge with) ====\n")
        for work_date, hours, text in sorted(manual_review, key=lambda x: x[0]):
            out.append(f"{work_date.isoformat()}\t{format_hours(hours)}\t{text}\n")

    grand_total = sum(h for _, h, _ in lines + manual_review)
    out.append(f"\n\nGrand total: {format_hours(grand_total)}h\n")

    messages = []
    total_in = sum(r["hours"] for r in rows)
    messages.append(f"{len(lines)} lines, {len(manual_review)} manual-review lines")
    messages.append(f"Total hours in export: {total_in:.2f}  |  Total hours in output: {grand_total:.2f}")
    if abs(total_in - grand_total) > 1e-6:
        messages.append("WARNING: totals do not match, check the script/data!")

    in_by_day = defaultdict(float)
    for r in rows:
        in_by_day[r["date"]] += r["hours"]
    out_by_day = defaultdict(float)
    for work_date, hours, _ in lines + manual_review:
        out_by_day[work_date] += hours
    for work_date in sorted(set(in_by_day) | set(out_by_day)):
        if abs(in_by_day[work_date] - out_by_day[work_date]) > 1e-6:
            messages.append(f"WARNING: {work_date.isoformat()} in={in_by_day[work_date]:.2f} "
                            f"out={out_by_day[work_date]:.2f} -- day total changed!")
    return "".join(out), messages


def convert_bytes(data, meeting_issues=None):
    """xls bytes -> (text, messages); entry point for the browser version.
    meeting_issues: iterable of issue keys treated as meetings (default WORK_DESCRIPTION_ISSUES)."""
    if hasattr(data, "to_py"):      # JS Uint8Array -> memoryview
        data = data.to_py()
    rows = load_worklogs(data=bytes(data))
    keys = set(meeting_issues) if meeting_issues else WORK_DESCRIPTION_ISSUES
    lines, manual_review = build_lines(rows, keys)
    return render_output(lines, manual_review, rows)


def write_output(path, lines, manual_review, rows):
    text, messages = render_output(lines, manual_review, rows)
    with open(path, "w", encoding="utf-8") as f:
        f.write(text)
    print(f"Wrote to {path}")
    for message in messages:
        print(message)


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    in_path = sys.argv[1]
    out_path = sys.argv[2] if len(sys.argv) > 2 else in_path.rsplit(".", 1)[0] + "_gekko.txt"

    rows = load_worklogs(in_path)
    lines, manual_review = build_lines(rows)
    write_output(out_path, lines, manual_review, rows)


if __name__ == "__main__":
    main()
