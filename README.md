# Tempo to Gekko

Converts a Jira Tempo "Worklogs" `.xls` export into plain text lines ready to paste
into Gekko, which has no import, no date field and no task field, only hours and a
free-text description per line.

It works two ways: as a static web page that runs entirely in the browser, or from the
command line.

## Output

One line per entry, with the hours padded to 5 characters, a `# <date>` header and a
`Total` line per day, and a final `Grand total`:

```
# 2026-09-08
4.5  Working on Jira - PROJ-101 - Specify the audit events
2    Working on Jira - PROJ-102 - Deploy release in production
1.5  Daily Meeting (30m), Follow-up meeting (1h)
Total: 8h
```

### Rules
- **Issue worklogs** are aggregated per day and issue. The text is
  `Working on Jira - <KEY> - <summary>`, or `Working on SN Ticket - ...` when the issue
  type is `SN Request`. Issues listed in `PLAIN_ISSUES` (the Bootcamp) use just the
  summary.
- **Meeting worklogs** (the general meetings issue) are one entry per row, and the text
  is the Work Description. "Working on issue ..." or an empty description becomes
  `Daily Meeting`.
- **Minimum 1h per line.** Short entries are merged, first with each other and then into
  the same day's largest line, so a day's total never changes. A merged line that
  contains a meeting shows each part's duration. Anything that cannot be merged goes in a
  `MANUAL REVIEW` section.
- The script checks that the hours in equal the hours out, in total and per day, and
  prints a warning if not.

## Command line

Requires Python 3 and `xlrd` (`pip install xlrd`).

```sh
python3 tempo_to_gekko.py export.xls [output.txt]
```

The default output is `<input>_gekko.txt`.

## Web page

`index.html` is a static page with no server and no build step. The conversion runs in the
browser through [Pyodide](https://pyodide.org), so the timesheet is never uploaded.

- Drop the `.xls` on the page, then use **Copy** or **Download .txt**.
- The **General meetings Jira issue** field holds the common issue you log meetings on
  (default `EHEALTH-9360`). Several keys can be separated by commas. The value is
  remembered in the browser, and changing it re-converts the loaded file.

Try it locally:

```sh
python3 -m http.server 8000
# open http://localhost:8000
```

### Publishing for free
This repo is already a static site. Serve it with GitHub Pages (Settings > Pages >
Deploy from branch `main`, folder `/`), or any other static host. The first load takes a few seconds while Pyodide and `xlrd` download from the
jsdelivr CDN, so users need internet access.

## Files

| Path | Purpose |
| --- | --- |
| `index.html` | Browser UI |
| `tempo_to_gekko.py` | The converter, loaded by the page and usable from the command line |

## Configuration

Set at the top of `tempo_to_gekko.py`:

| Name | Meaning |
| --- | --- |
| `WORK_DESCRIPTION_ISSUES` | Meeting issue keys (the web field overrides this) |
| `PLAIN_ISSUES` | Issues whose text is just the summary |
| `SN_ISSUE_TYPE` | Issue type that gets the `Working on SN Ticket` prefix |
| `STRIP_PREFIXES`, `TEXT_REWRITES` | Cleanups applied to meeting descriptions |
| `MIN_HOURS` | Minimum hours per Gekko line (1.0) |
