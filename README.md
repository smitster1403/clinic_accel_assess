# Scoreboard to JSON data extractor

Converts `Scoreboard Test.xlsx` into a structured, queryable `output.json`.

---

## How to run

```bash
pip install -r requirements.txt
python convert.py
```

On macOS (Homebrew Python), use `pip3` and `python3` instead:

```bash
pip3 install -r requirements.txt --break-system-packages
python3 convert.py
```

Output is written to `output.json` in the same directory.

---

## JSON shape

A list of weekly snapshots. Within each week, metrics are grouped by their **focus category**, so all Financial metrics sit under one key, all Caseload metrics under another, and so on. 

However, phone call metrics are pulled into their own `"Phone Performance"` section regardless of their individual focus tag, since they form a distinct operational area. Metrics with no assigned focus land in `"Other"`.

Only metrics that have a recorded value for a given week are included. Fields within a metric entry that carry no information (`target_value`, `role`, etc.) are omitted rather than written as `null`. The `focus` field is not repeated inside each metric entry — it is already captured by the parent group key.

```json
[
  {
    "week": "2026-02-09",
    "Financial": {
      "Total Revenue  - All Services": {
        "value": 42360.64,
        "source": "EMR",
        "role": "J",
        "target_label": "Target"
      }
    },
    "Phone Performance": {
      "Total Calls (5530)": { "value": 281, "source": "CallHero", "role": "J" },
      "Missed Calls":       { "value": 15,  "source": "CallHero", "role": "J" },
      "Answer Rate":        { "value": 0.94, "source": "CallHero", "role": "J" },
      "Booking Rate":       { "value": 0.73, "source": "CallHero", "role": "J" }
    },
    "Caseload": { "..." : "..." }
  }
]
```

**Why group by focus?**  
A flat `{metric: value}` dict forces a consumer to scan all metrics to find the ones they care about. Grouping by focus means a user of this data can navigate straight to `"Financial"` or `"Phone Performance"` without filtering. It also makes the structure self-documenting — the shape of the JSON mirrors how the clinic actually thinks about its performance areas.

**Why Phone Performance is its own section.**  
The spreadsheet marks a range of columns with the label `PHONE PERFORMANCE`. These metrics (Total Calls, Missed Calls, Answer Rate, Booking Rate, NBO, DNB, Winback) have granular focus tags (`Phone`, `Answer`, `Book`) that would fragment them across three separate top-level keys if grouped naively. Consolidating them under `"Phone Performance"` keeps the seven related metrics together and matches the intent of the original spreadsheet layout.

**Why nulls are omitted.**  
An earlier version of the output included every metric for every week regardless of whether it had data, with `null` filling the gaps. That approach produced a **72 KB** file. Because the most recent week (Feb 16) only had 30 of 124 metrics filled in, the bulk of the file was `"value": null` noise that added no information. Omitting metrics with no value and omitting null metadata fields reduced the output to **16 KB**, a 78% reduction, while preserving every real data point. A consumer can safely treat any absent metric as not recorded for that week.

---

## How the messy bits were handled

| Problem | Decision | Trade-off |
|---------|----------|-----------|
| **~18 spacer columns** (all-NaN visual dividers) | Dropped — detected by absence of a metric name in row 1 | None: they carry no data |
| **6 header rows** (metric, focus, source, role, target label, target value) | All preserved as per-metric metadata fields | Slightly larger JSON; worth it for self-describing output |
| **`\n` in cell text** | Stripped with `.replace("\n", " ").strip()` | Some metric names differ slightly from the spreadsheet display |
| **Duplicate metric names** (`Utilization`, `PVA (4 wk avg)`, `TP Utilization`, `Conversion`) | Prefixed with service section (`PT —`, `RMT —`, `CHIRO —`, `Pelvic Health —`) detected from anchor metrics. Remaining clashes get a column-index suffix as a last resort | Keys are slightly longer but unambiguous |
| **Artifact strings** (`` ` ``, `#REF!`, `Formula`, whitespace) | Converted to `null` | One real target value (`Total Revenue`) had a backtick placeholder — treated as missing |
| **Numeric strings** (Excel sometimes stores numbers as text) | Parsed to `float`/`int` | Applied universally; non-numeric strings like `"Target = 75%"` are preserved as-is |
| **Whole-number floats** (`42.0`) | Cast to `int` | Cosmetic — avoids `42.0` noise in the JSON |

---

## With extra time

- **Confirm missing metrics with the data owner** — the most recent week (Feb 16) only has 30 of 124 metrics recorded vs 75 for earlier weeks. The script omits those absent metrics rather than filling with `null`, but it is worth confirming whether Feb 16 is a partially-entered week or whether those metrics genuinely did not apply.
- **`#REF!` on `Pelvic Health — PVA (4 wk avg)`** — a broken formula reference in the source spreadsheet. Needs fixing at the source.
- **Normalise metric names** — several names have double spaces and minor inconsistencies from the spreadsheet layout. A slug-style key (`pt_utilization`) alongside the display name would make programmatic access cleaner.
- **Percentage values** — stored as raw decimals (`0.93`). Adding a `"unit"` field inferred from the metric name (anything ending in `%` or `Rate`) would remove ambiguity for dashboard consumers.
