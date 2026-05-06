import json
import math
import pandas as pd


INPUT_FILE  = "Scoreboard Test.xlsx"
OUTPUT_FILE = "output.json"

# ── Row indices in the raw sheet ──────────────────────────────────────────────
ROW_METRIC       = 1
ROW_FOCUS        = 2
ROW_SOURCE       = 3
ROW_ROLE         = 4
ROW_TARGET_LABEL = 5
ROW_TARGET_VALUE = 6
ROW_DATA_START   = 7


# Strings that carry no real information and should become null
_ARTIFACTS = {"", " ", "`", "#REF!", "Formula"}


def clean_value(v):
    """Return a JSON-safe Python value from a raw cell.

    Handles four cases beyond a plain pass-through:
      - NaN floats          → None
      - Whole-number floats → int  (42.0 → 42)
      - Artifact strings    → None (backtick, #REF!, whitespace, 'Formula')
      - Numeric strings     → float ('2.87' → 2.87)
    Everything else is returned as-is (meaningful strings like 'Target = 75%').
    """
    if v is None:
        return None
    if isinstance(v, float):
        if math.isnan(v):
            return None
        if v == int(v):
            return int(v)
        return v
    if isinstance(v, str):
        stripped = v.strip()
        if stripped in _ARTIFACTS:
            return None
        # Numeric string — Excel sometimes stores numbers as text
        try:
            f = float(stripped)
            return int(f) if f == int(f) else f
        except ValueError:
            pass
    return v


def clean_name(v):
    """Strip newlines and whitespace from a cell used as a key."""
    return str(v).replace("\n", " ").strip()


def build_catalogue(df):
    """
    Return a list of column descriptors — one per real data column.
    Spacer columns (no metric name in row 1) and the date column (col 0)
    are dropped.

    Duplicate metric names are qualified with their service-section prefix.
    Sections are detected by watching for known anchor metrics (the first
    metric in each service block, e.g. 'PT Total Revenue').  Any keys that
    are still duplicated after that get a column-index suffix so nothing
    is silently overwritten.
    """
    # Anchor metric to section label
    SECTION_ANCHORS = {
        "PT Total Revenue":           "PT",
        "RMT Total Revenue":          "RMT",
        "CHIRO  Total Revenue":       "CHIRO",
        "Pelvic Health Total Revenue": "Pelvic Health",
    }

    # Assign a section label to every column by watching for anchor metrics
    current_section = None
    sections = {}
    for i in range(1, df.shape[1]):
        v = df.iloc[ROW_METRIC, i]
        if not pd.isna(v):
            name = clean_name(v)
            if name in SECTION_ANCHORS:
                current_section = SECTION_ANCHORS[name]
        sections[i] = current_section

    # Count raw occurrences of each metric name
    raw_names = [
        clean_name(df.iloc[ROW_METRIC, i])
        for i in range(1, df.shape[1])
        if not pd.isna(df.iloc[ROW_METRIC, i])
    ]
    name_counts = {}
    for n in raw_names:
        name_counts[n] = name_counts.get(n, 0) + 1

    catalogue = []
    seen_keys = {}  # the key is count, to catch any remaining duplicates

    for col_idx in range(1, df.shape[1]):
        raw_name = df.iloc[ROW_METRIC, col_idx]
        if pd.isna(raw_name):
            continue
        base = clean_name(raw_name)

        if name_counts[base] > 1 and sections.get(col_idx):
            metric_key = f"{sections[col_idx]} — {base}"
        else:
            metric_key = base

        # Last-resort deduplication: append col index if still clashing
        if metric_key in seen_keys:
            metric_key = f"{metric_key} ({col_idx})"
        seen_keys[metric_key] = True

        catalogue.append({
            "col_index":    col_idx,
            "metric":       metric_key,
            "focus":        clean_name(df.iloc[ROW_FOCUS,        col_idx]) if not pd.isna(df.iloc[ROW_FOCUS,        col_idx]) else None,
            "source":       clean_name(df.iloc[ROW_SOURCE,       col_idx]) if not pd.isna(df.iloc[ROW_SOURCE,       col_idx]) else None,
            "role":         clean_name(df.iloc[ROW_ROLE,         col_idx]) if not pd.isna(df.iloc[ROW_ROLE,         col_idx]) else None,
            "target_label": clean_name(df.iloc[ROW_TARGET_LABEL, col_idx]) if not pd.isna(df.iloc[ROW_TARGET_LABEL, col_idx]) else None,
            "target_value": clean_value(df.iloc[ROW_TARGET_VALUE, col_idx]),
        })
    return catalogue


# Focus values that belong to the Phone Performance section
_PHONE_FOCUSES = {"Phone", "Answer", "Book"}


def _group_for(focus):
    """Map a focus value to its top-level JSON group name."""
    if focus in _PHONE_FOCUSES:
        return "Phone Performance"
    return focus or "Other"


def build_output(df, catalogue):
    """
    Return a list of week objects grouped by focus:
      { "week": "YYYY-MM-DD", "<Focus>": { <metric>: { value, source, role, ... } } }

    Phone Performance is its own group (focus values Phone / Answer / Book).
    Metrics with no focus land in "Other".
    Metrics with no value for a given week are omitted entirely.
    The focus field is dropped from individual metric entries since it is
    already captured by the parent group key.
    """
    meta_by_col = {c["col_index"]: c for c in catalogue}
    output = []

    for _, row in df.iloc[ROW_DATA_START:].iterrows():
        week_obj = {"week": pd.to_datetime(row.iloc[0]).strftime("%Y-%m-%d")}

        for col_idx, meta in meta_by_col.items():
            value = clean_value(row.iloc[col_idx])
            if value is None:
                continue

            group = _group_for(meta["focus"])
            if group not in week_obj:
                week_obj[group] = {}

            entry = {"value": value}
            for field in ("source", "role", "target_label", "target_value"):
                if meta[field] is not None:
                    entry[field] = meta[field]

            week_obj[group][meta["metric"]] = entry

        output.append(week_obj)

    return output


def main():
    print(f"Reading {INPUT_FILE} ...")
    df = pd.read_excel(INPUT_FILE, sheet_name=0, header=None)
    print(f"  Raw sheet: {df.shape[0]} rows × {df.shape[1]} columns")

    catalogue = build_catalogue(df)
    print(f"  Real metrics found: {len(catalogue)}  (spacer columns dropped)")

    output = build_output(df, catalogue)
    print(f"  Weeks of data: {len(output)}")
    for week in output:
        groups = {k: v for k, v in week.items() if k != "week"}
        total = sum(len(v) for v in groups.values())
        print(f"    {week['week']} — {total} metrics across {len(groups)} groups: {list(groups.keys())}")

    with open(OUTPUT_FILE, "w") as f:
        json.dump(output, f, indent=2, default=str)

    print(f"\nWrote {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
