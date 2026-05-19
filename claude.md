# Hour Loop — Monthly Ad Fund Performance Report

## Purpose
Automate the generation of Hour Loop's monthly Ad Fund Performance Reports (one Word `.docx` per vendor), using a fixed template format. The reports go to vendors and report on Amazon ad performance for a single month.

## Inputs

### 1. Data file (Excel)
Path: a monthly ad performance `.xlsx`, e.g. `Apr overall ad performance.xlsx`.

Columns:

| Column | Meaning |
|---|---|
| `vendor` | Vendor name (one row per vendor) |
| `month` | YYYY-MM (e.g. `2026-04`) |
| `marketplace` | e.g. `US` |
| `total_impressions` | Sum of impressions across all campaigns |
| `total_clicks` | Sum of clicks |
| `total_spend` | Total ad spend at the campaign level |
| `total_sales` | Total ad-attributed sales at the campaign level |
| `total_units_sold` | Units sold from ads |
| `ACOS`, `roas`, `ctr` | Computed by source |
| `fund` | Vendor's ad fund (may be blank/NaN) |
| `ad sales` | Sales attributed against the ad fund |

### 2. Template Word file
A previously approved `.docx` of the Hour Loop Monthly Ad Performance Report for any vendor. The current template uses Sangean's March 2026 report. It contains:

1. Title: "Hour Loop Monthly Ad Performance Report"
2. Month + Vendor Name
3. **Section 1 — Summary of Ad Performance** (table)
   - Total Ad Spend (prefixed with `$`)
   - Ad sales (prefixed with `$`)
   - ROAS
   - Click-Through Rate (CTR)  *— or TACOS for Spoontiques*
4. **Section 2 — Key Metrics Explanation** (formulas, narrative)
5. Hour Loop logo at the bottom

## Calculation rule (per vendor)

For each vendor, determine the calculation logic based on the `fund` column:

| Case | Condition | Total Ad Spend | Ad Sales |
|---|---|---|---|
| 1 | Vendor HAS a value in `fund` (> 0) | `fund` | `ad sales` column |
| 2 | Vendor has NO value in `fund` (null/blank/0) | `total_spend` column | `total_sales` column |

The rule OVERRIDES any default aggregation. Apply consistently across all calculations and tables. Do NOT mix logic within a single vendor.

Derived metrics:

- `ROAS = Ad Sales / Total Ad Spend`, rounded to 2 decimals
- `CTR  = total_clicks / total_impressions × 100`, rounded to 2 decimals (shown as `X.YY%`)
- `TACOS` (Spoontiques only) = `Total Ad Spend / Total Sales × 100` — **left blank in the report** so it can be filled in manually

Formatting:

- `Total Ad Spend` and `Ad Sales` are rounded to whole dollars with thousands separator, prefixed with `$` (e.g. `$3,324`).

## Output format

- One `.docx` per vendor named `Hour Loop Monthly Ad Fund Performance Report - <Month> (<Vendor>).docx`
- Match the template's structure, headings, fonts, table layout, and logo exactly
- Save into the user's selected folder (e.g. `Apr claude/`)
- Optionally convert each `.docx` to `.pdf` afterwards (LibreOffice)

## Vendor-specific notes

- **Spoontiques** — Going forward, uses a TACOS variant of the format. The Click-Through Rate row in Section 1 is replaced with a **TACOS** row (value blank). Section 2's CTR explanation is replaced with the TACOS explanation. The user fills in the TACOS value manually.

## Workflow

1. Read the data file and the template `.docx`.
2. For each vendor row, compute the four summary metrics using the fund rule.
3. Take a fresh copy of the template, unpack it (`unpack.py`), surgically replace the values inside `word/document.xml`, then repack (`pack.py`). This preserves all formatting byte-for-byte and only changes the values.
4. Save the result to the output folder.
5. (Optional) Convert each `.docx` to `.pdf` via LibreOffice.

The full Python script implementing this is provided as `generate_reports.py`. It takes command-line arguments — no code edits needed between runs.

## Running the script

Minimal invocation (uses defaults that match the March-Sangean template):

```bash
python3 generate_reports.py \
    --data "Apr overall ad performance.xlsx" \
    --template "Hour Loop Monthly Ad Fund Performance Report- Mar (Sangean).docx" \
    --output-dir "./out" \
    --template-month March \
    --target-month April
```

Other options:
- `--tacos-vendors "Spoontiques,OtherVendor"` — comma-separated list of vendors that use the TACOS variant (default: `Spoontiques`)
- `--skip-vendors "Allsop,Kalalou"` — skip these vendors
- `--no-pdf` — generate only `.docx`, skip PDF conversion
- `--template-vendor`, `--template-spend`, `--template-sales`, `--template-roas`, `--template-ctr-frac`, `--template-ctr-text` — override the values the script searches for inside the template (needed only if you use a template other than Sangean March 2026)
- `--unpack`, `--pack`, `--soffice` — explicit paths to the docx-skill helper scripts (the script auto-discovers them by default; specify these only if auto-discovery fails)

The script auto-discovers `unpack.py`, `pack.py`, and `soffice.py` from the built-in docx skill, so it works across Cowork sessions without any path edits.

## Known quirks

- LibreOffice's batch PDF conversion can fail silently when run rapidly. Convert one file at a time with a small pause between, or use a temp filename and rename when writing into the user's mounted folder.
- The `unpack.py` / `pack.py` helpers live under `~/.claude/skills/docx/scripts/office/`. They preserve formatting and validate the docx after repacking.
