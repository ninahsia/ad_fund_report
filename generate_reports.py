#!/usr/bin/env python3
"""
Hour Loop — Monthly Ad Fund Performance Report generator.

Generates one Word (.docx) report per vendor matching an approved template
format. Only month, vendor name, and four summary values change.

USAGE
-----
    python3 generate_reports.py \
        --data "Apr overall ad performance.xlsx" \
        --template "Hour Loop ... Mar (Sangean).docx" \
        --output-dir "./out" \
        --template-month March \
        --target-month April

For a template that isn't the March-Sangean one, also pass the values that
appear in it so the script knows what to find-and-replace:

    --template-vendor Sangean \
    --template-spend 500 --template-sales 6,833 \
    --template-roas 13.67 --template-ctr-frac 79 --template-ctr-text 0.79

CALCULATION RULE (per vendor)
-----------------------------
Case 1 — fund column has a value (>0):
    Total Ad Spend = fund
    Ad Sales       = "ad sales" column
Case 2 — fund column is blank/NaN/0:
    Total Ad Spend = total_spend
    Ad Sales       = total_sales
ROAS = Ad Sales / Total Ad Spend
CTR  = total_clicks / total_impressions * 100

SPOONTIQUES VARIANT
-------------------
Vendors listed in --tacos-vendors use a TACOS row instead of CTR. The value
is left blank in the report (user fills in manually). The CTR explanation
block in Section 2 is replaced with a TACOS explanation.
"""

from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

try:
    import pandas as pd
except ImportError:
    sys.stderr.write("This script requires pandas. Install with:\n    pip install pandas openpyxl\n")
    sys.exit(1)


# ---------------------------------------------------------------------------
# Helper script discovery — the docx skill ships with unpack.py / pack.py /
# soffice.py. They live at a different path in every Cowork session, so look
# them up rather than hard-coding.
# ---------------------------------------------------------------------------

def find_skill_scripts() -> dict[str, Path]:
    """Locate unpack.py, pack.py, soffice.py from the docx skill."""
    needed = ["unpack.py", "pack.py", "soffice.py"]
    search_roots = [
        Path.home() / ".claude" / "skills",
        Path("/mnt/.claude/skills"),
        Path("/sessions"),
    ]
    found: dict[str, Path] = {}
    for root in search_roots:
        if not root.exists():
            continue
        for name in needed:
            if name in found:
                continue
            for hit in root.rglob(f"docx/scripts/office/{name}"):
                found[name] = hit
                break
    missing = [n for n in needed if n not in found]
    if missing:
        raise SystemExit(
            f"Could not find docx skill scripts: {missing}\n"
            f"Searched: {[str(r) for r in search_roots]}\n"
            f"Pass --unpack/--pack/--soffice with explicit paths."
        )
    return found


# ---------------------------------------------------------------------------
# Number formatters
# ---------------------------------------------------------------------------

def fmt_dollar(v: float) -> str:
    return f"${round(v):,}"


def fmt_thousands(v: float) -> str:
    return f"{round(v):,}"


def fmt_roas(v: float) -> str:
    return f"{v:.2f}"


# ---------------------------------------------------------------------------
# Calculation rule
# ---------------------------------------------------------------------------

def compute_metrics(row) -> dict:
    """Apply the fund-based calculation rule and return formatted strings."""
    fund = row["fund"]
    has_fund = pd.notna(fund) and float(fund) > 0

    if has_fund:
        spend = float(fund)
        sales = float(row["ad sales"])
        case = 1
    else:
        spend = float(row["total_spend"])
        sales = float(row["total_sales"])
        case = 2

    roas = sales / spend if spend else 0.0
    impressions = row["total_impressions"]
    ctr_pct = (row["total_clicks"] / impressions) * 100 if impressions else 0.0

    return {
        "case": case,
        "spend_value": spend,
        "sales_value": sales,
        "roas_value": roas,
        "ctr_value": ctr_pct,
        "spend_str": fmt_dollar(spend),
        "sales_str": fmt_thousands(sales),
        "roas_str": fmt_roas(roas),
        "ctr_str": f"{ctr_pct:.2f}",
    }


# ---------------------------------------------------------------------------
# XML substitution
# ---------------------------------------------------------------------------

def apply_replacements(content: str, vendor: str, m: dict, is_tacos: bool, tpl: argparse.Namespace) -> str:
    """Surgically replace template values in word/document.xml."""
    # 1. Month
    content = content.replace(
        f'<w:t xml:space="preserve"> {tpl.template_month}</w:t>',
        f'<w:t xml:space="preserve"> {tpl.target_month}</w:t>',
    )
    # 2. Vendor name
    content = content.replace(
        f'<w:t>{tpl.template_vendor}</w:t>',
        f'<w:t>{vendor}</w:t>',
    )
    # 3. Total Ad Spend (with $ prefix)
    content = content.replace(
        f'<w:t>{tpl.template_spend}</w:t>',
        f'<w:t>{m["spend_str"]}</w:t>',
    )
    # 4. Ad Sales (number only — $ sits in a separate run in the template)
    content = content.replace(
        f'<w:t>{tpl.template_sales}</w:t>',
        f'<w:t>{m["sales_str"]}</w:t>',
    )
    # 5. ROAS — three occurrences in the standard template
    content = content.replace(
        f'<w:t>{tpl.template_roas}</w:t>',
        f'<w:t>{m["roas_str"]}</w:t>',
    )
    content = content.replace(
        f'<w:t xml:space="preserve">{tpl.template_roas} </w:t>',
        f'<w:t xml:space="preserve">{m["roas_str"]} </w:t>',
    )

    if is_tacos:
        return apply_tacos_variant(content, m, tpl)
    return apply_ctr_value(content, m, tpl)


def apply_ctr_value(content: str, m: dict, tpl: argparse.Namespace) -> str:
    """Replace CTR digits in both the table cell and the explanation sentence."""
    ctr = m["ctr_value"]
    int_part = int(ctr)
    frac_digits = f"{ctr:.2f}".split(".")[1]

    content = content.replace(
        f'<w:t>{tpl.template_ctr_text}</w:t>',
        f'<w:t>{m["ctr_str"]}</w:t>',
    )
    if int_part != 0:
        content = content.replace(
            '<w:t>0.</w:t>',
            f'<w:t>{int_part}.</w:t>',
        )
    content = content.replace(
        f'<w:t>{tpl.template_ctr_frac}</w:t>',
        f'<w:t>{frac_digits}</w:t>',
    )
    return content


def apply_tacos_variant(content: str, m: dict, tpl: argparse.Namespace) -> str:
    """Convert the CTR row + explanation block into a TACOS variant with a
    blank value cell. User fills in the TACOS percentage manually."""
    content = content.replace(
        '<w:t>Click-Through Rate (CTR)</w:t>',
        '<w:t>TACOS</w:t>',
    )
    content = content.replace(
        '<w:t>0.</w:t>',
        '<w:t xml:space="preserve"></w:t>',
    )
    content = content.replace(
        f'<w:t>{tpl.template_ctr_frac}</w:t>',
        '<w:t xml:space="preserve"></w:t>',
    )
    content = content.replace(
        '<w:t>%</w:t>',
        '<w:t xml:space="preserve"></w:t>',
    )
    content = content.replace('<w:t>CTR = (</w:t>', '<w:t>TACOS = (</w:t>')
    content = content.replace(
        '<w:t>Number of Clicks / Number of Impressions) x 100</w:t>',
        '<w:t>Total Ad Spend / Total Sales) x 100</w:t>',
    )
    content = content.replace(
        '<w:t>Click-Through</w:t>',
        '<w:t>Total Advertising Cost</w:t>',
    )
    content = content.replace(
        '<w:t>Rate measures the percentage of people who clicked on your ad after seeing it.</w:t>',
        '<w:t>of Sales measures what percentage of total revenue is spent on advertising. A lower TACOS indicates more efficient overall advertising performance.</w:t>',
    )
    content = content.replace(
        '<w:t xml:space="preserve">For every 100 times </w:t>',
        '<w:t xml:space="preserve">For every $100 in total sales generated</w:t>',
    )
    content = content.replace('<w:t>the</w:t>', '<w:t xml:space="preserve"></w:t>')
    content = content.replace(
        '<w:t xml:space="preserve"> ad was shown (impressions) </w:t>',
        '<w:t xml:space="preserve"></w:t>',
    )
    content = content.replace(
        f'<w:t>{tpl.template_ctr_text}</w:t>',
        '<w:t xml:space="preserve">$____</w:t>',
    )
    content = content.replace(
        '<w:t xml:space="preserve"> people clicked on it</w:t>',
        '<w:t xml:space="preserve"> was spent on advertising</w:t>',
    )
    content = content.replace(
        '<w:t>that is both engaging (as shown by the CTR) and effective (as indicated by the ROAS).</w:t>',
        '<w:t>that is both efficient (as shown by the TACOS) and effective (as indicated by the ROAS).</w:t>',
    )
    return content


# ---------------------------------------------------------------------------
# Build one report
# ---------------------------------------------------------------------------

def build_report(vendor: str, metrics: dict, args, work_dir: Path,
                 scripts: dict, is_tacos: bool) -> Path:
    """Produce one .docx for a vendor. Returns the output path."""
    safe_vendor = vendor.replace("/", "-")
    short_month = args.target_month[:3]
    out_name = f"Hour Loop Monthly Ad Fund Performance Report - {short_month} ({safe_vendor}).docx"
    out_path = args.output_dir / out_name

    work_copy = work_dir / f"{safe_vendor}.docx"
    unpacked = work_dir / f"{safe_vendor}_unpacked"
    if unpacked.exists():
        shutil.rmtree(unpacked)
    shutil.copy(args.template, work_copy)

    subprocess.run(
        [sys.executable, str(scripts["unpack.py"]), str(work_copy), str(unpacked)],
        check=True, capture_output=True, text=True,
    )

    doc_xml = unpacked / "word" / "document.xml"
    content = doc_xml.read_text(encoding="utf-8")
    content = apply_replacements(content, vendor, metrics, is_tacos, args)
    doc_xml.write_text(content, encoding="utf-8")

    # Atomic write — pack to a temp filename, verify, then rename. This is
    # robust on Cowork's mounted filesystem where in-place overwrites can
    # leave stale entries. The temp name keeps the .docx suffix so pack.py
    # accepts it as a valid output.
    tmp_out = out_path.with_name("_tmp_" + out_path.name)
    res = subprocess.run(
        [sys.executable, str(scripts["pack.py"]), str(unpacked), str(tmp_out),
         "--original", str(work_copy)],
        capture_output=True, text=True,
    )
    if res.returncode != 0:
        raise RuntimeError(f"pack failed for {vendor}:\nSTDOUT:\n{res.stdout}\nSTDERR:\n{res.stderr}")
    with zipfile.ZipFile(tmp_out):
        pass
    if out_path.exists():
        try: out_path.unlink()
        except: pass
    tmp_out.rename(out_path)
    return out_path


def convert_to_pdf(docx_path: Path, args, work_dir: Path, scripts: dict):
    pdf_in_work = work_dir / (docx_path.stem + ".pdf")
    if pdf_in_work.exists():
        pdf_in_work.unlink()
    res = subprocess.run(
        [sys.executable, str(scripts["soffice.py"]), "--headless", "--convert-to", "pdf",
         str(docx_path), "--outdir", str(work_dir)],
        capture_output=True, text=True,
    )
    if not pdf_in_work.exists():
        print(f"  PDF conversion failed: {res.stderr.strip()[:200]}")
        return None
    out_pdf = args.output_dir / pdf_in_work.name
    tmp_pdf = out_pdf.with_suffix(".pdf.new")
    tmp_pdf.write_bytes(pdf_in_work.read_bytes())
    if out_pdf.exists():
        try: out_pdf.unlink()
        except: pass
    tmp_pdf.rename(out_pdf)
    return out_pdf


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="Generate Hour Loop monthly ad performance reports.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    p.add_argument("--data", required=True, type=Path,
                   help="Path to the monthly ad performance .xlsx data file")
    p.add_argument("--template", required=True, type=Path,
                   help="Path to an approved prior-month .docx that will be used as the format template")
    p.add_argument("--output-dir", required=True, type=Path,
                   help="Folder where finished reports will be saved")
    p.add_argument("--template-month", required=True,
                   help="Month name in the template (e.g. 'March')")
    p.add_argument("--target-month", required=True,
                   help="Month name for the new reports (e.g. 'April')")

    # Template values — defaults match the Sangean March 2026 template
    p.add_argument("--template-vendor", default="Sangean",
                   help="Vendor name appearing in the template")
    p.add_argument("--template-spend", default="500",
                   help="Total Ad Spend value in the template (raw, no $)")
    p.add_argument("--template-sales", default="6,833",
                   help="Ad Sales value in the template (raw, no $)")
    p.add_argument("--template-roas", default="13.67",
                   help="ROAS value in the template")
    p.add_argument("--template-ctr-frac", default="79",
                   help="CTR fractional digits in the table cell (e.g. '79' for 0.79%%)")
    p.add_argument("--template-ctr-text", default="0.79",
                   help="CTR value as it appears in the 'Approximately X people clicked' sentence")

    p.add_argument("--tacos-vendors", default="Spoontiques",
                   help="Comma-separated vendor names that use the TACOS variant")
    p.add_argument("--skip-vendors", default="",
                   help="Comma-separated vendor names to skip")

    p.add_argument("--no-pdf", action="store_true",
                   help="Skip the .docx -> .pdf conversion")

    # Skill script overrides
    p.add_argument("--unpack", type=Path, help="Override path to unpack.py")
    p.add_argument("--pack", type=Path, help="Override path to pack.py")
    p.add_argument("--soffice", type=Path, help="Override path to soffice.py")

    return p.parse_args()


def main() -> int:
    args = parse_args()
    args.output_dir.mkdir(parents=True, exist_ok=True)

    # Resolve helper scripts
    scripts = find_skill_scripts() if not (args.unpack and args.pack and args.soffice) else {}
    if args.unpack: scripts["unpack.py"] = args.unpack
    if args.pack: scripts["pack.py"] = args.pack
    if args.soffice: scripts["soffice.py"] = args.soffice
    print("Using helper scripts:")
    for k, v in scripts.items():
        print(f"  {k}: {v}")

    tacos_vendors = {v.strip() for v in args.tacos_vendors.split(",") if v.strip()}
    skip_vendors = {v.strip() for v in args.skip_vendors.split(",") if v.strip()}

    df = pd.read_excel(args.data)
    print(f"\nLoaded {len(df)} rows from {args.data}")

    with tempfile.TemporaryDirectory(prefix="reports_work_") as work_str:
        work_dir = Path(work_str)
        for _, row in df.iterrows():
            vendor = row["vendor"]
            if vendor in skip_vendors:
                continue
            is_tacos = vendor in tacos_vendors
            m = compute_metrics(row)
            print(f"\n{vendor} (case {m['case']}{', TACOS' if is_tacos else ''})")
            print(f"  Total Ad Spend = {m['spend_str']}")
            print(f"  Ad Sales       = ${m['sales_str']}")
            print(f"  ROAS           = {m['roas_str']}")
            if not is_tacos:
                print(f"  CTR            = {m['ctr_str']}%")

            docx_path = build_report(vendor, m, args, work_dir, scripts, is_tacos)
            print(f"  -> {docx_path.name}")
            if not args.no_pdf:
                pdf = convert_to_pdf(docx_path, args, work_dir, scripts)
                if pdf:
                    print(f"  -> {pdf.name}")

    print("\nDone.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
