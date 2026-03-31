#!/usr/bin/env python3
from __future__ import annotations

import argparse
import os

from fill_pdf_with_json_gui import load_fill_pairs
from pdf_fill_from_json_gui import PDFSmartFiller


def build_default_output_path(pdf_path: str) -> str:
    base, ext = os.path.splitext(pdf_path)
    return f"{base}_filled{ext or '.pdf'}"


def main() -> int:
    parser = argparse.ArgumentParser(description="Fill a PDF using values from a JSON file.")
    parser.add_argument("--pdf", required=True, help="Input PDF path")
    parser.add_argument("--json", required=True, help="Input JSON path")
    parser.add_argument("--out", help="Output PDF path (default: <input>_filled.pdf)")
    args = parser.parse_args()

    pdf_path = args.pdf.strip()
    json_path = args.json.strip()
    out_path = (args.out or "").strip() or build_default_output_path(pdf_path)

    if not os.path.exists(pdf_path):
        print(f"ERROR: PDF not found: {pdf_path}")
        return 1
    if not os.path.exists(json_path):
        print(f"ERROR: JSON not found: {json_path}")
        return 1

    try:
        pairs = load_fill_pairs(json_path)
    except Exception as exc:
        print(f"ERROR: Could not parse JSON: {exc}")
        return 1

    if not pairs:
        print("ERROR: No usable field/value pairs were found in the JSON file.")
        return 1

    filler = None
    try:
        filler = PDFSmartFiller(pdf_path)
        results = filler.fill_pairs(pairs)
        filler.save(out_path)
    except Exception as exc:
        print(f"ERROR: Fill failed: {exc}")
        return 1
    finally:
        if filler is not None:
            filler.close()

    filled = sum(1 for item in results if item.status == "filled")
    not_found = sum(1 for item in results if item.status == "not found")
    errors = sum(1 for item in results if item.status == "error")
    skipped = sum(1 for item in results if item.status == "skipped")

    print(f"Saved: {out_path}")
    print(
        "Summary: "
        f"filled={filled}, not_found={not_found}, errors={errors}, skipped={skipped}, total={len(results)}"
    )

    if not_found or errors:
        print("\nDetails:")
        for item in results:
            if item.status in {"not found", "error"}:
                print(f"- [{item.status}] {item.location} -> {item.value} | {item.details}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
