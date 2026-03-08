#!/usr/bin/env python3
from __future__ import annotations

import argparse
import os

from pdf_fill_from_json_gui import PDFSmartFiller, load_pairs_from_json


def build_output_path(input_pdf: str, output_pdf: str | None) -> str:
    if output_pdf:
        return output_pdf
    base, ext = os.path.splitext(input_pdf)
    return f"{base}_filled{ext}"


def main() -> int:
    parser = argparse.ArgumentParser(description="Fill a PDF from a JSON mapping file (non-GUI).")
    parser.add_argument("--pdf", required=True, help="Input PDF path")
    parser.add_argument("--json", required=True, help="Input JSON path")
    parser.add_argument("--out", help="Output PDF path (default: <input>_filled.pdf)")
    args = parser.parse_args()

    inp = args.pdf.strip()
    js = args.json.strip()
    out = build_output_path(inp, args.out.strip() if args.out else None)

    if not os.path.exists(inp):
        print(f"ERROR: input PDF not found: {inp}")
        return 1
    if not os.path.exists(js):
        print(f"ERROR: input JSON not found: {js}")
        return 1

    try:
        pairs = load_pairs_from_json(js)
    except Exception as exc:
        print(f"ERROR: could not parse JSON: {exc}")
        return 1

    if not pairs:
        print("ERROR: no usable mappings found in JSON.")
        return 1

    filler = None
    try:
        filler = PDFSmartFiller(inp)
        results = filler.fill_pairs(pairs)
        filler.save(out)
    except Exception as exc:
        print(f"ERROR: fill failed: {exc}")
        return 1
    finally:
        if filler:
            filler.close()

    filled = sum(1 for r in results if r.status == "filled")
    not_found = sum(1 for r in results if r.status == "not found")
    errors = sum(1 for r in results if r.status == "error")
    skipped = sum(1 for r in results if r.status == "skipped")

    print(f"Saved: {out}")
    print(f"Summary: filled={filled}, not_found={not_found}, errors={errors}, skipped={skipped}, total={len(results)}")

    if not_found or errors:
        print("\nDetails:")
        for r in results:
            if r.status in {"not found", "error"}:
                print(f"- [{r.status}] {r.location} -> {r.value} | {r.details}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
