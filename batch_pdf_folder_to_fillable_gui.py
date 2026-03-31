#!/usr/bin/env python3
"""
batch_pdf_folder_to_fillable_gui.py

GUI tool to process all PDFs in a folder (including subfolders):
- If a PDF is static (no form fields), auto-detect fields and create a fillable PDF in the same folder.
- For every PDF, create a JSON file with placeholder names for text, checkbox, and radio fields.

Local update note:
- Uses nearest visible text when exported widget labels are missing.
- Uses the updated local static_pdf_to_fillable.py detector from this folder.
- For already fillable PDFs, placeholder labels are compared with the
  nearest visible text and replaced when the embedded label is not relevant.
"""

from __future__ import annotations

import json
import os
import re
import traceback
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, messagebox
from typing import Dict, List

try:
    import fitz  # PyMuPDF
except ImportError as exc:
    raise SystemExit("Missing dependency: PyMuPDF. Install with: pip install pymupdf") from exc

from static_pdf_to_fillable import FieldSpec, auto_detect_fields, convert_pdf


def _normalize_gui_label(text: str) -> str:
    text = re.sub(r"\s+", " ", (text or "").strip())
    text = re.sub(r"[_\.]{2,}", " ", text)
    text = re.sub(r"[\[\]\(\)]", " ", text)
    return text.strip(" -:\t,.;")


def _candidate_gui_label(text: str) -> str:
    cleaned = _normalize_gui_label(text)
    if not cleaned:
        return ""
    m = re.findall(r"([A-Za-z][A-Za-z0-9 /&().,'-]{1,80})\s*:", cleaned)
    if m:
        return _normalize_gui_label(m[-1])
    return cleaned


def _normalize_compare_text(text: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", (text or "").lower()).strip()


def _looks_like_auto_label(label: str, field_name: str) -> bool:
    norm_label = _normalize_compare_text(label)
    norm_name = _normalize_compare_text(field_name)
    if not norm_label:
        return True
    if norm_label == norm_name and norm_name:
        return True
    if norm_label.startswith(("txt", "chk", "fld", "field", "unnamed")):
        return True
    return False


def _labels_are_relevant(current_label: str, nearest_label: str) -> bool:
    a = _normalize_compare_text(current_label)
    b = _normalize_compare_text(nearest_label)
    if not a or not b:
        return False
    if a == b or a in b or b in a:
        return True

    a_tokens = {tok for tok in a.split() if tok}
    b_tokens = {tok for tok in b.split() if tok}
    if not a_tokens or not b_tokens:
        return False

    overlap = len(a_tokens & b_tokens)
    shortest = min(len(a_tokens), len(b_tokens))
    return overlap >= max(1, shortest - 1)


def _choose_placeholder_label(field_name: str, current_label: str, nearest_label: str) -> str:
    current = _candidate_gui_label(current_label)
    nearest = _candidate_gui_label(nearest_label)
    if not nearest:
        return current
    if not current:
        return nearest
    if _looks_like_auto_label(current, field_name):
        return nearest
    if _labels_are_relevant(current, nearest):
        return current
    return nearest


def _rect_center(rect: "fitz.Rect") -> tuple[float, float]:
    return ((rect.x0 + rect.x1) / 2.0, (rect.y0 + rect.y1) / 2.0)


def _nearest_text_label(page, target_rect: "fitz.Rect") -> str:
    text_dict = page.get_text("dict")
    best = ""
    best_score = float("inf")
    tx, ty = _rect_center(target_rect)

    for block in text_dict.get("blocks", []):
        if block.get("type", 0) != 0:
            continue
        for line in block.get("lines", []):
            spans = line.get("spans", [])
            line_text = "".join(span.get("text", "") for span in spans).strip()
            label = _candidate_gui_label(line_text)
            if len(label) < 2:
                continue

            if spans:
                x0 = min(float(span["bbox"][0]) for span in spans)
                y0 = min(float(span["bbox"][1]) for span in spans)
                x1 = max(float(span["bbox"][2]) for span in spans)
                y1 = max(float(span["bbox"][3]) for span in spans)
            else:
                b = line.get("bbox", [0, 0, 0, 0])
                x0, y0, x1, y1 = map(float, b)

            rect = fitz.Rect(x0, y0, x1, y1)
            lx, ly = _rect_center(rect)
            if rect.x1 < target_rect.x0:
                dx = target_rect.x0 - rect.x1
            elif rect.x0 > target_rect.x1:
                dx = rect.x0 - target_rect.x1
            else:
                dx = abs(tx - lx) * 0.15
            dy = abs(ty - ly)
            score = dx * 1.25 + dy * 2.6
            if rect.intersects(target_rect):
                score -= 18.0
            elif abs(rect.y0 - target_rect.y0) < max(6.0, target_rect.height):
                score -= 6.0

            if score < best_score:
                best_score = score
                best = label

    return best


def widget_type_name(widget_type: int) -> str:
    if widget_type == getattr(fitz, "PDF_WIDGET_TYPE_TEXT", 7):
        return "text"
    if widget_type == getattr(fitz, "PDF_WIDGET_TYPE_CHECKBOX", 2):
        return "checkbox"
    if widget_type == getattr(fitz, "PDF_WIDGET_TYPE_RADIOBUTTON", 5):
        return "radio"
    return "other"


def find_pdfs_recursive(root_folder: str) -> List[str]:
    pdfs: List[str] = []
    for base, _, files in os.walk(root_folder):
        for name in files:
            if name.lower().endswith(".pdf"):
                pdfs.append(os.path.join(base, name))
    pdfs.sort()
    return pdfs


def delete_json_recursive(root_folder: str) -> tuple[int, int, int]:
    found = 0
    deleted = 0
    failed = 0
    for base, _, files in os.walk(root_folder):
        for name in files:
            if not name.lower().endswith(".json"):
                continue
            found += 1
            p = os.path.join(base, name)
            try:
                os.remove(p)
                deleted += 1
            except Exception:
                failed += 1
    return found, deleted, failed


def _extract_pdf_string_tokens(src: str, key: str) -> List[str]:
    if not src:
        return []
    vals: List[str] = []
    for m in re.finditer(rf"/{key}\((.*?)\)", src, flags=re.S):
        v = (m.group(1) or "").strip()
        if v:
            vals.append(v)
    return vals


def _extract_parent_xref(src: str) -> int | None:
    if not src:
        return None
    m = re.search(r"/Parent\s+(\d+)\s+0\s+R", src)
    if not m:
        return None
    try:
        return int(m.group(1))
    except Exception:
        return None


def _widget_states(w) -> List[str]:
    vals: List[str] = []
    if not hasattr(w, "button_states"):
        return vals
    try:
        states = w.button_states()
        if isinstance(states, dict):
            for it in states.values():
                if isinstance(it, (list, tuple)):
                    vals.extend(str(x).lstrip("/") for x in it)
                else:
                    vals.append(str(it).lstrip("/"))
        elif isinstance(states, (list, tuple)):
            vals.extend(str(x).lstrip("/") for x in states)
    except Exception:
        return vals
    out: List[str] = []
    seen = set()
    for v in vals:
        vv = (v or "").strip()
        if not vv or vv.lower() == "off":
            continue
        k = vv.lower()
        if k in seen:
            continue
        seen.add(k)
        out.append(vv)
    return out


def _radio_group_name(doc, w) -> str:
    # Prefer parent field title (/T) for radio groups.
    xref = getattr(w, "xref", None)
    if xref:
        try:
            src = doc.xref_object(xref, compressed=False)
        except Exception:
            src = ""
        parent = _extract_parent_xref(src)
        if parent:
            try:
                psrc = doc.xref_object(parent, compressed=False)
            except Exception:
                psrc = ""
            for key in ("T", "TU"):
                toks = _extract_pdf_string_tokens(psrc, key)
                if toks:
                    return toks[0]
        for key in ("T", "TU"):
            toks = _extract_pdf_string_tokens(src, key)
            if toks:
                return toks[0]
    return (getattr(w, "field_name", "") or "").strip()


def _widget_current_value(w) -> str:
    raw = getattr(w, "field_value", None)
    if raw is None:
        return ""
    value = str(raw).strip()
    if not value or value.lower() == "off":
        return ""
    return value


def extract_existing_widget_placeholders(pdf_path: str) -> List[Dict]:
    items: List[Dict] = []
    doc = fitz.open(pdf_path)
    try:
        seen = set()
        radio_groups: Dict[str, Dict] = {}
        for pidx in range(doc.page_count):
            page = doc.load_page(pidx)
            for w in list(page.widgets() or []):
                name = (w.field_name or "").strip() or f"unnamed_p{pidx+1}_{len(items)+1}"
                ftype = widget_type_name(getattr(w, "field_type", -1))
                rect = getattr(w, "rect", None)
                nearest_label = _nearest_text_label(page, rect) if rect else ""
                chosen_label = _choose_placeholder_label(
                    name,
                    getattr(w, "field_label", "") or "",
                    nearest_label,
                )
                current_value = _widget_current_value(w)
                if ftype == "radio":
                    gname = _radio_group_name(doc, w).strip() or name
                    grp = radio_groups.setdefault(
                        gname.lower(),
                        {
                            "name": gname,
                            "type": "radio",
                            "page": pidx + 1,
                            "label": chosen_label,
                            "value": current_value,
                            "values": [],
                            "_seen": set(),
                        },
                    )
                    if chosen_label:
                        grp["label"] = _choose_placeholder_label(gname, grp["label"], chosen_label)
                    if current_value:
                        grp["value"] = current_value
                    for v in _widget_states(w):
                        vk = v.lower()
                        if vk in grp["_seen"]:
                            continue
                        grp["_seen"].add(vk)
                        grp["values"].append(v)
                    continue
                key = (name, ftype, pidx + 1)
                if key in seen:
                    continue
                seen.add(key)
                items.append(
                    {
                        "name": name,
                        "type": ftype,
                        "page": pidx + 1,
                        "label": chosen_label or "",
                        "value": current_value,
                    }
                )
        for grp in radio_groups.values():
            vals = grp.pop("values", [])
            grp.pop("_seen", None)
            if vals:
                for v in vals:
                    items.append(
                        {
                            "name": grp["name"],
                            "type": "radio",
                            "page": grp["page"],
                            "label": grp["label"],
                            "value": grp.get("value", ""),
                            "export_value": v,
                        }
                    )
            else:
                items.append(
                    {
                        "name": grp["name"],
                        "type": "radio",
                        "page": grp["page"],
                        "label": grp["label"],
                        "value": grp.get("value", ""),
                    }
                )
    finally:
        doc.close()
    return items


def fieldspecs_to_placeholder_list(fields: List[FieldSpec]) -> List[Dict]:
    out: List[Dict] = []
    for spec in fields:
        if spec.field_type == "radio" and spec.options:
            for opt in spec.options:
                out.append(
                    {
                        "name": spec.name,
                        "type": "radio",
                        "page": opt.get("page") or spec.page,
                        "label": opt.get("label") or spec.label or "",
                        "value": "",
                        "export_value": opt.get("value"),
                    }
                )
        else:
            out.append(
                {
                    "name": spec.name,
                    "type": spec.field_type,
                    "page": spec.page,
                    "label": spec.label or "",
                    "value": "",
                }
            )
    return out


def count_types(items: List[Dict]) -> Dict[str, int]:
    counts = {"text": 0, "checkbox": 0, "radio": 0, "other": 0}
    for it in items:
        t = (it.get("type") or "other").lower()
        counts[t] = counts.get(t, 0) + 1
    counts["total"] = len(items)
    return counts


def write_placeholder_json(pdf_path: str, payload: Dict) -> str:
    base, _ = os.path.splitext(pdf_path)
    out_json = base + "_placeholders.json"
    with open(out_json, "w", encoding="utf-8") as f:
        json.dump(payload, f, indent=2)
    return out_json


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Batch PDF Folder -> Fillable + Placeholder JSON")
        self.geometry("980x700")

        self.folder_var = tk.StringVar()
        self.use_ocr_var = tk.BooleanVar(value=False)
        self.ocr_lang_var = tk.StringVar(value="eng")
        self.text_dx_var = tk.StringVar(value="0")
        self.text_dy_var = tk.StringVar(value="0")
        self.box_dx_var = tk.StringVar(value="0")
        self.box_dy_var = tk.StringVar(value="0")

        self._build_ui()

    def _build_ui(self):
        top = tk.Frame(self)
        top.pack(fill="x", padx=12, pady=(12, 8))

        tk.Label(top, text="Root Folder").grid(row=0, column=0, sticky="w")
        tk.Entry(top, textvariable=self.folder_var, width=90).grid(row=0, column=1, padx=8, sticky="we")
        tk.Button(top, text="Browse", command=self._pick_folder).grid(row=0, column=2)
        top.grid_columnconfigure(1, weight=1)

        opts = tk.LabelFrame(self, text="Auto-Detect Options (for static PDFs)")
        opts.pack(fill="x", padx=12, pady=(0, 8))

        tk.Checkbutton(opts, text="Use OCR (for scanned/image PDFs)", variable=self.use_ocr_var).grid(row=0, column=0, padx=8, pady=8, sticky="w")
        tk.Label(opts, text="OCR lang").grid(row=0, column=1, sticky="w")
        tk.Entry(opts, textvariable=self.ocr_lang_var, width=10).grid(row=0, column=2, padx=(6, 18), sticky="w")

        tk.Label(opts, text="Text offset X/Y").grid(row=1, column=0, padx=8, pady=(0, 8), sticky="w")
        tk.Entry(opts, textvariable=self.text_dx_var, width=8).grid(row=1, column=1, pady=(0, 8), sticky="w")
        tk.Entry(opts, textvariable=self.text_dy_var, width=8).grid(row=1, column=2, padx=(6, 18), pady=(0, 8), sticky="w")

        tk.Label(opts, text="Box offset X/Y").grid(row=2, column=0, padx=8, pady=(0, 8), sticky="w")
        tk.Entry(opts, textvariable=self.box_dx_var, width=8).grid(row=2, column=1, pady=(0, 8), sticky="w")
        tk.Entry(opts, textvariable=self.box_dy_var, width=8).grid(row=2, column=2, padx=(6, 18), pady=(0, 8), sticky="w")

        action = tk.Frame(self)
        action.pack(fill="x", padx=12, pady=(0, 8))
        tk.Button(action, text="Process Folder", command=self._process_folder, height=2).pack(side="left")

        log_frame = tk.LabelFrame(self, text="Status")
        log_frame.pack(fill="both", expand=True, padx=12, pady=(0, 12))

        self.log = tk.Text(log_frame, state="disabled")
        self.log.pack(fill="both", expand=True, padx=8, pady=8)

    def _pick_folder(self):
        path = filedialog.askdirectory(title="Choose Root Folder")
        if path:
            self.folder_var.set(path)

    def _write_log(self, line: str):
        self.log.configure(state="normal")
        self.log.insert(tk.END, line + "\n")
        self.log.see(tk.END)
        self.log.configure(state="disabled")

    def _process_folder(self):
        root = self.folder_var.get().strip()
        if not root:
            messagebox.showerror("Missing folder", "Please choose a root folder.")
            return
        if not os.path.isdir(root):
            messagebox.showerror("Invalid folder", "Selected folder does not exist.")
            return

        try:
            text_dx = float(self.text_dx_var.get().strip() or "0")
            text_dy = float(self.text_dy_var.get().strip() or "0")
            box_dx = float(self.box_dx_var.get().strip() or "0")
            box_dy = float(self.box_dy_var.get().strip() or "0")
        except Exception:
            messagebox.showerror("Invalid offsets", "Offsets must be numeric values.")
            return

        use_ocr = bool(self.use_ocr_var.get())
        ocr_lang = self.ocr_lang_var.get().strip() or "eng"

        self._write_log("Deleting all JSON files first...")
        found_json, deleted_json, failed_json = delete_json_recursive(root)
        self._write_log(
            f"JSON cleanup: found={found_json}, deleted={deleted_json}, failed={failed_json}"
        )

        pdf_paths = find_pdfs_recursive(root)
        if not pdf_paths:
            messagebox.showwarning("No PDFs", "No PDF files found in selected folder/subfolders.")
            return

        self._write_log(f"Found {len(pdf_paths)} PDF(s). Starting...")

        converted = 0
        skipped_existing = 0
        skipped_no_detect = 0
        errors = 0

        for idx, pdf_path in enumerate(pdf_paths, start=1):
            self._write_log(f"[{idx}/{len(pdf_paths)}] {pdf_path}")
            try:
                existing_fields = extract_existing_widget_placeholders(pdf_path)
                had_form_fields = len(existing_fields) > 0

                out_pdf = None
                placeholders: List[Dict]

                if had_form_fields:
                    placeholders = existing_fields
                    skipped_existing += 1
                    self._write_log(f"  - Existing form fields detected: {len(placeholders)} (no conversion)")
                else:
                    with fitz.open(pdf_path) as doc:
                        specs = auto_detect_fields(
                            doc,
                            use_ocr=use_ocr,
                            ocr_lang=ocr_lang,
                            text_dx=text_dx,
                            text_dy=text_dy,
                            button_dx=box_dx,
                            button_dy=box_dy,
                        )

                    if not specs:
                        placeholders = []
                        skipped_no_detect += 1
                        self._write_log("  - No static fields detected")
                    else:
                        base, _ = os.path.splitext(pdf_path)
                        out_pdf = base + "_fillable.pdf"
                        convert_pdf(pdf_path, out_pdf, specs)
                        placeholders = fieldspecs_to_placeholder_list(specs)
                        converted += 1
                        self._write_log(f"  - Converted -> {out_pdf}")

                payload = {
                    "source_pdf": pdf_path,
                    "generated_at": datetime.now().isoformat(timespec="seconds"),
                    "had_form_fields": had_form_fields,
                    "generated_fillable_pdf": out_pdf,
                    "field_counts": count_types(placeholders),
                    "placeholders": placeholders,
                }

                json_path = write_placeholder_json(pdf_path, payload)
                self._write_log(f"  - JSON -> {json_path}")

            except Exception as exc:
                errors += 1
                self._write_log(f"  - ERROR: {exc}")
                self._write_log(traceback.format_exc().strip())

        summary = (
            f"Done. total={len(pdf_paths)}, converted={converted}, "
            f"existing_fillable={skipped_existing}, no_detect={skipped_no_detect}, errors={errors}"
        )
        self._write_log(summary)
        messagebox.showinfo("Completed", summary)


if __name__ == "__main__":
    app = App()
    app.mainloop()
