#!/usr/bin/env python3
from __future__ import annotations

import json
import os
import re
import tkinter as tk
from dataclasses import dataclass
from tkinter import filedialog, messagebox
from typing import Dict, Iterable, List, Optional, Set, Tuple

try:
    import fitz  # PyMuPDF
except ImportError as exc:
    raise SystemExit("Missing dependency: PyMuPDF. Install with: pip install pymupdf") from exc


def normalize_text(v: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", (v or "").lower())


@dataclass
class FillResult:
    location: str
    value: str
    status: str
    details: str


class PDFSmartFiller:
    def __init__(self, pdf_path: str):
        self.doc = fitz.open(pdf_path)

    def close(self) -> None:
        self.doc.close()

    def save(self, output_path: str) -> None:
        if os.path.exists(output_path):
            os.remove(output_path)
        self.doc.save(output_path)

    def fill_pairs(self, pairs: List[Tuple[str, str, str]]) -> List[FillResult]:
        out: List[FillResult] = []
        for kind, location, value in pairs:
            k = (kind or "auto").strip().lower() or "auto"
            loc = (location or "").strip()
            val = (value or "").strip()
            if not loc:
                out.append(FillResult(loc, val, "skipped", "Empty location"))
                continue

            try:
                if k in {"checkbox", "radio"}:
                    hits = self._fill_buttons(k, loc, val)
                else:
                    hits = self._fill_text(k, loc, val)
            except Exception as exc:
                out.append(FillResult(loc, val, "error", str(exc)))
                continue

            if hits > 0:
                out.append(FillResult(loc, val, "filled", f"Matched {hits} field(s)"))
            else:
                out.append(FillResult(loc, val, "not found", "No field/label matched"))
        return out

    def _iter_page_widgets(self, page):
        seen = set()
        out = []
        try:
            for w in (page.widgets() or []):
                x = getattr(w, "xref", None)
                key = ("x", x) if x is not None else ("i", id(w))
                if key not in seen:
                    seen.add(key)
                    out.append(w)
        except Exception:
            pass
        try:
            w = getattr(page, "first_widget", None)
            while w is not None:
                x = getattr(w, "xref", None)
                key = ("x", x) if x is not None else ("i", id(w))
                if key not in seen:
                    seen.add(key)
                    out.append(w)
                w = getattr(w, "next", None)
        except Exception:
            pass
        return out

    def _iter_widgets(self) -> Iterable[Tuple[int, object]]:
        for i in range(self.doc.page_count):
            p = self.doc.load_page(i)
            for w in self._iter_page_widgets(p):
                yield i, w

    def _looks_like_field_id(self, s: str) -> bool:
        v = (s or "").strip().lower()
        return bool(v) and ("_" in v or any(ch.isdigit() for ch in v) or v.startswith(("txt", "fld", "field")))

    def _label_variants(self, s: str) -> List[str]:
        base = (s or "").strip()
        core = base.rstrip(":").rstrip(".").strip()
        vals = [base, core, f"{core}.", f"{core}:", f"{core}.:"]
        out = []
        seen = set()
        for v in vals:
            if not v:
                continue
            k = v.lower()
            if k in seen:
                continue
            seen.add(k)
            out.append(v)
        return out

    def _extract_parent_xref(self, src: str) -> Optional[int]:
        if not src:
            return None
        m = re.search(r"/Parent\s+(\d+)\s+0\s+R", src)
        if not m:
            return None
        try:
            return int(m.group(1))
        except Exception:
            return None

    def _pdf_obj_tokens(self, src: str, key: str) -> List[str]:
        if not src:
            return []
        out: List[str] = []
        for m in re.finditer(rf"/{key}\((.*?)\)", src, flags=re.S):
            v = (m.group(1) or "").strip()
            if v:
                out.append(v)
        return out

    def _radio_group_name(self, w) -> str:
        t = getattr(w, "field_type", None)
        radio = getattr(fitz, "PDF_WIDGET_TYPE_RADIOBUTTON", 5)
        if t != radio:
            return ""
        xref = getattr(w, "xref", None)
        if xref:
            try:
                src = self.doc.xref_object(xref, compressed=False)
            except Exception:
                src = ""
            parent_xref = self._extract_parent_xref(src)
            if parent_xref:
                try:
                    psrc = self.doc.xref_object(parent_xref, compressed=False)
                except Exception:
                    psrc = ""
                for key in ("T", "TU"):
                    toks = self._pdf_obj_tokens(psrc, key)
                    if toks:
                        return toks[0]
            for key in ("T", "TU"):
                toks = self._pdf_obj_tokens(src, key)
                if toks:
                    return toks[0]
        return ""

    def _widget_keys(self, w) -> Set[str]:
        keys: Set[str] = set()

        def add(raw: str):
            n = normalize_text(raw)
            if not n:
                return
            keys.add(n)
            for pref in ("txt", "fld", "field", "input"):
                if n.startswith(pref):
                    tail = n[len(pref):]
                    if tail:
                        keys.add(tail)
            t = re.sub(r"\d+$", "", n)
            if t:
                keys.add(t)

        add((getattr(w, "field_name", "") or "").strip())
        add((getattr(w, "field_label", "") or "").strip())
        add(self._radio_group_name(w))
        xref = getattr(w, "xref", None)
        if xref:
            try:
                src = self.doc.xref_object(xref, compressed=False)
            except Exception:
                src = ""
            for token in re.findall(r"/(?:T|TU)\((.*?)\)", src):
                add(token)
        return keys

    def _key_match(self, target: str, keys: Set[str], strict: bool) -> bool:
        for k in keys:
            if strict:
                if target == k:
                    return True
            else:
                if target == k or target in k:
                    return True
        return False

    def _type_matches(self, w, kind: str) -> bool:
        if kind in {"auto", ""}:
            return True
        t = getattr(w, "field_type", None)
        cbox = getattr(fitz, "PDF_WIDGET_TYPE_CHECKBOX", 2)
        radio = getattr(fitz, "PDF_WIDGET_TYPE_RADIOBUTTON", 5)
        text = getattr(fitz, "PDF_WIDGET_TYPE_TEXT", 7)
        if kind == "checkbox":
            return t == cbox
        if kind == "radio":
            return t == radio
        if kind == "text":
            return t == text or t not in {cbox, radio}
        return True

    def _fill_text(self, kind: str, location: str, value: str) -> int:
        if self._looks_like_field_id(location):
            # Field IDs (e.g. 25a) must be exact, otherwise 25a can
            # accidentally match 5a in loose mode.
            hits = self._fill_form_fields(kind, location, value, strict=True)
            if hits:
                return hits
        else:
            for v in self._label_variants(location):
                hits = self._fill_form_fields(kind, v, value, strict=True)
                if hits:
                    return hits
        if self._fill_by_label(location, value):
            return 1
        return self._fill_form_fields(kind, location, value, strict=False)

    def _fill_buttons(self, kind: str, location: str, value: str) -> int:
        hits = 0
        for v in self._label_variants(location):
            hits += self._fill_form_fields(kind, v, value, strict=True)
        if hits:
            return hits
        return self._fill_form_fields(kind, location, value, strict=False)

    def _fill_form_fields(self, kind: str, location: str, value: str, strict: bool) -> int:
        target = normalize_text(location)
        if not target:
            return 0
        hits = 0
        seen_radio_groups: Set[str] = set()
        for page_no, w in self._iter_widgets():
            if not self._type_matches(w, kind):
                continue
            if not self._key_match(target, self._widget_keys(w), strict):
                continue
            if kind == "radio":
                g = normalize_text(self._radio_group_name(w) or (getattr(w, "field_name", "") or ""))
                if g and g in seen_radio_groups:
                    continue
                if g:
                    seen_radio_groups.add(g)
            self._set_widget_retry(page_no, w, value, kind)
            hits += 1
        return hits

    def _find_label_rects(self, page, location: str):
        for txt in self._label_variants(location):
            for cand in (txt, txt.lower(), txt.upper(), txt.title()):
                rects = page.search_for(cand)
                if rects:
                    return rects
        raw = location.rstrip(":").rstrip(".")
        words = [w for w in re.split(r"\s+", raw) if len(w) >= 3]
        words.sort(key=len, reverse=True)
        for w in words:
            rects = page.search_for(w)
            if rects:
                return rects
        return []

    def _fill_by_label(self, location: str, value: str) -> bool:
        text_type = getattr(fitz, "PDF_WIDGET_TYPE_TEXT", 7)
        for page_no in range(self.doc.page_count):
            page = self.doc.load_page(page_no)
            anchors = self._find_label_rects(page, location)
            if not anchors:
                continue
            for a in anchors:
                best = self._nearest_text_widget(page, a, text_type)
                if best is not None:
                    self._set_widget_retry(page_no, best, value, "text")
                    return True
                rect = fitz.Rect(a.x1 + 6, a.y0 - 1, min(page.rect.x1 - 20, a.x1 + 230), a.y1 + 2)
                self._insert_fit_text(page, rect, value)
                return True
        return False

    def _nearest_text_widget(self, page, anchor, text_type):
        cands = []
        ax = (anchor.x0 + anchor.x1) / 2.0
        ay = (anchor.y0 + anchor.y1) / 2.0
        for w in self._iter_page_widgets(page):
            if getattr(w, "field_type", None) != text_type:
                continue
            r = w.rect
            wx = (r.x0 + r.x1) / 2.0
            wy = (r.y0 + r.y1) / 2.0
            d = ((wx - ax) ** 2 + (wy - ay) ** 2) ** 0.5
            if d > 320:
                continue
            same_line_right = r.x0 >= anchor.x1 - 6 and abs(wy - ay) <= 24 and wy >= ay
            immediate_below = r.y0 >= anchor.y1 - 2 and r.y0 <= anchor.y1 + 40
            below = r.y0 >= anchor.y1 - 4 and r.y0 <= anchor.y1 + 90
            above = r.y1 <= anchor.y1 + 2 and r.y0 >= anchor.y0 - 36
            if same_line_right:
                bucket = 0
            elif immediate_below:
                bucket = 1
            elif below:
                bucket = 2
            elif above:
                bucket = 3
            else:
                bucket = 4
            score = -d + min(70.0, r.width / 4.0)
            cands.append((bucket, score, w))
        if not cands:
            return None
        cands.sort(key=lambda x: (x[0], -x[1]))
        return cands[0][2]

    def _set_widget_retry(self, page_no: int, w, value: str, kind: str):
        xref = getattr(w, "xref", None)
        name = (getattr(w, "field_name", "") or "").strip()
        rect = self._safe_rect(w)
        try:
            self._set_widget(w, value, kind)
            return
        except Exception as exc:
            if "not bound to a page" not in str(exc).lower():
                raise
        wr = self._resolve_widget(page_no, xref, name, rect)
        if wr is None:
            for i in range(self.doc.page_count):
                wr = self._resolve_widget(i, xref, name, rect)
                if wr is not None:
                    break
        if wr is None:
            raise RuntimeError("Widget detached from page")
        self._set_widget(wr, value, kind)

    def _resolve_widget(self, page_no: int, xref, name: str, rect):
        page = self.doc.load_page(page_no)
        widgets = self._iter_page_widgets(page)
        if xref is not None:
            for w in widgets:
                if getattr(w, "xref", None) == xref:
                    return w
        if name:
            n = normalize_text(name)
            for w in widgets:
                if normalize_text((getattr(w, "field_name", "") or "").strip()) == n:
                    return w
        if rect is not None:
            for w in widgets:
                r = w.rect
                if abs(r.x0 - rect.x0) < 0.5 and abs(r.y0 - rect.y0) < 0.5 and abs(r.x1 - rect.x1) < 0.5 and abs(r.y1 - rect.y1) < 0.5:
                    return w
        return None

    def _safe_rect(self, w):
        try:
            return fitz.Rect(w.rect)
        except Exception:
            return None

    def _set_widget(self, w, value: str, kind: str):
        req = (kind or "auto").strip().lower()
        t = getattr(w, "field_type", None)
        cbox = getattr(fitz, "PDF_WIDGET_TYPE_CHECKBOX", 2)
        radio = getattr(fitz, "PDF_WIDGET_TYPE_RADIOBUTTON", 5)
        if req == "checkbox" or (req == "auto" and t == cbox):
            self._set_checkbox(w, value)
        elif req == "radio" or (req == "auto" and t == radio):
            self._set_radio(w, value)
        else:
            self._set_text(w, value)
        w.update()

    def _set_text(self, w, value: str):
        w.field_value = value
        if hasattr(w, "text_fontsize"):
            try:
                w.text_fontsize = self._choose_font_size(w.rect, value)
            except Exception:
                pass

    def _set_checkbox(self, w, value: str):
        v = (value or "").strip().lower()
        checked = v in {"1", "true", "yes", "y", "on", "checked", "check"} or (v not in {"0", "false", "no", "n", "off", "unchecked", "uncheck"} and bool(v))
        on_state = "Yes"
        if hasattr(w, "button_states"):
            try:
                states = w.button_states()
                vals = []
                if isinstance(states, dict):
                    for it in states.values():
                        if isinstance(it, (list, tuple)):
                            vals.extend(str(x).lstrip("/") for x in it)
                        else:
                            vals.append(str(it).lstrip("/"))
                elif isinstance(states, (list, tuple)):
                    vals.extend(str(x).lstrip("/") for x in states)
                vals = [x for x in vals if x and x.lower() != "off"]
                if vals:
                    on_state = vals[0]
            except Exception:
                pass
        w.field_value = on_state if checked else "Off"

    def _set_radio(self, w, value: str):
        target = (value or "").strip()
        if not target:
            return
        if hasattr(w, "button_states"):
            try:
                states = w.button_states()
                vals = []
                if isinstance(states, dict):
                    for it in states.values():
                        if isinstance(it, (list, tuple)):
                            vals.extend(str(x).lstrip("/") for x in it)
                        else:
                            vals.append(str(it).lstrip("/"))
                elif isinstance(states, (list, tuple)):
                    vals.extend(str(x).lstrip("/") for x in states)
                vals = [x for x in vals if x and x.lower() != "off"]
                t = normalize_text(target)
                for s in vals:
                    if normalize_text(s) == t:
                        w.field_value = s
                        return
            except Exception:
                pass
        w.field_value = target

    def _choose_font_size(self, rect, value: str) -> float:
        for s in range(12, 5, -1):
            if fitz.get_text_length(value, fontname="helv", fontsize=float(s)) <= max(10.0, rect.width - 4.0):
                return float(s)
        return 6.0

    def _insert_fit_text(self, page, rect, value: str):
        for s in range(12, 5, -1):
            sh = page.new_shape()
            spare = sh.insert_textbox(rect, value, fontsize=float(s), fontname="helv", color=(0, 0, 0), align=fitz.TEXT_ALIGN_LEFT)
            if spare >= 0:
                sh.commit()
                return
        sh = page.new_shape()
        sh.insert_textbox(rect, value, fontsize=6.0, fontname="helv", color=(0, 0, 0), align=fitz.TEXT_ALIGN_LEFT)
        sh.commit()


def load_pairs_from_json(json_path: str) -> List[Tuple[str, str, str]]:
    with open(json_path, "r", encoding="utf-8") as f:
        raw = json.load(f)

    items = None
    if isinstance(raw, dict):
        for key in ("mappings", "pairs", "placeholders", "fields", "data"):
            if isinstance(raw.get(key), list):
                items = raw[key]
                break
    elif isinstance(raw, list):
        items = raw

    pairs: List[Tuple[str, str, str]] = []
    if items is not None:
        for it in items:
            if not isinstance(it, dict):
                continue
            kind = str(it.get("type", "auto")).strip().lower() or "auto"
            location = ""
            for k in ("location", "name", "field_name", "field", "key", "label"):
                v = it.get(k)
                if isinstance(v, str) and v.strip():
                    location = v.strip()
                    break
            value = it.get("value", it.get("text", it.get("selected", it.get("choice", ""))))
            if value is None:
                value = ""
            value = str(value)
            if (not value.strip()) and kind in {"text", "auto", ""}:
                # Placeholder files for existing fillable PDFs often store a sample
                # in "label" instead of "value". Use it as a fallback.
                lv = it.get("label", "")
                if lv is not None:
                    value = str(lv)
            if kind in {"text", "auto", ""} and not value.strip():
                # Avoid wiping fields with empty placeholder rows.
                continue
            if location:
                pairs.append((kind, location, value))
        return pairs

    if isinstance(raw, dict):
        for k, v in raw.items():
            if isinstance(v, (dict, list)):
                continue
            pairs.append(("auto", str(k), "" if v is None else str(v)))
    return pairs


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDF Fill From JSON")
        self.geometry("1220x840")

        self.pdf_var = tk.StringVar()
        self.json_var = tk.StringVar()
        self.out_var = tk.StringVar()

        self.in_page = 0
        self.out_page = 0
        self.in_total_pages = 0
        self.out_total_pages = 0
        self._in_img = None
        self._out_img = None

        self._build()

    def _build(self):
        top = tk.Frame(self, padx=10, pady=10)
        top.pack(fill="x")

        tk.Label(top, text="Input PDF").grid(row=0, column=0, sticky="w")
        tk.Entry(top, textvariable=self.pdf_var, width=92).grid(row=0, column=1, sticky="we", padx=6)
        tk.Button(top, text="Browse", command=self._pick_pdf).grid(row=0, column=2)

        tk.Label(top, text="Input JSON").grid(row=1, column=0, sticky="w")
        tk.Entry(top, textvariable=self.json_var, width=92).grid(row=1, column=1, sticky="we", padx=6)
        tk.Button(top, text="Browse", command=self._pick_json).grid(row=1, column=2)

        tk.Label(top, text="Output PDF").grid(row=2, column=0, sticky="w")
        tk.Entry(top, textvariable=self.out_var, width=92).grid(row=2, column=1, sticky="we", padx=6)
        tk.Button(top, text="Browse", command=self._pick_out).grid(row=2, column=2)

        action = tk.Frame(top)
        action.grid(row=3, column=0, columnspan=3, sticky="w", pady=(8, 0))
        tk.Button(action, text="Fill + Save", command=self._run_fill).pack(side="left")

        top.grid_columnconfigure(1, weight=1)

        prev = tk.LabelFrame(self, text="PDF Viewer", padx=10, pady=10)
        prev.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        in_frame = tk.LabelFrame(prev, text="Input PDF", padx=6, pady=6)
        in_frame.grid(row=0, column=0, sticky="nsew")
        in_toolbar = tk.Frame(in_frame)
        in_toolbar.pack(fill="x")
        tk.Button(in_toolbar, text="Prev", command=lambda: self._change_page("in", -1)).pack(side="left")
        tk.Button(in_toolbar, text="Next", command=lambda: self._change_page("in", 1)).pack(side="left", padx=(6, 0))
        self.in_page_label = tk.Label(in_toolbar, text="Page 0 / 0")
        self.in_page_label.pack(side="left", padx=10)
        in_canvas_wrap = tk.Frame(in_frame)
        in_canvas_wrap.pack(fill="both", expand=True, pady=(6, 0))
        self.in_canvas = tk.Canvas(in_canvas_wrap, bg="#e8e8e8", highlightthickness=0)
        in_v = tk.Scrollbar(in_canvas_wrap, orient="vertical", command=self.in_canvas.yview)
        in_h = tk.Scrollbar(in_canvas_wrap, orient="horizontal", command=self.in_canvas.xview)
        self.in_canvas.configure(yscrollcommand=in_v.set, xscrollcommand=in_h.set)
        self.in_canvas.grid(row=0, column=0, sticky="nsew")
        in_v.grid(row=0, column=1, sticky="ns")
        in_h.grid(row=1, column=0, sticky="ew")
        in_canvas_wrap.grid_rowconfigure(0, weight=1)
        in_canvas_wrap.grid_columnconfigure(0, weight=1)

        out_frame = tk.LabelFrame(prev, text="Output PDF", padx=6, pady=6)
        out_frame.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
        out_toolbar = tk.Frame(out_frame)
        out_toolbar.pack(fill="x")
        tk.Button(out_toolbar, text="Prev", command=lambda: self._change_page("out", -1)).pack(side="left")
        tk.Button(out_toolbar, text="Next", command=lambda: self._change_page("out", 1)).pack(side="left", padx=(6, 0))
        self.out_page_label = tk.Label(out_toolbar, text="Page 0 / 0")
        self.out_page_label.pack(side="left", padx=10)
        out_canvas_wrap = tk.Frame(out_frame)
        out_canvas_wrap.pack(fill="both", expand=True, pady=(6, 0))
        self.out_canvas = tk.Canvas(out_canvas_wrap, bg="#e8e8e8", highlightthickness=0)
        out_v = tk.Scrollbar(out_canvas_wrap, orient="vertical", command=self.out_canvas.yview)
        out_h = tk.Scrollbar(out_canvas_wrap, orient="horizontal", command=self.out_canvas.xview)
        self.out_canvas.configure(yscrollcommand=out_v.set, xscrollcommand=out_h.set)
        self.out_canvas.grid(row=0, column=0, sticky="nsew")
        out_v.grid(row=0, column=1, sticky="ns")
        out_h.grid(row=1, column=0, sticky="ew")
        out_canvas_wrap.grid_rowconfigure(0, weight=1)
        out_canvas_wrap.grid_columnconfigure(0, weight=1)

        prev.grid_columnconfigure(0, weight=1)
        prev.grid_columnconfigure(1, weight=1)
        prev.grid_rowconfigure(0, weight=1)

    def _pick_pdf(self):
        p = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")])
        if not p:
            return
        self.pdf_var.set(p)
        b, e = os.path.splitext(p)
        self.out_var.set(f"{b}_filled{e}")
        self.in_page = 0
        self.out_page = 0
        self._refresh_viewers()

    def _pick_json(self):
        p = filedialog.askopenfilename(filetypes=[("JSON files", "*.json"), ("All files", "*.*")])
        if p:
            self.json_var.set(p)

    def _pick_out(self):
        p = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if p:
            self.out_var.set(p)
            self.out_page = 0
            self._refresh_viewers()

    def _run_fill(self):
        inp = self.pdf_var.get().strip()
        js = self.json_var.get().strip()
        out = self.out_var.get().strip()

        if not inp:
            messagebox.showerror("Missing input", "Choose input PDF.")
            return
        if not os.path.exists(inp):
            messagebox.showerror("Invalid input", "Input PDF not found.")
            return
        if not js:
            messagebox.showerror("Missing JSON", "Choose input JSON.")
            return
        if not os.path.exists(js):
            messagebox.showerror("Invalid JSON", "Input JSON not found.")
            return
        if not out:
            messagebox.showerror("Missing output", "Choose output PDF.")
            return

        try:
            pairs = load_pairs_from_json(js)
        except Exception as exc:
            messagebox.showerror("JSON error", f"Could not parse JSON:\n{exc}")
            return

        if not pairs:
            messagebox.showerror("No mappings", "No usable mappings found in JSON.")
            return

        filler = None
        try:
            filler = PDFSmartFiller(inp)
            results = filler.fill_pairs(pairs)
            filler.save(out)

            self.out_page = 0
            self._refresh_viewers()

            errs = sum(1 for r in results if r.status == "error")
            miss = sum(1 for r in results if r.status == "not found")
            filled = sum(1 for r in results if r.status == "filled")

            if errs:
                messagebox.showwarning("Done with errors", f"Filled={filled}, not_found={miss}, errors={errs}")
            elif miss:
                messagebox.showwarning("Done with missing", f"Filled={filled}, not_found={miss}, errors={errs}")
            else:
                messagebox.showinfo("Done", f"PDF filled successfully. Fields filled: {filled}")
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
        finally:
            if filler:
                filler.close()

    def _change_page(self, which: str, delta: int):
        if which == "in":
            if self.in_total_pages <= 0:
                return
            self.in_page = max(0, min(self.in_total_pages - 1, self.in_page + delta))
        else:
            if self.out_total_pages <= 0:
                return
            self.out_page = max(0, min(self.out_total_pages - 1, self.out_page + delta))
        self._refresh_viewers()

    def _refresh_viewers(self):
        self._render_pdf_to_canvas("in", self.pdf_var.get().strip(), self.in_page)
        self._render_pdf_to_canvas("out", self.out_var.get().strip(), self.out_page)

    def _render_pdf_to_canvas(self, which: str, path: str, page_index: int):
        canvas = self.in_canvas if which == "in" else self.out_canvas
        label = self.in_page_label if which == "in" else self.out_page_label

        canvas.delete("all")
        if not path:
            label.config(text="Page 0 / 0")
            canvas.create_text(20, 20, anchor="nw", text="No file selected.", fill="#333333")
            return
        if not os.path.exists(path):
            label.config(text="Page 0 / 0")
            canvas.create_text(20, 20, anchor="nw", text=f"File not found:\n{path}", fill="#333333")
            return
        if not path.lower().endswith(".pdf"):
            label.config(text="Page 0 / 0")
            canvas.create_text(20, 20, anchor="nw", text=f"Not a PDF file:\n{path}", fill="#333333")
            return

        doc = None
        try:
            doc = fitz.open(path)
            total = doc.page_count
            if which == "in":
                self.in_total_pages = total
                self.in_page = max(0, min(max(total - 1, 0), page_index))
                page_index = self.in_page
            else:
                self.out_total_pages = total
                self.out_page = max(0, min(max(total - 1, 0), page_index))
                page_index = self.out_page

            if total <= 0:
                label.config(text="Page 0 / 0")
                canvas.create_text(20, 20, anchor="nw", text="PDF has no pages.", fill="#333333")
                return

            page = doc.load_page(page_index)
            pix = page.get_pixmap(matrix=fitz.Matrix(1.35, 1.35), alpha=False)
            img = tk.PhotoImage(data=pix.tobytes("ppm"))
            canvas.create_image(0, 0, anchor="nw", image=img)
            canvas.config(scrollregion=(0, 0, img.width(), img.height()))
            label.config(text=f"Page {page_index + 1} / {total}")
            if which == "in":
                self._in_img = img
            else:
                self._out_img = img
        except Exception as exc:
            label.config(text="Page 0 / 0")
            canvas.create_text(20, 20, anchor="nw", text=f"Could not render PDF:\n{exc}", fill="#333333")
        finally:
            if doc is not None:
                doc.close()


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
