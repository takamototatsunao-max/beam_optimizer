#!/usr/bin/env python3
"""
Extract wide/medium/narrow H-section rows from chapter-5 PDF and emit CSV.

Default target pages are 2..7 (chapter table pages for H-sections).
"""

from __future__ import annotations

import argparse
import csv
import re
import sys
from pathlib import Path
from typing import Dict, List, Optional, Tuple

try:
    import fitz  # PyMuPDF
except Exception as exc:  # pragma: no cover
    raise SystemExit(f"PyMuPDF is required: {exc}")


NUM_RE = re.compile(r"^[0-9][0-9,]*(?:\.[0-9]+)?$")
NOM_RE = re.compile(r"^\d+[^0-9,.\-]+\d+$")
COL = {
    "nominal": 30.0,
    "h": 74.0,
    "b": 95.0,
    "tw": 120.0,
    "tf": 140.0,
    "r_mm": 159.0,
    "a_cm2": 179.0,
    "unit_mass": 213.0,
    "ix_cm4": 240.0,
    "zx_cm3": 357.0,
}


def _to_float(tok: Optional[str]) -> Optional[float]:
    if tok is None:
        return None
    try:
        return float(str(tok).replace(",", ""))
    except Exception:
        return None


def _pick_numeric(band_words: List[Tuple], x_center: float, tol: float = 10.0) -> Optional[str]:
    cands: List[Tuple[float, str]] = []
    for w in band_words:
        x0, _y0, _x1, _y1, text, *_ = w
        if abs(float(x0) - x_center) <= tol and NUM_RE.match(str(text)):
            cands.append((abs(float(x0) - x_center), str(text)))
    if not cands:
        return None
    cands.sort(key=lambda x: x[0])
    return cands[0][1]


def _pick_numeric_multi(band_words: List[Tuple], x_centers: List[float], tol: float = 10.0) -> Optional[str]:
    best: Optional[Tuple[float, str]] = None
    for xc in x_centers:
        tok = _pick_numeric(band_words, xc, tol=tol)
        if tok is None:
            continue
        for w in band_words:
            x0, _y0, _x1, _y1, text, *_ = w
            if str(text) == str(tok):
                d = abs(float(x0) - xc)
                if best is None or d < best[0]:
                    best = (d, tok)
                break
    return None if best is None else best[1]


def _series_of_page(page_no: int) -> str:
    if page_no in (2, 3):
        return "WIDE"
    if page_no in (4, 5):
        return "MEDIUM"
    if page_no in (6, 7):
        return "NARROW"
    return "H"


def _format_g(v: float) -> str:
    return f"{float(v):g}"


def _section_name(h: float, b: float, tw: float, tf: float) -> str:
    return f"H-{_format_g(h)}x{_format_g(b)}x{_format_g(tw)}x{_format_g(tf)}"


def _parse_page_rows(doc, page_no: int) -> List[Dict[str, object]]:
    page = doc[page_no - 1]
    words = page.get_text("words")

    ys: List[float] = []
    for w in words:
        x0, y0, _x1, _y1, text, *_ = w
        if abs(float(x0) - COL["a_cm2"]) <= 10.0 and NUM_RE.match(str(text)) and 95.0 <= float(y0) <= 380.0:
            ys.append(float(y0))
    ys.sort()

    merged_ys: List[float] = []
    for y in ys:
        if not merged_ys or abs(y - merged_ys[-1]) > 1.2:
            merged_ys.append(y)

    out: List[Dict[str, object]] = []
    last_nominal: Optional[str] = None
    for y in merged_ys:
        band = [w for w in words if abs(float(w[1]) - y) <= 0.9]
        nominal = None
        for w in band:
            x0, _y0, _x1, _y1, text, *_ = w
            t = str(text)
            if abs(float(x0) - COL["nominal"]) <= 10.0 and NOM_RE.match(t):
                nominal = t.replace("\u00d7", "x").replace("X", "x")
                break
        if nominal:
            last_nominal = nominal
        if not last_nominal:
            continue

        h = _to_float(_pick_numeric(band, COL["h"]))
        b = _to_float(_pick_numeric(band, COL["b"]))
        tw = _to_float(_pick_numeric(band, COL["tw"]))
        tf = _to_float(_pick_numeric(band, COL["tf"]))
        r_mm = _to_float(_pick_numeric(band, COL["r_mm"], tol=12.0))
        a_cm2 = _to_float(_pick_numeric(band, COL["a_cm2"], tol=10.0))
        ix_cm4 = _to_float(_pick_numeric_multi(band, [COL["ix_cm4"], 250.0], tol=12.5))
        zx_cm3 = _to_float(_pick_numeric(band, COL["zx_cm3"]))
        unit_mass = _to_float(_pick_numeric(band, COL["unit_mass"]))

        if None in (h, b, tw, tf, a_cm2, zx_cm3):
            continue

        out.append({
            "Series": _series_of_page(page_no),
            "SourcePage": page_no,
            "Nominal": last_nominal,
            "h_mm": h,
            "b_mm": b,
            "tw_mm": tw,
            "tf_mm": tf,
            "r_mm": r_mm,
            "A_cm2": a_cm2,
            "Ix_cm4": ix_cm4,
            "Zx_cm3": zx_cm3,
            "UnitMass_kg_m": unit_mass,
        })
    return out


def _dedupe_rows(rows: List[Dict[str, object]]) -> List[Dict[str, object]]:
    best_by_key: Dict[Tuple[float, float, float, float], Dict[str, object]] = {}
    for row in rows:
        key = (
            float(row["h_mm"]),
            float(row["b_mm"]),
            float(row["tw_mm"]),
            float(row["tf_mm"]),
        )
        score = int(row.get("Ix_cm4") is not None) + int(row.get("UnitMass_kg_m") is not None)
        prev = best_by_key.get(key)
        if prev is None:
            best_by_key[key] = row
            continue
        prev_score = int(prev.get("Ix_cm4") is not None) + int(prev.get("UnitMass_kg_m") is not None)
        if score > prev_score:
            best_by_key[key] = row

    out = list(best_by_key.values())
    for row in out:
        area_cm2 = float(row["A_cm2"])
        unit_mass = row.get("UnitMass_kg_m")
        if unit_mass is not None:
            wg = float(unit_mass) * 9.80665 / 1000.0
        else:
            wg = area_cm2 * 100.0 * 1e-6 * 76.98
        row["w_g_kN_m"] = wg
        row["SectionName"] = _section_name(
            float(row["h_mm"]),
            float(row["b_mm"]),
            float(row["tw_mm"]),
            float(row["tf_mm"]),
        )

    out.sort(key=lambda d: (float(d["w_g_kN_m"]), str(d["SectionName"])))
    for i, row in enumerate(out, start=1):
        row["Rank"] = i
    return out


def _parse_pages_arg(text: str) -> List[int]:
    text = str(text).strip()
    if "-" in text:
        a, b = text.split("-", 1)
        ia = int(a)
        ib = int(b)
        if ia > ib:
            ia, ib = ib, ia
        return list(range(ia, ib + 1))
    out: List[int] = []
    for tok in text.split(","):
        tok = tok.strip()
        if tok:
            out.append(int(tok))
    return out


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", required=True, help="source PDF path")
    ap.add_argument("--out", required=True, help="output CSV path")
    ap.add_argument("--pages", default="2-7", help="target pages (default: 2-7)")
    args = ap.parse_args()

    pdf_path = Path(args.pdf)
    out_path = Path(args.out)
    pages = _parse_pages_arg(args.pages)

    if not pdf_path.exists():
        raise SystemExit(f"PDF not found: {pdf_path}")

    doc = fitz.open(str(pdf_path))
    all_rows: List[Dict[str, object]] = []
    for p in pages:
        if 1 <= p <= doc.page_count:
            all_rows.extend(_parse_page_rows(doc, p))

    rows = _dedupe_rows(all_rows)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    fields = [
        "Rank",
        "SectionName",
        "Series",
        "SourcePage",
        "Nominal",
        "h_mm",
        "b_mm",
        "tw_mm",
        "tf_mm",
        "r_mm",
        "A_cm2",
        "Ix_cm4",
        "Zx_cm3",
        "UnitMass_kg_m",
        "w_g_kN_m",
    ]
    with out_path.open("w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        for row in rows:
            w.writerow(row)

    print(f"rows={len(rows)} written={out_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
