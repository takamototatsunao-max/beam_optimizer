#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Beam layout enumerator and PDF renderer (ReportLab)

Implements the user's rules:
- Start from a rectangle (0..Lx, 0..Ly). Stop splitting when short side <= threshold.
- At each step choose split mode: LONG or SHORT, and ratio k=2 or 3.
  - LONG: split along the rectangle's long dimension (reduces long side).
  - SHORT: split along the rectangle's short dimension (reduces short side).
  - A decision (mode,k) is applied to *all* active rectangles in that step.
- Pruning: if, for a given state and mode, k=2 already makes every *resulting* sub-rectangle satisfy short<=threshold,
  then k=3 for that mode is not enumerated.
- Beam types: X = horizontal (y = const), Y = vertical (x = const).
- Beam segments are created per rectangle; later beams terminate at intersections with previously created beams
  because their endpoints are the rectangle boundaries at insertion time.
- Pin markers are drawn ONLY on the inserted beam ends (both ends of each inserted segment).

Output:
- Multi-page PDF, one page per enumerated case.
- Each page: diagram (step-colored) + coordinate table (ID/Step/Type/Const/Start/End).

Usage:
  python beam_layout_enum.py --Lx 14 --Ly 9 --threshold 3 --out beam_layouts.pdf

"""

from __future__ import annotations

import math
import argparse
import datetime
from dataclasses import dataclass
from typing import List, Tuple

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.pdfbase.pdfmetrics import stringWidth


# ----------------------------
# Geometry / data structures
# ----------------------------

@dataclass(frozen=True)
class Rect:
    x0: float
    x1: float
    y0: float
    y1: float

    def w(self) -> float:
        return self.x1 - self.x0

    def h(self) -> float:
        return self.y1 - self.y0

    def short(self) -> float:
        return min(self.w(), self.h())

    def long(self) -> float:
        return max(self.w(), self.h())


@dataclass
class BeamSeg:
    id: str
    step: int
    typ: str  # 'X' (horizontal y const) or 'Y' (vertical x const)
    const: float
    x0: float
    y0: float
    x1: float
    y1: float


STEP_COLORS = [
    colors.HexColor("#1f77b4"),  # blue
    colors.HexColor("#ff7f0e"),  # orange
    colors.HexColor("#2ca02c"),  # green
    colors.HexColor("#d62728"),  # red
    colors.HexColor("#9467bd"),  # purple
    colors.HexColor("#8c564b"),  # brown
    colors.HexColor("#e377c2"),  # pink
    colors.HexColor("#7f7f7f"),  # gray
    colors.HexColor("#bcbd22"),  # olive
    colors.HexColor("#17becf"),  # cyan
]


# ----------------------------
# Enumeration logic
# ----------------------------

def split_rect(rect: Rect, mode: str, k: int) -> Tuple[List[Rect], List[Tuple[str, float, Rect]]]:
    """Split a rectangle and return (subrects, beam_specs).

    beam_specs: list of (typ, const, parent_rect) for each inserted beam segment.
    """
    w, h = rect.w(), rect.h()

    if mode == "LONG":
        split_along_x = (w >= h)  # reduce X if X is long
    elif mode == "SHORT":
        split_along_x = (w <= h)  # reduce X if X is short
    else:
        raise ValueError(f"Unknown mode: {mode}")

    subrects: List[Rect] = []
    beams: List[Tuple[str, float, Rect]] = []

    if split_along_x:
        xs = [rect.x0 + w * i / k for i in range(1, k)]
        x_edges = [rect.x0] + xs + [rect.x1]
        for a, b in zip(x_edges[:-1], x_edges[1:]):
            subrects.append(Rect(a, b, rect.y0, rect.y1))
        for x in xs:
            beams.append(("Y", x, rect))
    else:
        ys = [rect.y0 + h * i / k for i in range(1, k)]
        y_edges = [rect.y0] + ys + [rect.y1]
        for a, b in zip(y_edges[:-1], y_edges[1:]):
            subrects.append(Rect(rect.x0, rect.x1, a, b))
        for y in ys:
            beams.append(("X", y, rect))

    return subrects, beams


def apply_step(
    rects: List[Rect],
    beams: List[BeamSeg],
    step: int,
    mode: str,
    k: int,
    threshold: float,
) -> Tuple[List[Rect], List[BeamSeg]]:
    new_rects: List[Rect] = []
    new_beams: List[BeamSeg] = list(beams)

    next_id = len(new_beams) + 1
    for r in rects:
        if r.short() <= threshold:
            new_rects.append(r)
            continue

        subs, bspecs = split_rect(r, mode, k)
        new_rects.extend(subs)

        for typ, const, rr in bspecs:
            if typ == "Y":
                seg = BeamSeg(f"B{next_id}", step, "Y", const, const, rr.y0, const, rr.y1)
            else:
                seg = BeamSeg(f"B{next_id}", step, "X", const, rr.x0, const, rr.x1, const)
            new_beams.append(seg)
            next_id += 1

    return new_rects, new_beams


def allow_k3(rects: List[Rect], mode: str, threshold: float) -> bool:
    """Pruning rule:
    If k=2 already makes every resulting sub-rectangle satisfy short<=threshold for this mode, then k=3 is unnecessary.
    We return True if k=3 should be considered (i.e., there exists at least one active rect where k=2 does NOT finish).
    """
    for r in rects:
        if r.short() <= threshold:
            continue
        subs, _ = split_rect(r, mode, 2)
        if any(sr.short() > threshold for sr in subs):
            return True
    return False


def enumerate_cases(Lx: float, Ly: float, threshold: float, max_steps: int = 12):
    start_rects = [Rect(0.0, Lx, 0.0, Ly)]
    cases = []

    stack = [(start_rects, [], [])]  # (rects, beams, decisions)
    while stack:
        rects, beams, dec = stack.pop()

        if all(r.short() <= threshold for r in rects):
            cases.append((dec, beams, rects))
            continue
        if len(dec) >= max_steps:
            continue

        for mode in ("LONG", "SHORT"):
            ks = [2]
            if allow_k3(rects, mode, threshold):
                ks.append(3)
            for k in ks:
                new_rects, new_beams = apply_step(rects, beams, len(dec) + 1, mode, k, threshold)
                stack.append((new_rects, new_beams, dec + [(mode, k)]))

    cases.sort(key=lambda x: (len(x[0]), x[0]))
    return cases



# ----------------------------
# De-duplication
# ----------------------------

def _round(x: float, tol: float) -> float:
    if tol <= 0:
        return x
    return round(x / tol) * tol


def layout_signature(beams: List[BeamSeg], tol: float = 1e-6) -> Tuple[Tuple, ...]:
    """Return a hashable signature for a beam layout, ignoring IDs and steps.

    Equivalence definition:
    - Same set of geometric beam segments (typ + endpoints), within tolerance.
    - Step (generation order) and IDs are ignored.

    Note:
    - We do NOT merge collinear adjacent segments; segments are compared as-is.
      If you later introduce segment merging, do it here to preserve equivalence.
    """
    sig = []
    for b in beams:
        x0 = _round(b.x0, tol); y0 = _round(b.y0, tol)
        x1 = _round(b.x1, tol); y1 = _round(b.y1, tol)
        # canonical endpoint ordering
        if (x1, y1) < (x0, y0):
            x0, y0, x1, y1 = x1, y1, x0, y0
        sig.append((b.typ, x0, y0, x1, y1))
    sig.sort()
    return tuple(sig)


def deduplicate_cases(cases, tol: float = 1e-6):
    """Return (unique_cases, dup_groups).

    dup_groups: list of dicts: {"signature_index": i, "members": [case_indices], "decisions": [...]}
    case_indices are 1-based in the original ordering of `cases`.
    """
    sig_to_first = {}
    unique = []
    groups = {}
    for idx, (dec, beams, rects) in enumerate(cases, start=1):
        sig = layout_signature(beams, tol=tol)
        if sig not in sig_to_first:
            sig_to_first[sig] = idx
            unique.append((dec, beams, rects))
            groups[idx] = [idx]
        else:
            first = sig_to_first[sig]
            groups[first].append(idx)

    dup_groups = []
    for first_idx, members in groups.items():
        if len(members) > 1:
            dup_groups.append({
                "signature_index": first_idx,
                "members": members,
                "decisions": [cases[i-1][0] for i in members],
            })
    return unique, dup_groups

# ----------------------------
# PDF rendering
# ----------------------------

def world_to_canvas(x: float, y: float, origin: Tuple[float, float], scale: float) -> Tuple[float, float]:
    ox, oy = origin
    return ox + x * scale, oy + y * scale


def draw_pin(c: canvas.Canvas, x: float, y: float, x2: float, y2: float, inset: float, r: float) -> None:
    """Draw a filled circle slightly inset from endpoint (x,y) towards (x2,y2)."""
    dx, dy = x2 - x, y2 - y
    L = math.hypot(dx, dy)
    if L == 0:
        return
    ux, uy = dx / L, dy / L
    px, py = x + ux * inset, y + uy * inset
    c.circle(px, py, r, stroke=0, fill=1)


def render_pdf(outpath: str, Lx: float, Ly: float, cases, threshold: float) -> None:
    pagesize = landscape(A4)
    W, H = pagesize
    c = canvas.Canvas(outpath, pagesize=pagesize)

    margin = 12 * mm
    header_h = 18 * mm
    footer_h = 10 * mm
    table_w = 115 * mm
    gap = 6 * mm

    diagram_w = W - 2 * margin - table_w - gap
    diagram_h = H - 2 * margin - header_h - footer_h
    scale = min(diagram_w / Lx, diagram_h / Ly)
    diag_origin = (margin, margin + footer_h)

    col_widths = [16 * mm, 10 * mm, 10 * mm, 16 * mm, 30 * mm, 30 * mm]  # sum ~112mm

    for case_idx, (dec, beams, _rects) in enumerate(cases, start=1):
        # Header
        c.setFillColor(colors.black)
        c.setFont("Helvetica-Bold", 13)
        c.drawString(margin, H - margin - 14, f"Beam layout enumeration  Case {case_idx}/{len(cases)}")
        c.setFont("Helvetica", 9)
        c.drawString(margin, H - margin - 28, f"Span: {Lx:.2f}m (X) x {Ly:.2f}m (Y)   Stop: short side <= {threshold:.2f}m")
        dec_str = " -> ".join([f"{m}:{k}" for m, k in dec]) if dec else "(no split)"
        c.drawString(margin, H - margin - 40, f"Step decisions: {dec_str}")

        # Legend
        beams_steps = max([b.step for b in beams], default=1)
        c.setFont("Helvetica", 8)
        leg_x = margin
        leg_y = H - margin - 52
        c.drawString(leg_x, leg_y, "Step colors:")
        lx = leg_x + 55
        for s in range(1, min(6, beams_steps) + 1):
            c.setStrokeColor(STEP_COLORS[(s - 1) % len(STEP_COLORS)])
            c.setLineWidth(2)
            c.line(lx, leg_y + 3, lx + 18, leg_y + 3)
            c.setFillColor(colors.black)
            c.drawString(lx + 22, leg_y, f"S{s}")
            lx += 45

        # Diagram frame
        ox, oy = diag_origin
        c.setStrokeColor(colors.black)
        c.setLineWidth(0.8)
        c.rect(ox, oy, Lx * scale, Ly * scale, stroke=1, fill=0)

        # Axis ticks (every 2m)
        c.setLineWidth(0.4)
        c.setFont("Helvetica", 6)
        for xm in range(0, int(math.floor(Lx)) + 1, 2):
            tx, _ = world_to_canvas(xm, 0, diag_origin, scale)
            c.line(tx, oy, tx, oy - 3)
            c.drawCentredString(tx, oy - 10, str(xm))
        for ym in range(0, int(math.floor(Ly)) + 1, 2):
            _, ty = world_to_canvas(0, ym, diag_origin, scale)
            c.line(ox, ty, ox - 3, ty)
            c.drawRightString(ox - 5, ty - 2, str(ym))

        # Beams
        beams_sorted = sorted(beams, key=lambda b: (b.step, int(b.id[1:]) if b.id[1:].isdigit() else 10**9))
        c.setLineWidth(1.4)

        used_boxes = []
        def place_label(text: str, x: float, y: float) -> None:
            c.setFont("Helvetica", 7)
            tw = stringWidth(text, "Helvetica", 7)
            th = 8
            candidates = [(0, 0), (8, 8), (-8, 8), (8, -8), (-8, -8), (12, 0), (-12, 0), (0, 12), (0, -12)]
            for dx, dy in candidates:
                bx, by = x + dx, y + dy
                box = (bx, by, bx + tw, by + th)
                if any(not (box[2] < ub[0] or box[0] > ub[2] or box[3] < ub[1] or box[1] > ub[3]) for ub in used_boxes):
                    continue
                used_boxes.append(box)
                c.setFillColor(colors.black)
                c.drawString(bx, by, text)
                return
            c.setFillColor(colors.black)
            c.drawString(x, y, text)

        for b in beams_sorted:
            col = STEP_COLORS[(b.step - 1) % len(STEP_COLORS)]
            c.setStrokeColor(col)
            c.setFillColor(col)

            cx0, cy0 = world_to_canvas(b.x0, b.y0, diag_origin, scale)
            cx1, cy1 = world_to_canvas(b.x1, b.y1, diag_origin, scale)
            c.line(cx0, cy0, cx1, cy1)

            inset = 0.12 * scale
            r = 1.5
            draw_pin(c, cx0, cy0, cx1, cy1, inset=inset, r=r)
            draw_pin(c, cx1, cy1, cx0, cy0, inset=inset, r=r)

            lbl = f"{b.id}  {'y' if b.typ=='X' else 'x'}={b.const:.2f}"
            mx, my = (cx0 + cx1) / 2, (cy0 + cy1) / 2
            place_label(lbl, mx + 2, my + 2)

        # Table
        table_x = margin + Lx * scale + gap
        table_y = margin + footer_h
        table_h = diagram_h

        c.setStrokeColor(colors.black)
        c.setLineWidth(0.8)
        c.rect(table_x, table_y, table_w, table_h, stroke=1, fill=0)

        # Title with white background to avoid grid overlap
        c.setFillColor(colors.white)
        c.rect(table_x + 1, table_y + table_h - 18, table_w - 2, 14, stroke=0, fill=1)
        c.setFillColor(colors.black)
        c.setFont("Helvetica-Bold", 9)
        c.drawString(table_x + 4, table_y + table_h - 14, "Beam coordinate list (m)")

        start_y = table_y + table_h - 40
        headers = ["ID", "Step", "Type", "Const", "Start (x,y)", "End (x,y)"]
        x0 = table_x + 3
        xs = [x0]
        for wcol in col_widths:
            xs.append(xs[-1] + wcol)

        c.setFont("Helvetica-Bold", 7)
        c.setLineWidth(0.4)
        for xv in xs:
            c.line(xv, table_y + 10, xv, start_y + 18)
        c.line(table_x, start_y + 18, table_x + table_w, start_y + 18)

        for i, hdr in enumerate(headers):
            c.drawString(xs[i] + 2, start_y + 22, hdr)

        c.setFont("Helvetica", 7)
        y = start_y + 18
        row_h = 9
        max_rows = int((y - (table_y + 12)) / row_h)

        c.setStrokeColor(colors.HexColor("#dddddd"))
        for b in beams_sorted[:max_rows]:
            y -= row_h
            c.setFillColor(colors.black)
            c.drawString(xs[0] + 2, y + 2, b.id)
            c.drawString(xs[1] + 2, y + 2, str(b.step))
            c.drawString(xs[2] + 2, y + 2, b.typ)
            c.drawRightString(xs[4] - 4, y + 2, f"{b.const:.2f}")
            c.drawString(xs[4] + 2, y + 2, f"({b.x0:.2f},{b.y0:.2f})")
            c.drawString(xs[5] + 2, y + 2, f"({b.x1:.2f},{b.y1:.2f})")
            c.line(table_x, y, table_x + table_w, y)

        c.setStrokeColor(colors.black)
        if len(beams_sorted) > max_rows:
            c.setFont("Helvetica-Oblique", 7)
            c.drawString(table_x + 4, table_y + 4, f"... {len(beams_sorted) - max_rows} more beams omitted")

        # Footer
        c.setFont("Helvetica", 8)
        c.setFillColor(colors.black)
        c.drawRightString(W - margin, margin - 2, f"Generated {datetime.date.today().isoformat()}  Page {case_idx}")

        c.showPage()

    c.save()


# ----------------------------
# CLI
# ----------------------------

def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--Lx", type=float, required=True, help="Span length in X [m]")
    ap.add_argument("--Ly", type=float, required=True, help="Span length in Y [m]")
    ap.add_argument("--threshold", type=float, default=3.0, help="Stop when short side <= threshold [m]")
    ap.add_argument("--dedup", action="store_true", default=True, help="Remove duplicate final layouts (default: on)")
    ap.add_argument("--no-dedup", dest="dedup", action="store_false", help="Keep duplicates (debug)")
    ap.add_argument("--dedup_tol", type=float, default=1e-6, help="Tolerance for duplicate detection [m]")
    ap.add_argument("--max_steps", type=int, default=12, help="Safety cap on recursion depth")
    ap.add_argument("--out", type=str, default="beam_layouts.pdf", help="Output PDF path")
    args = ap.parse_args()

    cases = enumerate_cases(args.Lx, args.Ly, args.threshold, max_steps=args.max_steps)
    if not cases:

        raise RuntimeError("No cases generated. Increase --max_steps or check parameters.")

    out_cases = cases
    dup_groups = []
    if args.dedup:
        out_cases, dup_groups = deduplicate_cases(cases, tol=args.dedup_tol)

    render_pdf(args.out, args.Lx, args.Ly, out_cases, args.threshold)

    # Duplicate report (text)
    if args.dedup and dup_groups:
        rep_path = args.out.rsplit('.', 1)[0] + "_dedup_report.txt"
        with open(rep_path, "w", encoding="utf-8") as f:
            f.write(f"Original cases: {len(cases)}\\n")
            f.write(f"Unique layouts : {len(out_cases)}\\n")
            f.write(f"Tolerance [m]  : {args.dedup_tol}\\n\\n")
            for g in dup_groups:
                f.write(f"Group (kept case #{g['signature_index']}): members={g['members']}\\n")
                for j, d in zip(g['members'], g['decisions']):
                    ds = ' -> '.join([f"{m}:{k}" for m,k in d]) if d else '(no split)'
                    f.write(f"  case #{j}: {ds}\\n")
                f.write("\n")
        print(f"Dedup report written to: {rep_path}")

    # lightweight verification
    print(f"Original cases: {len(cases)}")
    print(f"Output cases  : {len(out_cases)}")
    print(f"PDF written to: {args.out}")
    print(f"First output steps: {out_cases[0][0]}")
    print(f"Last  output steps: {out_cases[-1][0]}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())