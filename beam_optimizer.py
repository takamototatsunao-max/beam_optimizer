#!/usr/bin/env python3
from __future__ import annotations

import argparse
import logging
import math
import random
import tempfile
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image as XLImage

G = 9.80665
EPS = 1e-9

LOG = logging.getLogger("beam_optimizer")


@dataclass(frozen=True)
class Parameters:
    Lx: float
    Ly: float
    q: float
    E: float
    density: float
    allow_sigma: float
    allow_tau: float
    deflection_limit_ratio: float
    unit_system: str
    objective: str
    random_seed: int


@dataclass(frozen=True)
class Options:
    search_mode: str
    time_limit_sec: Optional[float]
    parallel: bool
    n_workers: int
    prune_level: int


@dataclass(frozen=True)
class PointLoad:
    name: str
    P: float
    x: float
    y: float


@dataclass(frozen=True)
class PointLoadLocal:
    name: str
    P: float
    x: float
    y: float


@dataclass(frozen=True)
class Section:
    name: str
    sec_type: str
    A: float
    I: float
    Z: float
    unit_weight: float
    cost_per_m: float

    @property
    def self_weight_kN_m(self) -> float:
        return self.unit_weight * G / 1000.0

    def metric(self, objective: str) -> float:
        if objective == "weight":
            return self.unit_weight
        return self.cost_per_m


@dataclass(frozen=True)
class CandSpec:
    row_id: str
    direction: str
    pitch_start: float
    pitch_end: float
    pitch_step: float
    offset_options: str
    max_count: int
    allowed_sections: Tuple[str, ...]
    min_edge_clear: float
    allow_mixed_xy: bool


@dataclass(frozen=True)
class PatternOption:
    key: str
    direction: str
    positions: Tuple[float, ...]
    allowed_sections: Tuple[str, ...]
    allow_mixed_xy: bool


@dataclass(frozen=True)
class LevelLayoutOption:
    key: str
    x_pattern: Optional[PatternOption]
    y_pattern: Optional[PatternOption]


@dataclass(frozen=True)
class EdgeTypes:
    left: str
    right: str
    bottom: str
    top: str


@dataclass(frozen=True)
class BoundaryLoad:
    edge: str
    s: float
    P: float
    source: str


@dataclass(frozen=True)
class AppliedPointLoad:
    P: float
    a: float
    source: str
    x: float
    y: float


@dataclass
class BeamResponse:
    Ra: float
    Rb: float
    Mmax: float
    Vmax: float
    ymax_mm: float
    sigma: float
    tau: float
    utilization: float
    x_samples: List[float]
    V_samples: List[float]
    M_samples: List[float]
    y_samples_mm: List[float]


@dataclass
class BeamRecord:
    level: int
    direction: str
    pos: float
    span_start: float
    span_end: float
    length: float
    allowed_sections: Tuple[str, ...]
    section_name: str
    unit_weight: float
    cost_per_m: float
    w_line: float
    point_loads: List[AppliedPointLoad]
    response: BeamResponse
    beam_id: str = ""


@dataclass
class TraceRow:
    candidate_key: str
    objective: float
    max_util: float
    infeasible_reason: str
    elapsed_sec: float


@dataclass
class PanelResult:
    feasible: bool
    objective: float
    total_weight: float
    total_cost: float
    max_util: float
    beams: List[BeamRecord]
    boundary_loads: List[BoundaryLoad]
    layout_key: str
    infeasible_reason: str = ""

    @property
    def beam_count(self) -> int:
        return len(self.beams)


@dataclass
class LocalBeamProto:
    beam_id: str
    level: int
    direction: str
    pos: float
    x0: float
    x1: float
    y0: float
    y1: float
    length: float
    allowed_sections: Tuple[str, ...]


@dataclass
class LocalPanel:
    x0: float
    x1: float
    y0: float
    y1: float
    left_id: str
    right_id: str
    bottom_id: str
    top_id: str
    edge_types: EdgeTypes

    @property
    def width(self) -> float:
        return self.x1 - self.x0

    @property
    def height(self) -> float:
        return self.y1 - self.y0


@dataclass
class SearchStats:
    l1_candidates_total: int = 0
    l1_candidates_evaluated: int = 0
    l1_candidates_pruned: int = 0
    l1_candidates_feasible: int = 0
    l2_cache_hit: int = 0
    l2_cache_miss: int = 0
    l3_cache_hit: int = 0
    l3_cache_miss: int = 0


@dataclass
class SearchContext:
    params: Parameters
    options: Options
    section_map: Dict[str, Section]
    all_sections: List[Section]
    cand_l1: List[CandSpec]
    cand_l2: List[CandSpec]
    cand_l3: List[CandSpec]
    traces: List[TraceRow] = field(default_factory=list)
    stats: SearchStats = field(default_factory=SearchStats)
    cache_l2: Dict[Tuple, PanelResult] = field(default_factory=dict)
    cache_l3: Dict[Tuple, PanelResult] = field(default_factory=dict)
    started_at: float = field(default_factory=time.perf_counter)

    def elapsed(self) -> float:
        return time.perf_counter() - self.started_at

    def check_timeout(self) -> None:
        if self.options.time_limit_sec is None:
            return
        if self.elapsed() > self.options.time_limit_sec:
            raise TimeoutError(f"time limit exceeded: {self.options.time_limit_sec}s")


@dataclass
class OptimizationResult:
    feasible: bool
    beams: List[BeamRecord]
    objective: float
    total_weight: float
    total_cost: float
    max_util: float
    worst_beam_id: str
    traces: List[TraceRow]
    stats: SearchStats
    elapsed_sec: float
    timed_out: bool
    message: str


def normalize_header(v: object) -> str:
    if v is None:
        return ""
    return str(v).strip().lower().replace(" ", "")


def to_bool(v: object, default: bool = False) -> bool:
    if v is None:
        return default
    s = str(v).strip().lower()
    if s in {"1", "true", "yes", "y", "on"}:
        return True
    if s in {"0", "false", "no", "n", "off"}:
        return False
    return default


def to_float(v: object, name: str) -> float:
    if v is None or str(v).strip() == "":
        raise ValueError(f"{name} is required")
    try:
        return float(v)
    except Exception as exc:
        raise ValueError(f"{name} must be numeric: {v}") from exc


def to_int(v: object, name: str) -> int:
    if v is None or str(v).strip() == "":
        raise ValueError(f"{name} is required")
    try:
        return int(float(v))
    except Exception as exc:
        raise ValueError(f"{name} must be integer-like: {v}") from exc


def split_list_text(v: object) -> Tuple[str, ...]:
    if v is None:
        return tuple()
    s = str(v).strip()
    if not s:
        return tuple()
    return tuple(x.strip() for x in s.split(";") if x.strip())


def read_table_rows(ws, required_cols: Sequence[str]) -> List[Dict[str, object]]:
    header_cells = list(ws.iter_rows(min_row=1, max_row=1, values_only=True))
    if not header_cells:
        raise ValueError(f"{ws.title}: row 1 must contain headers")
    headers = list(header_cells[0])
    norm_to_idx = {normalize_header(h): i for i, h in enumerate(headers) if normalize_header(h)}
    missing = [c for c in required_cols if normalize_header(c) not in norm_to_idx]
    if missing:
        raise ValueError(f"{ws.title}: missing required columns: {', '.join(missing)}")

    rows: List[Dict[str, object]] = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        data = {}
        non_empty = False
        for c in required_cols:
            idx = norm_to_idx[normalize_header(c)]
            val = row[idx] if idx < len(row) else None
            data[c] = val
            if val is not None and str(val).strip() != "":
                non_empty = True
        if non_empty:
            rows.append(data)
    return rows


def require_sheet(wb, name: str):
    if name not in wb.sheetnames:
        raise ValueError(f"missing required sheet: {name}")
    return wb[name]


def read_input(path: Path) -> Tuple[Parameters, List[PointLoad], Dict[str, Section], List[CandSpec], List[CandSpec], List[CandSpec], Options]:
    wb = load_workbook(path, data_only=True)

    ws_param = require_sheet(wb, "PARAM")
    param_rows = read_table_rows(
        ws_param,
        [
            "Lx[m]",
            "Ly[m]",
            "q[kN/m2]",
            "E[N/mm2]",
            "density[kg/m3]",
            "allow_sigma[N/mm2]",
            "allow_tau[N/mm2]",
            "deflection_limit_ratio",
            "unit_system",
            "objective",
            "random_seed",
        ],
    )
    if not param_rows:
        raise ValueError("PARAM: at least one data row is required")
    r = param_rows[0]
    params = Parameters(
        Lx=to_float(r["Lx[m]"], "PARAM.Lx[m]"),
        Ly=to_float(r["Ly[m]"], "PARAM.Ly[m]"),
        q=to_float(r["q[kN/m2]"], "PARAM.q[kN/m2]"),
        E=to_float(r["E[N/mm2]"], "PARAM.E[N/mm2]"),
        density=to_float(r["density[kg/m3]"], "PARAM.density[kg/m3]"),
        allow_sigma=to_float(r["allow_sigma[N/mm2]"], "PARAM.allow_sigma[N/mm2]"),
        allow_tau=to_float(r["allow_tau[N/mm2]"], "PARAM.allow_tau[N/mm2]"),
        deflection_limit_ratio=to_float(r["deflection_limit_ratio"], "PARAM.deflection_limit_ratio"),
        unit_system=str(r["unit_system"]).strip() if r["unit_system"] is not None else "",
        objective=(str(r["objective"]).strip().lower() if r["objective"] is not None else ""),
        random_seed=to_int(r["random_seed"], "PARAM.random_seed"),
    )
    if params.Lx <= 0 or params.Ly <= 0:
        raise ValueError("PARAM: Lx/Ly must be positive")
    if params.q < 0:
        raise ValueError("PARAM: q[kN/m2] must be non-negative")
    if params.E <= 0:
        raise ValueError("PARAM: E[N/mm2] must be positive")
    if params.allow_sigma <= 0 or params.allow_tau <= 0:
        raise ValueError("PARAM: allow_sigma/allow_tau must be positive")
    if params.deflection_limit_ratio <= 1.0:
        raise ValueError("PARAM: deflection_limit_ratio must be > 1")
    if params.objective not in {"weight", "cost"}:
        raise ValueError("PARAM.objective must be 'weight' or 'cost'")

    ws_point = require_sheet(wb, "POINT_LOADS")
    point_rows = read_table_rows(ws_point, ["Name", "P[kN]", "x[m]", "y[m]"])
    point_loads: List[PointLoad] = []
    for rr in point_rows:
        pl = PointLoad(
            name=str(rr["Name"]).strip(),
            P=to_float(rr["P[kN]"], "POINT_LOADS.P[kN]"),
            x=to_float(rr["x[m]"], "POINT_LOADS.x[m]"),
            y=to_float(rr["y[m]"], "POINT_LOADS.y[m]"),
        )
        if pl.P < 0:
            raise ValueError(f"POINT_LOADS {pl.name}: P[kN] must be non-negative")
        if not (0 - EPS <= pl.x <= params.Lx + EPS and 0 - EPS <= pl.y <= params.Ly + EPS):
            raise ValueError(f"POINT_LOADS {pl.name}: coordinate out of slab bounds")
        point_loads.append(pl)

    ws_sec = require_sheet(wb, "SECTIONS")
    sec_rows = read_table_rows(
        ws_sec,
        ["SectionName", "Type", "A[mm2]", "I[mm4]", "Z[mm3]", "unit_weight[kg/m]", "cost_per_m"],
    )
    if not sec_rows:
        raise ValueError("SECTIONS: at least one row is required")
    section_map: Dict[str, Section] = {}
    for rr in sec_rows:
        name = str(rr["SectionName"]).strip()
        if not name:
            raise ValueError("SECTIONS: SectionName cannot be blank")
        if name in section_map:
            raise ValueError(f"SECTIONS: duplicate SectionName: {name}")
        A = to_float(rr["A[mm2]"], f"SECTIONS.{name}.A[mm2]")
        I = to_float(rr["I[mm4]"], f"SECTIONS.{name}.I[mm4]")
        Z = to_float(rr["Z[mm3]"], f"SECTIONS.{name}.Z[mm3]")
        if A <= 0 or I <= 0 or Z <= 0:
            raise ValueError(f"SECTIONS.{name}: A/I/Z must be positive")
        unit_w_raw = rr["unit_weight[kg/m]"]
        if unit_w_raw is None or str(unit_w_raw).strip() == "":
            unit_w = params.density * A * 1e-6
        else:
            unit_w = to_float(unit_w_raw, f"SECTIONS.{name}.unit_weight[kg/m]")
        if unit_w <= 0:
            raise ValueError(f"SECTIONS.{name}: unit_weight[kg/m] must be positive")
        cp_raw = rr["cost_per_m"]
        cp = 0.0
        if cp_raw is not None and str(cp_raw).strip() != "":
            cp = to_float(cp_raw, f"SECTIONS.{name}.cost_per_m")
        if params.objective == "cost" and cp <= 0:
            raise ValueError(f"SECTIONS.{name}: cost_per_m must be positive when objective=cost")
        section_map[name] = Section(
            name=name,
            sec_type=str(rr["Type"]).strip() if rr["Type"] is not None else "",
            A=A,
            I=I,
            Z=Z,
            unit_weight=unit_w,
            cost_per_m=cp,
        )

    def parse_cands(sheet_name: str) -> List[CandSpec]:
        ws = require_sheet(wb, sheet_name)
        rows = read_table_rows(
            ws,
            [
                "Direction",
                "PitchStart[m]",
                "PitchEnd[m]",
                "PitchStep[m]",
                "OffsetOptions",
                "MaxCount",
                "AllowedSections",
                "MinEdgeClear[m]",
                "AllowMixedXY",
            ],
        )
        cands: List[CandSpec] = []
        for idx, rr in enumerate(rows, start=2):
            direction = str(rr["Direction"]).strip().upper()
            if direction not in {"X", "Y"}:
                raise ValueError(f"{sheet_name} row {idx}: Direction must be X or Y")
            ps = to_float(rr["PitchStart[m]"], f"{sheet_name}.PitchStart[m]")
            pe = to_float(rr["PitchEnd[m]"], f"{sheet_name}.PitchEnd[m]")
            st = to_float(rr["PitchStep[m]"], f"{sheet_name}.PitchStep[m]")
            mc = to_int(rr["MaxCount"], f"{sheet_name}.MaxCount")
            min_edge = to_float(rr["MinEdgeClear[m]"], f"{sheet_name}.MinEdgeClear[m]")
            if ps <= 0 or pe <= 0 or st <= 0:
                raise ValueError(f"{sheet_name} row {idx}: pitch values must be positive")
            if pe + EPS < ps:
                raise ValueError(f"{sheet_name} row {idx}: PitchEnd < PitchStart")
            if mc < 0:
                raise ValueError(f"{sheet_name} row {idx}: MaxCount must be >= 0")
            if min_edge < 0:
                raise ValueError(f"{sheet_name} row {idx}: MinEdgeClear[m] must be >= 0")
            allowed = split_list_text(rr["AllowedSections"])
            if allowed:
                unknown = [a for a in allowed if a not in section_map]
                if unknown:
                    raise ValueError(f"{sheet_name} row {idx}: unknown sections in AllowedSections: {unknown}")
            cands.append(
                CandSpec(
                    row_id=f"{sheet_name}_R{idx}",
                    direction=direction,
                    pitch_start=ps,
                    pitch_end=pe,
                    pitch_step=st,
                    offset_options=str(rr["OffsetOptions"]).strip() if rr["OffsetOptions"] is not None else "",
                    max_count=mc,
                    allowed_sections=allowed,
                    min_edge_clear=min_edge,
                    allow_mixed_xy=to_bool(rr["AllowMixedXY"], default=False),
                )
            )
        return cands

    cand_l1 = parse_cands("CAND_L1")
    cand_l2 = parse_cands("CAND_L2")
    cand_l3 = parse_cands("CAND_L3")

    ws_opt = require_sheet(wb, "OPTIONS")
    opt_rows = read_table_rows(ws_opt, ["SearchMode", "TimeLimitSec", "Parallel", "NWorkers", "PruneLevel"])
    if not opt_rows:
        raise ValueError("OPTIONS: at least one row is required")
    ro = opt_rows[0]
    sm = str(ro["SearchMode"]).strip().upper()
    if sm != "DP":
        raise ValueError("OPTIONS.SearchMode must be DP")
    tl = None
    if ro["TimeLimitSec"] is not None and str(ro["TimeLimitSec"]).strip() != "":
        tl = to_float(ro["TimeLimitSec"], "OPTIONS.TimeLimitSec")
        if tl <= 0:
            raise ValueError("OPTIONS.TimeLimitSec must be positive")
    nw = 1
    if ro["NWorkers"] is not None and str(ro["NWorkers"]).strip() != "":
        nw = max(1, to_int(ro["NWorkers"], "OPTIONS.NWorkers"))
    pl = to_int(ro["PruneLevel"], "OPTIONS.PruneLevel")
    if pl < 0 or pl > 2:
        raise ValueError("OPTIONS.PruneLevel must be 0..2")
    opts = Options(
        search_mode=sm,
        time_limit_sec=tl,
        parallel=to_bool(ro["Parallel"], default=False),
        n_workers=nw,
        prune_level=pl,
    )
    return params, point_loads, section_map, cand_l1, cand_l2, cand_l3, opts


def float_range(start: float, end: float, step: float) -> List[float]:
    vals: List[float] = []
    x = start
    guard = 0
    while x <= end + 1e-9:
        vals.append(round(x, 6))
        x += step
        guard += 1
        if guard > 100000:
            raise ValueError("pitch range produced too many values")
    if not vals:
        vals.append(round(start, 6))
    return vals


def parse_offset_options(text: str, pitch: float) -> List[float]:
    if not text.strip():
        return [0.0]
    out: List[float] = []
    for tok in text.split(";"):
        s = tok.strip().upper()
        if not s:
            continue
        if s.endswith("P"):
            coef_txt = s[:-1].strip()
            coef = float(coef_txt) if coef_txt else 1.0
            out.append(coef * pitch)
        else:
            out.append(float(s))
    if not out:
        return [0.0]
    return sorted(set(round(v, 6) for v in out))


def generate_positions(width: float, pitch: float, offset: float, min_edge_clear: float, max_count: int) -> List[float]:
    if max_count <= 0:
        return []
    if width <= 2 * min_edge_clear + EPS:
        return []
    pos0 = offset
    while pos0 < min_edge_clear - EPS:
        pos0 += pitch
    positions: List[float] = []
    x = pos0
    while x <= width - min_edge_clear + EPS and len(positions) < max_count:
        if x >= min_edge_clear - EPS and x <= width - min_edge_clear + EPS:
            positions.append(round(x, 6))
        x += pitch
    uniq: List[float] = []
    for p in positions:
        if not uniq or abs(uniq[-1] - p) > 1e-6:
            uniq.append(p)
    return uniq


def make_patterns(specs: List[CandSpec], direction: str, width: float, prune_level: int) -> List[PatternOption]:
    patterns: List[PatternOption] = []
    seen = set()
    for spec in specs:
        if spec.direction != direction:
            continue
        pitches = float_range(spec.pitch_start, spec.pitch_end, spec.pitch_step)
        for pitch in pitches:
            offsets = parse_offset_options(spec.offset_options, pitch)
            for off in offsets:
                base = generate_positions(width, pitch, off, spec.min_edge_clear, spec.max_count)
                nmax = min(spec.max_count, len(base))
                for n in range(1, nmax + 1):
                    pos = tuple(round(v, 6) for v in base[:n])
                    key = (pos, spec.allowed_sections, spec.allow_mixed_xy)
                    if key in seen:
                        continue
                    seen.add(key)
                    p_key = f"{spec.row_id}:{direction}:P{pitch:.4f}:O{off:.4f}:N{n}"
                    patterns.append(
                        PatternOption(
                            key=p_key,
                            direction=direction,
                            positions=pos,
                            allowed_sections=spec.allowed_sections,
                            allow_mixed_xy=spec.allow_mixed_xy,
                        )
                    )
    patterns.sort(key=lambda p: (len(p.positions), p.positions, p.key))
    cap = {0: 800, 1: 400, 2: 180}[prune_level]
    if len(patterns) > cap:
        patterns = patterns[:cap]
    return patterns


def build_layout_options(width: float, height: float, specs: List[CandSpec], prune_level: int) -> List[LevelLayoutOption]:
    x_patterns = make_patterns(specs, "X", height, prune_level)
    y_patterns = make_patterns(specs, "Y", width, prune_level)
    options: List[LevelLayoutOption] = [LevelLayoutOption(key="EMPTY", x_pattern=None, y_pattern=None)]
    for xp in x_patterns:
        options.append(LevelLayoutOption(key=f"X[{xp.key}]", x_pattern=xp, y_pattern=None))
    for yp in y_patterns:
        options.append(LevelLayoutOption(key=f"Y[{yp.key}]", x_pattern=None, y_pattern=yp))
    for xp in x_patterns:
        for yp in y_patterns:
            if xp.allow_mixed_xy and yp.allow_mixed_xy:
                options.append(LevelLayoutOption(key=f"MIX[{xp.key}]+[{yp.key}]", x_pattern=xp, y_pattern=yp))
    dedup: Dict[Tuple, LevelLayoutOption] = {}
    for op in options:
        k = (
            tuple(op.x_pattern.positions) if op.x_pattern else tuple(),
            tuple(op.y_pattern.positions) if op.y_pattern else tuple(),
            op.x_pattern.allowed_sections if op.x_pattern else tuple(),
            op.y_pattern.allowed_sections if op.y_pattern else tuple(),
        )
        if k not in dedup or op.key < dedup[k].key:
            dedup[k] = op
    out = sorted(dedup.values(), key=lambda o: (o.key != "EMPTY", o.key))
    cap = {0: 1200, 1: 500, 2: 220}[prune_level]
    if len(out) > cap:
        out = out[:cap]
    return out


def min_section_metric(allowed: Tuple[str, ...], sections: Dict[str, Section], objective: str) -> float:
    names = allowed if allowed else tuple(sorted(sections.keys()))
    vals = [sections[n].metric(objective) for n in names]
    return min(vals) if vals else math.inf


def tributary_widths(positions: Sequence[float], width: float) -> Dict[float, float]:
    s = sorted(float(p) for p in positions)
    if not s:
        return {}
    out: Dict[float, float] = {}
    for i, p in enumerate(s):
        left = 0.0 if i == 0 else 0.5 * (s[i - 1] + p)
        right = width if i == len(s) - 1 else 0.5 * (p + s[i + 1])
        out[p] = max(0.0, right - left)
    return out


def assign_slab_udl(beams: List[LocalBeamProto], q: float, width: float, height: float) -> Dict[str, float]:
    out = {b.beam_id: 0.0 for b in beams}
    x_beams = [b for b in beams if b.direction == "X"]
    y_beams = [b for b in beams if b.direction == "Y"]
    if x_beams and not y_beams:
        tw = tributary_widths([b.pos for b in x_beams], height)
        for b in x_beams:
            out[b.beam_id] = q * tw.get(b.pos, 0.0)
    elif y_beams and not x_beams:
        tw = tributary_widths([b.pos for b in y_beams], width)
        for b in y_beams:
            out[b.beam_id] = q * tw.get(b.pos, 0.0)
    elif x_beams and y_beams:
        twx = tributary_widths([b.pos for b in x_beams], height)
        twy = tributary_widths([b.pos for b in y_beams], width)
        for b in x_beams:
            out[b.beam_id] = 0.5 * q * twx.get(b.pos, 0.0)
        for b in y_beams:
            out[b.beam_id] = 0.5 * q * twy.get(b.pos, 0.0)
    return out


def analyze_simply_supported(length: float, w_line: float, points: List[AppliedPointLoad], E: float, I: float) -> Tuple[float, float, float, float, float, List[float], List[float], List[float], List[float]]:
    if length <= 0:
        raise ValueError("beam length must be positive")
    pts = sorted(points, key=lambda p: p.a)
    Ra = 0.5 * w_line * length + sum(p.P * (length - p.a) / length for p in pts)
    Rb = 0.5 * w_line * length + sum(p.P * p.a / length for p in pts)

    n = 201
    xs = [length * i / (n - 1) for i in range(n)]
    Vs: List[float] = []
    Ms: List[float] = []
    for x in xs:
        v = Ra - w_line * x
        m = Ra * x - 0.5 * w_line * x * x
        for p in pts:
            if p.a <= x + EPS:
                v -= p.P
                m -= p.P * (x - p.a)
        Vs.append(v)
        Ms.append(m)

    Vmax = max(abs(v) for v in Vs) if Vs else 0.0
    Mmax = max(abs(m) for m in Ms) if Ms else 0.0

    x_mm = [x * 1000.0 for x in xs]
    L_mm = length * 1000.0
    curv = [(m * 1e6) / (E * I) for m in Ms]

    i1 = 0.0
    for i in range(1, n):
        s0 = x_mm[i - 1]
        s1 = x_mm[i]
        f0 = (L_mm - s0) * curv[i - 1]
        f1 = (L_mm - s1) * curv[i]
        i1 += 0.5 * (f0 + f1) * (s1 - s0)
    theta0 = -i1 / L_mm

    ys: List[float] = []
    for i in range(n):
        xi = x_mm[i]
        integ = 0.0
        for j in range(1, i + 1):
            s0 = x_mm[j - 1]
            s1 = x_mm[j]
            f0 = (xi - s0) * curv[j - 1]
            f1 = (xi - s1) * curv[j]
            integ += 0.5 * (f0 + f1) * (s1 - s0)
        y = theta0 * xi + integ
        ys.append(y)
    ymax_mm = max(abs(y) for y in ys) if ys else 0.0
    return Ra, Rb, Mmax, Vmax, ymax_mm, xs, Vs, Ms, ys


def choose_section(
    beam: LocalBeamProto,
    udl_external: float,
    point_loads: List[AppliedPointLoad],
    ctx: SearchContext,
    ks: float = 0.6,
) -> Tuple[Optional[BeamRecord], Optional[str]]:
    names = beam.allowed_sections if beam.allowed_sections else tuple(s.name for s in ctx.all_sections)
    if not names:
        return None, "no candidate sections"
    candidates = [ctx.section_map[n] for n in names]
    candidates.sort(key=lambda s: (s.metric(ctx.params.objective), -s.Z, -s.I, s.name))
    best_fail_util = -1.0
    best_fail_reason = "all sections failed"
    for sec in candidates:
        w_total = udl_external + sec.self_weight_kN_m
        Ra, Rb, Mmax, Vmax, ymax_mm, xs, Vs, Ms, ys = analyze_simply_supported(
            beam.length, w_total, point_loads, ctx.params.E, sec.I
        )
        sigma = Mmax * 1e6 / sec.Z
        tau = Vmax * 1e3 / (ks * sec.A)
        lim_mm = beam.length * 1000.0 / ctx.params.deflection_limit_ratio
        util = max(sigma / ctx.params.allow_sigma, tau / ctx.params.allow_tau, ymax_mm / lim_mm)
        if util <= 1.0 + 1e-9:
            response = BeamResponse(
                Ra=Ra,
                Rb=Rb,
                Mmax=Mmax,
                Vmax=Vmax,
                ymax_mm=ymax_mm,
                sigma=sigma,
                tau=tau,
                utilization=util,
                x_samples=xs,
                V_samples=Vs,
                M_samples=Ms,
                y_samples_mm=ys,
            )
            if beam.direction == "X":
                span_start, span_end = beam.x0, beam.x1
            else:
                span_start, span_end = beam.y0, beam.y1
            rec = BeamRecord(
                level=beam.level,
                direction=beam.direction,
                pos=beam.pos,
                span_start=span_start,
                span_end=span_end,
                length=beam.length,
                allowed_sections=beam.allowed_sections,
                section_name=sec.name,
                unit_weight=sec.unit_weight,
                cost_per_m=sec.cost_per_m,
                w_line=w_total,
                point_loads=sorted(point_loads, key=lambda p: p.a),
                response=response,
            )
            return rec, None
        if util > best_fail_util:
            best_fail_util = util
            best_fail_reason = f"max utilization {util:.3f} (section={sec.name})"
    return None, best_fail_reason


def distribute_points_to_beams(points: Sequence[PointLoadLocal], beams: List[LocalBeamProto]) -> Dict[str, List[AppliedPointLoad]]:
    assigned: Dict[str, List[AppliedPointLoad]] = {b.beam_id: [] for b in beams}
    if not beams:
        return assigned
    for p in points:
        dlist: List[Tuple[float, LocalBeamProto]] = []
        for b in beams:
            d = abs(p.y - b.pos) if b.direction == "X" else abs(p.x - b.pos)
            dlist.append((d, b))
        dlist.sort(key=lambda t: (t[0], t[1].beam_id))
        picked = dlist[:2] if len(dlist) >= 2 else dlist
        if len(picked) == 1:
            shares = [(picked[0][1], p.P)]
        else:
            d1, b1 = picked[0]
            d2, b2 = picked[1]
            if d1 + d2 < 1e-12:
                shares = [(b1, 0.5 * p.P), (b2, 0.5 * p.P)]
            else:
                shares = [(b1, p.P * d2 / (d1 + d2)), (b2, p.P * d1 / (d1 + d2))]
        for b, pp in shares:
            a = p.x - b.x0 if b.direction == "X" else p.y - b.y0
            a = min(max(0.0, a), b.length)
            assigned[b.beam_id].append(AppliedPointLoad(P=pp, a=a, source=f"POINT:{p.name}", x=p.x, y=p.y))
    return assigned


def distribute_points_to_edges(points: Sequence[PointLoadLocal], width: float, height: float) -> List[BoundaryLoad]:
    out: List[BoundaryLoad] = []
    for p in points:
        d = [
            ("left", p.x),
            ("right", width - p.x),
            ("bottom", p.y),
            ("top", height - p.y),
        ]
        d.sort(key=lambda t: (t[1], t[0]))
        e1, d1 = d[0]
        e2, d2 = d[1]
        if d1 + d2 < 1e-12:
            shares = [(e1, 0.5 * p.P), (e2, 0.5 * p.P)]
        else:
            shares = [(e1, p.P * d2 / (d1 + d2)), (e2, p.P * d1 / (d1 + d2))]
        for e, pp in shares:
            s = p.y if e in {"left", "right"} else p.x
            out.append(BoundaryLoad(edge=e, s=s, P=pp, source=f"POINT:{p.name}"))
    return out


def distribute_uniform_to_edges(q: float, width: float, height: float) -> List[BoundaryLoad]:
    w = q * width * height
    if w <= 0:
        return []
    return [
        BoundaryLoad(edge="left", s=0.5 * height, P=0.25 * w, source="Q_UNIF"),
        BoundaryLoad(edge="right", s=0.5 * height, P=0.25 * w, source="Q_UNIF"),
        BoundaryLoad(edge="bottom", s=0.5 * width, P=0.25 * w, source="Q_UNIF"),
        BoundaryLoad(edge="top", s=0.5 * width, P=0.25 * w, source="Q_UNIF"),
    ]


def support_ok(level: int, direction: str, edge_types: EdgeTypes) -> bool:
    if level == 1:
        allowed = {"PERIM"}
    elif level == 2:
        allowed = {"PERIM", "L1"}
    elif level == 3:
        allowed = {"PERIM", "L2"}
    else:
        return False
    if direction == "X":
        return edge_types.left in allowed and edge_types.right in allowed
    return edge_types.bottom in allowed and edge_types.top in allowed


def build_level_beams(level: int, option: LevelLayoutOption, width: float, height: float) -> List[LocalBeamProto]:
    beams: List[LocalBeamProto] = []
    if option.x_pattern:
        for i, y in enumerate(option.x_pattern.positions, start=1):
            beams.append(
                LocalBeamProto(
                    beam_id=f"X{i}",
                    level=level,
                    direction="X",
                    pos=y,
                    x0=0.0,
                    x1=width,
                    y0=y,
                    y1=y,
                    length=width,
                    allowed_sections=option.x_pattern.allowed_sections,
                )
            )
    if option.y_pattern:
        for i, x in enumerate(option.y_pattern.positions, start=1):
            beams.append(
                LocalBeamProto(
                    beam_id=f"Y{i}",
                    level=level,
                    direction="Y",
                    pos=x,
                    x0=x,
                    x1=x,
                    y0=0.0,
                    y1=height,
                    length=height,
                    allowed_sections=option.y_pattern.allowed_sections,
                )
            )
    return beams


def split_panels(parent: LocalPanel, beams: List[LocalBeamProto], interior_type: str) -> List[LocalPanel]:
    x_beams = sorted([b for b in beams if b.direction == "Y"], key=lambda b: b.pos)
    y_beams = sorted([b for b in beams if b.direction == "X"], key=lambda b: b.pos)
    xs = [parent.x0] + [b.pos for b in x_beams] + [parent.x1]
    ys = [parent.y0] + [b.pos for b in y_beams] + [parent.y1]

    x_ids: Dict[float, str] = {parent.x0: parent.left_id, parent.x1: parent.right_id}
    y_ids: Dict[float, str] = {parent.y0: parent.bottom_id, parent.y1: parent.top_id}
    x_types: Dict[float, str] = {parent.x0: parent.edge_types.left, parent.x1: parent.edge_types.right}
    y_types: Dict[float, str] = {parent.y0: parent.edge_types.bottom, parent.y1: parent.edge_types.top}
    for b in x_beams:
        x_ids[b.pos] = b.beam_id
        x_types[b.pos] = interior_type
    for b in y_beams:
        y_ids[b.pos] = b.beam_id
        y_types[b.pos] = interior_type

    out: List[LocalPanel] = []
    for ix in range(len(xs) - 1):
        for iy in range(len(ys) - 1):
            x0 = xs[ix]
            x1 = xs[ix + 1]
            y0 = ys[iy]
            y1 = ys[iy + 1]
            if x1 - x0 <= EPS or y1 - y0 <= EPS:
                continue
            out.append(
                LocalPanel(
                    x0=x0,
                    x1=x1,
                    y0=y0,
                    y1=y1,
                    left_id=x_ids[x0],
                    right_id=x_ids[x1],
                    bottom_id=y_ids[y0],
                    top_id=y_ids[y1],
                    edge_types=EdgeTypes(
                        left=x_types[x0],
                        right=x_types[x1],
                        bottom=y_types[y0],
                        top=y_types[y1],
                    ),
                )
            )
    return out


def point_in_panel_local(p: PointLoadLocal, panel: LocalPanel, is_last_x: bool, is_last_y: bool) -> bool:
    in_x = (panel.x0 - EPS <= p.x <= panel.x1 + EPS) if is_last_x else (panel.x0 - EPS <= p.x < panel.x1 - EPS)
    in_y = (panel.y0 - EPS <= p.y <= panel.y1 + EPS) if is_last_y else (panel.y0 - EPS <= p.y < panel.y1 - EPS)
    return in_x and in_y


def shift_beam_records(beams: List[BeamRecord], dx: float, dy: float) -> List[BeamRecord]:
    out: List[BeamRecord] = []
    for b in beams:
        nb = BeamRecord(
            level=b.level,
            direction=b.direction,
            pos=(b.pos + dy if b.direction == "X" else b.pos + dx),
            span_start=(b.span_start + dx if b.direction == "X" else b.span_start + dy),
            span_end=(b.span_end + dx if b.direction == "X" else b.span_end + dy),
            length=b.length,
            allowed_sections=b.allowed_sections,
            section_name=b.section_name,
            unit_weight=b.unit_weight,
            cost_per_m=b.cost_per_m,
            w_line=b.w_line,
            point_loads=[
                AppliedPointLoad(P=pl.P, a=pl.a, source=pl.source, x=pl.x + dx, y=pl.y + dy)
                for pl in b.point_loads
            ],
            response=b.response,
            beam_id="",
        )
        out.append(nb)
    return out


def boundary_load_global_xy(subpanel: LocalPanel, load: BoundaryLoad) -> Tuple[float, float]:
    if load.edge == "left":
        return subpanel.x0, subpanel.y0 + load.s
    if load.edge == "right":
        return subpanel.x1, subpanel.y0 + load.s
    if load.edge == "bottom":
        return subpanel.x0 + load.s, subpanel.y0
    if load.edge == "top":
        return subpanel.x0 + load.s, subpanel.y1
    raise ValueError(f"unknown edge: {load.edge}")


def map_subpanel_edge_to_support(subpanel: LocalPanel, edge: str) -> str:
    if edge == "left":
        return subpanel.left_id
    if edge == "right":
        return subpanel.right_id
    if edge == "bottom":
        return subpanel.bottom_id
    if edge == "top":
        return subpanel.top_id
    raise ValueError(edge)


def map_global_to_parent_edge(x: float, y: float, parent: LocalPanel, support_id: str) -> Optional[BoundaryLoad]:
    if support_id == parent.left_id:
        return BoundaryLoad(edge="left", s=y - parent.y0, P=0.0, source="")
    if support_id == parent.right_id:
        return BoundaryLoad(edge="right", s=y - parent.y0, P=0.0, source="")
    if support_id == parent.bottom_id:
        return BoundaryLoad(edge="bottom", s=x - parent.x0, P=0.0, source="")
    if support_id == parent.top_id:
        return BoundaryLoad(edge="top", s=x - parent.x0, P=0.0, source="")
    return None


def panel_cache_key(level: int, width: float, height: float, edge_types: EdgeTypes, points: Sequence[PointLoadLocal], prune_level: int, objective: str) -> Tuple:
    pts = tuple((round(p.x, 5), round(p.y, 5), round(p.P, 5), p.name) for p in sorted(points, key=lambda x: (x.x, x.y, x.P, x.name)))
    return (
        level,
        round(width, 5),
        round(height, 5),
        edge_types.left,
        edge_types.right,
        edge_types.bottom,
        edge_types.top,
        pts,
        prune_level,
        objective,
    )


def compare_result(a: PanelResult, b: PanelResult) -> bool:
    ka = (a.objective, a.max_util, a.beam_count, a.layout_key)
    kb = (b.objective, b.max_util, b.beam_count, b.layout_key)
    return ka < kb


def lower_bound_layout(level: int, option: LevelLayoutOption, width: float, height: float, sections: Dict[str, Section], objective: str) -> float:
    lb = 0.0
    if option.x_pattern:
        m = min_section_metric(option.x_pattern.allowed_sections, sections, objective)
        lb += len(option.x_pattern.positions) * width * m
    if option.y_pattern:
        m = min_section_metric(option.y_pattern.allowed_sections, sections, objective)
        lb += len(option.y_pattern.positions) * height * m
    return lb


def make_panel_result_infeasible(layout_key: str, reason: str) -> PanelResult:
    return PanelResult(
        feasible=False,
        objective=math.inf,
        total_weight=math.inf,
        total_cost=math.inf,
        max_util=math.inf,
        beams=[],
        boundary_loads=[],
        layout_key=layout_key,
        infeasible_reason=reason,
    )


def solve_l3_panel(width: float, height: float, edge_types: EdgeTypes, points: Sequence[PointLoadLocal], ctx: SearchContext) -> PanelResult:
    ctx.check_timeout()
    key = panel_cache_key(3, width, height, edge_types, points, ctx.options.prune_level, ctx.params.objective)
    if key in ctx.cache_l3:
        ctx.stats.l3_cache_hit += 1
        return ctx.cache_l3[key]
    ctx.stats.l3_cache_miss += 1

    options = build_layout_options(width, height, ctx.cand_l3, ctx.options.prune_level)
    best: Optional[PanelResult] = None
    for opt in options:
        t0 = time.perf_counter()
        ctx.check_timeout()

        if ctx.options.prune_level >= 1 and best is not None:
            lb = lower_bound_layout(3, opt, width, height, ctx.section_map, ctx.params.objective)
            if lb >= best.objective - 1e-9:
                ctx.traces.append(
                    TraceRow(
                        candidate_key=f"L3:{round(width,4)}x{round(height,4)}:{opt.key}",
                        objective=math.inf,
                        max_util=math.inf,
                        infeasible_reason=f"pruned_by_bound(lb={lb:.3f})",
                        elapsed_sec=time.perf_counter() - t0,
                    )
                )
                continue

        beams = build_level_beams(3, opt, width, height)
        bad_support = False
        for b in beams:
            if not support_ok(3, b.direction, edge_types):
                bad_support = True
                break
        if bad_support:
            ctx.traces.append(
                TraceRow(
                    candidate_key=f"L3:{round(width,4)}x{round(height,4)}:{opt.key}",
                    objective=math.inf,
                    max_util=math.inf,
                    infeasible_reason="support_rule_violation",
                    elapsed_sec=time.perf_counter() - t0,
                )
            )
            continue

        if ctx.options.prune_level >= 2 and beams:
            approx_ok = True
            strong = max(ctx.all_sections, key=lambda s: (s.Z, s.I, s.A))
            udls = assign_slab_udl(beams, ctx.params.q, width, height)
            for b in beams:
                w = udls.get(b.beam_id, 0.0) + strong.self_weight_kN_m
                m = w * b.length * b.length / 8.0
                sigma = (m * 1e6) / strong.Z
                if sigma / ctx.params.allow_sigma > 1.8:
                    approx_ok = False
                    break
            if not approx_ok:
                ctx.traces.append(
                    TraceRow(
                        candidate_key=f"L3:{round(width,4)}x{round(height,4)}:{opt.key}",
                        objective=math.inf,
                        max_util=math.inf,
                        infeasible_reason="pruned_quick_strength_check",
                        elapsed_sec=time.perf_counter() - t0,
                    )
                )
                continue

        if not beams:
            boundary_loads = distribute_uniform_to_edges(ctx.params.q, width, height)
            boundary_loads.extend(distribute_points_to_edges(points, width, height))
            res = PanelResult(
                feasible=True,
                objective=0.0,
                total_weight=0.0,
                total_cost=0.0,
                max_util=0.0,
                beams=[],
                boundary_loads=boundary_loads,
                layout_key=opt.key,
            )
            if best is None or compare_result(res, best):
                best = res
            ctx.traces.append(
                TraceRow(
                    candidate_key=f"L3:{round(width,4)}x{round(height,4)}:{opt.key}",
                    objective=res.objective,
                    max_util=res.max_util,
                    infeasible_reason="",
                    elapsed_sec=time.perf_counter() - t0,
                )
            )
            continue

        udl = assign_slab_udl(beams, ctx.params.q, width, height)
        point_map = distribute_points_to_beams(points, beams)
        out_beams: List[BeamRecord] = []
        out_boundary: List[BoundaryLoad] = []
        feasible = True
        reason = ""
        total_weight = 0.0
        total_cost = 0.0
        max_util = 0.0
        for b in beams:
            b_points = point_map.get(b.beam_id, [])
            chosen, fail_reason = choose_section(b, udl.get(b.beam_id, 0.0), b_points, ctx)
            if chosen is None:
                feasible = False
                reason = f"{b.beam_id}: {fail_reason}"
                break
            out_beams.append(chosen)
            total_weight += chosen.length * chosen.unit_weight
            total_cost += chosen.length * chosen.cost_per_m
            max_util = max(max_util, chosen.response.utilization)
            if b.direction == "X":
                out_boundary.append(BoundaryLoad(edge="left", s=b.pos, P=chosen.response.Ra, source=f"R_{b.beam_id}_A"))
                out_boundary.append(BoundaryLoad(edge="right", s=b.pos, P=chosen.response.Rb, source=f"R_{b.beam_id}_B"))
            else:
                out_boundary.append(BoundaryLoad(edge="bottom", s=b.pos, P=chosen.response.Ra, source=f"R_{b.beam_id}_A"))
                out_boundary.append(BoundaryLoad(edge="top", s=b.pos, P=chosen.response.Rb, source=f"R_{b.beam_id}_B"))

        if not feasible:
            ctx.traces.append(
                TraceRow(
                    candidate_key=f"L3:{round(width,4)}x{round(height,4)}:{opt.key}",
                    objective=math.inf,
                    max_util=math.inf,
                    infeasible_reason=reason,
                    elapsed_sec=time.perf_counter() - t0,
                )
            )
            continue

        obj = total_weight if ctx.params.objective == "weight" else total_cost
        res = PanelResult(
            feasible=True,
            objective=obj,
            total_weight=total_weight,
            total_cost=total_cost,
            max_util=max_util,
            beams=out_beams,
            boundary_loads=out_boundary,
            layout_key=opt.key,
        )
        if best is None or compare_result(res, best):
            best = res
        ctx.traces.append(
            TraceRow(
                candidate_key=f"L3:{round(width,4)}x{round(height,4)}:{opt.key}",
                objective=res.objective,
                max_util=res.max_util,
                infeasible_reason="",
                elapsed_sec=time.perf_counter() - t0,
            )
        )

    if best is None:
        best = make_panel_result_infeasible("NO_LAYOUT", "no feasible L3 layout")
    ctx.cache_l3[key] = best
    return best


def solve_l2_panel(width: float, height: float, edge_types: EdgeTypes, points: Sequence[PointLoadLocal], ctx: SearchContext) -> PanelResult:
    ctx.check_timeout()
    key = panel_cache_key(2, width, height, edge_types, points, ctx.options.prune_level, ctx.params.objective)
    if key in ctx.cache_l2:
        ctx.stats.l2_cache_hit += 1
        return ctx.cache_l2[key]
    ctx.stats.l2_cache_miss += 1

    options = build_layout_options(width, height, ctx.cand_l2, ctx.options.prune_level)
    parent_panel = LocalPanel(
        x0=0.0,
        x1=width,
        y0=0.0,
        y1=height,
        left_id="left",
        right_id="right",
        bottom_id="bottom",
        top_id="top",
        edge_types=edge_types,
    )
    best: Optional[PanelResult] = None
    for opt in options:
        t0 = time.perf_counter()
        ctx.check_timeout()
        if ctx.options.prune_level >= 1 and best is not None:
            lb = lower_bound_layout(2, opt, width, height, ctx.section_map, ctx.params.objective)
            if lb >= best.objective - 1e-9:
                ctx.traces.append(
                    TraceRow(
                        candidate_key=f"L2:{round(width,4)}x{round(height,4)}:{opt.key}",
                        objective=math.inf,
                        max_util=math.inf,
                        infeasible_reason=f"pruned_by_bound(lb={lb:.3f})",
                        elapsed_sec=time.perf_counter() - t0,
                    )
                )
                continue

        beams = build_level_beams(2, opt, width, height)
        bad_support = False
        for b in beams:
            if not support_ok(2, b.direction, edge_types):
                bad_support = True
                break
        if bad_support:
            ctx.traces.append(
                TraceRow(
                    candidate_key=f"L2:{round(width,4)}x{round(height,4)}:{opt.key}",
                    objective=math.inf,
                    max_util=math.inf,
                    infeasible_reason="support_rule_violation",
                    elapsed_sec=time.perf_counter() - t0,
                )
            )
            continue

        subpanels = split_panels(parent_panel, beams, interior_type="L2")
        beam_map = {b.beam_id: b for b in beams}
        beam_point_map: Dict[str, List[AppliedPointLoad]] = {b.beam_id: [] for b in beams}
        upper_boundary_loads: List[BoundaryLoad] = []
        child_beams_total: List[BeamRecord] = []
        child_weight = 0.0
        child_cost = 0.0
        child_max_util = 0.0
        infeasible_reason = ""
        feasible = True

        for sp in subpanels:
            is_last_x = abs(sp.x1 - parent_panel.x1) <= 1e-7
            is_last_y = abs(sp.y1 - parent_panel.y1) <= 1e-7
            p_in: List[PointLoadLocal] = []
            for p in points:
                if point_in_panel_local(p, sp, is_last_x, is_last_y):
                    p_in.append(PointLoadLocal(name=p.name, P=p.P, x=p.x - sp.x0, y=p.y - sp.y0))
            child = solve_l3_panel(sp.width, sp.height, sp.edge_types, p_in, ctx)
            if not child.feasible:
                feasible = False
                infeasible_reason = f"child(L3) infeasible: {child.infeasible_reason}"
                break
            child_weight += child.total_weight
            child_cost += child.total_cost
            child_max_util = max(child_max_util, child.max_util)
            child_beams_total.extend(shift_beam_records(child.beams, sp.x0, sp.y0))

            for bl in child.boundary_loads:
                gx, gy = boundary_load_global_xy(sp, bl)
                support_id = map_subpanel_edge_to_support(sp, bl.edge)
                if support_id in beam_map:
                    bb = beam_map[support_id]
                    a = gx - bb.x0 if bb.direction == "X" else gy - bb.y0
                    a = min(max(0.0, a), bb.length)
                    beam_point_map[support_id].append(AppliedPointLoad(P=bl.P, a=a, source=bl.source, x=gx, y=gy))
                else:
                    pe = map_global_to_parent_edge(gx, gy, parent_panel, support_id)
                    if pe is None:
                        continue
                    upper_boundary_loads.append(BoundaryLoad(edge=pe.edge, s=pe.s, P=bl.P, source=bl.source))

        if not feasible:
            ctx.traces.append(
                TraceRow(
                    candidate_key=f"L2:{round(width,4)}x{round(height,4)}:{opt.key}",
                    objective=math.inf,
                    max_util=math.inf,
                    infeasible_reason=infeasible_reason,
                    elapsed_sec=time.perf_counter() - t0,
                )
            )
            continue

        own_beams: List[BeamRecord] = []
        own_weight = 0.0
        own_cost = 0.0
        own_max_util = 0.0
        reaction_loads: List[BoundaryLoad] = []
        feasible = True
        reason = ""
        for b in beams:
            pts = beam_point_map.get(b.beam_id, [])
            chosen, fail_reason = choose_section(b, 0.0, pts, ctx)
            if chosen is None:
                feasible = False
                reason = f"{b.beam_id}: {fail_reason}"
                break
            own_beams.append(chosen)
            own_weight += chosen.length * chosen.unit_weight
            own_cost += chosen.length * chosen.cost_per_m
            own_max_util = max(own_max_util, chosen.response.utilization)
            if b.direction == "X":
                reaction_loads.append(BoundaryLoad(edge="left", s=b.pos, P=chosen.response.Ra, source=f"R_{b.beam_id}_A"))
                reaction_loads.append(BoundaryLoad(edge="right", s=b.pos, P=chosen.response.Rb, source=f"R_{b.beam_id}_B"))
            else:
                reaction_loads.append(BoundaryLoad(edge="bottom", s=b.pos, P=chosen.response.Ra, source=f"R_{b.beam_id}_A"))
                reaction_loads.append(BoundaryLoad(edge="top", s=b.pos, P=chosen.response.Rb, source=f"R_{b.beam_id}_B"))
        if not feasible:
            ctx.traces.append(
                TraceRow(
                    candidate_key=f"L2:{round(width,4)}x{round(height,4)}:{opt.key}",
                    objective=math.inf,
                    max_util=math.inf,
                    infeasible_reason=reason,
                    elapsed_sec=time.perf_counter() - t0,
                )
            )
            continue

        total_weight = own_weight + child_weight
        total_cost = own_cost + child_cost
        obj = total_weight if ctx.params.objective == "weight" else total_cost
        beams_total = own_beams + child_beams_total
        res = PanelResult(
            feasible=True,
            objective=obj,
            total_weight=total_weight,
            total_cost=total_cost,
            max_util=max(own_max_util, child_max_util),
            beams=beams_total,
            boundary_loads=upper_boundary_loads + reaction_loads,
            layout_key=opt.key,
        )
        if best is None or compare_result(res, best):
            best = res
        ctx.traces.append(
            TraceRow(
                candidate_key=f"L2:{round(width,4)}x{round(height,4)}:{opt.key}",
                objective=res.objective,
                max_util=res.max_util,
                infeasible_reason="",
                elapsed_sec=time.perf_counter() - t0,
            )
        )

    if best is None:
        best = make_panel_result_infeasible("NO_LAYOUT", "no feasible L2 layout")
    ctx.cache_l2[key] = best
    return best


def evaluate_l1_layout(option: LevelLayoutOption, points: List[PointLoad], ctx: SearchContext) -> PanelResult:
    p = ctx.params
    root = LocalPanel(
        x0=0.0,
        x1=p.Lx,
        y0=0.0,
        y1=p.Ly,
        left_id="PERIM_LEFT",
        right_id="PERIM_RIGHT",
        bottom_id="PERIM_BOTTOM",
        top_id="PERIM_TOP",
        edge_types=EdgeTypes(left="PERIM", right="PERIM", bottom="PERIM", top="PERIM"),
    )
    beams = build_level_beams(1, option, p.Lx, p.Ly)
    if not beams:
        return make_panel_result_infeasible(option.key, "L1 requires at least one beam")
    for b in beams:
        if not support_ok(1, b.direction, root.edge_types):
            return make_panel_result_infeasible(option.key, "L1 support rule violation")
    beam_map = {b.beam_id: b for b in beams}
    beam_point_map: Dict[str, List[AppliedPointLoad]] = {b.beam_id: [] for b in beams}

    panels = split_panels(root, beams, interior_type="L1")
    child_beams: List[BeamRecord] = []
    child_weight = 0.0
    child_cost = 0.0
    child_max_util = 0.0

    for sp in panels:
        is_last_x = abs(sp.x1 - root.x1) <= 1e-7
        is_last_y = abs(sp.y1 - root.y1) <= 1e-7
        p_in: List[PointLoadLocal] = []
        for pl in points:
            ppl = PointLoadLocal(name=pl.name, P=pl.P, x=pl.x, y=pl.y)
            if point_in_panel_local(ppl, sp, is_last_x, is_last_y):
                p_in.append(PointLoadLocal(name=pl.name, P=pl.P, x=pl.x - sp.x0, y=pl.y - sp.y0))
        child = solve_l2_panel(sp.width, sp.height, sp.edge_types, p_in, ctx)
        if not child.feasible:
            return make_panel_result_infeasible(option.key, f"L2 panel infeasible: {child.infeasible_reason}")
        child_beams.extend(shift_beam_records(child.beams, sp.x0, sp.y0))
        child_weight += child.total_weight
        child_cost += child.total_cost
        child_max_util = max(child_max_util, child.max_util)

        for bl in child.boundary_loads:
            gx, gy = boundary_load_global_xy(sp, bl)
            support_id = map_subpanel_edge_to_support(sp, bl.edge)
            if support_id in beam_map:
                bb = beam_map[support_id]
                a = gx - bb.x0 if bb.direction == "X" else gy - bb.y0
                a = min(max(0.0, a), bb.length)
                beam_point_map[support_id].append(AppliedPointLoad(P=bl.P, a=a, source=bl.source, x=gx, y=gy))

    own_beams: List[BeamRecord] = []
    own_weight = 0.0
    own_cost = 0.0
    own_max_util = 0.0
    for b in beams:
        chosen, fail = choose_section(b, 0.0, beam_point_map.get(b.beam_id, []), ctx)
        if chosen is None:
            return make_panel_result_infeasible(option.key, f"L1 {b.beam_id}: {fail}")
        own_beams.append(chosen)
        own_weight += chosen.length * chosen.unit_weight
        own_cost += chosen.length * chosen.cost_per_m
        own_max_util = max(own_max_util, chosen.response.utilization)

    total_weight = own_weight + child_weight
    total_cost = own_cost + child_cost
    obj = total_weight if p.objective == "weight" else total_cost
    all_beams = own_beams + child_beams
    max_util = max(own_max_util, child_max_util)
    return PanelResult(
        feasible=True,
        objective=obj,
        total_weight=total_weight,
        total_cost=total_cost,
        max_util=max_util,
        beams=all_beams,
        boundary_loads=[],
        layout_key=option.key,
    )


def assign_beam_ids(beams: List[BeamRecord]) -> Tuple[List[BeamRecord], str]:
    ordered = sorted(
        beams,
        key=lambda b: (
            b.level,
            b.direction,
            round(b.pos, 6),
            round(b.span_start, 6),
            round(b.span_end, 6),
            b.section_name,
        ),
    )
    worst_id = ""
    worst_util = -1.0
    for i, b in enumerate(ordered, start=1):
        b.beam_id = f"B{b.level:01d}-{i:04d}"
        if b.response.utilization > worst_util:
            worst_util = b.response.utilization
            worst_id = b.beam_id
    return ordered, worst_id


def optimize(
    params: Parameters,
    point_loads: List[PointLoad],
    section_map: Dict[str, Section],
    cand_l1: List[CandSpec],
    cand_l2: List[CandSpec],
    cand_l3: List[CandSpec],
    options: Options,
) -> OptimizationResult:
    random.seed(params.random_seed)
    all_sections = sorted(section_map.values(), key=lambda s: (s.metric(params.objective), s.name))
    ctx = SearchContext(
        params=params,
        options=options,
        section_map=section_map,
        all_sections=all_sections,
        cand_l1=cand_l1,
        cand_l2=cand_l2,
        cand_l3=cand_l3,
    )

    l1_options = build_layout_options(params.Lx, params.Ly, cand_l1, options.prune_level)
    ctx.stats.l1_candidates_total = len(l1_options)
    best: Optional[PanelResult] = None
    timed_out = False
    for opt in l1_options:
        t0 = time.perf_counter()
        try:
            ctx.check_timeout()
        except TimeoutError:
            timed_out = True
            break
        if options.prune_level >= 1 and best is not None:
            lb = lower_bound_layout(1, opt, params.Lx, params.Ly, section_map, params.objective)
            if lb >= best.objective - 1e-9:
                ctx.stats.l1_candidates_pruned += 1
                ctx.traces.append(
                    TraceRow(
                        candidate_key=f"L1:{opt.key}",
                        objective=math.inf,
                        max_util=math.inf,
                        infeasible_reason=f"pruned_by_bound(lb={lb:.3f})",
                        elapsed_sec=time.perf_counter() - t0,
                    )
                )
                continue
        ctx.stats.l1_candidates_evaluated += 1
        try:
            res = evaluate_l1_layout(opt, point_loads, ctx)
        except TimeoutError:
            timed_out = True
            break
        if res.feasible:
            ctx.stats.l1_candidates_feasible += 1
            if best is None or compare_result(res, best):
                best = res
            ctx.traces.append(
                TraceRow(
                    candidate_key=f"L1:{opt.key}",
                    objective=res.objective,
                    max_util=res.max_util,
                    infeasible_reason="",
                    elapsed_sec=time.perf_counter() - t0,
                )
            )
        else:
            ctx.traces.append(
                TraceRow(
                    candidate_key=f"L1:{opt.key}",
                    objective=math.inf,
                    max_util=math.inf,
                    infeasible_reason=res.infeasible_reason,
                    elapsed_sec=time.perf_counter() - t0,
                )
            )

    if best is None:
        return OptimizationResult(
            feasible=False,
            beams=[],
            objective=math.inf,
            total_weight=math.inf,
            total_cost=math.inf,
            max_util=math.inf,
            worst_beam_id="",
            traces=ctx.traces,
            stats=ctx.stats,
            elapsed_sec=ctx.elapsed(),
            timed_out=timed_out,
            message="no feasible solution found",
        )

    beams, worst_beam_id = assign_beam_ids(best.beams)
    return OptimizationResult(
        feasible=True,
        beams=beams,
        objective=best.objective,
        total_weight=best.total_weight,
        total_cost=best.total_cost,
        max_util=best.max_util,
        worst_beam_id=worst_beam_id,
        traces=ctx.traces,
        stats=ctx.stats,
        elapsed_sec=ctx.elapsed(),
        timed_out=timed_out,
        message="ok",
    )


def point_summary(point_loads: List[AppliedPointLoad], max_items: int = 6) -> str:
    if not point_loads:
        return ""
    rows = [f"{pl.source}:{pl.P:.2f}@{pl.a:.2f}" for pl in sorted(point_loads, key=lambda x: x.a)]
    if len(rows) <= max_items:
        return "; ".join(rows)
    return "; ".join(rows[:max_items]) + f"; ... ({len(rows)} loads)"


def draw_layout_png(
    png_path: Path,
    params: Parameters,
    beams: List[BeamRecord],
    point_loads: List[PointLoad],
    show_point_labels: bool = True,
) -> None:
    lx = max(params.Lx, 1e-6)
    ly = max(params.Ly, 1e-6)
    fig_w = 10.0
    fig_h = max(4.0, min(14.0, fig_w * ly / lx))
    fig, ax = plt.subplots(figsize=(fig_w, fig_h), dpi=300)
    ax.plot([0, lx, lx, 0, 0], [0, 0, ly, ly, 0], color="black", linewidth=1.5)
    style = {
        1: {"ls": "-", "lw": 2.4},
        2: {"ls": "--", "lw": 1.8},
        3: {"ls": ":", "lw": 1.5},
    }
    for b in beams:
        st = style.get(b.level, {"ls": "-", "lw": 1.0})
        if b.direction == "X":
            ax.plot([b.span_start, b.span_end], [b.pos, b.pos], color="black", linestyle=st["ls"], linewidth=st["lw"])
            xm = 0.5 * (b.span_start + b.span_end)
            ym = b.pos
        else:
            ax.plot([b.pos, b.pos], [b.span_start, b.span_end], color="black", linestyle=st["ls"], linewidth=st["lw"])
            xm = b.pos
            ym = 0.5 * (b.span_start + b.span_end)
        jitter = ((hash(b.beam_id) % 11) - 5) * 0.004 * min(lx, ly)
        ax.text(xm + jitter, ym + jitter, b.beam_id, fontsize=5.0, color="black")
    for pl in point_loads:
        ax.plot(pl.x, pl.y, marker="x", markersize=6, color="red", linestyle="None")
        if show_point_labels:
            ax.text(pl.x, pl.y, f"{pl.name}({pl.P:.2f}kN)", fontsize=6, color="red")
    ax.set_xlim(-0.03 * lx, 1.03 * lx)
    ax.set_ylim(-0.03 * ly, 1.03 * ly)
    ax.set_aspect("equal", adjustable="box")
    ax.set_xlabel("x [m]")
    ax.set_ylabel("y [m]")
    ax.set_title("Beam Layout (L1 solid / L2 dashed / L3 dotted)")
    ax.grid(True, linestyle=":", linewidth=0.4, alpha=0.7)
    fig.tight_layout()
    fig.savefig(png_path, dpi=300)
    plt.close(fig)


def write_output(
    out_path: Path,
    params: Parameters,
    point_loads: List[PointLoad],
    result: OptimizationResult,
    keep_png: bool = True,
) -> Path:
    wb = Workbook()
    wb.remove(wb.active)

    ws_result = wb.create_sheet("RESULT")
    ws_layout = wb.create_sheet("LAYOUT")
    ws_trace = wb.create_sheet("TRACE")
    ws_sample = wb.create_sheet("SAMPLE")
    ws_fig = wb.create_sheet("LAYOUT_FIG")

    ws_result.append(["Item", "Value"])
    ws_result.append(["Status", "FEASIBLE" if result.feasible else "INFEASIBLE"])
    ws_result.append(["ObjectiveType", params.objective])
    ws_result.append(["ObjectiveValue", result.objective if result.feasible else None])
    ws_result.append(["TotalWeight[kg]", result.total_weight if result.feasible else None])
    ws_result.append(["TotalCost", result.total_cost if result.feasible else None])
    ws_result.append(["MaxUtilization", result.max_util if result.feasible else None])
    ws_result.append(["WorstBeamID", result.worst_beam_id if result.feasible else ""])
    ws_result.append(["ElapsedSec", result.elapsed_sec])
    ws_result.append(["TimedOut", result.timed_out])
    ws_result.append(["Message", result.message])
    ws_result.append(["L1_Candidates_Total", result.stats.l1_candidates_total])
    ws_result.append(["L1_Candidates_Evaluated", result.stats.l1_candidates_evaluated])
    ws_result.append(["L1_Candidates_Pruned", result.stats.l1_candidates_pruned])
    ws_result.append(["L1_Candidates_Feasible", result.stats.l1_candidates_feasible])
    ws_result.append(["L2_Cache_Hit", result.stats.l2_cache_hit])
    ws_result.append(["L2_Cache_Miss", result.stats.l2_cache_miss])
    ws_result.append(["L3_Cache_Hit", result.stats.l3_cache_hit])
    ws_result.append(["L3_Cache_Miss", result.stats.l3_cache_miss])

    ws_result.append([])
    ws_result.append(
        [
            "Level",
            "BeamID",
            "Direction",
            "Pos[m]",
            "SpanStart",
            "SpanEnd",
            "Length[m]",
            "SectionName",
            "w_line[kN/m]",
            "point_loads_summary",
            "Mmax[kN*m]",
            "Vmax[kN]",
            "ymax[mm]",
            "utilization",
        ]
    )
    for b in result.beams:
        ws_result.append(
            [
                b.level,
                b.beam_id,
                b.direction,
                b.pos,
                b.span_start,
                b.span_end,
                b.length,
                b.section_name,
                b.w_line,
                point_summary(b.point_loads),
                b.response.Mmax,
                b.response.Vmax,
                b.response.ymax_mm,
                b.response.utilization,
            ]
        )

    ws_layout.append(["BeamID", "Level", "Direction", "Pos[m]", "SpanStart", "SpanEnd"])
    for b in result.beams:
        ws_layout.append([b.beam_id, b.level, b.direction, b.pos, b.span_start, b.span_end])

    ws_trace.append(["Metric", "Value"])
    ws_trace.append(["L2_Cache_Hit", result.stats.l2_cache_hit])
    ws_trace.append(["L2_Cache_Miss", result.stats.l2_cache_miss])
    ws_trace.append(["L3_Cache_Hit", result.stats.l3_cache_hit])
    ws_trace.append(["L3_Cache_Miss", result.stats.l3_cache_miss])
    ws_trace.append([])
    ws_trace.append(["CandidateKey", "objective", "max_util", "infeasible_reason", "time"])
    traces_sorted = sorted(
        result.traces,
        key=lambda t: (
            math.isinf(t.objective),
            t.objective if not math.isinf(t.objective) else 1e30,
            t.max_util if not math.isinf(t.max_util) else 1e30,
            t.candidate_key,
        ),
    )
    for tr in traces_sorted[:200]:
        ws_trace.append(
            [
                tr.candidate_key,
                None if math.isinf(tr.objective) else tr.objective,
                None if math.isinf(tr.max_util) else tr.max_util,
                tr.infeasible_reason,
                tr.elapsed_sec,
            ]
        )

    if result.feasible and result.beams:
        worst = max(result.beams, key=lambda b: b.response.utilization)
        ws_sample.append(["BeamID", worst.beam_id])
        ws_sample.append(["Level", worst.level])
        ws_sample.append(["Direction", worst.direction])
        ws_sample.append(["Section", worst.section_name])
        ws_sample.append(["Utilization", worst.response.utilization])
        ws_sample.append([])
        ws_sample.append(["x[m]", "V[kN]", "M[kN*m]", "y[mm]"])
        for x, v, m, y in zip(
            worst.response.x_samples,
            worst.response.V_samples,
            worst.response.M_samples,
            worst.response.y_samples_mm,
        ):
            ws_sample.append([x, v, m, y])
        ws_sample.append([])
        ws_sample.append(["PointLoadSource", "P[kN]", "a[m]", "x[m]", "y[m]"])
        for pl in sorted(worst.point_loads, key=lambda z: z.a):
            ws_sample.append([pl.source, pl.P, pl.a, pl.x, pl.y])
    else:
        ws_sample.append(["No feasible beam available"])

    if keep_png:
        png_path = out_path.with_suffix("").with_name(out_path.stem + "_layout.png")
    else:
        tmp = Path(tempfile.gettempdir())
        png_path = tmp / f"{out_path.stem}_layout.png"
    draw_layout_png(png_path, params, result.beams, point_loads, show_point_labels=True)
    ws_fig["A1"] = "Layout figure is embedded below."
    ws_fig.add_image(XLImage(str(png_path)), "A3")

    wb.save(out_path)
    return png_path


def write_template(path: Path) -> None:
    wb = Workbook()
    ws_param = wb.active
    ws_param.title = "PARAM"
    ws_param.append(
        [
            "Lx[m]",
            "Ly[m]",
            "q[kN/m2]",
            "E[N/mm2]",
            "density[kg/m3]",
            "allow_sigma[N/mm2]",
            "allow_tau[N/mm2]",
            "deflection_limit_ratio",
            "unit_system",
            "objective",
            "random_seed",
        ]
    )
    ws_param.append([12.0, 8.0, 3.5, 205000, 7850, 235, 135, 300, "kN-m-mm", "weight", 42])

    ws_pl = wb.create_sheet("POINT_LOADS")
    ws_pl.append(["Name", "P[kN]", "x[m]", "y[m]"])
    ws_pl.append(["PL1", 120.0, 6.0, 4.0])

    ws_sec = wb.create_sheet("SECTIONS")
    ws_sec.append(["SectionName", "Type", "A[mm2]", "I[mm4]", "Z[mm3]", "unit_weight[kg/m]", "cost_per_m"])
    ws_sec.append(["H-200", "H", 3910, 2.61e8, 2.61e6, 30.7, 3200])
    ws_sec.append(["H-250", "H", 5220, 5.03e8, 4.02e6, 41.0, 4300])
    ws_sec.append(["H-300", "H", 6940, 8.92e8, 5.95e6, 54.5, 5600])

    def add_cand(name: str, rows: List[List[object]]) -> None:
        ws = wb.create_sheet(name)
        ws.append(
            [
                "Direction",
                "PitchStart[m]",
                "PitchEnd[m]",
                "PitchStep[m]",
                "OffsetOptions",
                "MaxCount",
                "AllowedSections",
                "MinEdgeClear[m]",
                "AllowMixedXY",
            ]
        )
        for r in rows:
            ws.append(r)

    add_cand(
        "CAND_L1",
        [
            ["X", 2.0, 4.0, 1.0, "0;0.5P", 3, "H-250;H-300", 0.5, False],
            ["Y", 2.0, 4.0, 1.0, "0;0.5P", 3, "H-250;H-300", 0.5, False],
        ],
    )
    add_cand(
        "CAND_L2",
        [
            ["X", 1.5, 3.0, 0.5, "0;0.5P", 4, "H-200;H-250", 0.3, True],
            ["Y", 1.5, 3.0, 0.5, "0;0.5P", 4, "H-200;H-250", 0.3, True],
        ],
    )
    add_cand(
        "CAND_L3",
        [
            ["X", 1.0, 2.5, 0.5, "0;0.5P", 5, "H-200", 0.2, True],
            ["Y", 1.0, 2.5, 0.5, "0;0.5P", 5, "H-200", 0.2, True],
        ],
    )

    ws_opt = wb.create_sheet("OPTIONS")
    ws_opt.append(["SearchMode", "TimeLimitSec", "Parallel", "NWorkers", "PruneLevel"])
    ws_opt.append(["DP", 120, False, 1, 1])
    wb.save(path)


def build_cli() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="Beam optimizer (L1/L2/L3 + DP + Excel I/O + PNG layout embedding)")
    p.add_argument("input_xlsx", nargs="?", help="input template xlsx")
    p.add_argument("output_xlsx", nargs="?", help="output xlsx")
    p.add_argument("--log", default="INFO", help="log level: DEBUG/INFO/WARNING/ERROR")
    p.add_argument("--time-limit", type=float, default=None, help="override OPTIONS.TimeLimitSec")
    p.add_argument("--parallel", action="store_true", help="reserved option; deterministic sequential solver is used")
    p.add_argument("--workers", type=int, default=None, help="reserved worker count")
    p.add_argument("--keep-png", action="store_true", help="keep generated layout png near output")
    p.add_argument("--make-template", default=None, help="write template xlsx and exit")
    return p


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = build_cli().parse_args(argv)
    logging.basicConfig(level=getattr(logging, str(args.log).upper(), logging.INFO), format="%(asctime)s [%(levelname)s] %(message)s")

    if args.make_template:
        out = Path(args.make_template).resolve()
        out.parent.mkdir(parents=True, exist_ok=True)
        write_template(out)
        LOG.info("template created: %s", out)
        return 0

    if not args.input_xlsx or not args.output_xlsx:
        raise SystemExit("usage: python beam_optimizer.py input.xlsx output.xlsx")

    in_path = Path(args.input_xlsx).resolve()
    out_path = Path(args.output_xlsx).resolve()
    if not in_path.exists():
        raise FileNotFoundError(in_path)

    params, point_loads, sections, cand1, cand2, cand3, opts = read_input(in_path)
    if args.time_limit is not None:
        opts = Options(
            search_mode=opts.search_mode,
            time_limit_sec=args.time_limit,
            parallel=opts.parallel,
            n_workers=opts.n_workers,
            prune_level=opts.prune_level,
        )
    if args.parallel:
        LOG.warning("--parallel requested. Current implementation runs deterministic sequential DP.")
    if args.workers:
        LOG.warning("--workers=%s requested. Current implementation runs deterministic sequential DP.", args.workers)

    LOG.info("start optimize: objective=%s, Lx=%.3f, Ly=%.3f", params.objective, params.Lx, params.Ly)
    res = optimize(params, point_loads, sections, cand1, cand2, cand3, opts)
    png = write_output(out_path, params, point_loads, res, keep_png=args.keep_png)
    LOG.info("done: feasible=%s, objective=%s, elapsed=%.2fs", res.feasible, res.objective, res.elapsed_sec)
    LOG.info("output: %s", out_path)
    LOG.info("layout png: %s", png)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
