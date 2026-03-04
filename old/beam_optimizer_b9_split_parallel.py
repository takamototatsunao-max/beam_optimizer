#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""beam_optimizer_b9_split_parallel.py

beam_optimizer_b9.py を「インプット→前準備→並列計算→後処理→アウトプット」に
分割した“組立型”エントリポイント。

設計方針
--------
- 並列化するのは候補（direction×pitch）単位の純計算（solve_layout）のみ
- openpyxl による Excel I/O は単一プロセス（競合回避）
- 再現性: 同一入力で同一 best が出るよう、候補順序と同点規則を固定
- 検証用出力（A〜E相当）/TRACE/DEBUG_* は全候補分を収集して RESULT に出力

実行
----
python beam_optimizer_b9_split_parallel.py input.xlsx output.xlsx

備考
----
本ファイルは元コードを“再利用”し、既存の Excel 出力仕様を維持する。
"""

from __future__ import annotations

import os
import sys
import math
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple

from concurrent.futures import ProcessPoolExecutor, as_completed


# 既存実装を丸ごと取り込む（同ディレクトリに beam_optimizer_b9.py がある前提）
import beam_optimizer_b9 as core


@dataclass
class CandidateSpec:
    cand_no: int
    cand_id: str
    direction: str
    pitch: float


@dataclass
class WorkerResult:
    cand_no: int
    row: core.CandidateRow
    sol: core.Solution
    # debug bundles (per candidate)
    trace: List[Dict[str, object]]
    dbg_trials: List[Dict[str, object]]
    dbg_main_geoms: List[Dict[str, object]]
    dbg_trans_defs: List[Dict[str, object]]
    dbg_member_final: List[Dict[str, object]]
    dbg_alloc_final: List[Dict[str, object]]


# -----------------------------
# 1) インプット
# -----------------------------
def input_stage(in_path: str) -> Tuple[core.Config, core.Material, core.SolverSettings, List[core.Section]]:
    cfg, mat, setts, sections = core.read_input_xlsx(in_path)

    # basic validation (same as original main)
    core.ensure_positive("Lx", cfg.Lx)
    core.ensure_positive("Ly", cfg.Ly)
    core.ensure_positive("q", cfg.q)
    core.ensure_positive("E", mat.E_kN_m2)
    core.ensure_positive("fb", mat.fb_kN_m2)
    core.ensure_positive("fv", mat.fv_kN_m2)
    core.ensure_positive("deflection_limit", mat.deflection_limit)

    return cfg, mat, setts, sections


# -----------------------------
# 2) 並列計算前準備
# -----------------------------
def _fill_section_props(sections: List[core.Section]) -> List[core.Section]:
    """断面諸元が欠損している場合、文字列から推定して埋める（前準備）。

    core.Section は frozen dataclass なので in-place 変更せず、新しいリストを返す。
    """
    out: List[core.Section] = []
    for s in sections:
        h, b, tw, tf = s.h, s.b, s.tw, s.tf
        A_mm2, I_mm4, Z_mm3, Av_mm2 = s.A_mm2, s.I_mm4, s.Z_mm3, s.Av_mm2

        if any(v is None for v in (h, b, tw, tf)):
            dims = core.parse_h_section_dims(s.name)
            if dims:
                h, b, tw, tf = dims

        if any(v is None for v in (A_mm2, I_mm4, Z_mm3, Av_mm2)) and all(v is not None for v in (h, b, tw, tf)):
            A, I, Z, Av = core.approx_h_section_props_mm(float(h), float(b), float(tw), float(tf))
            A_mm2 = A_mm2 if A_mm2 is not None else A
            I_mm4 = I_mm4 if I_mm4 is not None else I
            Z_mm3 = Z_mm3 if Z_mm3 is not None else Z
            Av_mm2 = Av_mm2 if Av_mm2 is not None else Av

        out.append(core.replace(
            s,
            h=h, b=b, tw=tw, tf=tf,
            A_mm2=A_mm2, I_mm4=I_mm4, Z_mm3=Z_mm3, Av_mm2=Av_mm2,
        ))
    return out


def prepare_stage(cfg: core.Config, sections: List[core.Section]) -> Tuple[List[CandidateSpec], List[core.Section]]:
    """候補（cand）生成＋明らかな不成立の除外。順序は固定（再現性）。"""
    sections2 = _fill_section_props(sections)

    pitches = core.make_pitch_candidates(cfg)
    directions: List[str] = []
    if cfg.enable_x:
        directions.append("X")
    if cfg.enable_y:
        directions.append("Y")
    if not directions:
        raise ValueError("Both X and Y directions are disabled.")

    s_axis = core.short_side_axis(cfg.Lx, cfg.Ly)

    specs: List[CandidateSpec] = []
    cand_no = 0
    for d in directions:
        pdir = core.pitch_direction_of(d)
        for pitch in pitches:
            # short-side pitch limit
            if pdir == s_axis and pitch > cfg.short_pitch_limit + 1e-9:
                continue

            cand_no += 1
            specs.append(CandidateSpec(cand_no=cand_no, cand_id=f"C{cand_no:03d}", direction=d, pitch=float(pitch)))

    return specs, sections2


# -----------------------------
# 3) 並列計算（ワーカー）
# -----------------------------
def _worker_eval(args) -> WorkerResult:
    """トップレベル関数（Windowsのmultiprocessingでpickle可能にする）。"""
    cfg, mat, setts, sections, spec = args

    # per-candidate collectors are globals inside core; clear them first
    core._clear_debug()

    sol = core.solve_layout(cfg, mat, setts, sections, spec.direction, spec.pitch, cand_id=spec.cand_id)

    n_main = len([m for m in sol.member_checks if m.member_type == "MAIN"])
    n_trans = len([m for m in sol.member_checks if m.member_type == "TRANS"])
    row = core.CandidateRow(
        direction=spec.direction,
        pitch=spec.pitch,
        n_main=n_main,
        n_trans=n_trans,
        max_rank_used=sol.max_rank_used if sol.ok else 0,
        total_weight=sol.total_weight,
        Mmax=sol.Mmax,
        Vmax=sol.Vmax,
        dmax_mm=sol.dmax * 1000.0,
        util_max=sol.util_max,
        ok=sol.ok,
        ng_reason=sol.ng_reason,
    )

    # copy debug bundles out of the worker process
    return WorkerResult(
        cand_no=spec.cand_no,
        row=row,
        sol=sol,
        trace=list(core._TRACE),
        dbg_trials=list(core._DBG_MEMBER_TRIALS),
        dbg_main_geoms=list(core._DBG_MAIN_GEOMS),
        dbg_trans_defs=list(core._DBG_TRANS_DEFS),
        dbg_member_final=list(core._DBG_MEMBER_FINAL),
        dbg_alloc_final=list(core._DBG_ALLOC_FINAL),
    )


def parallel_stage(
    cfg: core.Config,
    mat: core.Material,
    setts: core.SolverSettings,
    sections: List[core.Section],
    specs: List[CandidateSpec],
    max_workers: Optional[int] = None,
) -> List[WorkerResult]:
    """cand 単位の並列評価。返却順は cand_no で後処理時に固定。"""
    if max_workers is None:
        # Keep some headroom; Windows desktop utilization stability
        cpu = os.cpu_count() or 1
        max_workers = max(1, min(cpu, 8))

    # NOTE: 共有データ（cfg/mat/setts/sections）をそのままpickleで各プロセスへ送る。
    #       sections は前準備で諸元を埋めているのでワーカー側の補完負荷は小さい。
    tasks = [(cfg, mat, setts, sections, s) for s in specs]

    results: List[WorkerResult] = []
    with ProcessPoolExecutor(max_workers=max_workers) as ex:
        futs = [ex.submit(_worker_eval, t) for t in tasks]
        for fut in as_completed(futs):
            results.append(fut.result())

    # determinism
    results.sort(key=lambda x: x.cand_no)
    return results


# -----------------------------
# 4) 並列計算後作業
# -----------------------------
def postprocess_stage(worker_results: List[WorkerResult]) -> Tuple[List[core.CandidateRow], Optional[core.Solution]]:
    """候補の集計、best選定、全候補のデバッグ出力を core のグローバルへ統合。"""

    # merge debug bundles into core globals (for Excel verbose sheets)
    core._clear_debug()

    merged_trace: List[Dict[str, object]] = []
    for wr in worker_results:
        merged_trace.extend(wr.trace)
        core._DBG_MEMBER_TRIALS.extend(wr.dbg_trials)
        core._DBG_MAIN_GEOMS.extend(wr.dbg_main_geoms)
        core._DBG_TRANS_DEFS.extend(wr.dbg_trans_defs)
        core._DBG_MEMBER_FINAL.extend(wr.dbg_member_final)
        core._DBG_ALLOC_FINAL.extend(wr.dbg_alloc_final)

    # Renumber TRACE.seq to be unique and stable
    for i, e in enumerate(merged_trace, start=1):
        e["seq"] = i
    core._TRACE.extend(merged_trace)

    # candidate rows in deterministic order
    cand_rows = [wr.row for wr in worker_results]

    # select best (exclude NG)
    best: Optional[core.Solution] = None
    for wr in worker_results:
        sol = wr.sol
        if not sol.ok:
            continue
        if best is None:
            best = sol
            continue
        if sol.total_weight < best.total_weight - 1e-12:
            best = sol
            continue
        if abs(sol.total_weight - best.total_weight) <= 1e-12:
            # tie-breaker fixed: util_max smaller wins
            if sol.util_max < best.util_max - 1e-12:
                best = sol

    return cand_rows, best


# -----------------------------
# 5) アウトプット
# -----------------------------
def output_stage(
    in_path: str,
    out_path: str,
    cfg: core.Config,
    mat: core.Material,
    setts: core.SolverSettings,
    sections: List[core.Section],
    cand_rows: List[core.CandidateRow],
    best: Optional[core.Solution],
) -> None:
    core.write_result_xlsx(in_path, out_path, cfg, mat, setts, sections, cand_rows, best)


# -----------------------------
# Orchestrator
# -----------------------------
def run(in_path: str, out_path: str, max_workers: Optional[int] = None) -> None:
    cfg, mat, setts, sections = input_stage(in_path)
    specs, sections2 = prepare_stage(cfg, sections)
    worker_results = parallel_stage(cfg, mat, setts, sections2, specs, max_workers=max_workers)
    cand_rows, best = postprocess_stage(worker_results)
    output_stage(in_path, out_path, cfg, mat, setts, sections2, cand_rows, best)


def main(argv: List[str]) -> int:
    if len(argv) < 2:
        print("Usage: python beam_optimizer_b9_split_parallel.py input.xlsx output.xlsx [max_workers]", file=sys.stderr)
        return 2

    in_path = argv[1]
    out_path = argv[2] if len(argv) >= 3 else (in_path.replace(".xlsx", "") + "_out.xlsx")
    max_workers = int(argv[3]) if len(argv) >= 4 else None

    run(in_path, out_path, max_workers=max_workers)
    print(f"Done. Output written to: {out_path}")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main(sys.argv))
    except PermissionError as e:
        print(f"[ERROR] Permission denied. Close the target Excel file and retry: {e}", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"[ERROR] {e}", file=sys.stderr)
        sys.exit(1)
