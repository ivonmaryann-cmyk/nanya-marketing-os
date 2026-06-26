from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

import pandas as pd

from . import price_calculator_v3 as pc


@dataclass
class ConstraintDecision:
    blocked: bool
    grade: str
    summary: str
    reasons: list[str]


def _norm_text(value: object) -> str:
    return str(value if value is not None else "").strip()


def _classify_desc(desc: str) -> str:
    if re.search(r"\d+M/Roll", desc, re.IGNORECASE):
        return "roll"
    is_pp_prefix = bool(re.match(r"^PP\s+", desc, re.IGNORECASE))
    is_pp_implicit = (
        not is_pp_prefix
        and bool(re.search(r"RC\s*\d+\s*%", desc, re.IGNORECASE))
        and not bool(re.search(r"\d+\.?\d*\s*mm", desc, re.IGNORECASE))
        and not bool(re.search(r"\d+M/Roll", desc, re.IGNORECASE))
    )
    return "pp" if (is_pp_prefix or is_pp_implicit) else "ccl"


def _pp_metadata(desc: str, df_price: pd.DataFrame) -> dict:
    desc = pc.normalize_str(desc)
    mode = _classify_desc(desc)
    pp_rows = df_price[df_price["CCL"].astype(str).str.strip() == "PP"].copy()

    match = re.match(r"PP\s+([\w\-\(\)\.]+)", desc)
    if match:
        raw_glue = match.group(1).strip()
        lam_match = re.search(r"PP\s+[\w\-\(\)\.]+\s+(\d+)\s+RC", desc, re.IGNORECASE)
    else:
        match2 = re.match(r"([\w\-\(\)\.]+)\s+(\d+)\s+RC", desc, re.IGNORECASE)
        raw_glue = match2.group(1).strip() if match2 else ""
        lam_match = match2
    laminate_type = lam_match.group(1) if match else (lam_match.group(2) if lam_match else "")
    rc_percent = pc.parse_rc_percent(desc)
    converted_glue = pc._try_convert_pp_glue(raw_glue) if raw_glue else raw_glue

    def candidates(glue: str, contains: bool) -> pd.DataFrame:
        if not glue:
            return pp_rows.iloc[0:0]
        thickness_col = pp_rows["不含铜板厚/（mm)"].astype(str).str.strip()
        mask = pp_rows["型号"].astype(str).str.strip() == glue
        if contains:
            mask &= thickness_col.str.contains(str(laminate_type), na=False)
        else:
            mask &= thickness_col == str(laminate_type)
        return pp_rows[mask]

    raw_exact = candidates(raw_glue, contains=False)
    raw_contains = candidates(raw_glue, contains=True)
    conv_exact = candidates(converted_glue, contains=False)
    conv_contains = candidates(converted_glue, contains=True)

    exact_count = 0
    range_count = 0
    chosen = raw_exact
    glue_fallback = False

    for candidate_set, used_fallback in [
        (raw_exact, False),
        (raw_contains, False),
        (conv_exact, converted_glue != raw_glue),
        (conv_contains, converted_glue != raw_glue),
    ]:
        if candidate_set.empty:
            continue
        exact_rows = candidate_set[pd.to_numeric(candidate_set["铜厚"], errors="coerce") == rc_percent]
        if not exact_rows.empty:
            exact_count = len(exact_rows)
            chosen = exact_rows
            glue_fallback = used_fallback
            break
        range_rows = []
        for idx, row in candidate_set.iterrows():
            if pc._in_cu_range(_norm_text(row["铜厚"]), rc_percent):
                range_rows.append(idx)
        if range_rows:
            range_count = len(range_rows)
            chosen = candidate_set.loc[range_rows]
            glue_fallback = used_fallback
            break

    return {
        "mode": mode,
        "glue_fallback": glue_fallback,
        "rc_range_fallback": exact_count == 0 and range_count > 0,
        "multi_candidate": exact_count > 1 or range_count > 1,
        "roll_length_unchecked": mode == "roll",
        "chosen_count": len(chosen),
    }


def _ccl_match_rows(
    ccl_rows: pd.DataFrame,
    *,
    glue: str,
    thickness: str,
    cu: str,
    foil: str,
    laminate: str,
    contains_laminate: bool,
) -> pd.DataFrame:
    mask = (
        (ccl_rows["型号"].astype(str).str.strip() == glue)
        & (ccl_rows["不含铜板厚/（mm)"].astype(str).str.strip() == thickness)
        & (ccl_rows["铜厚"].astype(str).str.strip() == cu)
        & (ccl_rows["铜箔"].astype(str).str.strip() == foil)
    )
    if contains_laminate:
        mask &= ccl_rows["叠构"].astype(str).str.contains(laminate, na=False)
    else:
        mask &= ccl_rows["叠构"].astype(str).str.strip() == laminate
    return ccl_rows[mask]


def _ccl_metadata(desc: str, df_price: pd.DataFrame, df_account: pd.DataFrame) -> dict:
    desc = pc.normalize_str(desc)
    ccl_rows = df_price[df_price["CCL"].astype(str).str.strip() == "CCL"].copy()

    glue_match = re.match(r"([\w\-\(\)\.]+)\s+\d", desc)
    glue = glue_match.group(1).strip() if glue_match else desc.split()[0]
    thickness = pc.parse_thickness(desc)
    thickness_str = pc.format_thickness(thickness) if thickness is not None else ""
    cu_thick = pc.parse_copper_thickness(desc) or ""
    foil_type = pc.parse_foil_type(desc, cu_thick) or ""
    laminate = pc.parse_laminate(desc) or ""
    size = pc.parse_size(desc)

    foil_candidates = [foil_type]
    if "/" in foil_type and cu_thick:
        foil_parts = foil_type.split("/")
        cu_parts = cu_thick.split("/")
        if len(foil_parts) == 2 and len(cu_parts) == 2:
            chosen = foil_parts[1] if pc.cu_thick_value(cu_parts[1]) >= pc.cu_thick_value(cu_parts[0]) else foil_parts[0]
            foil_candidates.append(chosen)

    cu_candidates = [cu_thick]
    if cu_thick in pc.CU_THICK_FALLBACK:
        cu_candidates.append(pc.CU_THICK_FALLBACK[cu_thick])

    exact_rows = ccl_rows.iloc[0:0]
    contains_rows = ccl_rows.iloc[0:0]
    used_fallback_chain = False
    multi_candidate = False

    for foil in foil_candidates:
        for cu in cu_candidates:
            rows = _ccl_match_rows(
                ccl_rows,
                glue=glue,
                thickness=thickness_str,
                cu=cu,
                foil=foil,
                laminate=laminate,
                contains_laminate=False,
            )
            if not rows.empty:
                exact_rows = rows
                used_fallback_chain = foil != foil_type or cu != cu_thick
                multi_candidate = len(rows) > 1
                break
        if not exact_rows.empty:
            break

    if exact_rows.empty:
        for foil in foil_candidates:
            for cu in cu_candidates:
                rows = _ccl_match_rows(
                    ccl_rows,
                    glue=glue,
                    thickness=thickness_str,
                    cu=cu,
                    foil=foil,
                    laminate=laminate,
                    contains_laminate=True,
                )
                if not rows.empty:
                    contains_rows = rows
                    used_fallback_chain = foil != foil_type or cu != cu_thick
                    multi_candidate = len(rows) > 1
                    break
            if not contains_rows.empty:
                break

    nearest_thickness = False
    if thickness is not None and exact_rows.empty and contains_rows.empty:
        cands = ccl_rows[
            (ccl_rows["型号"].astype(str).str.strip() == glue)
            & (ccl_rows["铜厚"].astype(str).str.strip().isin(cu_candidates))
            & (ccl_rows["铜箔"].astype(str).str.strip().isin(foil_candidates))
            & (ccl_rows["叠构"].astype(str).str.contains(laminate, na=False))
        ]
        if not cands.empty:
            thick_vals = pd.to_numeric(cands["不含铜板厚/（mm)"], errors="coerce").dropna()
            nearest_thickness = not thick_vals.empty and thickness_str not in {pc.format_thickness(v) for v in thick_vals.tolist()}

    account_multi_path = False
    is_standard = False
    if size:
        is_standard = pc.is_standard_size(size[0], size[1])
        if not is_standard:
            code = pc.size_to_code(size[0], size[1])
            account_matches = df_account[df_account["品名"].astype(str).str.contains(code, na=False)]
            account_multi_path = len(account_matches) > 1

    return {
        "mode": "ccl",
        "standard_size": is_standard,
        "laminate_contains": exact_rows.empty and not contains_rows.empty,
        "nearest_thickness": nearest_thickness,
        "account_multi_path": account_multi_path,
        "multi_candidate": multi_candidate,
        "used_generic_fallback": used_fallback_chain,
    }


def _build_metadata(desc: str, df_price: pd.DataFrame, df_account: pd.DataFrame) -> dict:
    mode = _classify_desc(desc)
    if mode in {"pp", "roll"}:
        return _pp_metadata(desc, df_price)
    return _ccl_metadata(desc, df_price, df_account)


def _note_glue_fallback(note: str) -> bool:
    match = re.search(r"原始胶系=([^|]+)→匹配=([^|]+)", note)
    if not match:
        return False
    return match.group(1).strip() != match.group(2).strip()


def apply_confirmation_constraints(
    desc: str,
    price: float,
    note: str,
    *,
    df_price: pd.DataFrame,
    df_account: pd.DataFrame,
    confirm_bundle: dict,
) -> ConstraintDecision:
    policy = confirm_bundle["policy"]
    allow_flags = policy["allow_flags"]
    block_flags = policy["block_flags"]
    metadata = _build_metadata(desc, df_price, df_account)

    reasons: list[str] = []

    if metadata["mode"] in {"pp", "roll"}:
        glue_fallback = _note_glue_fallback(note)
        if glue_fallback and block_flags.get("BLOCK-001", False) and not allow_flags.get("ALLOW-007", False):
            reasons.append("触发 BLOCK-001：PP 胶系发生回退，当前规则不允许自动写价")
        if metadata.get("rc_range_fallback") and block_flags.get("BLOCK-002", False):
            reasons.append("触发 BLOCK-002：RC 未精确命中，只命中了近似范围")
        if metadata.get("roll_length_unchecked") and block_flags.get("BLOCK-003", False):
            reasons.append("触发 BLOCK-003：PP 卷料规格未做精确规则确认，当前默认转人工")
        if metadata.get("multi_candidate") and block_flags.get("BLOCK-007", False):
            reasons.append("触发 BLOCK-007：同条件命中多条价格候选，当前不允许自动写价")
    else:
        if metadata.get("laminate_contains") and block_flags.get("BLOCK-004", False):
            reasons.append("触发 BLOCK-004：叠构仅包含匹配，未精确命中")
        if metadata.get("nearest_thickness") and block_flags.get("BLOCK-005", False):
            reasons.append("触发 BLOCK-005：厚度仅能取最近值，当前不允许自动放行")
        if metadata.get("account_multi_path") and block_flags.get("BLOCK-006", False):
            reasons.append("触发 BLOCK-006：非标尺寸存在多条基板拆分路径")
        if metadata.get("multi_candidate") and block_flags.get("BLOCK-007", False):
            reasons.append("触发 BLOCK-007：同条件命中多条价格候选，当前不允许自动写价")
        if metadata.get("used_generic_fallback") and block_flags.get("BLOCK-011", False):
            reasons.append("触发 BLOCK-011：当前命中依赖未批准的兜底/回退链路")

    if reasons:
        return ConstraintDecision(
            blocked=True,
            grade="C-待人工确认",
            summary=f"规则确认表拦截：{'; '.join(reasons)}",
            reasons=reasons,
        )

    grade = "A-自动通过"
    if metadata.get("used_generic_fallback") or metadata.get("glue_fallback"):
        grade = "B-规则通过"
    return ConstraintDecision(
        blocked=False,
        grade=grade,
        summary=f"{grade} | 规则确认表允许自动写价",
        reasons=[],
    )
