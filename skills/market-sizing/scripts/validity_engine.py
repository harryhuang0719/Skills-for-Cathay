#!/usr/bin/env python3
"""Market Sizing Validity Engine v3.2 — Execution Integrity Enforcement."""

from enum import IntEnum
from dataclasses import dataclass, field
from typing import List, Dict, Any, Optional


class Severity(IntEnum):
    L1_COSMETIC = 1
    L2_COMMENTARY = 2
    L3_MECHANISM = 3
    L4_VALIDITY = 4


@dataclass
class Issue:
    severity: Severity
    rule: str
    message: str
    field: str = ""


# ── Field Taxonomy ───────────────────────────────────────────────────────────

IDENTITY_FIELDS = [
    "sizing_objective", "measurement_basis", "realization_basis",
    "time_horizon_type", "billing_unit", "market_boundary",
]

MECHANISM_FIELDS = [
    "primary_archetype", "formula_contract.volume_driver",
    "formula_contract.price_driver", "formula_contract.competition_driver",
    "generator_type", "model_governance_bias",
]

COMMENTARY_FIELDS = [
    "archetype_rationale", "investment_conclusion.notes",
    "secondary_archetype",
]

REQUIRED_STATE_VARS = {
    "population": ["target_base", "penetration", "units_per_adopter"],
    "installed_base": ["compatible_installed_base", "attach_rate", "replacement_rate", "eligibility_ceiling"],
    "throughput": ["node_base", "utilization", "throughput_per_node", "price_per_unit"],
    "project": ["project_pipeline", "conversion", "avg_ticket", "recognition_timing"],
    "substitution": ["legacy_base", "substitution_rate", "switching_friction", "relative_asp_bridge"],
    "commodity": ["demand_volume", "supply_volume", "inventory_buffer", "price_formation_rule"],
}


def _get_nested(cfg: dict, dotpath: str):
    parts = dotpath.split(".")
    v = cfg
    for p in parts:
        if not isinstance(v, dict):
            return None
        v = v.get(p)
    return v


def _is_empty(v) -> bool:
    if v is None:
        return True
    if isinstance(v, str) and not v.strip():
        return True
    if isinstance(v, (list, dict)) and len(v) == 0:
        return True
    return False


# ── Pre-flight (blocks generation) ──────────────────────────────────────────

def preflight_check(cfg: dict) -> List[Issue]:
    issues = []
    # A-class: identity fields
    for f in IDENTITY_FIELDS:
        if _is_empty(_get_nested(cfg, f)):
            issues.append(Issue(Severity.L4_VALIDITY, "IDENTITY_FIELD", f"{f} is empty — model cannot start", f))
    # B-class: mechanism fields
    for f in MECHANISM_FIELDS:
        if _is_empty(_get_nested(cfg, f)):
            issues.append(Issue(Severity.L3_MECHANISM, "MECHANISM_FIELD", f"{f} is empty — model not publishable", f))
    # C-class: commentary
    for f in COMMENTARY_FIELDS:
        if _is_empty(_get_nested(cfg, f)):
            issues.append(Issue(Severity.L2_COMMENTARY, "COMMENTARY_FIELD", f"{f} is empty", f))
    # Minimum state variables
    arch = cfg.get("primary_archetype", "")
    if arch in REQUIRED_STATE_VARS:
        msv = cfg.get("minimum_state_variables", {})
        for v in REQUIRED_STATE_VARS[arch]:
            if not msv.get(v):
                issues.append(Issue(Severity.L4_VALIDITY, "STATE_VAR", f"Missing state variable {v} for {arch}", v))
    return issues


# ── No Silent Fallback ──────────────────────────────────────────────────────

def check_no_silent_fallback(cfg: dict) -> List[Issue]:
    issues = []
    yn = len(cfg.get("years", []))
    # Top-down estimates
    td = cfg.get("demand", {}).get("top_down_estimates", [])
    if not td or all(e.get("value", 0) == 0 for e in td):
        issues.append(Issue(Severity.L4_VALIDITY, "NO_SILENT_FALLBACK",
                            "top_down_estimates empty/all-zero — cannot silently skip TD gate"))
    # Price mechanism vs asp_mechanism
    pm = cfg.get("demand", {}).get("price_mechanism", {})
    asp_mech = cfg.get("asp_mechanism", "")
    if asp_mech == "gap_driven":
        if pm.get("shortage_elasticity", 0) == 0 and pm.get("surplus_elasticity", 0) == 0:
            issues.append(Issue(Severity.L4_VALIDITY, "NO_SILENT_FALLBACK",
                                "gap_driven ASP but all elasticities=0 — silent fallback to flat ASP"))
    # Player revenue with capacity but zero revenue
    for p in cfg.get("supply", {}).get("players", []):
        caps = p.get("capacity", [])
        revs = p.get("revenue_estimates", [])
        has_cap = any(c > 0 for c in caps)
        all_zero_rev = all((revs[j] if j < len(revs) else 0) == 0 for j in range(yn))
        if has_cap and all_zero_rev and yn > 0:
            issues.append(Issue(Severity.L3_MECHANISM, "NO_SILENT_FALLBACK",
                                f"{p['name']} has capacity but all revenue_estimates=0"))
    return issues


# ── Formula Realization ─────────────────────────────────────────────────────

def check_formula_realization(cfg: dict) -> List[Issue]:
    issues = []
    overrides = {o.get("scope", ""): o for o in cfg.get("field_overrides", [])}
    for p in cfg.get("supply", {}).get("players", []):
        scope = f"supply.{p['name']}.revenue_estimates"
        revs = p.get("revenue_estimates", [])
        if revs and any(r > 0 for r in revs):
            if scope not in overrides:
                issues.append(Issue(Severity.L3_MECHANISM, "FORMULA_REALIZATION",
                                    f"{p['name']} revenue_estimates are hand-filled constants without override declaration",
                                    scope))
            else:
                otype = overrides[scope].get("type", "")
                if otype == "temporary_placeholder":
                    issues.append(Issue(Severity.L4_VALIDITY, "FORMULA_REALIZATION",
                                        f"{p['name']} revenue uses temporary_placeholder override — not allowed in final output",
                                        scope))
                elif otype == "expert_judgment":
                    issues.append(Issue(Severity.L2_COMMENTARY, "FORMULA_REALIZATION",
                                        f"{p['name']} revenue uses expert_judgment override — logged", scope))
    return issues


# ── V3.3: Demand Engine Realization ────────────────────────────────────────

def check_demand_engine(cfg: dict) -> List[Issue]:
    issues = []
    de = cfg.get("demand_engine")
    fc = cfg.get("formula_contract", {})
    arch = cfg.get("primary_archetype", "")

    if fc.get("volume_driver") and not de:
        issues.append(Issue(Severity.L3_MECHANISM, "ENGINE_REALIZATION",
                            f"formula_contract.volume_driver declared but no demand_engine — formula is metadata only",
                            "demand_engine"))

    if de:
        params = de.get("params", {})
        if not params:
            issues.append(Issue(Severity.L4_VALIDITY, "ENGINE_REALIZATION",
                                "demand_engine declared but params is empty", "demand_engine.params"))
        # Check allocation sums to ~100%
        segs = cfg.get("demand", {}).get("segments", [])
        yn = len(cfg.get("years", []))
        for j in range(yn):
            total_alloc = sum(s.get("allocation_pct", [0])[j] if j < len(s.get("allocation_pct", [])) else 0 for s in segs)
            if abs(total_alloc - 1.0) > 0.02:
                issues.append(Issue(Severity.L3_MECHANISM, "ENGINE_REALIZATION",
                                    f"Year {j+1} allocation_pct sum={total_alloc:.3f}, expected ~1.0",
                                    "segments.allocation_pct"))
                break  # one warning is enough
        # Check archetype consistency
        if arch and de.get("archetype") and arch != de.get("archetype"):
            issues.append(Issue(Severity.L3_MECHANISM, "ENGINE_REALIZATION",
                                f"primary_archetype={arch} but demand_engine.archetype={de.get('archetype')}",
                                "demand_engine.archetype"))
    return issues


# ── Gate Realization ────────────────────────────────────────────────────────

def check_gates(cfg: dict) -> List[Issue]:
    issues = []
    # Top-down gate
    td = cfg.get("demand", {}).get("top_down_estimates", [])
    if not td:
        issues.append(Issue(Severity.L4_VALIDITY, "GATE_REALIZATION",
                            "Top-down gate: trigger input missing — gate cannot default to PASS"))
    # Bridge validity gate
    forecast_yrs = sum(1 for y in cfg.get("years", []) if "E" in str(y))
    gbv = cfg.get("generator_bridge_validation", {})
    if cfg.get("generator_bridge_required") is True:
        true_count = sum(1 for v in gbv.values() if v)
        if true_count < 2:
            issues.append(Issue(Severity.L4_VALIDITY, "GATE_REALIZATION",
                                f"Bridge gate: only {true_count}/4 dimensions true — pseudo-bridge"))
    elif forecast_yrs > 3 and not cfg.get("generator_bridge_required"):
        issues.append(Issue(Severity.L3_MECHANISM, "GATE_REALIZATION",
                            f"Forecast {forecast_yrs}Y but generator_bridge_required not set"))
    # Archetype purity gate
    arch = cfg.get("primary_archetype", "")
    fc = cfg.get("formula_contract", {})
    if arch == "commodity" and fc.get("price_driver", "").lower() == "exogenous":
        issues.append(Issue(Severity.L4_VALIDITY, "GATE_REALIZATION",
                            "Archetype purity: commodity requires gap-driven ASP, not exogenous"))
    # Competition denominator gate
    cd = cfg.get("competition_denominator_basis", {})
    if cd:
        boundary = cfg.get("market_boundary", "")
        geo = cd.get("geography", "")
        if geo and boundary and geo.lower() not in boundary.lower():
            issues.append(Issue(Severity.L3_MECHANISM, "GATE_REALIZATION",
                                f"Competition denominator geography={geo} may not match market_boundary"))
    return issues


# ── Dimension & Basis Integrity ─────────────────────────────────────────────

def check_dimensions(cfg: dict) -> List[Issue]:
    issues = []
    uc = cfg.get("unit_contract", {})
    vol_u = uc.get("volume_unit", cfg.get("unit", ""))
    rev_u = uc.get("revenue_unit", cfg.get("revenue_unit", ""))
    # Unit integrity: K units * $/unit should need /1000 to get 
    if vol_u.startswith("K") and rev_u == "":
        bridge = uc.get("scale_bridge", "")
        if not bridge:
            issues.append(Issue(Severity.L3_MECHANISM, "UNIT_INTEGRITY",
                                f"vol={vol_u}, rev={rev_u} but no scale_bridge declared"))
    # Basis integrity
    rb = cfg.get("realization_basis", "")
    mb = cfg.get("measurement_basis", "")
    if rb and mb:
        if "normalized" in mb and "realized" in rb:
            issues.append(Issue(Severity.L3_MECHANISM, "BASIS_INTEGRITY",
                                f"measurement_basis={mb} vs realization_basis={rb} — potential mismatch"))
    # Denominator integrity
    cd = cfg.get("competition_denominator_basis", {})
    if cd:
        time_b = cd.get("time_basis", "")
        boundary = cfg.get("market_boundary", "")
        if time_b and boundary:
            if time_b.lower() not in boundary.lower() and boundary.lower().find(time_b.lower()) == -1:
                issues.append(Issue(Severity.L2_COMMENTARY, "DENOMINATOR_INTEGRITY",
                                    f"competition time_basis={time_b} not explicit in market_boundary"))
    return issues


# ── Override Governance ─────────────────────────────────────────────────────

def check_overrides(cfg: dict) -> List[Issue]:
    issues = []
    for o in cfg.get("field_overrides", []):
        otype = o.get("type", "")
        scope = o.get("scope", "unknown")
        reason = o.get("reason", "")
        if not reason:
            issues.append(Issue(Severity.L3_MECHANISM, "OVERRIDE_GOVERNANCE",
                                f"Override {scope} has no reason", scope))
        if otype == "temporary_placeholder":
            issues.append(Issue(Severity.L4_VALIDITY, "OVERRIDE_GOVERNANCE",
                                f"temporary_placeholder override {scope} in output", scope))
        elif otype == "expert_judgment":
            issues.append(Issue(Severity.L2_COMMENTARY, "OVERRIDE_GOVERNANCE",
                                f"expert_judgment override: {scope}", scope))
    return issues


# ── Segment Heterogeneity ───────────────────────────────────────────────────

def check_segment_heterogeneity(cfg: dict) -> List[Issue]:
    issues = []
    segments = cfg.get("demand", {}).get("segments", [])
    if len(segments) < 2:
        return issues
    billing_units = set()
    price_formations = set()
    for seg in segments:
        bu = seg.get("segment_billing_unit", cfg.get("billing_unit", ""))
        pf = seg.get("segment_price_formation", "")
        if bu: billing_units.add(bu)
        if pf: price_formations.add(pf)
    diffs = 0
    if len(billing_units) > 1: diffs += 1
    if len(price_formations) > 1: diffs += 1
    if diffs >= 1:
        sh = cfg.get("segment_heterogeneity", {})
        if not sh.get("justification"):
            issues.append(Issue(Severity.L3_MECHANISM, "SEGMENT_HETEROGENEITY",
                                f"Segments differ on {diffs} dimensions but no heterogeneity justification"))
    return issues


# ── Three-Layer Validity Rollup ─────────────────────────────────────────────

def compute_validity(issues: List[Issue]) -> Dict[str, str]:
    structural_rules = {"IDENTITY_FIELD", "STATE_VAR", "GATE_REALIZATION"}
    mechanical_rules = {"FORMULA_REALIZATION", "NO_SILENT_FALLBACK", "UNIT_INTEGRITY",
                        "BASIS_INTEGRITY", "OVERRIDE_GOVERNANCE"}
    l4s = [i for i in issues if i.severity == Severity.L4_VALIDITY]
    l3s = [i for i in issues if i.severity == Severity.L3_MECHANISM]
    struct_fail = any(i.rule in structural_rules for i in l4s)
    mech_fail = any(i.rule in mechanical_rules for i in l4s)
    l3_count = len(l3s)
    override_count = sum(1 for i in issues if "OVERRIDE" in i.rule)
    if l3_count == 0 and override_count == 0:
        econ = "high"
    elif l3_count <= 2:
        econ = "medium"
    else:
        econ = "low"
    return {
        "structural": "invalid" if struct_fail else "valid",
        "mechanical": "invalid" if mech_fail else "valid",
        "economic": econ,
    }


# ── Full Validation Run ────────────────────────────────────────────────────

def validate_all(cfg: dict) -> tuple:
    all_issues = []
    all_issues.extend(preflight_check(cfg))
    all_issues.extend(check_no_silent_fallback(cfg))
    all_issues.extend(check_formula_realization(cfg))
    all_issues.extend(check_demand_engine(cfg))
    all_issues.extend(check_gates(cfg))
    all_issues.extend(check_dimensions(cfg))
    all_issues.extend(check_overrides(cfg))
    all_issues.extend(check_segment_heterogeneity(cfg))
    validity = compute_validity(all_issues)
    return all_issues, validity


def issues_to_legacy(issues: List[Issue]) -> tuple:
    fails = [f"[{i.rule}] {i.message}" for i in issues if i.severity >= Severity.L4_VALIDITY]
    warns = [f"[{i.rule}] {i.message}" for i in issues if i.severity in (Severity.L2_COMMENTARY, Severity.L3_MECHANISM)]
    return fails, warns
