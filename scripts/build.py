#!/usr/bin/env python3
"""
FUSO MEA Demand Planning Dashboard Builder
Reads FUSO_Advanced_Model_v2.xlsx → outputs data/dashboard_data.json
Run: python3 scripts/build.py
"""

import json
import math
import os
from datetime import datetime

try:
    import openpyxl
except ImportError:
    print("ERROR: pip3 install openpyxl")
    raise

# ─── PATHS ───────────────────────────────────────────────────────────────────
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
EXCEL_PATH = os.path.join(BASE_DIR, "data", "FUSO_Advanced_Model_v2.xlsx")
OUTPUT_PATH = os.path.join(BASE_DIR, "data", "dashboard_data.json")

# ─── CONSTANTS ───────────────────────────────────────────────────────────────
LEAD_TIMES = {"Japan": 45, "Chennai": 21, "GPC Halberstadt": 30}
LEAD_TIME_SIGMA = {"Japan": 2, "Chennai": 3, "GPC Halberstadt": 10}
REVIEW_PERIOD = 30  # days
PORTFOLIO_VALUE_AED = 38_000_000
SMOB_CURRENT_PCT = 0.18
SMOB_TARGET_PCT = 0.05
FORECAST_ACCURACY_CURRENT = 0.70
FORECAST_ACCURACY_TARGET = 0.85
FILL_RATE_CURRENT = 0.86
FILL_RATE_TARGET = 0.95
MARKETS = 58

# Z-scores by class
Z_SCORES = {"A-critical": 2.33, "A": 2.05, "B": 1.65, "C": 1.28}


def z_score_for_class(abc_class: str) -> float:
    if abc_class == "A":
        return Z_SCORES["A"]
    elif abc_class == "B":
        return Z_SCORES["B"]
    return Z_SCORES["C"]


# ─── LOAD RAW SKU DATA FROM EXCEL ────────────────────────────────────────────
def load_sku_data(wb) -> list[dict]:
    ws = wb["ABC_XYZ_MASTER"]
    skus = []
    for row in ws.iter_rows(min_row=3, max_row=52, values_only=True):
        cols = (list(row) + [None] * 21)[:21]
        pn = cols[0]
        desc = cols[1]
        model = cols[2]
        origin = cols[3]
        cost = cols[4]
        demand = cols[5]
        std_dev = cols[9]
        if not pn or not isinstance(cost, (int, float)):
            continue
        demand = demand if isinstance(demand, (int, float)) else 0
        cost = float(cost)
        std_dev = std_dev if isinstance(std_dev, (int, float)) else 0
        skus.append(
            {
                "pn": pn,
                "desc": desc or "",
                "model": model or "",
                "origin": origin or "Japan",
                "cost": float(cost),
                "annual_demand": float(demand),
                "std_dev_monthly": float(std_dev),
            }
        )
    return skus


# ─── ABC-XYZ CLASSIFICATION ───────────────────────────────────────────────────
def classify_abc_xyz(skus: list[dict]) -> list[dict]:
    # Compute annual value
    for sku in skus:
        sku["annual_value"] = sku["cost"] * sku["annual_demand"]

    total_value = sum(s["annual_value"] for s in skus)

    # Sort by annual value descending
    skus.sort(key=lambda x: x["annual_value"], reverse=True)

    cumul = 0.0
    for sku in skus:
        cumul += sku["annual_value"]
        pct = cumul / total_value if total_value else 0
        sku["cumul_value_pct"] = round(pct, 4)
        if pct <= 0.70:
            sku["abc"] = "A"
        elif pct <= 0.90:
            sku["abc"] = "B"
        else:
            sku["abc"] = "C"

    # XYZ by Coefficient of Variation
    for sku in skus:
        mean_monthly = sku["annual_demand"] / 12 if sku["annual_demand"] else 0
        sku["mean_monthly"] = round(mean_monthly, 2)
        cv = sku["std_dev_monthly"] / mean_monthly if mean_monthly > 0 else 99
        sku["cv"] = round(cv, 3)
        if cv < 0.10:
            sku["xyz"] = "X"
        elif cv <= 0.30:
            sku["xyz"] = "Y"
        else:
            sku["xyz"] = "Z"
        sku["combined"] = sku["abc"] + sku["xyz"]

    # Replenishment strategies
    strategies = {
        "AX": "Min-Max / Continuous Review",
        "AY": "(R,S) Periodic Review",
        "AZ": "Demand Sensing + Buffer",
        "BX": "Kanban / Fixed-Period",
        "BY": "Periodic Review Quarterly",
        "BZ": "Order-on-Demand + Emergency Stock",
        "CX": "Min-Max Light",
        "CY": "Periodic Review Annual",
        "CZ": "SMOB Candidate — Order-on-Demand",
    }
    for sku in skus:
        sku["strategy"] = strategies.get(sku["combined"], "Review Required")

    return skus


# ─── SAFETY STOCK CALCULATION ─────────────────────────────────────────────────
def compute_safety_stock(skus: list[dict]) -> list[dict]:
    for sku in skus:
        origin = sku.get("origin", "Japan")
        lt = LEAD_TIMES.get(origin, 45)
        sigma_lt = LEAD_TIME_SIGMA.get(origin, 2)
        z = z_score_for_class(sku["abc"])
        sigma_d = sku["std_dev_monthly"] / 30  # convert monthly to daily
        avg_d = sku["mean_monthly"] / 30

        # Enhanced SS formula: Z × √(LT×σ_d² + avg_d²×σ_LT²)
        ss_enhanced = z * math.sqrt(
            (lt * sigma_d**2) + (avg_d**2 * sigma_lt**2)
        )
        # Standard SS for comparison
        ss_standard = z * sigma_d * math.sqrt(lt + REVIEW_PERIOD)

        rop = (avg_d * lt) + ss_enhanced
        max_stock = ss_enhanced + (avg_d * (lt + REVIEW_PERIOD))

        sku["lead_time_days"] = lt
        sku["z_score"] = z
        sku["ss_standard"] = round(ss_standard, 1)
        sku["ss_enhanced"] = round(ss_enhanced, 1)
        sku["rop"] = round(rop, 1)
        sku["max_stock"] = round(max_stock, 1)
        sku["ss_value_aed"] = round(ss_enhanced * sku["cost"], 0)
    return skus


# ─── SMOB FLAGS ───────────────────────────────────────────────────────────────
SMOB_PARTS = {
    "FUSO-TIM-028": {"months_no_mv": 18, "action": "SCRAP"},
    "FUSO-PIS-030": {"months_no_mv": 16, "action": "SCRAP"},
    "FUSO-DPF-041": {"months_no_mv": 14, "action": "BUNDLE"},
    "FUSO-SCR-043": {"months_no_mv": 12, "action": "BUNDLE"},
    "FUSO-ACP-023": {"months_no_mv": 9, "action": "ROTATE"},
    "FUSO-DRV-022": {"months_no_mv": 8, "action": "ROTATE"},
    "FUSO-HDG-024": {"months_no_mv": 7, "action": "BUNDLE"},
    "FUSO-CRK-025": {"months_no_mv": 8, "action": "ROTATE"},
    "FUSO-CAM-026": {"months_no_mv": 11, "action": "SCRAP"},
    "FUSO-VLV-029": {"months_no_mv": 7, "action": "ROTATE"},
}


def flag_smob(skus: list[dict]) -> list[dict]:
    for sku in skus:
        smob = SMOB_PARTS.get(sku["pn"])
        if smob:
            sku["smob_flag"] = True
            sku["months_no_movement"] = smob["months_no_mv"]
            sku["disposition"] = smob["action"]
            if smob["months_no_mv"] > 12:
                sku["smob_status"] = "OBSOLETE"
            elif smob["months_no_mv"] > 6:
                sku["smob_status"] = "OBSOLETE RISK"
            else:
                sku["smob_status"] = "SLOW MOVING"
        else:
            sku["smob_flag"] = False
            sku["months_no_movement"] = 0
            sku["disposition"] = "RETAIN"
            sku["smob_status"] = "ACTIVE"
    return skus


# ─── AGGREGATE STATS ──────────────────────────────────────────────────────────
def compute_aggregates(skus: list[dict]) -> dict:
    total_skus = len(skus)
    total_annual_value = sum(s["annual_value"] for s in skus)

    # ABC counts and values
    abc_counts = {"A": 0, "B": 0, "C": 0}
    abc_values = {"A": 0.0, "B": 0.0, "C": 0.0}
    xyz_counts = {"X": 0, "Y": 0, "Z": 0}
    combined_counts = {}
    smob_skus = [s for s in skus if s["smob_flag"]]
    smob_value = sum(s["cost"] * max(s["annual_demand"] / 12, 1) for s in smob_skus)

    for sku in skus:
        abc_counts[sku["abc"]] += 1
        abc_values[sku["abc"]] += sku["annual_value"]
        xyz_counts[sku["xyz"]] += 1
        combined_counts[sku["combined"]] = combined_counts.get(sku["combined"], 0) + 1

    # Supply origin split
    origins = {}
    for sku in skus:
        o = sku["origin"]
        origins[o] = origins.get(o, 0) + 1

    total_ss_value = sum(s["ss_value_aed"] for s in skus)

    # Top 10 by value
    top10 = sorted(skus, key=lambda x: x["annual_value"], reverse=True)[:10]

    return {
        "meta": {
            "generated_at": datetime.utcnow().isoformat() + "Z",
            "excel_file": "FUSO_Advanced_Model_v2.xlsx",
            "total_skus": total_skus,
            "markets": MARKETS,
        },
        "kpis": {
            "forecast_accuracy_current": FORECAST_ACCURACY_CURRENT,
            "forecast_accuracy_target": FORECAST_ACCURACY_TARGET,
            "fill_rate_current": FILL_RATE_CURRENT,
            "fill_rate_target": FILL_RATE_TARGET,
            "smob_pct_current": SMOB_CURRENT_PCT,
            "smob_pct_target": SMOB_TARGET_PCT,
            "portfolio_value_aed": PORTFOLIO_VALUE_AED,
            "smob_value_aed": round(PORTFOLIO_VALUE_AED * SMOB_CURRENT_PCT),
            "smob_recovery_target_aed": 1_750_000,
            "total_ss_investment_aed": round(total_ss_value),
        },
        "abc_analysis": {
            "counts": abc_counts,
            "values_aed": {k: round(v) for k, v in abc_values.items()},
            "pct_skus": {
                k: round(abc_counts[k] / total_skus * 100, 1) for k in ["A", "B", "C"]
            },
            "pct_value": {
                k: round(abc_values[k] / total_annual_value * 100, 1)
                for k in ["A", "B", "C"]
            },
        },
        "xyz_analysis": {
            "counts": xyz_counts,
            "pct_skus": {
                k: round(xyz_counts[k] / total_skus * 100, 1) for k in ["X", "Y", "Z"]
            },
        },
        "combined_matrix": combined_counts,
        "supply_origins": origins,
        "smob_summary": {
            "total_smob_skus": len(smob_skus),
            "smob_pct_portfolio": SMOB_CURRENT_PCT,
            "smob_value_aed": round(PORTFOLIO_VALUE_AED * SMOB_CURRENT_PCT),
            "by_action": {
                "SCRAP": sum(1 for s in smob_skus if s["disposition"] == "SCRAP"),
                "BUNDLE": sum(1 for s in smob_skus if s["disposition"] == "BUNDLE"),
                "ROTATE": sum(1 for s in smob_skus if s["disposition"] == "ROTATE"),
            },
            "3year_targets": [
                {"year": "2026", "target_pct": 0.12, "label": "Year 1 — Stabilise"},
                {"year": "2027", "target_pct": 0.08, "label": "Year 2 — Optimise"},
                {"year": "2028", "target_pct": 0.05, "label": "Year 3 — Excellence"},
            ],
        },
        "roadmap": {
            "year1": {
                "label": "2026 — STABILISE",
                "forecast_target": 0.80,
                "fill_rate_target": 0.89,
                "milestones": [
                    "Complete ABC-XYZ classification (all 50 active SKUs)",
                    "Audit SAP APO master data — fix gaps & errors",
                    "Establish MAPE + Bias as baseline KPIs",
                    "Implement enhanced safety stock for top 500 A-X parts",
                    "Launch SMOB disposition — clear EOL model parts",
                    "Establish monthly S&OP cadence with all 58 markets",
                ],
            },
            "year2": {
                "label": "2027 — OPTIMISE",
                "forecast_target": 0.87,
                "fill_rate_target": 0.92,
                "milestones": [
                    "Deploy SBA engine alongside SAP APO baseline",
                    "Integrate fleet telematics (eCanter causal signals)",
                    "Optimise safety stock for all 58 market segments",
                    "Reduce SMOB from 18% to <10%",
                    "GPC Halberstadt bridge stock model live",
                    "Launch customer scorecards + fulfillment reports",
                ],
            },
            "year3": {
                "label": "2028 — EXCELLENCE",
                "forecast_target": 0.85,
                "fill_rate_target": 0.95,
                "milestones": [
                    "Fully automated demand sensing — minimal override",
                    "Real-time replenishment signals to Japan / Chennai / Germany",
                    "SMOB reduced to <5% of total portfolio value",
                    "BI dashboards replace all manual reporting",
                    "Predictive backorder alerts — 30-day advance warning",
                ],
            },
        },
        "top10_skus": [
            {
                "rank": i + 1,
                "pn": s["pn"],
                "desc": s["desc"],
                "model": s["model"],
                "abc": s["abc"],
                "xyz": s["xyz"],
                "combined": s["combined"],
                "annual_value_aed": round(s["annual_value"]),
                "ss_enhanced": s["ss_enhanced"],
                "strategy": s["strategy"],
            }
            for i, s in enumerate(top10)
        ],
        "all_skus": [
            {
                "pn": s["pn"],
                "desc": s["desc"],
                "model": s["model"],
                "origin": s["origin"],
                "cost": s["cost"],
                "annual_demand": s["annual_demand"],
                "annual_value_aed": round(s["annual_value"]),
                "abc": s["abc"],
                "xyz": s["xyz"],
                "combined": s["combined"],
                "cv": s["cv"],
                "mean_monthly": s["mean_monthly"],
                "ss_enhanced": s["ss_enhanced"],
                "rop": s["rop"],
                "strategy": s["strategy"],
                "smob_flag": s["smob_flag"],
                "smob_status": s["smob_status"],
                "disposition": s["disposition"],
            }
            for s in skus
        ],
    }


# ─── MAIN ─────────────────────────────────────────────────────────────────────
def main():
    print(f"Loading: {EXCEL_PATH}")
    wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)

    skus = load_sku_data(wb)
    print(f"  Loaded {len(skus)} SKUs")

    skus = classify_abc_xyz(skus)
    skus = compute_safety_stock(skus)
    skus = flag_smob(skus)

    data = compute_aggregates(skus)

    os.makedirs(os.path.dirname(OUTPUT_PATH), exist_ok=True)
    with open(OUTPUT_PATH, "w") as f:
        json.dump(data, f, indent=2)

    print(f"  Written: {OUTPUT_PATH}")
    print(f"  SKUs: {data['meta']['total_skus']}")
    print(f"  ABC → A:{data['abc_analysis']['counts']['A']}  B:{data['abc_analysis']['counts']['B']}  C:{data['abc_analysis']['counts']['C']}")
    xyz = data["xyz_analysis"]["counts"]
    print(f"  XYZ → X:{xyz['X']}  Y:{xyz['Y']}  Z:{xyz['Z']}")
    print(f"  SMOB SKUs: {data['smob_summary']['total_smob_skus']}")
    print("Done.")


if __name__ == "__main__":
    main()
