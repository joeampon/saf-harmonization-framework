"""
harmonization/engine.py
=======================
Core harmonization engine applying the five-step protocol to all 48 studies.

Five-step protocol
------------------
Step 1 : Convert MFSP to 2023 USD/GGE  (CPI + unit conversion)
Step 2 : Normalize allocation method to energy allocation (ISO 14044 §4.3.4.2)
Step 3 : Adjust system boundary to Well-to-Wake (+3.0 gCO2e/MJ for WtG studies)
Step 4 : Remove ILUC if study included it (restore ILUC-free baseline)
Step 5 : CRF normalization to 10% DR / 30-yr lifetime / 90% CF
         Uses pathway-specific CAPEX fraction (config.CAPEX_FRACTION) to avoid
         over-correcting feedstock-dominated pathways (e.g., HEFA).

Key fix vs v1.0
---------------
v1.0 used a uniform CAPEX_FRACTION=0.40 across all pathways.
HEFA MFSP is ~87 % feedstock cost; applying 40 % capex fraction
over-corrected the CRF adjustment by ~3x for HEFA studies.
This version uses pathway-specific fractions derived from the modal-parameter
cost breakdown (Section 4 of paper).
"""

import numpy as np
import pandas as pd
from typing import Dict, Any

import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from config import (
    CEPCI, CPI_TO_2023, EUR_USD, ALLOC_FACTORS, ILUC_GCO2E_MJ,
    CAPEX_FRACTION, HARMONIZED_DISCOUNT_RATE, HARMONIZED_CAPACITY_FACTOR,
    HARMONIZED_PLANT_LIFETIME, WTG_TO_WTWAKE_DELTA,
    L_JET_PER_GGE, JET_LHV_BTU_PER_GAL, GASOLINE_LHV_BTU_PER_GAL, L_PER_GAL,
)


# ---------------------------------------------------------------------------
# Helper: CEPCI interpolation
# ---------------------------------------------------------------------------

def _cepci_value(year: int) -> float:
    """Return CEPCI value for a given year with linear interpolation."""
    if year in CEPCI:
        return CEPCI[year]
    years = sorted(CEPCI.keys())
    if year < years[0]:
        return CEPCI[years[0]]
    if year > years[-1]:
        return CEPCI[years[-1]]
    idx = next(i for i, y in enumerate(years) if y > year)
    y1, y2 = years[idx - 1], years[idx]
    return CEPCI[y1] + (CEPCI[y2] - CEPCI[y1]) * (year - y1) / (y2 - y1)


# ---------------------------------------------------------------------------
# Helper: Capital Recovery Factor
# ---------------------------------------------------------------------------

def crf(rate: float, lifetime: int) -> float:
    """
    Capital recovery factor: annualises a lump-sum capital investment.

    CRF = r(1+r)^n / [(1+r)^n - 1]

    Parameters
    ----------
    rate     : Real discount rate (fraction, e.g. 0.10)
    lifetime : Plant economic lifetime (years)

    Returns
    -------
    float  CRF value
    """
    if rate == 0:
        return 1.0 / lifetime
    return (rate * (1 + rate) ** lifetime) / ((1 + rate) ** lifetime - 1)


# Precomputed reference CRF at harmonized basis
CRF_REF = crf(HARMONIZED_DISCOUNT_RATE, HARMONIZED_PLANT_LIFETIME)


# ---------------------------------------------------------------------------
# Step 1: MFSP unit conversion and CPI escalation
# ---------------------------------------------------------------------------

def _to_usd(cost: float, currency: str, year: int) -> float:
    """Convert cost to USD using historical exchange rates."""
    if currency == "EUR":
        rate = EUR_USD.get(year)
        if rate is None:
            raise ValueError(f"No EUR/USD rate for year {year}")
        return cost * rate
    return float(cost)


def _cpi_escalate(cost_usd: float, from_year: int) -> float:
    """Scale cost from `from_year` to 2023 USD using CPI."""
    factor = CPI_TO_2023.get(from_year)
    if factor is None:
        # Extrapolate modestly for years outside table
        factor = CPI_TO_2023.get(max(CPI_TO_2023.keys()), 1.0)
    return cost_usd * factor


def mfsp_to_2023_usd_per_gge(mfsp_raw: float, unit: str,
                               ref_year: int, currency: str) -> float:
    """
    Convert reported MFSP to 2023 USD/GGE.

    Supported units
    ---------------
    $/GGE   — direct; CPI escalation only
    $/L     — convert L jet -> GGE then CPI escalate
    EUR/L   — EUR->USD then L->GGE then CPI escalate

    Note: MFSP already amortises CAPEX, so CPI (not CEPCI) is used here.
    CEPCI correction is applied separately in Step 5 via CRF normalisation.
    """
    cost_usd = _to_usd(mfsp_raw, currency, ref_year)
    cost_2023 = _cpi_escalate(cost_usd, ref_year)

    unit_clean = unit.strip()
    if unit_clean == "$/GGE":
        return cost_2023
    elif unit_clean in ("$/L", "EUR/L"):
        return cost_2023 * L_JET_PER_GGE
    else:
        raise ValueError(f"Unknown MFSP unit: '{unit}'. Expected $/GGE, $/L, or EUR/L.")


# ---------------------------------------------------------------------------
# Step 5: CRF and capacity-factor normalisation of MFSP
# ---------------------------------------------------------------------------

def _normalise_mfsp_crf(mfsp_2023: float, pathway: str,
                          dr_study: float, lt_study: int,
                          cf_study: float) -> float:
    """
    Adjust MFSP for discount rate, plant lifetime, and capacity factor.

    Only the CAPEX-embedded portion of MFSP scales with CRF changes;
    O&M and feedstock costs are invariant to DR and lifetime.

    MFSP_harm = MFSP_raw * {(1 - f_cap) + f_cap * [CRF_ref/CRF_study] * [CF_study/CF_ref]}

    where f_cap = pathway-specific capital cost fraction of MFSP.

    Parameters
    ----------
    mfsp_2023  : MFSP in 2023 USD/GGE (after unit conversion + CPI)
    pathway    : "ATJ" | "HEFA" | "FT-SPK" | "PtL"
    dr_study   : Discount rate used in the study (fraction)
    lt_study   : Plant lifetime used in the study (years)
    cf_study   : Capacity factor used in the study (fraction)
    """
    f_cap = CAPEX_FRACTION.get(pathway, 0.40)
    crf_study = crf(dr_study, lt_study)

    if crf_study == 0:
        return mfsp_2023   # degenerate case — no correction

    mfsp_harm = mfsp_2023 * (
        (1 - f_cap)
        + f_cap * (CRF_REF / crf_study) * (cf_study / HARMONIZED_CAPACITY_FACTOR)
    )
    return mfsp_harm


# ---------------------------------------------------------------------------
# Main harmonization function
# ---------------------------------------------------------------------------

def harmonize_study(row: Dict[str, Any]) -> Dict[str, Any]:
    """
    Apply the five-step harmonization protocol to one literature study.

    Parameters
    ----------
    row : dict
        One row from LiteratureDatabase.STUDIES

    Returns
    -------
    dict with harmonized fields (appended to original row fields)
    """
    pathway  = row["pathway"]
    feedstock = row["feedstock"]

    # ── Step 1: MFSP -> 2023 USD/GGE ─────────────────────────────────────────
    mfsp_2023_raw = mfsp_to_2023_usd_per_gge(
        row["mfsp_raw"], row["mfsp_unit"], row["ref_year_cost"], row["currency"]
    )

    # ── Step 5: CRF normalisation ─────────────────────────────────────────────
    mfsp_harmonized = _normalise_mfsp_crf(
        mfsp_2023_raw,
        pathway,
        dr_study = row["discount_rate"] / 100.0,
        lt_study = row["plant_lifetime"],
        cf_study = row["capacity_factor"] / 100.0,
    )

    # ── Step 2: Allocation normalisation ─────────────────────────────────────
    alloc_table  = ALLOC_FACTORS.get(pathway, {})
    study_alloc  = alloc_table.get(row["allocation"], 1.0)
    ref_alloc    = alloc_table.get("energy", 1.0)

    ghg_raw = float(row["ghg_raw"])

    # Guard: PtL always has alloc_factor=1.0; avoid division by zero elsewhere
    if study_alloc > 0:
        alloc_correction = ref_alloc / study_alloc
    else:
        alloc_correction = 1.0

    ghg_harm = ghg_raw * alloc_correction

    # ── Step 3: System boundary correction (WtG -> WtWake) ───────────────────
    boundary_corrected = False
    if row["boundary"] in ("WtG", "GtG"):
        ghg_harm += WTG_TO_WTWAKE_DELTA
        boundary_corrected = True
    # WtW treated as equivalent to WtWake for bio-SAF (delta < 0.5 gCO2e/MJ)

    # ── Step 4: ILUC removal ──────────────────────────────────────────────────
    iluc_removed = 0.0
    if row["include_iluc"]:
        iluc_est = ILUC_GCO2E_MJ.get(feedstock, 0.0)
        ghg_harm = max(ghg_harm - iluc_est, 0.0)
        iluc_removed = iluc_est

    # ── GHG reduction relative to petroleum baseline ─────────────────────────
    from config import PETROLEUM_JET_GHG_WTW
    ghg_reduction_pct = (PETROLEUM_JET_GHG_WTW - ghg_harm) / PETROLEUM_JET_GHG_WTW * 100.0

    # ── CRF correction factor (for reporting) ─────────────────────────────────
    crf_study = crf(row["discount_rate"] / 100.0, row["plant_lifetime"])
    crf_correction = CRF_REF / crf_study if crf_study > 0 else 1.0

    return {
        "mfsp_2023_raw":             round(mfsp_2023_raw,   3),
        "mfsp_harmonized":           round(mfsp_harmonized, 3),
        "ghg_harmonized":            round(ghg_harm,        2),
        "ghg_reduction_pct":         round(ghg_reduction_pct, 1),
        "alloc_correction_factor":   round(alloc_correction, 4),
        "boundary_correction_applied": boundary_corrected,
        "iluc_removed_gco2e_mj":    round(iluc_removed,    2),
        "crf_correction_factor":     round(crf_correction,  4),
        "capex_fraction_used":       CAPEX_FRACTION.get(pathway, 0.40),
    }


def build_harmonized_dataset() -> pd.DataFrame:
    """
    Harmonize all 48 studies and return a merged DataFrame.

    Returns
    -------
    pd.DataFrame  with original columns + harmonization output columns
    """
    from data.literature_database import get_dataframe
    df = get_dataframe()

    harmonized_rows = []
    errors = []
    for idx, row in df.iterrows():
        try:
            h = harmonize_study(row.to_dict())
            harmonized_rows.append(h)
        except Exception as e:
            errors.append(f"Study {row.get('study_id', idx)}: {e}")
            harmonized_rows.append({})   # placeholder

    if errors:
        print(f"WARNING: {len(errors)} harmonization error(s):")
        for err in errors:
            print(f"  {err}")

    harm_df = pd.DataFrame(harmonized_rows)
    return pd.concat(
        [df.reset_index(drop=True), harm_df.reset_index(drop=True)], axis=1
    )
