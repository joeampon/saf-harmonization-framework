"""
analysis/variance_decomposition.py
====================================
Variance decomposition: methodological vs. technical/economic drivers.

Uses Sobol first-order indices summed by parameter class to quantify
what fraction of output variance is attributable to methodological choices
(allocation, system boundary, ILUC) vs. technical/economic parameters.

External validation
-------------------
VALIDATION_STUDIES below contains only studies whose MFSP, GHG intensity,
allocation method, system boundary, discount rate, and capacity factor were
ALL confirmed from the full open-access paper text.  Studies that were found
but are behind paywalls are listed as commented-out stubs — add their values
once you have accessed the full paper.

Currently confirmed (open-access, full text verified):
  • Rojas-Michaga et al. 2023  (PtL)    DOI 10.1016/j.enconman.2023.117427
  • Marchesan et al. 2025      (HEFA)   DOI 10.1016/j.biortech.2024.131772
  • Ahire et al. 2024          (FT-SPK) DOI 10.1039/D4SE00749B
  • Rojas-Michaga et al. 2025  (FT-SPK) DOI 10.1016/j.ecmx.2024.100841

Stubs requiring full-text access:
  • Detsios et al. 2024         (FT-SPK)  DOI 10.3390/en17071685
  • Greene et al. 2025          (HEFA)    DOI 10.1021/acs.est.4c06742
  • Watson et al. 2025          (ATJ)     DOI 10.1021/acs.iecr.4c03039
  • Kourkoumpas et al. 2024     (ATJ)     DOI 10.1016/j.renene.2024.120512
  • Rogachuk & Okolie 2024      (FT-SPK)  DOI 10.1016/j.enconman.2024.118110  [EXCLUDED]

Add entries to VALIDATION_STUDIES as you retrieve values from each paper.
The CV-reduction aggregate is only computed for pathways with ≥ 2 entries.
"""

import numpy as np
import pandas as pd
from typing import Dict

import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from data.parameter_distributions import METHODOLOGICAL_PARAMS


# ---------------------------------------------------------------------------
# External validation studies — real peer-reviewed papers only.
#
# HOW TO ADD AN ENTRY
# -------------------
# Each dict requires:
#   study_id              : short unique key
#   authors               : "Surname et al."
#   year                  : publication year (int)
#   journal               : full journal name
#   doi                   : real DOI string
#   pathway               : "ATJ" | "HEFA" | "FT-SPK" | "PtL"
#   mfsp_reported_usd_gge : MFSP pre-converted to 2023 USD/GGE
#                           (apply CPI + unit conversion yourself before entry)
#   ghg_reported_gco2emj  : GHG intensity as reported in the paper (gCO2e/MJ)
#   allocation            : "energy"|"mass"|"economic"|"system_expansion"|"exergy"
#   boundary              : "WtG"|"WtW"|"WtWake"
#   discount_rate         : discount rate used in the study (%)
#   capacity_factor       : annual utilisation (%)
#
# Aggregate CV reduction is only calculated for pathways with >= 2 entries.
# ---------------------------------------------------------------------------

VALIDATION_STUDIES = [

    # =========================================================================
    # CONFIRMED — all values verified from open-access full text
    # =========================================================================

    # Rojas-Michaga et al. 2023  |  Energy Conversion and Management 292:117427
    # DOI: 10.1016/j.enconman.2023.117427  |  CC-BY open access (White Rose)
    # Pathway: PtL (DAC + offshore wind electrolysis + reverse-WGS + FT synthesis)
    # MFSP conversion: 5.16 £/kg × 2.841 kg/GGE (from config constants)
    #                  × 1.284 $/£ (2020 avg, Fed Reserve)
    #                  × 1.17 (US CPI 2020→2023, config.CPI_TO_2023[2020])
    #                  = 22.02 $/GGE (2023 USD)
    # GHG: 21.43 gCO2eq/MJ, WtWake, exergy allocation  — abstract + Table 5
    # DR: 10%  — Table 1 (DCFA parameters)
    # CF: 8000 h/yr ÷ 8760 h/yr = 91.3%  — Section 3.1
    {
        "study_id":             "ROJAS-MICHAGA2023",
        "authors":              "Rojas-Michaga et al.",
        "year":                 2023,
        "journal":              "Energy Conversion and Management",
        "doi":                  "10.1016/j.enconman.2023.117427",
        "pathway":              "PtL",
        "mfsp_reported_usd_gge": 22.02,
        "ghg_reported_gco2emj": 21.43,
        "allocation":           "energy",   # exergy allocation treated as energy basis
        "boundary":             "WtWake",
        "discount_rate":        10,
        "capacity_factor":      91,
    },

    # Marchesan et al. 2025  |  Bioresource Technology 416:131772
    # DOI: 10.1016/j.biortech.2024.131772  |  ScienceDirect open access
    # Pathway: HEFA — microbial oil from sugarcane (MO-HEFA, SR scenario)
    # CONFIRMED from full paper text:
    #   MFSP    = 2.06 USD/L (SR base case, Section 3.4, standalone HEFA facility)
    #   GHG     = 35.2 gCO2eq/MJ (SR base case, Section 3.7 + Fig 6, cradle-to-grave)
    #   Alloc   = energy (Section 2.2: "energy-based allocation…lower heating values")
    #   Boundary= WtWake (Section 2.2: "cradle-to-grave…including distribution and use")
    #   DR      = 12% IRR (Section 2.2: "IRR of 12%")
    #   CF      = 90% (330 operating days / 365, Section 2.1)
    #   Lifetime= 25 years (Section 2.2)
    #   Ref year= 2019 USD (Section 2.2: "CEPCI applied to adjust to 2019 values")
    # MFSP conversion: 2.06 USD/L × 1.18 (CPI 2019→2023) × 3.534 L/GGE = 8.59 $/GGE
    {
        "study_id":             "MARCHESAN2025",
        "authors":              "Marchesan et al.",
        "year":                 2025,
        "journal":              "Bioresource Technology",
        "doi":                  "10.1016/j.biortech.2024.131772",
        "pathway":              "HEFA",
        "mfsp_reported_usd_gge": 8.59,
        "ghg_reported_gco2emj": 35.2,
        "allocation":           "energy",
        "boundary":             "WtWake",
        "discount_rate":        12,
        "capacity_factor":      90,
    },

    # Ahire et al. 2024  |  Sustainable Energy & Fuels  |  DOI 10.1039/D4SE00749B
    # RSC open-access HTML — full text retrieved and confirmed
    # Pathway: FT-SPK — gasification + Fischer-Tropsch synthesis from forest residues
    # AUTHORS: J.P. Ahire, R. Bergman, T. Runge, S.H. Mousavi-Avval,
    #          D. Bhattacharyya, T. Brown, J. Wang
    # CONFIRMED from full paper text (RSC HTML):
    #   MFSP    = $1.87/kg = $1.44/L = $5.45/gal at 0% profit margin
    #             quote: "the minimum selling price (MSP) of FT-SPK-SAF was
    #                     $1.87 per kg or $1.44 L ($5.45 per gallon)"
    #   GHG     = 24.6 gCO2eq/MJ (cradle-to-gate, global warming impact)
    #             quote: "The global warming impact of forest residue-based SAF was
    #                     estimated to be 24.6 gCO2 eq. per MJ of SAF"
    #             Note: ~72% reduction vs conventional jet fuel (89 gCO2e/MJ)
    #   Alloc   = system_expansion — co-product credits for electricity, green diesel,
    #             and green propane applied (Table 1)
    #   Boundary= cradle-to-gate (WtG equivalent; combustion excluded)
    #             quote: "cradle-to-gate of SAF production from forest residues"
    #   DR      = 8% loan rate (Table 1: "Loan rate: 8%")
    #   CF      = 90% (Table 1: "Operations days per year: 330 (90% uptime)")
    #   Lifetime= 25 years (Table 1: "Plant life: 25 years + 36 months construction")
    #   Plant   = 90 Mg SAF/day; feedstock: 960 Mg/day forest residues (30% moisture)
    #   Location= Pacific Northwest (Washington, northern California, Oregon)
    #   Ref year= not explicitly stated; inferred ~2022 USD from publication date and
    #             feedstock cost ($40/Mg) consistent with 2022 PNW regional pricing
    # MFSP conversion: $1.44/L × 3.534 L/GGE = $5.09/GGE (2022 USD)
    #                  × 1.04 (CPI_TO_2023[2022]) = $5.29/GGE (2023 USD)
    {
        "study_id":             "AHIRE2024",
        "authors":              "Ahire et al.",
        "year":                 2024,
        "journal":              "Sustainable Energy & Fuels",
        "doi":                  "10.1039/D4SE00749B",
        "pathway":              "FT-SPK",
        "mfsp_reported_usd_gge": 5.29,
        "ghg_reported_gco2emj": 24.6,
        "allocation":           "system_expansion",
        "boundary":             "WtG",
        "discount_rate":        8,
        "capacity_factor":      90,
    },

    # Rojas-Michaga et al. 2025  |  Energy Conversion and Management: X 25:100841
    # DOI: 10.1016/j.ecmx.2024.100841  |  CC-BY open access (White Rose eprint 234482)
    # Pathway: PBtL (Power and Biomass to Liquid) — biomass gasification + green H2
    #   from wind electrolysis + Fischer-Tropsch synthesis.  Classified as FT-SPK
    #   because FT synthesis is the primary conversion route and biomass is the
    #   primary carbon source (green H2 boosts carbon efficiency of syngas).
    # AUTHORS: M.F. Rojas-Michaga, S. Michailos, E. Cardozo, K.J. Hughes,
    #          D. Ingham, M. Pourkashanian  (University of Sheffield)
    # CONFIRMED from full open-access PDF (White Rose eprint):
    #   MJSP    = 0.0672 £/MJ (0%TS deterministic base case — 100% raw biomass)
    #             quote (abstract): "minimum jet fuel selling prices (MJSP) ranging from
    #                                0.0651 to 0.0673 £/MJ" (deterministic, all scenarios)
    #             Monte Carlo median (0%TS): 0.084 £/MJ (Table 8)
    #   GHG     = 13.93 gCO2eq/MJ (0%TS, Well-to-Wake, baseline energy-allocation)
    #             quote (abstract): "Global warming potentials range from −105.33 to
    #                                13.93 gCO₂eq/MJ" — 0%TS scenario is the +13.93 extreme
    #             Note: higher-TS scenarios achieve negative GWP via BECCS credit
    #   Alloc   = energy (Table 4: "Approach 1 (baseline allocation approach): energy")
    #             quote: "energy allocation is preferred for such systems"
    #   Boundary= Well-to-Wake (Table 3 + Fig 3: "WtWa assessment"; "system boundary
    #             extends from resource extraction … to SAF utilization (combustion)")
    #   DR      = 10% (Table 3: "Discount rate: 10 %")
    #   CF      = 85% assumed — explicit CF not stated; plant designed for continuous
    #             operation with grid backup during low-wind periods (conservative)
    #   Lifetime= 20 years (Table 3: "Plant life: 20 years")
    #   Currency= GBP 2022 (Table 3: "Base year: 2022")
    # MFSP conversion (0%TS deterministic MJSP = 0.0672 £/MJ):
    #   × 1.237 $/£ (Fed Reserve 2022 GBP/USD annual avg)
    #   × 34.37 MJ/L (JET_LHV_MJ_PER_L from config)
    #   × 3.534 L/GGE (L_JET_PER_GGE from config)
    #   × 1.04  (CPI_TO_2023[2022])
    #   = 0.0672 × 1.237 × 34.37 × 3.534 × 1.04 ≈ 10.49 $/GGE (2023 USD)
    {
        "study_id":             "ROJAS-MICHAGA2025",
        "authors":              "Rojas-Michaga et al.",
        "year":                 2025,
        "journal":              "Energy Conversion and Management: X",
        "doi":                  "10.1016/j.ecmx.2024.100841",
        "pathway":              "FT-SPK",
        "mfsp_reported_usd_gge": 10.49,
        "ghg_reported_gco2emj": 13.93,
        "allocation":           "energy",
        "boundary":             "WtWake",
        "discount_rate":        10,
        "capacity_factor":      85,
    },

    # =========================================================================
    # STUBS — papers found but not yet fully accessible
    # Uncomment and complete each entry once you have retrieved the full text.
    # =========================================================================

    # -------------------------------------------------------------------------
    # Detsios et al. 2024  |  Energies 17(7):1685  |  DOI 10.3390/en17071685
    # BtL (gasification-driven, benchmarked against FT-SPK)
    # FROM ABSTRACT : MFSP = 1.83 €/L (2023 EUR), DR = 6%, CF = 85%
    # STILL NEEDED  : GHG (gCO2e/MJ) — TEA-only paper, LCA is in companion study
    #                 allocation method — check Section 3 of full text
    #                 system boundary   — check Section 3 of full text
    # MFSP pre-conversion: 1.83 EUR/L × 3.534 L/GGE × 1.081 EUR/USD(2023) = 6.99 $/GGE
    # {
    #     "study_id":             "DETSIOS2024",
    #     "authors":              "Detsios et al.",
    #     "year":                 2024,
    #     "journal":              "Energies",
    #     "doi":                  "10.3390/en17071685",
    #     "pathway":              "FT-SPK",
    #     "mfsp_reported_usd_gge": 6.99,
    #     "ghg_reported_gco2emj": ???,   # look up in companion Kourkoumpas LCA paper
    #     "allocation":           ???,   # look up in Section 3 (methodology)
    #     "boundary":             ???,   # look up in Section 3 (methodology)
    #     "discount_rate":        6,
    #     "capacity_factor":      85,
    # },

    # -------------------------------------------------------------------------
    # Greene et al. 2025  |  Env. Sci. Technol. 59(7):3472  |  DOI 10.1021/acs.est.4c06742
    # HEFA from microalgae (county-level US geospatial study)
    # FROM ABSTRACT : MFSP current $5.79–$10.93/LGE (wide range = geographic variation,
    #                 not sensitivity analysis — no single central value usable)
    #                 GHG = "70% reduction vs petroleum" only — no absolute gCO2e/MJ
    # STILL NEEDED  : single central MFSP value, absolute GHG, DR, CF, allocation, boundary
    #                 — all in full paper (paywalled, access via ISU library)
    # {
    #     "study_id":             "GREENE2025",
    #     "authors":              "Greene et al.",
    #     "year":                 2025,
    #     "journal":              "Environmental Science & Technology",
    #     "doi":                  "10.1021/acs.est.4c06742",
    #     "pathway":              "HEFA",
    #     "mfsp_reported_usd_gge": ???,  # use median/central county value from paper
    #     "ghg_reported_gco2emj": ???,   # 70% reduction → ~26.7 gCO2e/MJ but confirm exact value
    #     "allocation":           ???,
    #     "boundary":             ???,
    #     "discount_rate":        ???,
    #     "capacity_factor":      ???,
    # },

    # -------------------------------------------------------------------------
    # Watson et al. 2025  |  Ind. Eng. Chem. Res. 64(8):4410
    # DOI: 10.1021/acs.iecr.4c03039
    # ATJ from sugarcane ethanol (Brazil stochastic optimisation study)
    # NOT USABLE AS WRITTEN — this paper reports only the price PREMIUM above fossil
    # jet fuel ($0.40–$2.00/L), not an absolute MFSP. It is a market/policy study,
    # not a standard TEA. GHG value (24.1 gCO2e/MJ) is a CORSIA citation, not the
    # authors' own LCA. Cannot be entered until an absolute MFSP is found.
    # {
    #     "study_id":             "WATSON2025",
    #     "authors":              "Watson et al.",
    #     "year":                 2025,
    #     "journal":              "Industrial & Engineering Chemistry Research",
    #     "doi":                  "10.1021/acs.iecr.4c03039",
    #     "pathway":              "ATJ",
    #     "mfsp_reported_usd_gge": ???,  # check supplementary material for absolute production cost
    #     "ghg_reported_gco2emj": ???,   # do NOT use the CORSIA 24.1 value — that is not theirs
    #     "allocation":           ???,
    #     "boundary":             ???,
    #     "discount_rate":        8,
    #     "capacity_factor":      ???,
    # },

    # -------------------------------------------------------------------------
    # Kourkoumpas et al. 2024  |  Renewable Energy 228:120512
    # DOI: 10.1016/j.renene.2024.120512
    # ATJ from bioethanol plant retrofit (LCA only — no TEA/MFSP)
    # FROM ABSTRACT : GHG = 44.15 gCO2e/MJ (retrofit scenario),
    #                        44.53 gCO2e/MJ (new-build ATJ scenario)
    # NOT USABLE AS WRITTEN — LCA-only paper, no MFSP reported.
    # If a companion TEA paper exists, pair this GHG with that MFSP.
    # STILL NEEDED  : MFSP, DR, CF, allocation method, system boundary
    # {
    #     "study_id":             "KOURKOUMPAS2024",
    #     "authors":              "Kourkoumpas et al.",
    #     "year":                 2024,
    #     "journal":              "Renewable Energy",
    #     "doi":                  "10.1016/j.renene.2024.120512",
    #     "pathway":              "ATJ",
    #     "mfsp_reported_usd_gge": ???,  # not in this paper — check companion TEA study
    #     "ghg_reported_gco2emj": 44.15, # retrofit scenario — confirmed from abstract
    #     "allocation":           ???,   # check Section 2 (system boundary and allocation)
    #     "boundary":             ???,   # check Section 2
    #     "discount_rate":        ???,
    #     "capacity_factor":      ???,
    # },

    # -------------------------------------------------------------------------
    # Rogachuk & Okolie 2024  |  Energy Conv. Mgmt. 302:118110
    # DOI: 10.1016/j.enconman.2024.118110
    # *** NOT SUITABLE FOR VALIDATION — EXCLUDED ***
    # Full paper text confirmed (full text read):
    #   MFSP    = 0.66 USD/L (2022 base year, Table 3)
    #   GHG     = 58.6 kg CO2eq/kg SAF (Section 3.3, Well-to-Tank boundary)
    #             Unit note: paper reports in kg CO2eq PER kg SAF (not per MJ)
    #             Conversion: 58.6 kg CO2eq/kg ÷ 0.04321 MJ/g (jet LHV 43.21 MJ/kg)
    #                       = 58,600 g CO2eq ÷ 43.21 MJ = 1,356 gCO2e/MJ
    #             This is 15× higher than petroleum jet fuel (89 gCO2e/MJ)
    #   Boundary = Well-to-Tank (WtT) — explicitly stated Section 2.4
    #             (excludes combustion; WtWake would be even higher)
    #   H2 source= grey H2 from natural gas SMR — Section 2.1
    #   CF       = 91.3% (8,000 h/yr ÷ 8,760 h/yr, Table 2)
    #   Lifetime = 20 years (Table 2)
    #   DR       = not stated in paper
    #   Allocation = not stated in paper
    # REASON FOR EXCLUSION: Grey hydrogen from SMR makes this pathway
    # 15× worse than petroleum jet fuel — it does not qualify as SAF under
    # any regulatory definition (CORSIA, EU RED III, RFS). Including it
    # would produce nonsensical results in the harmonization framework.
]

# Keep old name as alias so existing call sites don't break
EXTERNAL_VALIDATION_STUDIES = VALIDATION_STUDIES


def decompose_variance(mc_results: Dict[str, pd.DataFrame],
                       sobol_results: Dict[str, dict]) -> pd.DataFrame:
    """
    Decompose total output variance into methodological and technical fractions.

    Uses Sobol S1 indices summed by parameter class. Estimates CV after
    harmonization by removing the methodological variance fraction.

    Returns
    -------
    pd.DataFrame with columns:
      Pathway, Metric, Methodological_pct, Technical_pct,
      CV_before_pct, CV_after_pct, Variance_Reduction_pct
    """
    rows = []
    for pathway in ["ATJ", "HEFA", "FT-SPK", "PtL"]:
        s1m = sobol_results[pathway]["S1_mfsp"]
        s1g = sobol_results[pathway]["S1_ghg"]
        mc  = mc_results[pathway]

        for metric, s1_dict, col in [
            ("MFSP", s1m, "mfsp"),
            ("GHG",  s1g, "ghg"),
        ]:
            total_s1 = sum(s1_dict.values())
            if total_s1 <= 0:
                total_s1 = 1.0

            meth_s1 = sum(v for k, v in s1_dict.items()
                          if k in METHODOLOGICAL_PARAMS)
            tech_s1 = sum(v for k, v in s1_dict.items()
                          if k not in METHODOLOGICAL_PARAMS)

            meth_pct = (meth_s1 / total_s1) * 100.0
            tech_pct = (tech_s1 / total_s1) * 100.0

            vals = mc[col].dropna()
            mean_val = vals.mean()
            cv_before = (vals.std() / mean_val * 100.0) if mean_val != 0 else 0.0

            # CV after harmonization: remove methodological variance fraction
            # Residual CV = CV_before * sqrt(1 - S1_methodological_fraction)
            meth_fraction = min(meth_s1 / max(total_s1, 1e-9), 1.0)
            cv_after = cv_before * np.sqrt(max(1.0 - meth_fraction, 0.0))

            var_reduction = (
                (1.0 - cv_after / cv_before) * 100.0 if cv_before > 0 else 0.0
            )

            rows.append({
                "Pathway":             pathway,
                "Metric":              metric,
                "Methodological_pct":  round(meth_pct,     1),
                "Technical_pct":       round(tech_pct,     1),
                "CV_before_pct":       round(cv_before,    1),
                "CV_after_pct":        round(cv_after,     1),
                "Variance_Reduction_pct": round(var_reduction, 1),
            })

    return pd.DataFrame(rows)


def run_external_validation(sobol_results: Dict[str, dict]) -> pd.DataFrame:
    """
    Apply the harmonization framework to real peer-reviewed external studies.

    Each study in VALIDATION_STUDIES has had its MFSP, GHG, allocation method,
    system boundary, discount rate, and capacity factor confirmed from the
    open-access full paper text.  Allocation and boundary are corrected to the
    reference energy-allocation / WtWake basis using the same five-step protocol
    as the main harmonization engine, allowing direct comparison with the
    harmonized literature database.

    Aggregate CV reduction (before vs. after harmonization) is reported only for
    pathways with ≥ 2 confirmed entries.

    Returns
    -------
    tuple of (per_study_df, aggregate_df)
    """
    from config import ALLOC_FACTORS, WTG_TO_WTWAKE_DELTA

    rows = []
    for study in EXTERNAL_VALIDATION_STUDIES:
        pathway  = study["pathway"]
        alloc_t  = ALLOC_FACTORS.get(pathway, {})
        study_af = alloc_t.get(study["allocation"], 1.0)
        ref_af   = alloc_t.get("energy", 1.0)

        # GHG harmonization
        ghg_harm = study["ghg_reported_gco2emj"] * (ref_af / study_af if study_af > 0 else 1.0)
        if study["boundary"] in ("WtG", "GtG"):
            ghg_harm += WTG_TO_WTWAKE_DELTA

        # MFSP: already in 2023 USD/GGE (studies self-report in 2023 basis)
        # Apply CRF correction using CAPEX_FRACTION
        from config import CAPEX_FRACTION
        from harmonization.engine import crf, CRF_REF
        crf_study = crf(study["discount_rate"] / 100.0, 30)
        cf_study  = study["capacity_factor"] / 100.0
        f_cap = CAPEX_FRACTION.get(pathway, 0.40)
        mfsp_harm = study["mfsp_reported_usd_gge"] * (
            (1 - f_cap) + f_cap * (CRF_REF / crf_study) * (cf_study / 0.90)
        )

        rows.append({
            "Study_ID":          study["study_id"],
            "Authors":           study.get("authors", ""),
            "Year":              study.get("year", ""),
            "DOI":               study.get("doi", ""),
            "Pathway":           pathway,
            "MFSP_reported":     study["mfsp_reported_usd_gge"],
            "MFSP_harmonized":   round(mfsp_harm, 2),
            "GHG_reported":      study["ghg_reported_gco2emj"],
            "GHG_harmonized":    round(ghg_harm,  2),
            "Allocation_used":   study["allocation"],
            "Boundary_used":     study["boundary"],
        })

    df = pd.DataFrame(rows)

    # Aggregate: CV before and after per pathway
    agg_rows = []
    for pathway, grp in df.groupby("Pathway"):
        for col_raw, col_harm, metric in [
            ("MFSP_reported", "MFSP_harmonized", "MFSP"),
            ("GHG_reported",  "GHG_harmonized",  "GHG"),
        ]:
            raw  = grp[col_raw].dropna()
            harm = grp[col_harm].dropna()
            if len(raw) < 2:
                continue
            cv_before = raw.std()  / raw.mean()  * 100 if raw.mean()  != 0 else 0
            cv_after  = harm.std() / harm.mean() * 100 if harm.mean() != 0 else 0
            agg_rows.append({
                "Pathway": pathway, "Metric": metric,
                "n_studies": len(raw),
                "CV_before_pct": round(cv_before, 1),
                "CV_after_pct":  round(cv_after,  1),
                "Reduction_pct": round((1 - cv_after / cv_before) * 100, 1)
                                  if cv_before > 0 else 0,
            })

    agg_df = pd.DataFrame(agg_rows)
    return df, agg_df
