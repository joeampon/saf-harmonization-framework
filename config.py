"""
config.py
=========
Global constants, unit conversions, escalation tables, and reference values
for the SAF Harmonization Meta-Analysis Framework.

All monetary values are in 2023 USD unless stated otherwise.
All GHG values are in gCO2e/MJ on a Well-to-Wake (WtWake) basis.

References
----------
CEPCI        : Chemical Engineering Plant Cost Index, Chemical Engineering Magazine
CPI          : US Bureau of Labor Statistics CPI-U All Items (base: 2023=1.00)
EUR/USD      : US Federal Reserve H.10 release (annual averages)
Petroleum jet: ICAO CORSIA Doc 10164 (2022), 89 gCO2e/MJ WtWake
LHV values   : ASTM D4809 (jet fuel), DOE Energy Efficiency Handbook (gasoline)
"""

import os

# ---------------------------------------------------------------------------
# Unit conversions
# ---------------------------------------------------------------------------
GASOLINE_LHV_BTU_PER_GAL = 112_194        # BTU/gal  (DOE standard GGE basis)
JET_LHV_BTU_PER_GAL      = 120_200        # BTU/gal  (ASTM D4809 Jet-A)
JET_LHV_MJ_PER_L         = 34.37          # MJ/L
GASOLINE_LHV_MJ_PER_L    = 31.76          # MJ/L
L_PER_GAL                = 3.78541        # exact (US gallon)
JET_DENSITY_KG_PER_L     = 0.804          # kg/L Jet-A (ASTM D1655)

# L of jet fuel that equals 1 GGE (energy equivalent)
L_JET_PER_GGE = (GASOLINE_LHV_BTU_PER_GAL / JET_LHV_BTU_PER_GAL) * L_PER_GAL

# ---------------------------------------------------------------------------
# Harmonization reference basis
# ---------------------------------------------------------------------------
HARMONIZED_DISCOUNT_RATE   = 0.10      # 10 % real, Nth-plant (NREL convention)
HARMONIZED_CAPACITY_FACTOR = 0.90      # 90 %
HARMONIZED_PLANT_LIFETIME  = 30        # years
HARMONIZED_COST_YEAR       = 2023      # USD reference year
HARMONIZED_BOUNDARY        = "WtWake"
HARMONIZED_ALLOCATION      = "energy"  # ISO 14044 section 4.3.4.2

# ---------------------------------------------------------------------------
# Regulatory GHG baselines
# ---------------------------------------------------------------------------
PETROLEUM_JET_GHG_WTW    = 89.0   # gCO2e/MJ  (ICAO CORSIA Doc 10164, 2022)
CORSIA_TIER2_THRESHOLD   = 44.5   # gCO2e/MJ  50 % reduction vs petroleum
EU_RED3_THRESHOLD        = 31.15  # gCO2e/MJ  65 % reduction vs petroleum (EU RED III)
WTG_TO_WTWAKE_DELTA      = 3.0   # gCO2e/MJ  distribution phase (ICAO CORSIA default)

# ---------------------------------------------------------------------------
# Pathway-specific capital-cost fraction of MFSP
# Used in CRF normalisation step.
# MFSP_harm = MFSP * [(1 - f_cap) + f_cap * (CRF_ref / CRF_study) * (CF_study / CF_ref)]
# Values derived from modal-parameter cost breakdowns in each foundational study:
#   ATJ    : Capital 41 % — Tao et al. (2017) Green Chem. DOI 10.1039/C6GC02800D, 2000 tpd basis
#   HEFA   : Capital ~13 % — feedstock-dominated; Pearlson et al. (2013) DOI 10.1002/bbb.1378
#   FT-SPK : Capital ~55 % — gasification + FT reactors; Swanson et al. (2010) NREL/TP-6A20-48440
#   PtL    : Capital ~19 % — electricity cost dominant; Schmidt et al. (2018) Chem. Ing. Tech. DOI 10.1002/cite.201700129
# ---------------------------------------------------------------------------
CAPEX_FRACTION = {
    "ATJ":    0.41,
    "HEFA":   0.13,
    "FT-SPK": 0.55,
    "PtL":    0.19,
}

# ---------------------------------------------------------------------------
# CEPCI — Chemical Engineering Plant Cost Index
# Source: Chemical Engineering Magazine (2005-2022);
#         2023 estimated from Intratec PCI (widely used in biorefinery TEA)
# ---------------------------------------------------------------------------
CEPCI = {
    2005: 468.2, 2006: 499.6, 2007: 525.4, 2008: 575.4, 2009: 521.9,
    2010: 550.8, 2011: 585.7, 2012: 584.6, 2013: 567.3, 2014: 576.1,
    2015: 556.8, 2016: 541.7, 2017: 567.5, 2018: 603.1, 2019: 607.5,
    2020: 596.2, 2021: 708.8, 2022: 816.0, 2023: 798.0,
}

# ---------------------------------------------------------------------------
# CPI escalation factors to 2023 USD
# Source: US BLS CPI-U All Items; 2023 = 1.00 base
# ---------------------------------------------------------------------------
CPI_TO_2023 = {
    2005: 1.58, 2006: 1.54, 2007: 1.49, 2008: 1.44, 2009: 1.44,
    2010: 1.42, 2011: 1.37, 2012: 1.34, 2013: 1.32, 2014: 1.30,
    2015: 1.29, 2016: 1.27, 2017: 1.24, 2018: 1.21, 2019: 1.18,
    2020: 1.17, 2021: 1.11, 2022: 1.04, 2023: 1.00,
}

# ---------------------------------------------------------------------------
# EUR/USD historical exchange rates (annual averages)
# Source: US Federal Reserve H.10 statistical release
# ---------------------------------------------------------------------------
EUR_USD = {
    2005: 1.245, 2006: 1.256, 2007: 1.370, 2008: 1.471, 2009: 1.394,
    2010: 1.326, 2011: 1.392, 2012: 1.285, 2013: 1.328, 2014: 1.329,
    2015: 1.110, 2016: 1.107, 2017: 1.130, 2018: 1.181, 2019: 1.119,
    2020: 1.142, 2021: 1.183, 2022: 1.053, 2023: 1.081,
}

# ---------------------------------------------------------------------------
# Co-product allocation factors (fraction of total emissions assigned to jet)
# Based on jet fuel energy content relative to all products (LHV basis).
# Sources: ISO 14044 section 4.3.4.2; Han et al. (2017) Biotechnol. Biofuels;
#          Stratton et al. (2010) MIT LAE; Davis et al. (2018) NREL/TP-5100-71949
# ---------------------------------------------------------------------------
ALLOC_FACTORS = {
    "ATJ": {
        "energy":           0.71,
        "mass":             0.75,
        "economic":         0.69,
        "system_expansion": 0.60,
    },
    "HEFA": {
        "energy":           0.52,
        "mass":             0.57,
        "economic":         0.48,
        "system_expansion": 0.42,
    },
    "FT-SPK": {
        "energy":           0.66,
        "mass":             0.68,
        "economic":         0.63,
        "system_expansion": 0.55,
    },
    "PtL": {
        "energy":           1.00,
        "mass":             1.00,
        "economic":         1.00,
        "system_expansion": 1.00,
    },
}

# ---------------------------------------------------------------------------
# ILUC emission estimates (gCO2e/MJ fuel)
# Used to REMOVE ILUC from studies that included it.
# Sources: Searchinger et al. (2008) Science; ICAO CORSIA ILUC values (2022);
#          EU RED III Annex IX; IPCC AR6 WG3 Table 7.4 (2022)
# Zero ILUC for waste/residue feedstocks — no land displacement occurs.
# ---------------------------------------------------------------------------
ILUC_GCO2E_MJ = {
    "Corn stover":                0.0,
    "Wheat straw":                0.0,
    "Sugarcane bagasse":          0.0,
    "Municipal solid waste":      0.0,
    "Tallow":                     0.0,
    "Used cooking oil":           0.0,
    "Waste cooking oil":          0.0,
    "Forestry residue":           0.0,
    "Woody biomass":              0.0,
    "Straw":                      0.0,
    "Agricultural residues":      0.0,
    "Poplar":                     2.0,
    "Eucalyptus":                 2.0,
    "Miscanthus":                 3.0,
    "Switchgrass":                4.0,
    "Sugarcane":                  5.0,
    "Camelina oil":               3.0,
    "Carinata oil":               3.0,
    "Jatropha oil":               8.0,
    "Soybean oil":               15.0,
    "Palm oil":                  25.0,
    "Microalgae oil":             0.0,
    "Renewable electricity + CO2": 0.0,
    "Wind electricity + DAC CO2":  0.0,
    "Solar electricity + DAC CO2": 0.0,
    "Solar + DAC CO2":             0.0,
}

# ---------------------------------------------------------------------------
# Emission factors used in pathway models
# ---------------------------------------------------------------------------
NG_COMBUSTION_GCO2E_MJ  = 56.1    # gCO2e/MJ NG combustion (LHV) — GREET 2023
GREY_H2_GCO2E_KG        = 9_000   # gCO2e/kg H2, SMR without CCS — IEA 2021
US_GRID_GCO2E_KWH       = 386.0   # gCO2e/kWh US avg 2023 — EPA eGRID 2023
H2_LHV_MJ_PER_KG        = 120.0   # MJ/kg H2 — NIST
ELEC_KWH_PER_KG_H2      = 55.0    # kWh/kg H2 PEM electrolyzer — IEA 2023
CO2_KG_PER_KG_H2        = 5.5     # kg CO2 per kg H2 (DAC + FT stoichiometry)

# ---------------------------------------------------------------------------
# Output paths
# ---------------------------------------------------------------------------
PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR   = os.path.join(PROJECT_ROOT, "outputs")
FIGURES_DIR  = os.path.join(OUTPUT_DIR, "figures")
EXCEL_PATH   = os.path.join(OUTPUT_DIR, "SAF_MetaAnalysis_Harmonization.xlsx")

os.makedirs(FIGURES_DIR, exist_ok=True)
