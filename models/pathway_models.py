"""
models/pathway_models.py
=========================
Techno-economic and LCA pathway models for ATJ, HEFA, FT-SPK, and PtL.

Each model takes a dict of parameter values and returns:
  {"mfsp": float (2023 $/GGE), "ghg": float (gCO2e/MJ)}

All models are implemented as pure functions with no side effects.
Plant configurations are fixed at representative design-basis values;
uncertainty is introduced through the Monte Carlo parameter sampling.

References
----------
ATJ    : Tao et al. (2017) Green Chemistry 10.1039/C6GC02800D
         Han et al. (2017) Biotechnol. Biofuels 10.1186/s13068-017-0698-z
         Davis et al. (2018) NREL/TP-5100-71949
HEFA   : Pearlson et al. (2013) Biofuels Bioprod. Biorefining 10.1002/bbb.1414
         NREL 2020 HBio baseline (NREL/TP-5100-75060)
FT-SPK : Swanson et al. (2010) NREL/TP-6A20-48440
         Davis et al. (2018) NREL/TP-5100-71949
PtL    : Schmidt et al. (2018) Joule 10.1016/j.joule.2018.05.008
         Fasihi et al. (2019) Joule 10.1016/j.joule.2019.05.002
"""

import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from config import (
    JET_LHV_BTU_PER_GAL, GASOLINE_LHV_BTU_PER_GAL, L_PER_GAL,
    JET_LHV_MJ_PER_L, JET_DENSITY_KG_PER_L, L_JET_PER_GGE,
    HARMONIZED_PLANT_LIFETIME, NG_COMBUSTION_GCO2E_MJ,
    GREY_H2_GCO2E_KG, H2_LHV_MJ_PER_KG, ELEC_KWH_PER_KG_H2, CO2_KG_PER_KG_H2,
)
from harmonization.engine import crf


# ---------------------------------------------------------------------------
# ATJ — Alcohol-to-Jet (corn stover design basis)
# ---------------------------------------------------------------------------

# Fixed plant configuration (Tao 2017 / Davis 2018 basis)
_ATJ_PLANT_TPD     = 2_000     # dry tonne feedstock per day
_ATJ_DIESEL_FRAC   = 0.15      # gal diesel co-product per gal ethanol
_ATJ_GASOLINE_FRAC = 0.14      # gal gasoline co-product per gal ethanol
_DIESEL_LHV_BTU    = 129_488   # BTU/gal diesel


def atj_model(params: dict) -> dict:
    """
    ATJ techno-economic and LCA model.

    Parameters (all from Monte Carlo sampling)
    ------------------------------------------
    feedstock_cost   $/tonne (2023 USD)
    capex_2023       $ (2023 USD, CEPCI-escalated)
    ethanol_yield    gal ethanol / tonne dry feedstock
    jet_yield        fraction ethanol -> jet (remainder -> diesel/gasoline)
    capacity_factor  0-1
    discount_rate    0-1
    ng_use           MJ natural gas / MJ jet fuel produced
    elec_use         kWh electricity / MJ jet fuel produced
    feedstock_ghg    gCO2e/MJ fuel (upstream feedstock)
    grid_intensity   gCO2e/kWh
    alloc_factor     fraction of burden allocated to jet (methodological)
    boundary_offset  gCO2e/MJ (WtG -> WtWake delta; methodological)
    iluc_penalty     gCO2e/MJ (methodological; 0 in baseline)
    """
    p = params

    # Annual mass flows
    annual_feed_t  = _ATJ_PLANT_TPD * 365 * p["capacity_factor"]
    annual_etoh_gal = annual_feed_t * p["ethanol_yield"]
    annual_jet_gal  = annual_etoh_gal * p["jet_yield"]
    annual_diesel_gal = annual_etoh_gal * _ATJ_DIESEL_FRAC
    annual_gas_gal  = annual_etoh_gal * _ATJ_GASOLINE_FRAC

    # Convert all products to GGE for total energy denominator
    jet_gge     = annual_jet_gal    * (JET_LHV_BTU_PER_GAL / GASOLINE_LHV_BTU_PER_GAL)
    diesel_gge  = annual_diesel_gal * (_DIESEL_LHV_BTU    / GASOLINE_LHV_BTU_PER_GAL)
    gas_gge     = annual_gas_gal    # gasoline = 1.0 GGE/gal by definition
    total_gge   = jet_gge + diesel_gge + gas_gge

    if total_gge <= 0:
        return {"mfsp": float("nan"), "ghg": float("nan")}

    # Annualised CAPEX and OPEX
    crf_val       = crf(p["discount_rate"], HARMONIZED_PLANT_LIFETIME)
    annual_capex  = p["capex_2023"] * crf_val
    fixed_opex    = p["capex_2023"] * 0.04    # 4 % of TCI (NREL convention)
    var_opex      = p["capex_2023"] * 0.03    # 3 % of TCI
    feed_cost     = annual_feed_t * p["feedstock_cost"]
    total_cost    = annual_capex + fixed_opex + var_opex + feed_cost

    mfsp = total_cost / total_gge   # 2023 $/GGE

    # LCA: energy balance GHG contributions
    ng_ghg   = p["ng_use"]  * NG_COMBUSTION_GCO2E_MJ
    elec_ghg = p["elec_use"] * p["grid_intensity"]
    raw_ghg  = (p["feedstock_ghg"] + ng_ghg + elec_ghg) * p["alloc_factor"]
    # ATJ has no CCS; biological feedstock cannot yield net-negative GHG in this model.
    # Floor of 0 prevents unphysical negative values from extreme parameter combinations.
    ghg      = max(raw_ghg + p["boundary_offset"] + p["iluc_penalty"], 0.0)

    return {"mfsp": mfsp, "ghg": ghg}


# ---------------------------------------------------------------------------
# HEFA — Hydroprocessed Esters and Fatty Acids
# ---------------------------------------------------------------------------

_HEFA_PLANT_TPD  = 800    # tonne oil per day (representative mid-size)


def hefa_model(params: dict) -> dict:
    """
    HEFA techno-economic and LCA model.

    Additional parameters vs ATJ
    ------------------------------
    h2_use      tonne H2 per tonne feedstock oil
    h2_price    $/kg H2 (2023)
    process_ghg gCO2e/MJ (facility-level process emissions)
    """
    p = params

    annual_feed_t    = _HEFA_PLANT_TPD * 365 * p["capacity_factor"]
    annual_jet_kg    = annual_feed_t * 1_000 * p["jet_yield"]   # kg jet
    annual_jet_L     = annual_jet_kg / JET_DENSITY_KG_PER_L
    annual_jet_MJ    = annual_jet_L * JET_LHV_MJ_PER_L
    annual_jet_GGE   = annual_jet_L / L_JET_PER_GGE

    if annual_jet_GGE <= 0:
        return {"mfsp": float("nan"), "ghg": float("nan")}

    crf_val      = crf(p["discount_rate"], HARMONIZED_PLANT_LIFETIME)
    annual_capex = p["capex_2023"] * crf_val
    opex         = p["capex_2023"] * 0.07    # HEFA lower O&M than gasification
    feed_cost    = annual_feed_t * p["feedstock_cost"]

    # H2 cost: h2_use is tonne H2 / tonne feed
    h2_t         = annual_feed_t * p["h2_use"]
    h2_cost      = h2_t * 1_000 * p["h2_price"]   # $/kg * 1000 kg/tonne

    total_cost   = annual_capex + opex + feed_cost + h2_cost
    mfsp         = total_cost / annual_jet_GGE

    # LCA: H2 production GHG + process GHG
    # Grey H2 emission factor allocated per MJ jet fuel
    h2_ghg_alloc = (h2_t * 1_000 * GREY_H2_GCO2E_KG) / annual_jet_MJ
    raw_ghg      = (p["feedstock_ghg"] + h2_ghg_alloc + p["process_ghg"]) * p["alloc_factor"]
    ghg          = max(raw_ghg + p["boundary_offset"] + p["iluc_penalty"], 0.0)

    return {"mfsp": mfsp, "ghg": ghg}


# ---------------------------------------------------------------------------
# FT-SPK — Fischer-Tropsch Synthetic Paraffinic Kerosene
# ---------------------------------------------------------------------------

_FTSPK_PLANT_TPD    = 2_500   # tonne biomass per day
_BIOMASS_LHV_MJ_T   = 17_500  # MJ/tonne dry biomass (lignocellulosic average)
_JET_FRAC_FT_LIQUIDS = 0.65   # fraction of FT liquids that is jet-range


def ftspk_model(params: dict) -> dict:
    """
    FT-SPK gasification + Fischer-Tropsch model.

    Additional parameters vs ATJ
    ------------------------------
    ft_efficiency  MJ total fuel out / MJ biomass in (LHV basis)
    """
    p = params

    annual_feed_t  = _FTSPK_PLANT_TPD * 365 * p["capacity_factor"]
    energy_in_MJ   = annual_feed_t * _BIOMASS_LHV_MJ_T
    total_fuel_MJ  = energy_in_MJ * p["ft_efficiency"]
    jet_MJ         = total_fuel_MJ * _JET_FRAC_FT_LIQUIDS
    jet_L          = jet_MJ / JET_LHV_MJ_PER_L
    jet_GGE        = jet_L / L_JET_PER_GGE

    if jet_GGE <= 0:
        return {"mfsp": float("nan"), "ghg": float("nan")}

    crf_val      = crf(p["discount_rate"], HARMONIZED_PLANT_LIFETIME)
    annual_capex = p["capex_2023"] * crf_val
    opex         = p["capex_2023"] * 0.05
    feed_cost    = annual_feed_t * p["feedstock_cost"]
    total_cost   = annual_capex + opex + feed_cost
    mfsp         = total_cost / jet_GGE

    raw_ghg = (p["feedstock_ghg"] + p["process_ghg"]) * p["alloc_factor"]
    # No floor: FT-SPK with integrated CCS can be genuinely carbon-negative.
    # Larson et al. (2009) DOI 10.1002/bbb.137 reports −8.5 gCO2e/MJ for
    # switchgrass BTL with CCS — biogenic CO2 is captured, not emitted.
    ghg     = raw_ghg + p["boundary_offset"] + p["iluc_penalty"]

    return {"mfsp": mfsp, "ghg": ghg}


# ---------------------------------------------------------------------------
# PtL — Power-to-Liquid (electrolysis + DAC + FT)
# ---------------------------------------------------------------------------

_PTL_PLANT_MW_ELEC = 200      # MW electrolyzer nameplate capacity
_JET_FRAC_PTL      = 0.70     # fraction of FT output that is jet range


def ptl_model(params: dict) -> dict:
    """
    PtL techno-economic and LCA model.

    Additional parameters vs other pathways
    -----------------------------------------
    elec_cost_mwh    $/MWh electricity (2023)
    elec_capex_kw    $/kW electrolyzer installed (2023)
    ft_capex         $ FT synthesis + upgrading capex (2023)
    co2_capture_cost $/tonne CO2 captured (DAC or point source)
    grid_intensity   gCO2e/kWh (renewable electricity lifecycle)
    """
    p = params

    annual_MWh     = _PTL_PLANT_MW_ELEC * 8_760 * p["capacity_factor"]
    annual_kg_H2   = annual_MWh * 1_000 / ELEC_KWH_PER_KG_H2
    annual_CO2_t   = annual_kg_H2 * CO2_KG_PER_KG_H2 / 1_000

    h2_energy_MJ   = annual_kg_H2 * H2_LHV_MJ_PER_KG
    ft_out_MJ      = h2_energy_MJ * p["ft_efficiency"]
    jet_MJ         = ft_out_MJ * _JET_FRAC_PTL
    jet_L          = jet_MJ / JET_LHV_MJ_PER_L
    jet_GGE        = jet_L / L_JET_PER_GGE

    if jet_GGE <= 0 or jet_MJ <= 0:
        return {"mfsp": float("nan"), "ghg": float("nan")}

    crf_val        = crf(p["discount_rate"], HARMONIZED_PLANT_LIFETIME)
    elec_capex_tot = p["elec_capex_kw"] * _PTL_PLANT_MW_ELEC * 1_000
    capex_ann      = (elec_capex_tot + p["ft_capex"]) * crf_val
    opex           = (elec_capex_tot + p["ft_capex"]) * 0.04
    elec_cost      = annual_MWh * p["elec_cost_mwh"]
    co2_cost       = annual_CO2_t * p["co2_capture_cost"]
    total_cost     = capex_ann + opex + elec_cost + co2_cost
    mfsp           = total_cost / jet_GGE

    # LCA: electricity dominates; per-MJ-jet basis
    elec_ghg_per_mj = (annual_MWh * 1_000 * p["grid_intensity"]) / jet_MJ
    ghg = max(elec_ghg_per_mj + p["boundary_offset"] + p["iluc_penalty"], 0.0)

    return {"mfsp": mfsp, "ghg": ghg}


# ---------------------------------------------------------------------------
# Registry
# ---------------------------------------------------------------------------
PATHWAY_MODELS = {
    "ATJ":    atj_model,
    "HEFA":   hefa_model,
    "FT-SPK": ftspk_model,
    "PtL":    ptl_model,
}
