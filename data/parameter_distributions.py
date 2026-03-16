"""
data/parameter_distributions.py
================================
Monte Carlo parameter distributions for all four SAF pathways.

Each entry is a tuple: (distribution_type, *args)
  triangular(low, mode, high)  — asymmetric; preferred for engineering parameters
  normal(mean, std)            — symmetric bell curve
  uniform(low, high)           — flat; used for methodological parameters

Parameter classification
------------------------
TECHNICAL     : feedstock cost, capex, yields, efficiency, grid intensity
METHODOLOGICAL: alloc_factor, boundary_offset, iluc_penalty

Sources for parameter ranges
-----------------------------
ATJ feedstock cost : Tao (2017), Han (2017); USDA ERS corn stover 2015-2023
ATJ capex          : Tao (2017) $420M 2011 -> $572M 2023; upper bound from
                     NREL 2022 design cases showing 1.4-1.8x uncertainty
HEFA feedstock     : USDA NASS oilseed prices 2015-2023; FAO commodity data
HEFA capex         : Pearlson (2013), Klein (2018), NREL HBio 2020
FT capex           : Swanson (2010) $500M-$1.5B range for 2000-4000 tpd
PtL electricity    : IEA World Energy Outlook 2023; IRENA Renewable Power 2023
PtL electrolyzer   : IEA 2023 Electrolyser costs $400-$1200/kW range
Grid intensity     : GREET 2023 regional US average; EPA eGRID 2023
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Parameters classified as methodological vs technical
METHODOLOGICAL_PARAMS = {"alloc_factor", "boundary_offset", "iluc_penalty"}

# =============================================================================
# ATJ — Alcohol-to-Jet
# =============================================================================
ATJ_PARAMS = {
    # --- TECHNICAL ---
    # Feedstock cost: corn stover 2023 USD/tonne
    # USDA ERS: $50-110/dry tonne; with logistics: up to $180/tonne
    "feedstock_cost":   ("triangular",  50.0,   110.0,   220.0),

    # Total capital investment 2023 USD
    # Tao 2017: $420M (2011) -> $572M (2023); NREL upper bound ~$1.4B for Nth
    "capex_2023":       ("triangular",  550e6,  875e6,  1_400e6),

    # Ethanol yield: gallons intermediate per dry tonne feedstock
    # Tao 2017: 79 gal/tonne; range from fermentation efficiency
    "ethanol_yield":    ("triangular",   65.0,   79.0,    90.0),

    # Jet fuel yield: fraction ethanol converted to jet
    # Tao 2017: 0.71; range reflects catalyst and separation performance
    "jet_yield":        ("triangular",    0.60,   0.71,    0.80),

    # Capacity factor: annual average utilisation
    "capacity_factor":  ("triangular",    0.75,   0.90,    0.95),

    # Discount rate: reflects project finance risk
    "discount_rate":    ("triangular",    0.05,   0.10,    0.15),

    # Natural gas use: MJ NG per MJ jet fuel produced
    # Tao 2017 energy balance
    "ng_use":           ("triangular",    0.08,   0.15,    0.25),

    # Electricity use: kWh per MJ jet fuel produced
    "elec_use":         ("triangular",    0.01,   0.02,    0.04),

    # Feedstock GHG intensity: gCO2e per MJ fuel (upstream emissions)
    # GREET 2023 corn stover range
    "feedstock_ghg":    ("triangular",    8.0,   15.0,    28.0),

    # Grid electricity carbon intensity: gCO2e/kWh
    # Range: low-carbon biogas (30) to coal-heavy grid (700)
    "grid_intensity":   ("triangular",   30.0,  200.0,   700.0),

    # --- METHODOLOGICAL ---
    # Allocation factor: fraction of total emissions assigned to jet fuel
    # Range covers energy (0.71) to mass (0.75) to system expansion (0.60)
    "alloc_factor":     ("uniform",       0.60,   0.75),

    # Boundary offset: WtG -> WtWake delta in gCO2e/MJ
    # ICAO CORSIA default 3.0; range 0 (already WtWake) to 5.0 (high distribution)
    "boundary_offset":  ("uniform",       0.0,    5.0),

    # ILUC penalty: baseline scenario excludes ILUC
    "iluc_penalty":     ("uniform",       0.0,    0.0),
}

# =============================================================================
# HEFA — Hydroprocessed Esters and Fatty Acids
# =============================================================================
HEFA_PARAMS = {
    # --- TECHNICAL ---
    # Feedstock cost: vegetable/waste oil 2023 USD/tonne
    # UCO: $300-500/tonne; soybean oil: $700-1100/tonne; microalgae: $1000+
    "feedstock_cost":   ("triangular",  300.0,  700.0,  1_400.0),

    # Total capital investment 2023 USD
    # Smaller plant than ATJ; Pearlson 2013: ~$150M, Klein 2018: ~$200M
    "capex_2023":       ("triangular",  120e6,  220e6,    380e6),

    # Jet fuel yield: kg jet per kg feedstock oil
    # Depends on feedstock quality; typical range 0.55-0.75
    "jet_yield":        ("triangular",    0.55,   0.65,    0.75),

    "capacity_factor":  ("triangular",    0.80,   0.90,    0.95),
    "discount_rate":    ("triangular",    0.05,   0.10,    0.15),

    # H2 consumption: tonne H2 per tonne feedstock
    # Hydroprocessing requires H2 for deoxygenation
    "h2_use":           ("triangular",    0.025,  0.040,   0.060),

    # H2 price: 2023 USD/kg
    # Grey H2: $1.5-2.5/kg; Blue: $2-4/kg; Green: $4-8/kg
    "h2_price":         ("triangular",    1.5,    3.0,     6.0),

    # Feedstock GHG: gCO2e/MJ (highly variable by feedstock type)
    # UCO: ~12 gCO2e/MJ; soybean: ~35; palm: ~35-60
    "feedstock_ghg":    ("triangular",   12.0,   32.0,    75.0),

    # Process GHG: conversion facility direct emissions gCO2e/MJ
    "process_ghg":      ("triangular",    1.0,    3.0,     8.0),

    # Grid intensity for process electricity
    "grid_intensity":   ("triangular",   50.0,  386.0,   700.0),

    # --- METHODOLOGICAL ---
    # HEFA energy allocation: 0.42 (sys expansion) to 0.57 (mass)
    "alloc_factor":     ("uniform",       0.42,   0.57),
    "boundary_offset":  ("uniform",       0.0,    5.0),
    "iluc_penalty":     ("uniform",       0.0,    0.0),
}

# =============================================================================
# FT-SPK — Fischer-Tropsch Synthetic Paraffinic Kerosene
# =============================================================================
FTSPK_PARAMS = {
    # --- TECHNICAL ---
    # Feedstock cost: lignocellulosic biomass 2023 USD/tonne
    # Forest residue: $25-50/tonne; energy crops: $50-100/tonne
    "feedstock_cost":   ("triangular",   25.0,   55.0,   100.0),

    # Total capital investment 2023 USD
    # Gasification + FT reactors capital intensive
    # Swanson 2010: ~$530M; range for uncertainty in reactor scale
    "capex_2023":       ("triangular",  500e6,  850e6,  1_500e6),

    # FT overall efficiency: MJ fuel output / MJ biomass input (LHV basis)
    # Includes gasification, cleaning, FT synthesis, and upgrading
    # Swanson 2010: ~0.38; range 0.30-0.52
    "ft_efficiency":    ("triangular",    0.30,   0.40,    0.52),

    "capacity_factor":  ("triangular",    0.80,   0.90,    0.95),
    "discount_rate":    ("triangular",    0.05,   0.10,    0.15),

    # Feedstock GHG: well-to-plant boundary, gCO2e/MJ fuel
    # Residues: 2-6 gCO2e/MJ; energy crops: 6-14
    "feedstock_ghg":    ("triangular",    2.0,    6.0,    14.0),

    # Process GHG: facility combustion/utilities, gCO2e/MJ
    "process_ghg":      ("triangular",    1.0,    3.0,     7.0),

    # --- METHODOLOGICAL ---
    # FT-SPK energy allocation: 0.55 (system expansion) to 0.68 (mass)
    "alloc_factor":     ("uniform",       0.55,   0.68),
    "boundary_offset":  ("uniform",       0.0,    5.0),
    "iluc_penalty":     ("uniform",       0.0,    0.0),
}

# =============================================================================
# PtL — Power-to-Liquid
# =============================================================================
PTL_PARAMS = {
    # --- TECHNICAL ---
    # Electricity cost: 2023 USD/MWh
    # Dedicated renewable: $20-50/MWh (wind/solar LCOE 2023)
    # Grid-backed: up to $110/MWh
    "elec_cost_mwh":    ("triangular",   20.0,   50.0,   110.0),

    # Electrolyzer CAPEX: 2023 USD/kW installed capacity
    # IEA 2023: PEM $500-$1000/kW; range for scale uncertainty
    "elec_capex_kw":    ("triangular",  400.0,  700.0,  1_200.0),

    # FT synthesis + upgrading CAPEX: 2023 USD total
    "ft_capex":         ("triangular",   80e6,  130e6,   220e6),

    # FT efficiency: MJ jet / MJ H2 input (includes FT + upgrading)
    "ft_efficiency":    ("triangular",    0.60,   0.72,    0.82),

    "capacity_factor":  ("triangular",    0.75,   0.85,    0.95),

    # PtL discount rate: often lower (public/EU funding)
    "discount_rate":    ("triangular",    0.05,   0.08,    0.12),

    # CO2 capture cost: 2023 USD/tonne CO2
    # DAC: $200-400/tonne (2023); industrial point source: $50-100/tonne
    "co2_capture_cost": ("triangular",   50.0,  130.0,   280.0),

    # Grid carbon intensity for dedicated renewable electricity
    # Solar LCOE LCA: 2-15 gCO2e/kWh; wind: 5-25 gCO2e/kWh (IPCC AR6)
    "grid_intensity":   ("triangular",    2.0,    8.0,    30.0),

    # --- METHODOLOGICAL ---
    # PtL has no co-products; allocation factor always 1.0
    "alloc_factor":     ("uniform",       1.00,   1.00),

    # WtG to WtWake: smaller delta for PtL (no feedstock transport)
    "boundary_offset":  ("uniform",       0.0,    3.0),
    "iluc_penalty":     ("uniform",       0.0,    0.0),
}

# =============================================================================
# Registry: maps pathway name to parameter dictionary
# =============================================================================
PATHWAY_PARAMS = {
    "ATJ":    ATJ_PARAMS,
    "HEFA":   HEFA_PARAMS,
    "FT-SPK": FTSPK_PARAMS,
    "PtL":    PTL_PARAMS,
}
