"""
main.py
=======
SAF Harmonization Meta-Analysis Framework v2.0
Entry point — runs the full analysis pipeline end-to-end.

Pipeline
--------
1. Harmonize 48 literature studies (five-step protocol)
2. Run Monte Carlo simulation (10,000 iterations × 4 pathways)
3. Run Sobol sensitivity analysis (Jansen 1999 estimator)
4. Decompose variance (methodological vs. technical)
5. Run external validation (real peer-reviewed studies, open-access confirmed)
6. Generate all publication figures (16 figures, PNG + PDF)
7. Export comprehensive Excel workbook (9 sheets)

Usage
-----
    python main.py                    # full run
    python main.py --fast             # reduced MC/Sobol for quick testing
    python main.py --skip-figures     # skip figure generation
    python main.py --skip-excel       # skip Excel export

Output
------
    outputs/figures/  — 16 figures in PNG (300 DPI) and PDF (vector)
    outputs/SAF_MetaAnalysis_Harmonization.xlsx — full data workbook

Author : Joseph Amponsah, Iowa State University
Contact: joeampon@iastate.edu
"""

import argparse
import sys
import os
import time

# Ensure project root is on path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import numpy as np
import pandas as pd

from config import OUTPUT_DIR, FIGURES_DIR, EXCEL_PATH


# ---------------------------------------------------------------------------
# Pipeline steps
# ---------------------------------------------------------------------------

def step1_harmonize():
    print("\n[1/7] Harmonizing 48 literature studies...")
    from harmonization.engine import build_harmonized_dataset
    df = build_harmonized_dataset()
    print(f"      Studies     : {len(df)}")
    print(f"      Raw MFSP    : ${df['mfsp_2023_raw'].min():.2f} – "
          f"${df['mfsp_2023_raw'].max():.2f} / GGE (2023 USD)")
    print(f"      Harm. MFSP  : ${df['mfsp_harmonized'].min():.2f} – "
          f"${df['mfsp_harmonized'].max():.2f} / GGE")
    print(f"      Raw GHG     : {df['ghg_raw'].min():.1f} – "
          f"{df['ghg_raw'].max():.1f} gCO2e/MJ")
    print(f"      Harm. GHG   : {df['ghg_harmonized'].min():.1f} – "
          f"{df['ghg_harmonized'].max():.1f} gCO2e/MJ")
    print(f"      Mean GHG reduction vs petroleum: "
          f"{df['ghg_reduction_pct'].mean():.1f}%")
    return df


def step2_monte_carlo(n_iter: int):
    print(f"\n[2/7] Monte Carlo simulation ({n_iter:,} iterations × 4 pathways)...")
    t0 = time.time()
    from analysis.monte_carlo import run_monte_carlo, summarise_mc
    mc = run_monte_carlo(n_iter=n_iter, seed=42)
    summary = summarise_mc(mc)
    elapsed = time.time() - t0
    print(f"      Completed in {elapsed:.1f}s")
    for pw, df in mc.items():
        print(f"      {pw:8s}: MFSP median=${df['mfsp'].median():.2f}/GGE  "
              f"GHG median={df['ghg'].median():.1f} gCO2e/MJ  "
              f"n={len(df):,}")
    return mc, summary


def step3_sobol(n_sobol: int):
    print(f"\n[3/7] Sobol sensitivity analysis (N={n_sobol} base samples × 4 pathways)...")
    t0 = time.time()
    from analysis.sobol_analysis import run_sobol_analysis, sobol_summary_dataframe
    sobol = run_sobol_analysis(n_sobol=n_sobol, seed=42)
    sobol_df = sobol_summary_dataframe(sobol)
    elapsed = time.time() - t0
    print(f"      Completed in {elapsed:.1f}s")
    return sobol, sobol_df


def step4_variance(mc, sobol):
    print("\n[4/7] Variance decomposition (methodological vs. technical)...")
    from analysis.variance_decomposition import decompose_variance
    var_df = decompose_variance(mc, sobol)
    print(var_df[["Pathway","Metric","Methodological_pct","Technical_pct",
                  "CV_before_pct","CV_after_pct","Variance_Reduction_pct"]
                 ].to_string(index=False))
    return var_df


def step5_external_validation(sobol):
    print("\n[5/7] External validation (real open-access studies)...")
    from analysis.variance_decomposition import run_external_validation
    ext_df, ext_agg = run_external_validation(sobol)
    print("      Aggregate CV reduction by pathway:")
    print(ext_agg.to_string(index=False))
    return ext_df, ext_agg


def step6_figures(df_harm, mc, sobol, var_df, skip: bool):
    if skip:
        print("\n[6/7] Figure generation SKIPPED (--skip-figures)")
        return
    print(f"\n[6/7] Generating 16 publication figures -> {FIGURES_DIR}")
    from visualization.figures import generate_all_figures
    generate_all_figures(df_harm, mc, sobol, var_df)
    print(f"      All figures saved to {FIGURES_DIR}")


def step7_excel(df_harm, mc, sobol_df, var_df, ext_df, ext_agg,
                mc_summary, skip: bool):
    if skip:
        print("\n[7/7] Excel export SKIPPED (--skip-excel)")
        return
    print(f"\n[7/7] Exporting Excel workbook -> {EXCEL_PATH}")
    from data.generate_excel import create_workbook
    wb = create_workbook(df_harm, mc, sobol_df, var_df,
                         ext_df, ext_agg, mc_summary)
    wb.save(EXCEL_PATH)
    print(f"      Workbook saved: {EXCEL_PATH}")
    print(f"      Sheets: Literature Database | Harmonized Values | "
          f"Parameter Distributions | Monte Carlo Summary | Sobol Indices | "
          f"Variance Decomposition | External Validation | Key Findings | References")


# ---------------------------------------------------------------------------
# Summary printer
# ---------------------------------------------------------------------------

def print_summary(df_harm, mc, var_df):
    from config import PETROLEUM_JET_GHG_WTW
    print("\n" + "="*72)
    print("ANALYSIS SUMMARY")
    print("="*72)

    # Compliance counts
    print("\nGHG Compliance (harmonized values):")
    print(f"  {'Pathway':8s}  {'n':>4s}  {'<EU RED III (31.15)':>20s}  {'<50% cut (44.5)':>16s}")
    for pw, grp in df_harm.groupby("pathway"):
        n     = len(grp)
        ghg   = grp["ghg_harmonized"]
        n_red = (ghg <= 31.15).sum()
        n_50  = (ghg <= 44.5).sum()
        print(f"  {pw:8s}  {n:>4d}  {n_red:>4d}/{n} ({n_red/n*100:.0f}%)          "
              f"  {n_50:>4d}/{n} ({n_50/n*100:.0f}%)")

    print("\nMonte Carlo medians [P5 – P50 – P95]:")
    PATHWAY_ORDER = ["ATJ","HEFA","FT-SPK","PtL"]
    for pw in PATHWAY_ORDER:
        m = mc[pw]
        mp5, mp50, mp95 = np.percentile(m["mfsp"], [5,50,95])
        gp5, gp50, gp95 = np.percentile(m["ghg"],  [5,50,95])
        print(f"  {pw:8s}: MFSP ${mp5:.2f}–${mp50:.2f}–${mp95:.2f}/GGE  |  "
              f"GHG {gp5:.1f}–{gp50:.1f}–{gp95:.1f} gCO2e/MJ")

    print("\nVariance reduction from harmonization:")
    for _, row in var_df.iterrows():
        print(f"  {row.Pathway:8s} {row.Metric:5s}: "
              f"CV {row.CV_before_pct:.1f}% → {row.CV_after_pct:.1f}%  "
              f"(−{row.Variance_Reduction_pct:.0f}%)")

    print("\nOutput files:")
    print(f"  Figures : {FIGURES_DIR}")
    print(f"  Excel   : {EXCEL_PATH}")
    print("="*72 + "\n")


# ---------------------------------------------------------------------------
# Argument parsing and entry point
# ---------------------------------------------------------------------------

def parse_args():
    p = argparse.ArgumentParser(
        description="SAF Harmonization Meta-Analysis Framework v2.0",
        allow_abbrev=False,
    )
    p.add_argument("--fast", action="store_true",
                   help="Reduced iterations for quick testing "
                        "(MC=500, Sobol N=300)")
    p.add_argument("--skip-figures", action="store_true",
                   help="Skip figure generation")
    p.add_argument("--skip-excel", action="store_true",
                   help="Skip Excel export")
    # When running interactively (Jupyter / VS Code #%% cells), sys.argv
    # contains kernel arguments that argparse doesn't recognise → exit 2.
    # parse_known_args silently ignores them.
    args, _ = p.parse_known_args()
    return args


def main():
    args = parse_args()

    n_mc    = 500   if args.fast else 10_000
    n_sobol = 300   if args.fast else 1_500

    print("=" * 72)
    print("SAF HARMONIZATION META-ANALYSIS FRAMEWORK v2.0")
    print("Joseph Amponsah | Iowa State University | joeampon@iastate.edu")
    if args.fast:
        print("*** FAST MODE: reduced iterations for testing ***")
    print("=" * 72)

    t_start = time.time()

    df_harm           = step1_harmonize()
    mc, mc_summary    = step2_monte_carlo(n_mc)
    sobol, sobol_df   = step3_sobol(n_sobol)
    var_df            = step4_variance(mc, sobol)
    ext_df, ext_agg   = step5_external_validation(sobol)

    step6_figures(df_harm, mc, sobol, var_df, skip=args.skip_figures)
    step7_excel(df_harm, mc, sobol_df, var_df, ext_df, ext_agg,
                mc_summary, skip=args.skip_excel)

    print_summary(df_harm, mc, var_df)

    elapsed = time.time() - t_start
    print(f"Total runtime: {elapsed/60:.1f} min\n")

#%%
if __name__ == "__main__":
    main()
#%%