"""
analysis/sobol_analysis.py
==========================
Sobol variance-based global sensitivity analysis using the
Jansen (1999) first-order and total-order estimators.

Method
------
For each pathway and each parameter j, we construct a matrix C_j
identical to A except column j is replaced by column B.

First-order Sobol index (S1) — Jansen estimator:
  S1_j = 1 - E[(Y_B - Y_Cj)^2] / (2 * Var(Y))

Total-order Sobol index (ST) — Jansen estimator:
  ST_j = E[(Y_A - Y_Cj)^2] / (2 * Var(Y))

Reliability notes
-----------------
Indices S1 < 0.05 have significant sampling uncertainty at N=1500.
These are flagged as "negligible" in the output and should not be
interpreted as precise values.

Reference
---------
Jansen (1999) Analysis of variance designs for model output.
Computer Physics Communications 117(1-2), 35-43.
DOI: 10.1016/S0010-4655(98)00154-4
"""

import numpy as np
from typing import Dict, Tuple

import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

S1_NEGLIGIBLE_THRESHOLD = 0.05   # S1 below this has high sampling uncertainty


def _draw_matrix(param_defs: dict, N: int, rng: np.random.Generator) -> Dict[str, np.ndarray]:
    """Draw N samples for all parameters."""
    from analysis.monte_carlo import _draw_samples
    return {name: _draw_samples(spec, N, rng) for name, spec in param_defs.items()}


def _eval_matrix(sample_dict: dict, param_names: list,
                 model_fn, N: int) -> np.ndarray:
    """
    Evaluate model at N parameter sets from sample_dict.

    Returns
    -------
    np.ndarray, shape (N, 2)  columns: [mfsp, ghg]
    """
    out = np.full((N, 2), np.nan)
    for i in range(N):
        p = {nm: float(sample_dict[nm][i]) for nm in param_names}
        try:
            result = model_fn(p)
            out[i, 0] = result["mfsp"]
            out[i, 1] = result["ghg"]
        except Exception:
            pass
    return out


def jansen_sobol(model_fn, param_defs: dict,
                 N: int = 1_500,
                 seed: int = 42) -> Dict[str, Dict[str, float]]:
    """
    Compute first-order (S1) and total-order (ST) Sobol indices
    using the Jansen (1999) estimator.

    Parameters
    ----------
    model_fn    : callable — pathway model function
    param_defs  : dict — parameter distribution specifications
    N           : int — base sample size (total evaluations: N*(k+2))
    seed        : int — random seed

    Returns
    -------
    dict with keys "S1_mfsp", "S1_ghg", "ST_mfsp", "ST_ghg"
    Each value is a dict mapping parameter name -> float index
    """
    rng = np.random.default_rng(seed)
    names = list(param_defs.keys())
    k = len(names)

    # Draw two independent sample matrices A and B
    A = _draw_matrix(param_defs, N, rng)
    B = _draw_matrix(param_defs, N, rng)

    # Evaluate models at A and B
    Y_A = _eval_matrix(A, names, model_fn, N)
    Y_B = _eval_matrix(B, names, model_fn, N)

    # Compute total variance from combined A+B output
    Y_all = np.vstack([Y_A, Y_B])
    var_mfsp = np.nanvar(Y_all[:, 0])
    var_ghg  = np.nanvar(Y_all[:, 1])

    S1_mfsp, S1_ghg = {}, {}
    ST_mfsp, ST_ghg = {}, {}

    for j, name in enumerate(names):
        # Construct C_j: A with column j replaced by B
        C_j = {nm: A[nm].copy() for nm in names}
        C_j[name] = B[name].copy()

        Y_Cj = _eval_matrix(C_j, names, model_fn, N)

        # Jansen S1 estimator
        def _s1(Y_b, Y_c, var):
            diff_sq = np.nanmean((Y_b - Y_c) ** 2)
            return max(1.0 - diff_sq / (2.0 * var), 0.0) if var > 0 else 0.0

        def _st(Y_a, Y_c, var):
            diff_sq = np.nanmean((Y_a - Y_c) ** 2)
            return max(diff_sq / (2.0 * var), 0.0) if var > 0 else 0.0

        S1_mfsp[name] = _s1(Y_B[:, 0], Y_Cj[:, 0], var_mfsp)
        S1_ghg[name]  = _s1(Y_B[:, 1], Y_Cj[:, 1], var_ghg)
        ST_mfsp[name] = _st(Y_A[:, 0], Y_Cj[:, 0], var_mfsp)
        ST_ghg[name]  = _st(Y_A[:, 1], Y_Cj[:, 1], var_ghg)

    return {
        "S1_mfsp": S1_mfsp,
        "S1_ghg":  S1_ghg,
        "ST_mfsp": ST_mfsp,
        "ST_ghg":  ST_ghg,
        "var_mfsp": var_mfsp,
        "var_ghg":  var_ghg,
    }


def run_sobol_analysis(n_sobol: int = 1_500,
                       seed: int = 42) -> Dict[str, dict]:
    """
    Run Sobol analysis for all four pathways.

    Parameters
    ----------
    n_sobol : int  — base sample size per pathway (default 1500)
    seed    : int  — random seed (default 42)

    Returns
    -------
    dict  {pathway: sobol_result_dict}
    """
    from data.parameter_distributions import PATHWAY_PARAMS
    from models.pathway_models import PATHWAY_MODELS

    results = {}
    for pathway in ["ATJ", "HEFA", "FT-SPK", "PtL"]:
        print(f"  Sobol [{pathway}] N={n_sobol} base samples ...", end=" ", flush=True)
        results[pathway] = jansen_sobol(
            model_fn    = PATHWAY_MODELS[pathway],
            param_defs  = PATHWAY_PARAMS[pathway],
            N           = n_sobol,
            seed        = seed + hash(pathway) % 1000,
        )
        # Flag negligible indices
        for metric in ("S1_mfsp", "S1_ghg"):
            for param, val in results[pathway][metric].items():
                if 0 < val < S1_NEGLIGIBLE_THRESHOLD:
                    results[pathway][metric][param] = round(val, 4)
        print("done")

    return results


def sobol_summary_dataframe(sobol_results: Dict[str, dict]) -> "pd.DataFrame":
    """Convert Sobol results to a tidy DataFrame for export."""
    import pandas as pd
    from data.parameter_distributions import METHODOLOGICAL_PARAMS

    rows = []
    for pathway, res in sobol_results.items():
        for param in res["S1_mfsp"]:
            rows.append({
                "Pathway":      pathway,
                "Parameter":    param,
                "Type":         "Methodological" if param in METHODOLOGICAL_PARAMS
                                else "Technical",
                "S1_MFSP":     round(res["S1_mfsp"][param], 4),
                "S1_GHG":      round(res["S1_ghg"][param],  4),
                "ST_MFSP":     round(res["ST_mfsp"][param], 4),
                "ST_GHG":      round(res["ST_ghg"][param],  4),
                "Negligible_MFSP": res["S1_mfsp"][param] < S1_NEGLIGIBLE_THRESHOLD,
                "Negligible_GHG":  res["S1_ghg"][param]  < S1_NEGLIGIBLE_THRESHOLD,
            })
    return pd.DataFrame(rows)
