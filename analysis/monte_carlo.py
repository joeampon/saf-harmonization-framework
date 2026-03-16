"""
analysis/monte_carlo.py
=======================
Monte Carlo simulation for all four SAF pathways.

Runs 10,000 iterations per pathway, sampling from the parameter distributions
defined in data/parameter_distributions.py. Each iteration evaluates the
corresponding pathway model (models/pathway_models.py).

Returns a dict mapping pathway -> pd.DataFrame with columns:
  [all parameter names, "mfsp", "ghg"]

Seed is fixed at 42 for reproducibility.
"""

import numpy as np
import pandas as pd
from scipy import stats
from typing import Dict

import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def _draw_samples(dist_spec: tuple, n: int, rng: np.random.Generator) -> np.ndarray:
    """
    Draw n samples from a distribution specification.

    Spec format: (kind, *args)
      triangular(low, mode, high)
      normal(mean, std)
      uniform(low, high)
    """
    kind = dist_spec[0]

    if kind == "triangular":
        _, low, mode, high = dist_spec
        if high <= low:
            return np.full(n, mode)
        c = (mode - low) / (high - low)
        return stats.triang.rvs(c=c, loc=low, scale=high - low, size=n,
                                random_state=rng.integers(0, 2**31))

    elif kind == "normal":
        _, mean, std_val = dist_spec
        return rng.normal(loc=mean, scale=std_val, size=n)

    elif kind == "uniform":
        _, low, high = dist_spec
        if low == high:
            return np.full(n, low)
        return rng.uniform(low=low, high=high, size=n)

    else:
        raise ValueError(f"Unknown distribution kind: '{kind}'")


def run_monte_carlo(n_iter: int = 10_000,
                    seed: int = 42) -> Dict[str, pd.DataFrame]:
    """
    Run Monte Carlo simulation for ATJ, HEFA, FT-SPK, and PtL.

    Parameters
    ----------
    n_iter : int
        Number of Monte Carlo iterations per pathway (default 10,000)
    seed : int
        Random seed for reproducibility (default 42)

    Returns
    -------
    dict
        Keys: pathway names ("ATJ", "HEFA", "FT-SPK", "PtL")
        Values: pd.DataFrame with n_iter rows and columns for each
                parameter + "mfsp" + "ghg"
    """
    from data.parameter_distributions import PATHWAY_PARAMS
    from models.pathway_models import PATHWAY_MODELS

    rng = np.random.default_rng(seed)
    results: Dict[str, pd.DataFrame] = {}

    for pathway in ["ATJ", "HEFA", "FT-SPK", "PtL"]:
        param_defs = PATHWAY_PARAMS[pathway]
        model_fn   = PATHWAY_MODELS[pathway]

        # Draw all parameter samples in one pass
        samples: Dict[str, np.ndarray] = {}
        for name, spec in param_defs.items():
            samples[name] = _draw_samples(spec, n_iter, rng)

        mfsp_arr = np.full(n_iter, np.nan)
        ghg_arr  = np.full(n_iter, np.nan)

        for i in range(n_iter):
            p = {name: float(samples[name][i]) for name in samples}
            try:
                out = model_fn(p)
                mfsp_arr[i] = out["mfsp"]
                ghg_arr[i]  = out["ghg"]
            except Exception:
                pass   # leave as NaN

        df = pd.DataFrame(samples)
        df["mfsp"] = mfsp_arr
        df["ghg"]  = ghg_arr

        # Drop rows with NaN outputs (model failures)
        n_before = len(df)
        df = df.dropna(subset=["mfsp", "ghg"]).reset_index(drop=True)
        n_dropped = n_before - len(df)
        if n_dropped > 0:
            print(f"  [{pathway}] {n_dropped} iterations dropped (model errors)")

        # Clip extreme outliers (>5 IQR from median) to prevent plotting artefacts
        for col in ["mfsp", "ghg"]:
            q1, q3 = df[col].quantile([0.25, 0.75])
            iqr = q3 - q1
            df = df[
                (df[col] >= q1 - 5 * iqr) &
                (df[col] <= q3 + 5 * iqr)
            ]

        results[pathway] = df.reset_index(drop=True)

    return results


def summarise_mc(mc_results: Dict[str, pd.DataFrame]) -> pd.DataFrame:
    """
    Compute summary statistics from Monte Carlo results.

    Returns pd.DataFrame with columns:
      Pathway, Metric, P5, P25, Median, Mean, P75, P95, Std, CV_pct
    """
    rows = []
    for pathway, df in mc_results.items():
        for col, metric_label in [("mfsp", "MFSP ($/GGE)"),
                                   ("ghg",  "GHG (gCO2e/MJ)")]:
            v = df[col].dropna()
            rows.append({
                "Pathway":  pathway,
                "Metric":   metric_label,
                "P5":       round(np.percentile(v, 5),  2),
                "P25":      round(np.percentile(v, 25), 2),
                "Median":   round(np.percentile(v, 50), 2),
                "Mean":     round(v.mean(), 2),
                "P75":      round(np.percentile(v, 75), 2),
                "P95":      round(np.percentile(v, 95), 2),
                "Std":      round(v.std(),  2),
                "CV_pct":   round(v.std() / v.mean() * 100, 1) if v.mean() != 0 else 0,
            })
    return pd.DataFrame(rows)
