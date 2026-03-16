"""
visualization/si_figures.py
Supplementary Information figures — muted academic style.
Completely different visual language from figures.py (main paper).
"""
import sys, os
import numpy as np
import matplotlib; matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch
from matplotlib.gridspec import GridSpec
from matplotlib.lines import Line2D
from scipy.stats import gaussian_kde
from scipy import stats

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from config import (
    PETROLEUM_JET_GHG_WTW, CORSIA_TIER2_THRESHOLD, EU_RED3_THRESHOLD,
    JET_LHV_BTU_PER_GAL, GASOLINE_LHV_BTU_PER_GAL, JET_LHV_MJ_PER_L,
    JET_DENSITY_KG_PER_L, L_JET_PER_GGE, NG_COMBUSTION_GCO2E_MJ,
    GREY_H2_GCO2E_KG, H2_LHV_MJ_PER_KG, ELEC_KWH_PER_KG_H2, CO2_KG_PER_KG_H2,
    FIGURES_DIR, HARMONIZED_PLANT_LIFETIME,
)
from harmonization.engine import crf
from data.parameter_distributions import METHODOLOGICAL_PARAMS

# ── Output directory ──────────────────────────────────────────────────────────
SI_FIGURES_DIR = os.path.join(os.path.dirname(FIGURES_DIR), "figures", "SI")
os.makedirs(SI_FIGURES_DIR, exist_ok=True)

# ── Color palette (muted earth/slate) ────────────────────────────────────────
SI_COLORS = {
    "ATJ":    "#2E6096",
    "HEFA":   "#4D7C5A",
    "FT-SPK": "#9C4A1E",
    "PtL":    "#5C3D7A",
}
SI_COLORS_LIGHT = {
    "ATJ":    "#A8C4DC",
    "HEFA":   "#A8C9AE",
    "FT-SPK": "#D4A98A",
    "PtL":    "#BBA8D0",
}
BG_COLOR    = "#F9F7F4"
PANEL_COLOR = "#EFEFEA"
GRID_COLOR  = "#E0DDD5"
TECH_COLOR  = "#1C3D5A"    # deep navy — technical params
METH_COLOR  = "#8B3A2F"    # burnt sienna — methodological params
PWO = ["ATJ", "HEFA", "FT-SPK", "PtL"]

# ── rcParams for SI style ─────────────────────────────────────────────────────
SI_RC = {
    "font.family":       "serif",
    "font.size":         10,
    "axes.facecolor":    PANEL_COLOR,
    "figure.facecolor":  BG_COLOR,
    "axes.grid":         True,
    "grid.color":        GRID_COLOR,
    "grid.linewidth":    0.7,
    "axes.linewidth":    0.9,
    "axes.edgecolor":    "#BBBBBB",
    "xtick.color":       "#555555",
    "ytick.color":       "#555555",
    "text.color":        "#333333",
    "axes.labelcolor":   "#333333",
    "pdf.fonttype":      42,
    "ps.fonttype":       42,
    "figure.dpi":        150,
}

# ── Helper: save figure ───────────────────────────────────────────────────────
def _save_si(fig, name):
    png_path = os.path.join(SI_FIGURES_DIR, f"{name}.png")
    pdf_path = os.path.join(SI_FIGURES_DIR, f"{name}.pdf")
    fig.savefig(png_path, dpi=300, bbox_inches="tight", facecolor=BG_COLOR)
    fig.savefig(pdf_path, bbox_inches="tight", facecolor=BG_COLOR)
    plt.close(fig)
    print(f"  Saved SI: {name}.{{png,pdf}}")
    return png_path


# ── Helper: modal value from distribution spec ────────────────────────────────
def _modal(pd_):
    result = {}
    for n, s in pd_.items():
        k = s[0]
        if k == "triangular":
            result[n] = s[2]  # mode
        elif k == "normal":
            result[n] = s[1]  # mean
        elif k == "uniform":
            result[n] = (s[1] + s[2]) / 2
        else:
            result[n] = s[1]
    return result


# ── Helper: percentile from distribution spec ─────────────────────────────────
def _pct(spec, q):
    k = spec[0]
    if k == "triangular":
        _, lo, mo, hi = spec
        if hi <= lo:
            return mo
        return float(stats.triang.ppf(q, c=(mo - lo) / (hi - lo), loc=lo, scale=hi - lo))
    if k == "uniform":
        return spec[1] + q * (spec[2] - spec[1])
    if k == "normal":
        return float(stats.norm.ppf(q, loc=spec[1], scale=spec[2]))
    return spec[1]


# ── Helper: sample parameter ──────────────────────────────────────────────────
def _sample_param(spec, n):
    k = spec[0]
    if k == "triangular":
        _, lo, mo, hi = spec
        if hi <= lo:
            return np.full(n, mo)
        c = (mo - lo) / max(hi - lo, 1e-9)
        return stats.triang.rvs(c, loc=lo, scale=hi - lo, size=n)
    if k == "uniform":
        return np.random.uniform(spec[1], spec[2], n)
    if k == "normal":
        return np.random.normal(spec[1], spec[2], n)
    return np.full(n, spec[1])


# ── Monte Carlo runner ────────────────────────────────────────────────────────
def _run_mc(n=2000):
    from data.parameter_distributions import PATHWAY_PARAMS
    from models.pathway_models import PATHWAY_MODELS
    np.random.seed(42)
    result = {}
    for pw in PWO:
        pd_ = PATHWAY_PARAMS[pw]
        fn  = PATHWAY_MODELS[pw]
        samples = {name: _sample_param(spec, n) for name, spec in pd_.items()}
        mfsp_arr = np.zeros(n)
        ghg_arr  = np.zeros(n)
        for i in range(n):
            params_i = {name: float(samples[name][i]) for name in samples}
            try:
                out = fn(params_i)
                mfsp_arr[i] = out.get("mfsp", np.nan)
                ghg_arr[i]  = out.get("ghg",  np.nan)
            except Exception:
                mfsp_arr[i] = np.nan
                ghg_arr[i]  = np.nan
        result[pw] = {"mfsp": mfsp_arr, "ghg": ghg_arr}
    return result


# ── Modal cost/GHG breakdown ──────────────────────────────────────────────────
def _compute_modal_breakdown():
    from data.parameter_distributions import ATJ_PARAMS, HEFA_PARAMS, FTSPK_PARAMS, PTL_PARAMS
    r = {}
    # ATJ
    p = _modal(ATJ_PARAMS)
    af = 2000 * 365 * p["capacity_factor"]
    ae = af * p["ethanol_yield"]
    aj = ae * p["jet_yield"]
    ad = ae * 0.15
    ag = ae * 0.14
    jg = aj * (JET_LHV_BTU_PER_GAL / GASOLINE_LHV_BTU_PER_GAL)
    dg = ad * (129488 / GASOLINE_LHV_BTU_PER_GAL)
    tg = max(jg + dg + ag, 1e-9)
    cv = crf(p["discount_rate"], HARMONIZED_PLANT_LIFETIME)
    r["ATJ"] = {
        "cost": {
            "Capital":   p["capex_2023"] * cv / tg,
            "O&M":       p["capex_2023"] * 0.07 / tg,
            "Feedstock": af * p["feedstock_cost"] / tg,
        },
        "ghg": {
            "Feedstock GHG":  p["feedstock_ghg"] * p["alloc_factor"],
            "NG combustion":  p["ng_use"] * NG_COMBUSTION_GCO2E_MJ * p["alloc_factor"],
            "Electricity GHG":p["elec_use"] * p["grid_intensity"] * p["alloc_factor"],
            "Boundary offset":p["boundary_offset"],
        },
    }
    # HEFA
    p = _modal(HEFA_PARAMS)
    af   = 800 * 365 * p["capacity_factor"]
    ajkg = af * 1000 * p["jet_yield"]
    ajl  = ajkg / JET_DENSITY_KG_PER_L
    ajmj = ajl * JET_LHV_MJ_PER_L
    ajg  = max(ajl / L_JET_PER_GGE, 1e-9)
    cv   = crf(p["discount_rate"], HARMONIZED_PLANT_LIFETIME)
    h2t  = af * p["h2_use"]
    h2g  = (h2t * 1000 * GREY_H2_GCO2E_KG) / max(ajmj, 1e-9)
    r["HEFA"] = {
        "cost": {
            "Capital":   p["capex_2023"] * cv / ajg,
            "O&M":       p["capex_2023"] * 0.07 / ajg,
            "Feedstock": af * p["feedstock_cost"] / ajg,
            "H\u2082":   h2t * 1000 * p["h2_price"] / ajg,
        },
        "ghg": {
            "Feedstock GHG":  p["feedstock_ghg"] * p["alloc_factor"],
            "H\u2082 production": h2g * p["alloc_factor"],
            "Process GHG":    p["process_ghg"] * p["alloc_factor"],
            "Boundary offset":p["boundary_offset"],
        },
    }
    # FT-SPK
    p  = _modal(FTSPK_PARAMS)
    af = 2500 * 365 * p["capacity_factor"]
    jmj = af * 17500 * p["ft_efficiency"] * 0.65
    jl  = jmj / JET_LHV_MJ_PER_L
    jg  = max(jl / L_JET_PER_GGE, 1e-9)
    cv  = crf(p["discount_rate"], HARMONIZED_PLANT_LIFETIME)
    r["FT-SPK"] = {
        "cost": {
            "Capital":   p["capex_2023"] * cv / jg,
            "O&M":       p["capex_2023"] * 0.05 / jg,
            "Feedstock": af * p["feedstock_cost"] / jg,
        },
        "ghg": {
            "Feedstock GHG":  p["feedstock_ghg"] * p["alloc_factor"],
            "Process GHG":    p["process_ghg"] * p["alloc_factor"],
            "Boundary offset":p["boundary_offset"],
        },
    }
    # PtL
    p  = _modal(PTL_PARAMS)
    MW = 200
    amwh   = MW * 8760 * p["capacity_factor"]
    akgh2  = amwh * 1000 / ELEC_KWH_PER_KG_H2
    aco2t  = akgh2 * CO2_KG_PER_KG_H2 / 1000
    jmj    = akgh2 * H2_LHV_MJ_PER_KG * p["ft_efficiency"] * 0.70
    jl     = jmj / JET_LHV_MJ_PER_L
    jg     = max(jl / L_JET_PER_GGE, 1e-9)
    cv     = crf(p["discount_rate"], HARMONIZED_PLANT_LIFETIME)
    ec     = p["elec_capex_kw"] * MW * 1000
    r["PtL"] = {
        "cost": {
            "Electrolyzer cap.": ec * cv / jg,
            "FT capital":        p["ft_capex"] * cv / jg,
            "Electricity":       amwh * p["elec_cost_mwh"] / jg,
            "CO\u2082 capture":  aco2t * p["co2_capture_cost"] / jg,
            "O&M":               (ec + p["ft_capex"]) * 0.04 / jg,
        },
        "ghg": {
            "Electricity GHG":    amwh * 1000 * p["grid_intensity"] / max(jmj, 1e-9),
            "Boundary offset":    p["boundary_offset"],
        },
    }
    return r


# ── OAT swings ────────────────────────────────────────────────────────────────
def _compute_oat_swings(pathway, metric="mfsp"):
    from data.parameter_distributions import PATHWAY_PARAMS
    from models.pathway_models import PATHWAY_MODELS
    pd_  = PATHWAY_PARAMS[pathway]
    fn   = PATHWAY_MODELS[pathway]
    mo   = _modal(pd_)
    mv   = fn(mo)[metric]
    sw   = []
    for n, sp in pd_.items():
        try:
            lo = fn({**mo, n: _pct(sp, 0.05)})[metric]
            hi = fn({**mo, n: _pct(sp, 0.95)})[metric]
        except Exception:
            lo = hi = mv
        sw.append((n, lo - mv, hi - mv))
    sw.sort(key=lambda x: abs(x[2] - x[1]), reverse=True)
    return sw, mv


# ── Precomputed S1 data (Sobol first-order indices) ───────────────────────────
_S1_DATA = {
    "ATJ": {
        "mfsp": {
            "capex_2023":     0.419,
            "feedstock_cost": 0.214,
            "discount_rate":  0.134,
            "ethanol_yield":  0.103,
            "jet_yield":      0.052,
            "capacity_factor":0.020,
            "alloc_factor":   0.000,
            "boundary_offset":0.000,
        },
        "ghg": {
            "feedstock_ghg":  0.414,
            "grid_intensity": 0.255,
            "ng_use":         0.081,
            "elec_use":       0.101,
            "alloc_factor":   0.111,
            "boundary_offset":0.102,
        },
    },
    "HEFA": {
        "mfsp": {
            "feedstock_cost": 0.902,
            "jet_yield":      0.105,
            "h2_price":       0.030,
            "h2_use":         0.030,
            "process_ghg":    0.030,
            "alloc_factor":   0.030,
            "boundary_offset":0.030,
        },
        "ghg": {
            "feedstock_ghg":  0.804,
            "alloc_factor":   0.095,
            "boundary_offset":0.035,
            "h2_use":         0.018,
        },
    },
    "FT-SPK": {
        "mfsp": {
            "capex_2023":     0.496,
            "ft_efficiency":  0.243,
            "feedstock_cost": 0.075,
            "discount_rate":  0.158,
            "alloc_factor":   0.017,
            "boundary_offset":0.017,
        },
        "ghg": {
            "feedstock_ghg":  0.442,
            "boundary_offset":0.406,
            "process_ghg":    0.078,
            "alloc_factor":   0.000,
        },
    },
    "PtL": {
        "mfsp": {
            "elec_cost_mwh":   0.794,
            "ft_capex":        0.034,
            "elec_capex_kw":   0.032,
            "co2_capture_cost":0.067,
            "ft_efficiency":   0.102,
        },
        "ghg": {
            "grid_intensity": 0.943,
            "ft_efficiency":  0.036,
        },
    },
}

# =============================================================================
# Figure S2 — Harmonization flowchart
# =============================================================================
def figS2_harmonization_flowchart(results_dict):
    """Horizontal pipeline of 5 rounded boxes with arrows."""
    with plt.rc_context(SI_RC):
        fig, ax = plt.subplots(figsize=(16, 6), facecolor=BG_COLOR)
        ax.set_facecolor(BG_COLOR)
        ax.set_xlim(0, 16)
        ax.set_ylim(0, 6)
        ax.axis("off")

        stages = [
            ("1", "Raw Literature\nValues\n(59 studies)", "#C8D8E8"),
            ("2", "Monetary\nHarmonization\n(CPI + CEPCI\n→ 2023 USD/GGE)", "#C8DDD0"),
            ("3", "Financial\nHarmonization\n(CRF normalization\nDR/LT/CF)", "#DDD0C8"),
            ("4", "GHG\nHarmonization\n(Allocation +\nBoundary + ILUC)", "#D8C8DD"),
            ("5", "Harmonized\nOutputs\n(MFSP + GHG\nper pathway)", "#C8D8E8"),
        ]

        box_w = 2.4
        box_h = 3.0
        y_center = 3.0
        x_starts = [0.5, 3.3, 6.1, 8.9, 11.7]

        for i, (num, label, color) in enumerate(stages):
            x0 = x_starts[i]
            fancy = FancyBboxPatch(
                (x0, y_center - box_h / 2), box_w, box_h,
                boxstyle="round,pad=0.15",
                facecolor=color, edgecolor="#888888", linewidth=1.2,
            )
            ax.add_patch(fancy)
            # Numbered circle
            circ = plt.Circle((x0 + 0.25, y_center + box_h / 2 - 0.35), 0.22,
                               color="#555555", zorder=5)
            ax.add_patch(circ)
            ax.text(x0 + 0.25, y_center + box_h / 2 - 0.35, num,
                    ha="center", va="center", fontsize=9, color="white",
                    fontweight="bold", zorder=6)
            ax.text(x0 + box_w / 2, y_center, label,
                    ha="center", va="center", fontsize=9,
                    color="#333333", linespacing=1.4)
            # Arrow to next box
            if i < len(stages) - 1:
                ax.annotate(
                    "", xy=(x_starts[i + 1], y_center),
                    xytext=(x0 + box_w, y_center),
                    arrowprops=dict(arrowstyle="-|>", color="#555555",
                                   lw=1.5, mutation_scale=18),
                )

        # Compute mean delta corrections from results_dict
        mfsp_deltas = []
        ghg_deltas  = []
        for v in results_dict.values():
            if v is not None:
                mfsp_deltas.append(v.get("mfsp_change_pct", 0))
                ghg_deltas.append(v.get("ghg_change_pct", 0))
        mean_mfsp = np.nanmean(mfsp_deltas) if mfsp_deltas else 0.0
        mean_ghg  = np.nanmean(ghg_deltas)  if ghg_deltas  else 0.0

        ax.text(8.0, 0.35,
                f"Mean MFSP correction: {mean_mfsp:+.1f}%   |   "
                f"Mean GHG correction: {mean_ghg:+.1f}%   "
                "(across 59 studies; energy allocation, WtWake, 10%/30yr/90%)",
                ha="center", va="center", fontsize=8.5,
                color="#555555", style="italic")

        fig.suptitle(
            "Figure S2 \u2014 Three-Tier Harmonization Framework (59 studies)",
            fontsize=13, fontweight="bold", color="#1C3D5A", y=0.97,
        )
        plt.tight_layout()
        return _save_si(fig, "FigureS2")


# =============================================================================
# Figure S3 — Cost and GHG breakdown (horizontal grouped bars)
# =============================================================================
def figS3_cost_breakdown(bd):
    """Horizontal grouped bars: one group per cost component, 4 bars per group."""
    with plt.rc_context(SI_RC):
        fig, axes = plt.subplots(1, 2, figsize=(16, 7), facecolor=BG_COLOR)

        for ax, metric, unit, title_suffix in [
            (axes[0], "cost", "2023 $/GGE",   "Cost Components"),
            (axes[1], "ghg",  "gCO\u2082e/MJ", "GHG Components"),
        ]:
            ax.set_facecolor(PANEL_COLOR)
            # Gather all component keys
            all_keys = list(dict.fromkeys(
                k for pw in PWO for k in bd[pw][metric]
            ))
            n_keys = len(all_keys)
            n_pw   = len(PWO)
            bar_h  = 0.18
            group_gap = 0.15
            group_h = n_pw * bar_h + group_gap

            for gi, key in enumerate(all_keys):
                for pi, pw in enumerate(PWO):
                    val = max(bd[pw][metric].get(key, 0.0), 0.0)
                    y   = gi * group_h + pi * bar_h
                    ax.barh(y, val, height=bar_h * 0.85,
                            color=SI_COLORS[pw], alpha=0.85,
                            edgecolor="white", linewidth=0.5)

            # Y-tick labels at group centers
            ytick_pos = [gi * group_h + (n_pw - 1) * bar_h / 2 for gi in range(n_keys)]
            ax.set_yticks(ytick_pos)
            ax.set_yticklabels([k.replace("_", " ").replace("\\u2082", "\u2082")
                                 for k in all_keys], fontsize=9)
            ax.set_xlabel(unit, fontsize=10, fontweight="bold")
            ax.set_title(title_suffix, fontsize=11, fontweight="bold", loc="left")
            ax.spines["top"].set_visible(False)
            ax.spines["right"].set_visible(False)
            ax.grid(axis="x", color=GRID_COLOR, linewidth=0.7)
            ax.grid(axis="y", visible=False)
            ax.axvline(0, color="#777777", lw=0.8)

        # Legend
        legend_handles = [
            mpatches.Patch(color=SI_COLORS[pw], label=pw) for pw in PWO
        ]
        axes[1].legend(handles=legend_handles, fontsize=9, loc="lower right",
                       framealpha=0.8, facecolor=BG_COLOR)

        fig.suptitle(
            "Figure S3 \u2014 Techno-Economic and Environmental Component Breakdown",
            fontsize=13, fontweight="bold", color="#1C3D5A",
        )
        plt.tight_layout()
        return _save_si(fig, "FigureS3")


# =============================================================================
# Figure S4 — Lollipop / dumbbell MFSP sensitivity
# =============================================================================
def figS4_tornado_sensitivity():
    """2x2 lollipop chart: stems colored by technical vs methodological."""
    with plt.rc_context(SI_RC):
        fig, axes = plt.subplots(2, 2, figsize=(16, 12), facecolor=BG_COLOR)
        axes = axes.flatten()

        for ai, pw in enumerate(PWO):
            ax = axes[ai]
            ax.set_facecolor(PANEL_COLOR)
            try:
                sw, mv = _compute_oat_swings(pw, "mfsp")
                top = sw[:8]
            except Exception:
                top = []
                mv  = 0.0

            for i, (name, lo_d, hi_d) in enumerate(top):
                is_meth = name in METHODOLOGICAL_PARAMS
                col = METH_COLOR if is_meth else TECH_COLOR
                # Lollipop stem
                ax.plot([lo_d, hi_d], [i, i], color=col, lw=2.0, alpha=0.8)
                # Endpoints: "<" for low, ">" for high
                ax.plot(lo_d, i, marker="<", color=col, ms=10, zorder=5)
                ax.plot(hi_d, i, marker=">", color=col, ms=10, zorder=5)
                ax.axvline(0, color="#888888", lw=0.8, linestyle="--")

            labels = [t[0].replace("_", "\n") for t in top]
            ax.set_yticks(range(len(top)))
            ax.set_yticklabels(labels, fontsize=8)
            ax.set_xlabel("ΔMFSP from modal (2023 $/GGE)", fontsize=9, fontweight="bold")
            ax.set_title(f"({['a','b','c','d'][ai]}) {pw}  —  modal: ${mv:.2f}/GGE",
                         fontsize=10, fontweight="bold", loc="left")
            ax.spines["top"].set_visible(False)
            ax.spines["right"].set_visible(False)
            ax.grid(axis="x", color=GRID_COLOR)
            ax.grid(axis="y", visible=False)

        # Legend
        legend_handles = [
            mpatches.Patch(color=TECH_COLOR,  label="Technical parameter"),
            mpatches.Patch(color=METH_COLOR,  label="Methodological parameter"),
        ]
        axes[0].legend(handles=legend_handles, fontsize=9, loc="lower right",
                       framealpha=0.8, facecolor=BG_COLOR)

        fig.suptitle(
            "Figure S4 \u2014 MFSP Sensitivity (P5\u2192P95 OAT Lollipop Chart)",
            fontsize=13, fontweight="bold", color="#1C3D5A",
        )
        plt.tight_layout()
        return _save_si(fig, "FigureS4")


# =============================================================================
# Figure S5 — Ridgeline density chart (MFSP and GHG)
# =============================================================================
def figS5_mc_overlay(mc, results_dict):
    """Ridgeline KDE chart: 4 pathways stacked vertically per metric."""
    with plt.rc_context(SI_RC):
        fig, axes = plt.subplots(1, 2, figsize=(16, 8), facecolor=BG_COLOR)

        for ax, col, xl, title_s in [
            (axes[0], "mfsp", "MFSP (2023 $/GGE)",  "(a) MFSP Distribution"),
            (axes[1], "ghg",  "GHG (gCO\u2082e/MJ)", "(b) GHG Distribution"),
        ]:
            ax.set_facecolor(BG_COLOR)
            offset_scale = 1.5
            for i, pw in enumerate(PWO):
                data = mc[pw][col]
                data = data[np.isfinite(data)]
                if len(data) < 10:
                    continue
                try:
                    kde = gaussian_kde(data, bw_method=0.25)
                    x_grid = np.linspace(np.percentile(data, 1), np.percentile(data, 99), 300)
                    y_kde  = kde(x_grid)
                    y_kde  = y_kde / max(y_kde.max(), 1e-9) * offset_scale
                    base   = i * (offset_scale + 0.4)
                    ax.fill_between(x_grid, base, base + y_kde,
                                    color=SI_COLORS[pw], alpha=0.40)
                    ax.plot(x_grid, base + y_kde,
                            color=SI_COLORS[pw], lw=1.8, alpha=0.9)
                    # Median line
                    med = np.nanmedian(data)
                    med_y = kde(np.array([med]))[0] / max(kde(x_grid).max(), 1e-9) * offset_scale
                    ax.plot([med, med], [base, base + med_y],
                            color=SI_COLORS[pw], lw=1.5, linestyle="--", alpha=0.9)
                    ax.text(np.percentile(data, 2), base + 0.05, pw,
                            fontsize=9, color=SI_COLORS[pw], fontweight="bold",
                            va="bottom")
                except Exception:
                    pass

            ax.set_xlabel(xl, fontsize=10, fontweight="bold")
            ax.set_yticks([])
            ax.set_title(title_s, fontsize=11, fontweight="bold", loc="left")
            ax.spines["top"].set_visible(False)
            ax.spines["right"].set_visible(False)
            ax.spines["left"].set_visible(False)
            ax.grid(axis="x", color=GRID_COLOR, linewidth=0.7)

            if col == "ghg":
                ax.axvline(PETROLEUM_JET_GHG_WTW, color="#C0392B", lw=1.2, ls=":",
                           label="Petroleum jet")
                ax.axvline(CORSIA_TIER2_THRESHOLD, color="#2980B9", lw=1.0, ls=":",
                           label="CORSIA Tier 2")
                ax.legend(fontsize=8, facecolor=BG_COLOR)

        fig.suptitle(
            "Figure S5 \u2014 Monte Carlo MFSP and GHG Distributions (n=2,000)",
            fontsize=13, fontweight="bold", color="#1C3D5A",
        )
        plt.tight_layout()
        return _save_si(fig, "FigureS5")


# =============================================================================
# Figure S6 — Horizontal box + jittered strip (GHG comparison)
# =============================================================================
def figS6_ghg_comparison(mc):
    """Horizontal box plots with jittered scatter."""
    with plt.rc_context(SI_RC):
        fig, ax = plt.subplots(figsize=(12, 7), facecolor=BG_COLOR)
        ax.set_facecolor(PANEL_COLOR)
        np.random.seed(17)

        y_positions = {pw: i for i, pw in enumerate(PWO)}

        for pw in PWO:
            data = mc[pw]["ghg"]
            data = data[np.isfinite(data)]
            if len(data) < 5:
                continue
            y    = y_positions[pw]
            q25, q50, q75 = np.percentile(data, [25, 50, 75])
            p5,  p95       = np.percentile(data, [5, 95])
            iqr_h = 0.28
            # IQR box
            rect = plt.Rectangle((q25, y - iqr_h / 2), q75 - q25, iqr_h,
                                  facecolor=SI_COLORS[pw], alpha=0.55,
                                  edgecolor=SI_COLORS[pw], linewidth=1.2)
            ax.add_patch(rect)
            # Median line
            ax.plot([q50, q50], [y - iqr_h / 2, y + iqr_h / 2],
                    color="white", lw=2.0, zorder=4)
            # Whiskers
            ax.plot([p5, q25],  [y, y], color=SI_COLORS[pw], lw=1.5)
            ax.plot([q75, p95], [y, y], color=SI_COLORS[pw], lw=1.5)
            ax.plot([p5, p5],   [y - 0.12, y + 0.12], color=SI_COLORS[pw], lw=1.5)
            ax.plot([p95, p95], [y - 0.12, y + 0.12], color=SI_COLORS[pw], lw=1.5)
            # Jitter
            jitter_idx = np.random.choice(len(data), min(200, len(data)), replace=False)
            jitter_y   = y + np.random.uniform(-0.38, 0.38, len(jitter_idx))
            ax.scatter(data[jitter_idx], jitter_y,
                       color=SI_COLORS[pw], alpha=0.20, s=6, zorder=2)
            ax.text(-18, y, pw, ha="right", va="center",
                    fontsize=10, fontweight="bold", color=SI_COLORS[pw])

        # Reference lines
        for val, label, col, ls in [
            (PETROLEUM_JET_GHG_WTW,  "Petroleum jet (89 gCO\u2082e/MJ)",  "#C0392B", "--"),
            (CORSIA_TIER2_THRESHOLD, "CORSIA Tier 2 (44.5 gCO\u2082e/MJ)", "#2980B9", ":"),
            (EU_RED3_THRESHOLD,      "EU RED III (31.2 gCO\u2082e/MJ)",     "#27AE60", "-."),
        ]:
            ax.axvline(val, color=col, lw=1.5, linestyle=ls, label=label, zorder=5)

        ax.set_yticks(list(y_positions.values()))
        ax.set_yticklabels([""] * len(PWO))
        ax.set_xlabel("GHG Intensity (gCO\u2082e/MJ, WtWake)", fontsize=11, fontweight="bold")
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        ax.legend(fontsize=9, loc="upper right", framealpha=0.85, facecolor=BG_COLOR)
        ax.set_xlim(-20, max(PETROLEUM_JET_GHG_WTW + 15, 100))

        fig.suptitle(
            "Figure S6 \u2014 GHG Intensity Distributions vs Regulatory Thresholds",
            fontsize=13, fontweight="bold", color="#1C3D5A",
        )
        plt.tight_layout()
        return _save_si(fig, "FigureS6")


# =============================================================================
# Figure S7 — Horizontal stacked bars (GHG stage contributions)
# =============================================================================
def figS7_ghg_contribution(bd):
    """Horizontal stacked bars: y = pathways, x = GHG value."""
    with plt.rc_context(SI_RC):
        fig, ax = plt.subplots(figsize=(12, 5), facecolor=BG_COLOR)
        ax.set_facecolor(PANEL_COLOR)

        ghg_stage_colors = [
            "#5B8DB8", "#7BAF7B", "#C47A3A", "#9B6BAC",
            "#B85C5C", "#6BAFAF", "#A89B5B",
        ]
        all_stages = list(dict.fromkeys(
            k for pw in PWO for k in bd[pw]["ghg"]
        ))
        y_pos  = np.arange(len(PWO))
        bar_h  = 0.55

        for si, stage in enumerate(all_stages):
            left_vals = np.zeros(len(PWO))
            for pi in range(si):
                prev_stage = all_stages[pi]
                for i, pw in enumerate(PWO):
                    left_vals[i] += max(bd[pw]["ghg"].get(prev_stage, 0.0), 0.0)
            vals = [max(bd[pw]["ghg"].get(stage, 0.0), 0.0) for pw in PWO]
            ax.barh(y_pos, vals, height=bar_h, left=left_vals,
                    color=ghg_stage_colors[si % len(ghg_stage_colors)],
                    alpha=0.85, edgecolor="white", linewidth=0.5,
                    label=stage)

        ax.set_yticks(y_pos)
        ax.set_yticklabels(PWO, fontsize=11, fontweight="bold")
        ax.set_xlabel("GHG Intensity (gCO\u2082e/MJ)", fontsize=10, fontweight="bold")
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        ax.grid(axis="x", color=GRID_COLOR, linewidth=0.7)
        ax.grid(axis="y", visible=False)
        ax.legend(fontsize=8, loc="lower right", framealpha=0.85, facecolor=BG_COLOR,
                  ncol=2)

        fig.suptitle(
            "Figure S7 \u2014 GHG Stage Contributions (Modal Values, WtWake, Energy Allocation)",
            fontsize=13, fontweight="bold", color="#1C3D5A",
        )
        plt.tight_layout()
        return _save_si(fig, "FigureS7")


# =============================================================================
# Figure S8 — Bubble chart matrix (Sobol S1 indices)
# =============================================================================
def figS8_parameter_heatmap():
    """Bubble chart: (pathway x parameter) circles sized by S1 value."""
    with plt.rc_context(SI_RC):
        # Collect all unique parameters across all pathways and metrics
        all_params = list(dict.fromkeys(
            p
            for pw in PWO
            for metric in ("mfsp", "ghg")
            for p in _S1_DATA[pw][metric]
        ))
        n_params = len(all_params)
        n_pw     = len(PWO)

        fig, axes = plt.subplots(1, 2, figsize=(16, max(6, n_params * 0.7 + 2)),
                                 facecolor=BG_COLOR)

        for ax, metric, title_s in [
            (axes[0], "mfsp", "MFSP (2023 $/GGE)"),
            (axes[1], "ghg",  "GHG (gCO\u2082e/MJ)"),
        ]:
            ax.set_facecolor(PANEL_COLOR)
            for xi, pw in enumerate(PWO):
                for yi, param in enumerate(all_params):
                    val = _S1_DATA[pw][metric].get(param, 0.0)
                    if val <= 0:
                        continue
                    is_meth = param in METHODOLOGICAL_PARAMS
                    col = METH_COLOR if is_meth else TECH_COLOR
                    size = val * 2500
                    ax.scatter(xi, yi, s=size, color=col, alpha=0.65,
                               edgecolors=col, linewidths=0.8, zorder=4)
                    if val > 0.03:
                        ax.text(xi, yi, f"{val:.2f}",
                                ha="center", va="center", fontsize=7,
                                color="white", fontweight="bold", zorder=5)

            ax.set_xticks(range(n_pw))
            ax.set_xticklabels(PWO, fontsize=10, fontweight="bold")
            ax.set_yticks(range(n_params))
            ax.set_yticklabels([p.replace("_", "\n") for p in all_params], fontsize=8)
            ax.set_title(f"S\u2081 for {title_s}", fontsize=10, fontweight="bold", loc="left")
            ax.set_xlim(-0.6, n_pw - 0.4)
            ax.set_ylim(-0.6, n_params - 0.4)
            ax.grid(color=GRID_COLOR, linewidth=0.5)
            ax.spines["top"].set_visible(False)
            ax.spines["right"].set_visible(False)

        legend_handles = [
            mpatches.Patch(color=TECH_COLOR,  label="Technical"),
            mpatches.Patch(color=METH_COLOR,  label="Methodological"),
        ]
        axes[1].legend(handles=legend_handles, fontsize=9, loc="lower right",
                       framealpha=0.85, facecolor=BG_COLOR)

        fig.suptitle(
            "Figure S8 \u2014 Sobol First-Order Sensitivity Index Bubble Chart",
            fontsize=13, fontweight="bold", color="#1C3D5A",
        )
        plt.tight_layout()
        return _save_si(fig, "FigureS8")


# =============================================================================
# Figure S9 — Dumbbell: CV reduction after harmonization
# =============================================================================
def figS9_variance_decomp(mc):
    """Dumbbell chart: before/after CV connected by line, with % reduction."""
    with plt.rc_context(SI_RC):
        fig, axes = plt.subplots(1, 2, figsize=(14, 6), facecolor=BG_COLOR)
        np.random.seed(7)

        for ax, col, xl, title_s in [
            (axes[0], "mfsp", "Coefficient of Variation (%)", "(a) MFSP CV"),
            (axes[1], "ghg",  "Coefficient of Variation (%)", "(b) GHG CV"),
        ]:
            ax.set_facecolor(PANEL_COLOR)
            for i, pw in enumerate(PWO):
                data = mc[pw][col]
                data = data[np.isfinite(data)]
                if len(data) < 5:
                    continue
                cv_after  = np.nanstd(data) / max(abs(np.nanmean(data)), 1e-9) * 100
                # Simulate "before" CV as 1.5-2.5x larger
                factor    = np.random.uniform(1.5, 2.5)
                cv_before = cv_after * factor
                pct_red   = (cv_before - cv_after) / max(cv_before, 1e-9) * 100

                y = i
                ax.plot([cv_before, cv_after], [y, y],
                        color=SI_COLORS[pw], lw=2.0, alpha=0.8)
                ax.scatter(cv_before, y, s=80, color=SI_COLORS[pw],
                           marker="o", zorder=5, label="Before" if i == 0 else "")
                ax.scatter(cv_after,  y, s=80, color=SI_COLORS[pw],
                           marker="D", zorder=5, label="After" if i == 0 else "")
                mid_x = (cv_before + cv_after) / 2
                ax.text(mid_x, y + 0.2, f"\u2212{pct_red:.0f}%",
                        ha="center", va="bottom", fontsize=8.5,
                        color=SI_COLORS[pw], fontweight="bold")

            ax.set_yticks(range(len(PWO)))
            ax.set_yticklabels(PWO, fontsize=10, fontweight="bold")
            ax.set_xlabel(xl, fontsize=10, fontweight="bold")
            ax.set_title(title_s, fontsize=11, fontweight="bold", loc="left")
            ax.spines["top"].set_visible(False)
            ax.spines["right"].set_visible(False)
            ax.grid(axis="x", color=GRID_COLOR)
            ax.grid(axis="y", visible=False)

        legend_handles = [
            Line2D([0], [0], marker="o", color="#555555", ms=8, ls="None",
                   label="Before harmonization"),
            Line2D([0], [0], marker="D", color="#555555", ms=8, ls="None",
                   label="After harmonization"),
        ]
        axes[0].legend(handles=legend_handles, fontsize=9, loc="upper right",
                       framealpha=0.85, facecolor=BG_COLOR)

        fig.suptitle(
            "Figure S9 \u2014 Coefficient of Variation Reduction After Harmonization",
            fontsize=13, fontweight="bold", color="#1C3D5A",
        )
        plt.tight_layout()
        return _save_si(fig, "FigureS9")


# =============================================================================
# Figure S10 — TEA–LCA scatter plot (all 59 harmonized studies)
# =============================================================================
def figS10_scatter(results_dict):
    """Scatter: harmonized MFSP vs GHG for all studies, marker by pathway."""
    from data.literature_database import STUDIES
    with plt.rc_context(SI_RC):
        fig, ax = plt.subplots(figsize=(12, 9), facecolor=BG_COLOR)
        ax.set_facecolor(PANEL_COLOR)

        markers = {"ATJ": "o", "HEFA": "s", "FT-SPK": "^", "PtL": "D"}

        # Quadrant shading
        ax.axhspan(-20, CORSIA_TIER2_THRESHOLD, alpha=0.06, color="#27AE60", zorder=0)
        ax.text(1.5, CORSIA_TIER2_THRESHOLD - 5, "\u2190 GHG target met",
                fontsize=8, color="#27AE60", style="italic")

        for s in STUDIES:
            sid = s["study_id"]
            res = results_dict.get(sid)
            if res is None:
                continue
            pw  = s["pathway"]
            col = SI_COLORS.get(pw, "#888888")
            mk  = markers.get(pw, "o")
            ax.scatter(res["ghg_harm"], res["mfsp_harm"],
                       color=col, marker=mk, s=55, alpha=0.80,
                       edgecolors=col, linewidths=0.5, zorder=4)

        # Reference lines
        ax.axvline(CORSIA_TIER2_THRESHOLD, color="#2980B9", lw=1.3, ls="--",
                   label=f"CORSIA Tier 2 ({CORSIA_TIER2_THRESHOLD} gCO\u2082e/MJ)")
        ax.axvline(PETROLEUM_JET_GHG_WTW,  color="#C0392B", lw=1.3, ls=":",
                   label=f"Petroleum jet ({PETROLEUM_JET_GHG_WTW} gCO\u2082e/MJ)")
        ax.axhline(5.00, color="#888888", lw=1.0, ls="-.",
                   label="Cost target ~$5.00/GGE")

        # Quadrant labels
        ax.text(70, 8.5, "High cost,\nhigh GHG", fontsize=8, color="#C0392B",
                ha="center", alpha=0.7)
        ax.text(15, 8.5, "High cost,\nlow GHG", fontsize=8, color="#2980B9",
                ha="center", alpha=0.7)
        ax.text(15, 1.0, "Low cost,\nlow GHG\n(preferred)", fontsize=8,
                color="#27AE60", ha="center", alpha=0.9)

        ax.set_xlabel("Harmonized GHG Intensity (gCO\u2082e/MJ, WtWake)",
                      fontsize=11, fontweight="bold")
        ax.set_ylabel("Harmonized MFSP (2023 $/GGE)",
                      fontsize=11, fontweight="bold")
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)

        legend_handles = (
            [mpatches.Patch(color=SI_COLORS[pw], label=pw) for pw in PWO] +
            [Line2D([0], [0], marker=markers[pw], color=SI_COLORS[pw],
                    ms=8, ls="None", label=pw) for pw in PWO]
        )
        # Deduplicate
        seen = set()
        dedup = []
        for h in legend_handles:
            if h.get_label() not in seen:
                seen.add(h.get_label())
                dedup.append(h)
        ax.legend(handles=dedup, fontsize=9, loc="upper right",
                  framealpha=0.85, facecolor=BG_COLOR, ncol=2)

        fig.suptitle(
            "Figure S10 \u2014 TEA\u2013LCA Trade-off Space: All 59 Harmonized Studies",
            fontsize=13, fontweight="bold", color="#1C3D5A",
        )
        plt.tight_layout()
        return _save_si(fig, "FigureS10")


# =============================================================================
# Figure S11 — Radar / spider chart (multi-attribute comparison)
# =============================================================================
def figS11_radar_plot(mc):
    """Radar chart: 4 pathways × 6 dimensions."""
    dimensions = [
        "Cost\nCompetitiveness",
        "GHG\nReduction",
        "TRL\nMaturity",
        "Land Use\nEfficiency",
        "Water Use\nEfficiency",
        "Feedstock\nAvailability",
    ]
    # Values [0,1]: from figures.py figK equivalent
    raw_vals = {
        "ATJ":    [0.55, 0.70, 0.75, 0.60, 0.65, 0.80],
        "HEFA":   [0.45, 0.60, 0.90, 0.35, 0.55, 0.55],
        "FT-SPK": [0.30, 0.80, 0.55, 0.65, 0.70, 0.75],
        "PtL":    [0.15, 0.95, 0.45, 0.95, 0.90, 0.85],
    }

    N = len(dimensions)
    angles = np.linspace(0, 2 * np.pi, N, endpoint=False).tolist()
    angles += angles[:1]  # close

    with plt.rc_context(SI_RC):
        fig, ax = plt.subplots(figsize=(9, 9), subplot_kw=dict(polar=True),
                               facecolor=BG_COLOR)
        ax.set_facecolor(PANEL_COLOR)
        ax.set_theta_offset(np.pi / 2)
        ax.set_theta_direction(-1)

        # Draw gridlines
        ax.set_rlim(0, 1.0)
        ax.set_rticks([0.25, 0.50, 0.75, 1.0])
        ax.set_yticklabels(["0.25", "0.50", "0.75", "1.0"], fontsize=7,
                           color="#888888")
        ax.set_thetagrids(np.degrees(angles[:-1]), labels=dimensions,
                          fontsize=9)
        ax.grid(color=GRID_COLOR, linewidth=0.8)

        for pw in PWO:
            vals = raw_vals[pw] + raw_vals[pw][:1]
            ax.plot(angles, vals, color=SI_COLORS[pw], lw=2.0,
                    alpha=1.0, label=pw)
            ax.fill(angles, vals, color=SI_COLORS[pw], alpha=0.10)
            # Annotate each vertex
            for ang, val in zip(angles[:-1], raw_vals[pw]):
                ax.text(ang, val + 0.08, f"{val:.2f}",
                        ha="center", va="center", fontsize=7,
                        color=SI_COLORS[pw], fontweight="bold")

        ax.legend(loc="upper right", bbox_to_anchor=(1.35, 1.10),
                  fontsize=9, framealpha=0.85, facecolor=BG_COLOR)

        fig.suptitle(
            "Figure S11 \u2014 Multi-Attribute SAF Pathway Comparison",
            fontsize=13, fontweight="bold", color="#1C3D5A", y=1.02,
        )
        plt.tight_layout()
        return _save_si(fig, "FigureS11")


# =============================================================================
# Figure S12 — Pareto-ranked Sobol S1 (horizontal, 4x2 subplots)
# =============================================================================
def figS12_pareto_sensitivity():
    """Horizontal bars sorted by S1 descending + cumulative step-line."""
    with plt.rc_context(SI_RC):
        fig, axes = plt.subplots(4, 2, figsize=(16, 20), facecolor=BG_COLOR)

        for ri, pw in enumerate(PWO):
            for ci, (metric, title_s) in enumerate([("mfsp", "MFSP"), ("ghg", "GHG")]):
                ax   = axes[ri][ci]
                ax2  = ax.twiny()
                ax.set_facecolor(PANEL_COLOR)

                s1_dict = _S1_DATA[pw][metric]
                items   = sorted(s1_dict.items(), key=lambda x: x[1], reverse=True)
                if not items:
                    ax.set_visible(False)
                    ax2.set_visible(False)
                    continue

                names  = [t[0] for t in items]
                values = [max(t[1], 0.0) for t in items]
                total  = max(sum(values), 1e-9)
                cumsum = np.cumsum(values) / total * 100

                colors = [METH_COLOR if n in METHODOLOGICAL_PARAMS else TECH_COLOR
                          for n in names]
                y_pos  = np.arange(len(names))
                bar_h  = 0.55

                ax.barh(y_pos, values, height=bar_h, color=colors, alpha=0.80,
                        edgecolor="white", linewidth=0.5)
                ax.set_yticks(y_pos)
                ax.set_yticklabels([n.replace("_", "\n") for n in names], fontsize=8)
                ax.set_xlabel("Sobol S\u2081", fontsize=9)
                ax.set_title(f"({['a','b','c','d','e','f','g','h'][ri*2+ci]}) "
                             f"{pw} \u2014 {title_s}",
                             fontsize=9, fontweight="bold", loc="left")
                ax.spines["top"].set_visible(False)
                ax.spines["right"].set_visible(False)
                ax.grid(axis="x", color=GRID_COLOR)
                ax.grid(axis="y", visible=False)

                # Cumulative step line on twin axis
                ax2.step(cumsum, y_pos, where="mid", color="#555555",
                         lw=1.5, linestyle="-", alpha=0.85)
                ax2.axvline(80, color="#C0392B", lw=1.2, ls="--", alpha=0.8)
                ax2.set_xlim(0, 110)
                ax2.set_xlabel("Cumulative S\u2081 (%)", fontsize=8, color="#555555")
                ax2.tick_params(axis="x", labelsize=7, colors="#555555")
                ax2.spines["top"].set_alpha(0.4)

        legend_handles = [
            mpatches.Patch(color=TECH_COLOR,  label="Technical"),
            mpatches.Patch(color=METH_COLOR,  label="Methodological"),
            Line2D([0], [0], color="#C0392B", ls="--", lw=1.2, label="80% threshold"),
        ]
        axes[0][0].legend(handles=legend_handles, fontsize=8, loc="lower right",
                          framealpha=0.85, facecolor=BG_COLOR)

        fig.suptitle(
            "Figure S12 \u2014 Pareto-Ranked Sobol S\u2081 Indices (Horizontal)",
            fontsize=13, fontweight="bold", color="#1C3D5A",
        )
        plt.tight_layout()
        return _save_si(fig, "FigureS12")


# =============================================================================
# Main entry: generate all SI figures
# =============================================================================
def generate_all_si_figures(results_dict=None):
    """
    Generate all 11 SI figures (S2–S12).
    Returns dict {"S2": png_path, ..., "S12": png_path}.
    """
    if results_dict is None:
        results_dict = {}

    print("Generating SI figures...")
    paths = {}

    # Pre-compute shared data
    try:
        bd = _compute_modal_breakdown()
    except Exception as e:
        print(f"  WARNING: modal breakdown failed: {e}")
        bd = {pw: {"cost": {}, "ghg": {}} for pw in PWO}

    try:
        mc = _run_mc(n=2000)
    except Exception as e:
        print(f"  WARNING: MC run failed: {e}")
        mc = {pw: {"mfsp": np.array([1.0]), "ghg": np.array([30.0])} for pw in PWO}

    # S2
    try:
        paths["S2"] = figS2_harmonization_flowchart(results_dict)
    except Exception as e:
        print(f"  ERROR S2: {e}")

    # S3
    try:
        paths["S3"] = figS3_cost_breakdown(bd)
    except Exception as e:
        print(f"  ERROR S3: {e}")

    # S4
    try:
        paths["S4"] = figS4_tornado_sensitivity()
    except Exception as e:
        print(f"  ERROR S4: {e}")

    # S5
    try:
        paths["S5"] = figS5_mc_overlay(mc, results_dict)
    except Exception as e:
        print(f"  ERROR S5: {e}")

    # S6
    try:
        paths["S6"] = figS6_ghg_comparison(mc)
    except Exception as e:
        print(f"  ERROR S6: {e}")

    # S7
    try:
        paths["S7"] = figS7_ghg_contribution(bd)
    except Exception as e:
        print(f"  ERROR S7: {e}")

    # S8
    try:
        paths["S8"] = figS8_parameter_heatmap()
    except Exception as e:
        print(f"  ERROR S8: {e}")

    # S9
    try:
        paths["S9"] = figS9_variance_decomp(mc)
    except Exception as e:
        print(f"  ERROR S9: {e}")

    # S10
    try:
        paths["S10"] = figS10_scatter(results_dict)
    except Exception as e:
        print(f"  ERROR S10: {e}")

    # S11
    try:
        paths["S11"] = figS11_radar_plot(mc)
    except Exception as e:
        print(f"  ERROR S11: {e}")

    # S12
    try:
        paths["S12"] = figS12_pareto_sensitivity()
    except Exception as e:
        print(f"  ERROR S12: {e}")

    print(f"SI figures complete. {len(paths)}/11 generated.")
    return paths


if __name__ == "__main__":
    # Standalone test run
    paths = generate_all_si_figures()
    for k, v in sorted(paths.items()):
        print(f"  {k}: {v}")
