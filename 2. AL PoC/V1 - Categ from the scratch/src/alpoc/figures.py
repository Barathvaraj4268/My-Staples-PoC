"""PNG figures (matplotlib), dataviz reference palette (validated with validate_palette.js).

Retailer identity is fixed across every chart: Staples = blue, competitor = orange.
Decision groups: Approve = aqua, Review/hold = yellow, Reject = violet (contrast relief: legends + tables).
"""
from __future__ import annotations

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from scipy.stats import gaussian_kde

INK = {"primary": "#0b0b0b", "secondary": "#52514e", "muted": "#898781", "grid": "#e1e0d9",
       "axis": "#c3c2b7", "surface": "#fcfcfb"}
C_STAPLES, C_COMP = "#2a78d6", "#eb6834"
GROUP = {"Approve": "#1baf7a", "Review / hold": "#eda100", "Reject": "#4a3aa7"}
TIER = {"Strong": "#1baf7a", "Gap-led": "#4a3aa7", "Vector-led": "#eda100", "Weak": "#b5b3ab"}
SINGLE = "#4a3aa7"
DPI = 110

plt.rcParams.update({
    "font.family": ["Segoe UI", "DejaVu Sans"], "font.size": 9, "axes.edgecolor": INK["axis"],
    "axes.labelcolor": INK["secondary"], "xtick.color": INK["muted"], "ytick.color": INK["muted"],
    "axes.titlesize": 10.5, "axes.titleweight": "semibold", "axes.titlecolor": INK["primary"],
    "axes.titlelocation": "left", "axes.spines.top": False, "axes.spines.right": False,
    "figure.facecolor": INK["surface"], "axes.facecolor": INK["surface"], "savefig.facecolor": INK["surface"],
    "legend.frameon": False, "legend.fontsize": 8.5, "axes.grid": False,
})


def esc(s) -> str:
    """Matplotlib treats paired '$' as math; escape them in any text drawn from data."""
    return str(s).replace("$", r"\$")


def _grid(ax, axis="y"):
    ax.grid(axis=axis, color=INK["grid"], linewidth=0.8)
    ax.set_axisbelow(True)


def _save(fig, path):
    fig.tight_layout()
    fig.savefig(path, dpi=DPI)
    plt.close(fig)
    return path


def _pct_axis(ax, axis="y"):
    from matplotlib.ticker import PercentFormatter
    (ax.yaxis if axis == "y" else ax.xaxis).set_major_formatter(PercentFormatter(1.0, decimals=0))


def price_bands(bands: pd.DataFrame, comp_name: str, path):
    fig, ax = plt.subplots(figsize=(6.4, 3.0))
    x = np.arange(len(bands))
    w = 0.36
    for off, col, colour, lab in ((-w / 2, "share_staples", C_STAPLES, "Staples"),
                                  (w / 2, "share_competitor", C_COMP, comp_name)):
        ax.bar(x + off, bands[col], width=w, color=colour, edgecolor=INK["surface"], linewidth=1.5, label=lab)
    for i, r in enumerate(bands.itertuples()):
        if r.credibility >= 0.9 and r.share_competitor - r.share_staples >= 0.05:
            ax.annotate("gap", (i + w / 2, r.share_competitor), textcoords="offset points", xytext=(0, 3),
                        ha="center", fontsize=7.5, color=INK["secondary"])
    ax.set_xticks(x, [esc(b) for b in bands["band"]], fontsize=8)
    _pct_axis(ax)
    _grid(ax)
    ax.set_ylabel("share of families")
    ax.set_title("Price-band coverage")
    ax.legend(loc="upper right", ncols=2)
    return _save(fig, path)


def dfi_density(s: pd.Series, c: pd.Series, comp_name: str, path):
    fig, ax = plt.subplots(figsize=(6.4, 2.8))
    xs = np.linspace(0, 1, 200)
    for vals, colour, lab in ((s, C_STAPLES, "Staples"), (c, C_COMP, comp_name)):
        vals = vals.dropna()
        if len(vals) < 5:
            continue
        y = gaussian_kde(vals, bw_method=0.25)(xs)
        ax.plot(xs, y, color=colour, linewidth=2, label=f"{lab} (median {vals.median():.2f})")
        ax.fill_between(xs, y, color=colour, alpha=0.10, linewidth=0)
    ax.axvline(0.6, color=INK["axis"], linewidth=1)
    ax.text(0.61, ax.get_ylim()[1] * 0.92, "design-forward →", fontsize=7.5, color=INK["muted"])
    ax.set_xlim(0, 1)
    ax.set_yticks([])
    ax.spines["left"].set_visible(False)
    ax.set_xlabel("Design-Forward Index (0 = utilitarian, 1 = design-led; pooled percentile)")
    ax.set_title("Design-forward distribution")
    ax.legend(loc="upper left")
    return _save(fig, path)


def attribute_jsd(df: pd.DataFrame, path):
    d = df.groupby("attribute")["jsd"].first().sort_values()
    fig, ax = plt.subplots(figsize=(6.4, 0.28 * len(d) + 0.9))
    ax.barh(d.index.str.replace("_", " "), d.values, height=0.6, color=SINGLE)
    for i, v in enumerate(d.values):
        ax.text(v + 0.005, i, f"{v:.2f}", va="center", fontsize=7.5, color=INK["secondary"])
    _grid(ax, "x")
    ax.set_xlabel("Jensen-Shannon divergence (0 = same mix, 1 = no overlap)")
    ax.set_title("How different is the mix, by attribute")
    return _save(fig, path)


def value_gaps(df: pd.DataFrame, comp_name: str, path, n_pos=8, n_neg=4):
    d = df[(df["credibility"] >= 0.9) | (df["credibility"] <= 0.1)]
    d = d[~d["descriptive_only"]]
    pos = d[d["delta"] > 0].nlargest(n_pos, "delta")
    neg = d[d["delta"] < 0].nsmallest(n_neg, "delta")
    d = pd.concat([neg, pos]).sort_values("delta")
    if d.empty:
        return None
    fig, ax = plt.subplots(figsize=(6.4, 0.3 * len(d) + 1.0))
    vals = d["value"].astype(str).str.replace(r"^([a-z])_", lambda m: m.group(1).upper() + "-", regex=True).str.replace("_", " ")
    labels = [esc(s) for s in (d["attribute"].str.replace("_", " ") + ": " + vals)]
    colours = np.where(d["delta"] > 0, C_COMP, C_STAPLES)
    ax.barh(labels, d["delta"], height=0.6, color=colours)
    ax.axvline(0, color=INK["axis"], linewidth=1)
    for i, (v, r) in enumerate(zip(d["delta"], d.itertuples())):
        ax.text(v + (0.004 if v > 0 else -0.004), i, f"{r.share_competitor:.0%} vs {r.share_staples:.0%}",
                va="center", ha="left" if v > 0 else "right", fontsize=7.5, color=INK["secondary"])
    _pct_axis(ax, "x")
    _grid(ax, "x")
    lim = max(abs(d["delta"]).max() * 1.6, 0.05)
    ax.set_xlim(-lim, lim)
    ax.set_xlabel(f"share gap ({comp_name} − Staples); labels = {comp_name} vs Staples")
    ax.set_title("Largest credible attribute-value gaps")
    from matplotlib.patches import Patch
    ax.legend(handles=[Patch(color=C_COMP, label=f"{comp_name} carries more"), Patch(color=C_STAPLES, label="Staples carries more")],
              loc="lower right")
    return _save(fig, path)


def decision_scatter(cand: pd.DataFrame, T: dict, v: dict, path):
    fig, ax = plt.subplots(figsize=(6.4, 3.8))
    for g in ["Reject", "Review / hold", "Approve"]:
        d = cand[cand["label_group"] == g]
        ax.scatter(d["crs"], d["aas"], s=16, color=GROUP[g], edgecolor=INK["surface"], linewidth=0.8, label=f"{g} ({len(d)})",
                   alpha=0.9)
    for xv in (v["crs_low"], v["crs_high"]):
        ax.axvline(xv, color=INK["axis"], linewidth=1)
    for yv in (T["aas_low"], T["aas_high"]):
        ax.axhline(yv, color=INK["axis"], linewidth=1)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.set_xlabel("CRS: cannibalisation risk (price- and look-agnostic)")
    ax.set_ylabel("AAS: fit with Staples")
    ax.set_title("Decision surface: every competitor product in this node")
    ax.legend(loc="lower left", fontsize=8)
    return _save(fig, path)


def agreement(fa: pd.DataFrame, path):
    fig, ax = plt.subplots(figsize=(6.4, 4.0))
    for t in ["Weak", "Vector-led", "Gap-led", "Strong"]:
        d = fa[fa["tier"] == t]
        e, ne = d[d["eligible"]], d[~d["eligible"]]
        ax.scatter(e["tg"], e["vos"], s=46, color=TIER[t], edgecolor=INK["surface"], linewidth=1.2, label=t)
        ax.scatter(ne["tg"], ne["vos"], s=40, facecolor="none", edgecolor=TIER[t], linewidth=1.2)
    sl = fa[fa["shortlisted"]].sort_values("final_rank")
    for r in sl.itertuples():
        ax.annotate(f"#{int(r.final_rank)}", (r.tg, r.vos), textcoords="offset points", xytext=(6, 4), fontsize=8,
                    color=INK["primary"], weight="semibold")
    _grid(ax, "both")
    ax.set_xlabel("TG: Total Gap (Method 2, attribute view)")
    ax.set_ylabel("VOS: Vector Opportunity (Method 1)")
    ax.set_title("Two methods, one view (hollow = failed the safety gate)")
    ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.16), ncols=4, fontsize=8)
    return _save(fig, path)


def label_mix(fa: pd.DataFrame, path, n=12):
    d = fa[fa["depth"] > 0].copy()
    d["sort"] = d["final"].fillna(-1)
    d = d.sort_values(["sort", "tg"], ascending=False).head(n).iloc[::-1]
    if d.empty:
        return None
    appr = d[["share_CURATE", "share_STYLE-EXTENSION", "share_TRADE-UP"]].sum(axis=1)
    hold = d[["share_REVIEW", "share_EDGE"]].sum(axis=1)
    rej = 1 - appr - hold
    fig, ax = plt.subplots(figsize=(6.4, 0.34 * len(d) + 1.0))
    y = np.arange(len(d))
    names = [esc(s if len(s) <= 52 else s[:50] + "…") for s in d["name"]]
    left = np.zeros(len(d))
    for vals, g in ((appr, "Approve"), (hold, "Review / hold"), (rej, "Reject")):
        ax.barh(y, vals, left=left, height=0.6, color=GROUP[g], edgecolor=INK["surface"], linewidth=1.5, label=g)
        left += vals.values
    ax.set_yticks(y, names, fontsize=7.5)
    _pct_axis(ax, "x")
    ax.set_xlim(0, 1)
    ax.set_title("Label mix of competitor products per archetype")
    ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.08), ncols=3, fontsize=8)
    return _save(fig, path)


def node_coverage(nodes: pd.DataFrame, comp_name: str, path):
    d = nodes[nodes["status"].isin(["scored", "thin"])].copy()
    d["label"] = d["node_id"].str.split(" > ").map(lambda p: " > ".join(p[2:]) if len(p) > 2 else p[-1])
    d = d.sort_values("n_competitor")
    fig, ax = plt.subplots(figsize=(7.2, 0.26 * len(d) + 1.2))
    y = np.arange(len(d))
    ax.barh(y - 0.2, d["n_staples"], height=0.38, color=C_STAPLES, label="Staples families")
    ax.barh(y + 0.2, d["n_competitor"], height=0.38, color=C_COMP, label=f"{comp_name} families")
    ax.set_yticks(y, [l + ("  (thin)" if s == "thin" else "") for l, s in zip(d["label"], d["status"])], fontsize=7.5)
    _grid(ax, "x")
    ax.set_xlabel("product families")
    ax.set_title("Families per analysis node")
    ax.legend(loc="lower right")
    return _save(fig, path)


def extractor_accuracy(val: dict, bar: float, path):
    d = pd.Series({k: v["accuracy_when_found"] for k, v in val.items() if v["accuracy_when_found"] is not None}).sort_values()
    if d.empty:
        return None
    fig, ax = plt.subplots(figsize=(6.4, 0.28 * len(d) + 1.0))
    ax.barh(d.index.str.replace("_", " "), d.values, height=0.6, color=SINGLE)
    ax.axvline(bar, color=INK["secondary"], linewidth=1)
    ax.text(bar, len(d) - 0.4, f" G2 bar {bar:.0%}", fontsize=7.5, color=INK["secondary"])
    for i, v in enumerate(d.values):
        ax.text(v + 0.01, i, f"{v:.0%}", va="center", fontsize=7.5, color=INK["secondary"])
    _pct_axis(ax, "x")
    ax.set_xlim(0, 1.1)
    ax.set_title("Text instrument vs Staples specs (accuracy when found)")
    return _save(fig, path)


# score-component colours: kept clear of the retailer blue/orange
COMP_COL = {"vw": "#4a3aa7", "aas": "#1baf7a", "ad": "#eda100",
            "lsr": "#4a3aa7", "ppg": "#1baf7a", "cg": "#eda100", "msg": "#c2477f", "dfg": "#6f8f2e"}


def _component_bars(d: pd.DataFrame, parts: list[tuple[str, str, str]], score: str, title: str, xlabel: str,
                    path, note: pd.Series | None = None):
    """Horizontal stacked bars, highest score on top; failed-gate archetypes drawn faded."""
    d = d.iloc[::-1]
    fig, ax = plt.subplots(figsize=(6.8, 0.34 * len(d) + 1.5))
    y = np.arange(len(d))
    alpha = np.where(d["eligible"].fillna(False).astype(bool), 1.0, 0.35)
    left = np.zeros(len(d))
    for col, colour, lab in parts:
        vals = d[col].fillna(0).to_numpy()
        for i in range(len(d)):
            ax.barh(y[i], vals[i], left=left[i], height=0.62, color=colour, alpha=alpha[i],
                    edgecolor=INK["surface"], linewidth=1.2)
        left += vals
    for i, (s, ok) in enumerate(zip(d[score], d["eligible"].fillna(False))):
        extra = "" if note is None else f"  {note.iloc[i]}"
        ax.text(left[i] + left.max() * 0.015, i, f"{s:.0f}" + extra + ("" if ok else " · failed gate"), va="center",
                fontsize=7.5, color=INK["primary"] if ok else INK["muted"])
    names = [esc(s if len(s) <= 50 else s[:48] + "…") for s in d["name"]]
    ax.set_yticks(y, names, fontsize=7.5)
    ax.set_xlim(0, float(left.max()) * (1.55 if note is not None else 1.3))
    _grid(ax, "x")
    ax.set_xlabel(xlabel)
    ax.set_title(title)
    from matplotlib.patches import Patch
    ax.legend(handles=[Patch(color=c, label=l) for _, c, l in parts], loc="upper center",
              bbox_to_anchor=(0.5, -0.1 if len(d) > 6 else -0.2), ncols=3, fontsize=8)
    return _save(fig, path)


def vos_components(fa: pd.DataFrame, w: dict, gamma: float, path, n: int = 15):
    """VOS split into what VW, AAS and AD contribute after the CRS damping, plus the part CRS removed."""
    d = fa[fa["depth"] > 0].dropna(subset=["vos"]).nlargest(n, "vos").copy()
    if d.empty:
        return None
    damp = (1 - d["crs"] / 100) ** gamma
    d["c_vw"] = 100 * w["vw"] * d["pct_vw"] * damp
    d["c_aas"] = 100 * w["aas"] * d["aas"] / 100 * damp
    d["c_ad"] = 100 * w["ad"] * d["pct_ad"] * damp
    pre = 100 * (w["vw"] * d["pct_vw"] + w["aas"] * d["aas"] / 100 + w["ad"] * d["pct_ad"])
    note = pre.map(lambda x: f"({x:.0f} before risk)")
    return _component_bars(d, [("c_vw", COMP_COL["vw"], "Whitespace (VW)"), ("c_aas", COMP_COL["aas"], "Fit (AAS)"),
                               ("c_ad", COMP_COL["ad"], "Look difference (AD)")],
                           "vos", f"Method 1: VOS by archetype (top {len(d)})",
                           "VOS (0–100) = segments after the cannibalisation-risk discount", path, note=note.iloc[::-1])


def tg_components(fa: pd.DataFrame, w: dict, credible: float, shrink: float, path, n: int = 15):
    """TG split into the weighted percentile contribution of each gap metric."""
    d = fa[fa["depth"] > 0].dropna(subset=["tg"]).nlargest(n, "tg").copy()
    if d.empty:
        return None
    sh = np.where(d["lsr_credibility"] >= credible, 1.0, shrink)
    for k in ("lsr", "ppg", "cg", "msg", "dfg"):
        d[f"c_{k}"] = 100 * w[k] * d[f"pct_{k}"] * (sh if k == "lsr" else 1.0)
    return _component_bars(d, [("c_lsr", COMP_COL["lsr"], "Share (LSR)"), ("c_ppg", COMP_COL["ppg"], "Price (PPG)"),
                               ("c_cg", COMP_COL["cg"], "Colour (CG)"), ("c_msg", COMP_COL["msg"], "Material/style (MSG)"),
                               ("c_dfg", COMP_COL["dfg"], "Design-forward (DFG)")],
                           "tg", f"Method 2: TG by archetype (top {len(d)})", "TG (0–100); bar length = score", path)
