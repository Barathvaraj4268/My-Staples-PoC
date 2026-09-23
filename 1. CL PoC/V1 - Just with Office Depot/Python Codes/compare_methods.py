#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
=====================================================================================
 STAPLES MARKETPLACE PoC - CATEGORY LEVEL:  COMPARE METHOD V (vector) vs METHOD G (graph)
=====================================================================================
Run AFTER method_v_vector.py and method_g_graph.py (same folder). Needs only pandas,
numpy, openpyxl, matplotlib.

Why: the two methods measure the same things (gap, adjacency, cannibalization) with
independent machinery - word-meaning vectors vs wording + tree structure. Where they
agree, the call is robust to the method; where they disagree, a human looks.

Confidence tiers (per recommendation unit, i.e. a competitor L1-L3 category):
   CONSENSUS APPROVE   both methods label it CURATE
   LIKELY APPROVE      one CURATE, the other in the review zone or VERTICAL EXTENSION
   CONSENSUS 1P ROUTE  both say competitor is deeper in a CORE line -> fix via 1P, not 3P
   CONSENSUS HOLD      both say VERTICAL EXTENSION (phase 2) or both OFF-BRAND
   VERIFY              either method thinks it is a naming gap, or they disagree on presence
   SPLIT               anything else - human review

Outputs: outputs_compare/method_comparison.xlsx, consensus_top.png
"""
# %% [0] INSTALL  (Colab)
# !pip install -q pandas numpy openpyxl matplotlib

# %% [1] SETUP
import os

import numpy as np
import pandas as pd

V_FILE = "outputs_method_v/node_scores_method_v.csv"
G_FILE = "outputs_method_g/node_scores_method_g.csv"
OUT = "outputs_compare"
TOP_N = 20
os.makedirs(OUT, exist_ok=True)

TH = {"aas_hi": 60, "aas_lo": 40, "crs_hi": 60, "crs_lo": 40}   # keep in sync with the method scripts


def implied_label(row, sfx):
    """Label a node would get from one method even if that method did not pick it as a unit."""
    lab = row.get(f"label_{sfx}")
    if isinstance(lab, str):
        return lab
    gt = row.get(f"gap_type_{sfx}")
    if gt == "PARITY" or pd.isna(gt):
        return "NO GAP"
    a, c = row[f"AAS_{sfx}"], row[f"CRS_{sfx}"]
    if c >= TH["crs_hi"]:
        return "1P-CORE GAP"
    if c >= TH["crs_lo"]:
        return "REVIEW"
    if a >= TH["aas_hi"]:
        return "CURATE - " + gt
    if a >= TH["aas_lo"]:
        return "VERTICAL EXTENSION"
    return "OFF-BRAND"


def tier(lv, lg):
    cv, cg = lv.startswith("CURATE"), lg.startswith("CURATE")
    soft = ("REVIEW", "VERTICAL EXTENSION")
    if cv and cg:
        return "CONSENSUS APPROVE"
    if (cv and lg in soft) or (cg and lv in soft):
        return "LIKELY APPROVE"
    if lv == lg == "1P-CORE GAP":
        return "CONSENSUS 1P ROUTE"
    if lv == lg and lv in ("VERTICAL EXTENSION", "OFF-BRAND"):
        return "CONSENSUS HOLD"
    if "VERIFY" in (lv, lg) or (lv == "NO GAP") != (lg == "NO GAP"):
        return "VERIFY"
    return "SPLIT"


# %% [2] LOAD + ALIGN
def main():
    v = pd.read_csv(V_FILE)
    g = pd.read_csv(G_FILE)
    keys = ["competitor", "uid", "path_str", "level", "is_leaf", "comp_items"]
    m = v.merge(g, on=keys, how="outer", suffixes=("_v", "_g"))

    # ---- matching agreement (all nodes)
    same_top1 = (m["best_staples_v"] == m["best_staples_g"]).mean()
    pres_agree = (m["matched_v"] == m["matched_g"]).mean()
    gt_agree = (m["gap_type_v"] == m["gap_type_g"]).mean()
    corr = {k: m[[f"{k}_v", f"{k}_g"]].corr(method="spearman").iloc[0, 1] for k in ["DG", "AAS", "CRS", "coverage"]}

    # ---- recommendation units from either method
    units = m[m["label_v"].notna() | m["label_g"].notna()].copy()
    units["label_v_full"] = units.apply(lambda r: implied_label(r, "v"), axis=1)
    units["label_g_full"] = units.apply(lambda r: implied_label(r, "g"), axis=1)
    units["tier"] = [tier(a, b) for a, b in zip(units["label_v_full"], units["label_g_full"])]
    units["O_mean"] = units[["O_node_v", "O_node_g"]].mean(axis=1)
    units["unit_in"] = np.select([units.label_v.notna() & units.label_g.notna(), units.label_v.notna()],
                                 ["both", "V only"], "G only")
    order = ["CONSENSUS APPROVE", "LIKELY APPROVE", "CONSENSUS 1P ROUTE", "CONSENSUS HOLD", "VERIFY", "SPLIT"]
    units["tier"] = pd.Categorical(units["tier"], order, ordered=True)
    units = units.sort_values(["tier", "O_mean"], ascending=[True, False])

    # ---- de-duplicate nested units: if a parent and its child are both listed in the same
    # tier, keep the parent row and list the children as evidence
    keep, children = [], {}
    for _, r in units.iterrows():
        parent_hit = [k for k in keep if r["uid"].startswith(k + " > ") and units.loc[units.uid == k, "tier"].iloc[0] == r["tier"]]
        if parent_hit:
            children.setdefault(parent_hit[0], []).append(r["path_str"].split(" > ")[-1])
        else:
            keep.append(r["uid"])
    units["nested_children"] = units["uid"].map(lambda u: ", ".join(children.get(u, [])))
    top = units[units["uid"].isin(keep)]

    cols = ["tier", "path_str", "level", "comp_items", "label_v_full", "label_g_full", "unit_in",
            "DG_v", "DG_g", "AAS_v", "AAS_g", "CRS_v", "CRS_g", "O_node_v", "O_node_g", "O_mean",
            "best_staples_v", "best_staples_g", "nested_children"]
    agree_mx = pd.crosstab(units["label_v_full"], units["label_g_full"])
    summary = pd.DataFrame({
        "metric": ["Nodes compared", "Same top-1 Staples shelf", "Same confident-match decision",
                   "Same gap type (ENTER/DEEPEN/PARITY)", "Spearman DG", "Spearman AAS", "Spearman CRS",
                   "Spearman coverage", "Recommendation units (either method)"] +
                  [f"Tier: {t}" for t in order],
        "value": [len(m), round(same_top1, 3), round(pres_agree, 3), round(gt_agree, 3),
                  round(corr["DG"], 3), round(corr["AAS"], 3), round(corr["CRS"], 3), round(corr["coverage"], 3),
                  len(units)] + [int((top["tier"] == t).sum()) for t in order]})
    fp = os.path.join(OUT, "method_comparison.xlsx")
    with pd.ExcelWriter(fp, engine="openpyxl") as xw:
        summary.to_excel(xw, sheet_name="Summary", index=False)
        top[cols].round(2).to_excel(xw, sheet_name="Consensus_Units", index=False)
        units[cols].round(2).to_excel(xw, sheet_name="All_Units_Incl_Nested", index=False)
        agree_mx.to_excel(xw, sheet_name="Label_Agreement")
        for ws in xw.book.worksheets:
            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width = min(60, max(10, max(len(str(c.value or "")) for c in col[:200]) + 2))
    print(summary.to_string(index=False))
    print("\nTop consensus / likely approvals:")
    print(top[top["tier"].isin(order[:2])][["tier", "path_str", "comp_items", "label_v_full", "label_g_full",
                                            "O_mean"]].round(2).head(TOP_N).to_string(index=False))

    # ---- chart: approvals by tier, both methods' scores side by side
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    t2 = top[top["tier"].isin(order[:2])].head(TOP_N).iloc[::-1]
    if len(t2):
        ink, muted, grid = "#0b0b0b", "#52514e", "#e4e3df"
        fig, ax = plt.subplots(figsize=(9, 0.42 * len(t2) + 1.4))
        y = np.arange(len(t2))
        ax.barh(y + 0.19, t2["O_node_v"].fillna(0), height=0.34, color="#2a78d6", label="Method V (vector)")
        ax.barh(y - 0.19, t2["O_node_g"].fillna(0), height=0.34, color="#eb6834", label="Method G (graph)")
        ax.set_yticks(y)
        ax.set_yticklabels([" › ".join(p.split(" > ")[-2:]) + ("" if t.startswith("CONSENSUS") else "  (likely)")
                            for p, t in zip(t2["path_str"], t2["tier"])], fontsize=8, color=ink)
        ax.set_xlabel("Opportunity score O (same formula, each method's own measurements)", color=muted)
        ax.set_title("Approved by both methods (consensus) or one method (likely)", loc="left", fontsize=11, color=ink)
        ax.tick_params(colors=muted, labelsize=8)
        ax.tick_params(axis="y", labelcolor=ink)
        for sp in ["top", "right"]:
            ax.spines[sp].set_visible(False)
        for sp in ["left", "bottom"]:
            ax.spines[sp].set_color(grid)
        ax.legend(frameon=False, fontsize=8, loc="lower right")
        fig.tight_layout()
        fig.savefig(os.path.join(OUT, "consensus_top.png"), dpi=160)
        plt.close(fig)
    print(f"\nwrote {fp}")


if __name__ == "__main__":
    main()
