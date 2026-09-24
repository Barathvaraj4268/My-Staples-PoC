# -*- coding: utf-8 -*-
"""
framework_outputs.py - every table, figure and number used in the documents and the deck.
Called by run_framework.py (write_all). Figures are numbered; the documents cite them by
number, so a re-run with new data regenerates the whole evidence pack in one go.

 fig01 data availability          fig11 zone: OFF-BRAND
 fig02 pipeline flow              fig12 zone: VERIFY (findability)
 fig03 match quality (V, G)       fig13 themes x zones
 fig04 Staples coverage           fig14 peer consensus grid (vector DB)
 fig05 verification loop          fig15 vector map of the category space (vector DB)
 fig06 THE MATRIX (six zones)     fig16 graph view: where recommendations dock (graph DB)
 fig07 zone: CURATE               fig17 DEEPEN evidence by competitor
 fig08 zone: VERTICAL EXTENSION   fig18 stress test (500 runs)
 fig09 zone: REVIEW               fig19 method agreement (V vs G)
 fig10 zone: 1P-CORE GAP
"""
import os

import numpy as np
import pandas as pd

import poc_common as pc
from poc_common import CONFIG, FOCAL, INK, LABEL_COLOR, LABELS, STATUS_COLOR

OUT = "final"
ACTIONABLE = ["CURATE", "VERTICAL EXTENSION", "REVIEW", "1P-CORE GAP"]
ZONE_TITLE = {"CURATE": "CURATE - open to marketplace sellers now",
              "VERTICAL EXTENSION": "VERTICAL EXTENSION - phase 2, through a vertical Staples serves",
              "REVIEW": "REVIEW - merchant decision (sits next to a 1P line)",
              "1P-CORE GAP": "1P-CORE GAP - fix in 1P, keep out of the marketplace",
              "OFF-BRAND": "OFF-BRAND - real gap, wrong store: pass",
              "VERIFY": "VERIFY - flagged, but Staples already sells it (findability)"}
FIGNUM = {"CURATE": 8, "VERTICAL EXTENSION": 9, "REVIEW": 10, "1P-CORE GAP": 11, "OFF-BRAND": 12, "VERIFY": 13}
SHORT = {"OfficeDepot": "Office Depot", "WestElm": "West Elm", "Wayfair": "Wayfair", "Amazon": "Amazon",
         "Walmart": "Walmart", "Staples": "Staples"}


def fig_path(name):
    return pc.out_path(OUT, "figures", name)


def short(s, n=38):
    s = str(s)
    return s if len(s) <= n else s[:n - 1].rstrip() + "…"


def slug(lab):
    return lab.lower().replace(" ", "_").replace("-", "")


def unit_display(r):
    """Name as shown to Staples. When Staples' own shelf label is misleading ("Compasses" holds patio
    furniture), add what competitors actually stock there."""
    nm = str(r["name"])
    if " (competitors: " in nm:
        base, rest = nm.split(" (competitors: ", 1)
        return f"{base} (e.g. {rest.rstrip(')').split(', ')[0]})"
    return nm


def two_line(fam):
    fam = str(fam)
    if fam.startswith("Competitor dept: "):
        return "Competitors' " + short(fam.replace("Competitor dept: ", ""), 28)
    parts = fam.split(" > ")
    return f"{short(parts[-1], 30)}\n{short(parts[0], 30)}" if len(parts) > 1 else short(fam, 30)


def fam_name(f):
    return str(f).replace(" > ", " › ")


def recommendable(U):
    """Units that carry market consensus: every DEEPEN aisle (>= 2 competitors by construction), every
    ENTER concept seen at >= 2 competitors. Single-competitor
    concepts - including single-competitor VERIFY findings - stay in the unit tables as a watch list
    (evidence C) but do not drive aisle calls."""
    return U[(U.gap_type == "DEEPEN") | (U.k >= 2)]


def families(U):
    """Aisle-level (Staples L2) recommendation table: one row per zone x family (consensus units only)."""
    rows = []
    for (lab, fam), g in recommendable(U).groupby(["label", "family"]):
        g = g.sort_values("O", ascending=False)
        w = g["O"].clip(lower=0) + 0.01
        car = sorted({c for s in g["carriers"].dropna() for c in str(s).split(", ") if c})
        rows.append({"label": lab, "family": fam, "family_display": fam_name(fam), "n_units": len(g),
                     "n_enter": int((g.gap_type == "ENTER").sum()), "n_deepen": int((g.gap_type == "DEEPEN").sum()),
                     "O_sum": float(g.O.sum()), "O_max": float(g.O.max()),
                     "AAS": float(np.average(g.AAS, weights=w)), "CRS": float(np.average(g.CRS, weights=w)),
                     "k_max": int(g.k.max()), "competitors": ", ".join(SHORT.get(c, c) for c in car),
                     "n_competitors": len(car), "best_evidence": min(g.evidence),
                     "theme": g.theme.mode().iloc[0],
                     "examples": ", ".join(unit_display(r) for _, r in g.head(5).iterrows()),
                     "examples_list": [unit_display(r) for _, r in g.head(5).iterrows()],
                     "top_unit": unit_display(g.iloc[0])})
    F = pd.DataFrame(rows)
    F["rank_in_zone"] = F.groupby("label")["O_sum"].rank(ascending=False, method="first").astype(int)
    return F.sort_values(["label", "O_sum"], ascending=[True, False]).reset_index(drop=True)


# ---------------------------------------------------------------------------------------
def write_all(U, members, evid, m, nodes, focal, totals, cal_V, cal_G, gold, sf_items, sf_leaves, peer_edges,
              Vmap, files, theme_order, vstats, E1):
    plt = pc.set_style()
    U = U.copy()
    U["display"] = [unit_display(r) for _, r in U.iterrows()]
    S = {}                                                    # summary.json
    S.update(summary_basics(U, m, nodes, totals, cal_V, cal_G, gold, sf_items, sf_leaves, vstats))
    fig01_data(plt, nodes, totals, cal_V, gold, S)
    fig02_flow(plt, len(gold))
    fig03_quality(plt, cal_V, cal_G)
    fig04_coverage(plt, m, gold, S)
    fig05_verification(plt, gold, S)
    F = families(U)
    fig06_family_matrix(plt, F)
    fig07_unit_matrix(plt, U)
    for lab in LABELS:
        fig_zone(plt, F, U, lab)
        fig_zone_slide(plt, F, lab)
    fig14_themes(plt, U, theme_order, S)
    fig15_peers(plt, U, members, evid)
    fig16_vector_map(plt, U, focal, Vmap)
    fig17_graph(plt, U)
    fig18_deepen(plt, U, evid)
    fig19_sensitivity(plt, U, S)
    fig20_agreement(plt, U, S)
    S["zones"] = zone_summary(U)
    S["top_units"] = {lab: top_rows(U, lab, 15) for lab in LABELS}
    ev = evid.copy()
    ev["ratio"] = np.exp(ev["depth_gap"])
    dd = U[(U.gap_type == "DEEPEN") & U.label.isin(ACTIONABLE + ["VERTICAL EXTENSION"])]
    S["deepen_detail"] = [{"name": unit_display(r), "path": r.example_paths, "label": r.label,
                           "staples_items": float(r.get("staples_items", np.nan)) if pd.notna(r.get("staples_items", np.nan)) else None,
                           "staples_shelves": int(r.staples_shelves) if pd.notna(r.staples_shelves) else None,
                           "by_competitor": {c: {"ratio": round(float(x.ratio), 2), "basis": x.basis,
                                                 "competitor_size": float(x.competitor_size)}
                                             for c, x in ev[ev.staples_path == r.example_paths].set_index("competitor").iterrows()}}
                          for _, r in dd.sort_values("O", ascending=False).iterrows()]
    S["families"] = {lab: F[F.label == lab].head(12).round(3).to_dict("records") for lab in LABELS}
    S["family_counts"] = F.groupby("label").size().to_dict()
    S["recommendable_units_by_zone"] = recommendable(U).label.value_counts().to_dict()
    S["watchlist_single_competitor"] = int(((U.gap_type == "ENTER") & (U.k < 2)).sum())
    ob = F[F.label == "OFF-BRAND"]
    S["offbrand_groups"] = {"competitor_departments": int(ob.family.str.startswith("Competitor dept:").sum()),
                            "staples_aisles": int((~ob.family.str.startswith("Competitor dept:")).sum())}
    # verification coverage: model-flagged member shelves of actionable consensus ENTER units that were checked
    act = recommendable(U)
    act = act[(act.gap_type == "ENTER") & act.label.isin(ACTIONABLE + ["VERIFY"])]
    mm = members[members.unit_id.isin(act.unit_id) & (members.label_source != "peer search")]
    S["verification_coverage"] = {"member_shelves": int(len(mm)),
                                  "labelled": int((mm.label_source == "labelled").sum()),
                                  "unchecked": int((mm.label_source == "model").sum()),
                                  "share_labelled": float((mm.label_source == "labelled").mean()) if len(mm) else None}
    th = S.get("themes", {})
    core = th.get("Core office & business supplies", {}).get("by_zone", {})
    S["core_office_theme"] = {"total": int(sum(core.values())) if core else 0, "one_p": int(core.get("1P-CORE GAP", 0))}
    workbook(U, F, members, evid, m, nodes, totals, cal_V, cal_G, gold, S)
    pd.DataFrame(peer_edges, columns=["source", "target", "score"]).assign(rel="SIMILAR_TO").to_csv(
        pc.out_path(OUT, "graph", "peer_edges.csv"), index=False)
    pc.save_json(S, OUT, "summary.json")
    U.to_csv(pc.out_path(OUT, "units_all.csv"), index=False)
    pc.log(f"      figures: {len([x for x in os.listdir(os.path.dirname(fig_path('x'))) if x.endswith('.png')])} "
           f"(+ deck versions in figures/slides) | workbook + summary.json written")


# ---------------------------------------------------------------------------------------
def summary_basics(U, m, nodes, totals, cal_V, cal_G, gold, sf_items, sf_leaves, vstats):
    S = {"retailers": {}, "calibration": {}, "verification": {}, "size_factors": {"items": sf_items, "shelves": sf_leaves}}
    for rc in pc.RETAILERS:
        r = rc["name"]
        ins = nodes[(nodes.retailer == r) & nodes.in_scope]
        S["retailers"][r] = {"role": rc["role"], "count_mode": rc["count_mode"], "rows_in_file": None,
                             "shelves_in_scope": int(ins["unit_ok"].sum()), "leaf_shelves": totals[r]["leaves"],
                             "items": None if pd.isna(totals[r]["items"]) else float(totals[r]["items"]),
                             "max_depth": int(ins["depth"].max())}
    g_cal = gold[~gold.source.astype(str).str.contains("verification")]
    for _, r in cal_V.iterrows():
        c = r["competitor"]
        rg = cal_G[cal_G.competitor == c].iloc[0]
        rnd = g_cal[(g_cal.competitor == c) & g_cal.source.eq("stratified sample")]
        S["calibration"][c] = {"n_labelled": int((g_cal.competitor == c).sum()),
                               "tau_V": r["tau"], "auc_V": r["auc"], "top1_V": r["top1"],
                               "tau_G": rg["tau"], "auc_G": rg["auc"], "top1_G": rg["top1"],
                               "top1_G_wording_only": rg.get("top1_wording_only", np.nan),
                               "base_rate_carried": float(rnd["staples_carries"].mean()) if len(rnd) else None}
    ver = gold[gold.source.astype(str).str.contains("verification")]
    S["verification"] = {"checked": int(len(ver)), "carried_under_other_name": int(ver.staples_carries.sum()),
                         "confirmed_gaps": int((ver.staples_carries == 0).sum()), **vstats,
                         "by_competitor": ver.groupby("competitor").staples_carries.agg(["size", "sum"]).rename(
                             columns={"size": "checked", "sum": "carried"}).to_dict("index")}
    try:
        b = pd.read_csv(pc.out_path("V", "encoder_bakeoff.csv"))
        ok = b[b.status == "ok"]
        S["encoder_bakeoff"] = ok.groupby("encoder").apply(
            lambda d: float(np.average(d.top1_hybrid, weights=d.n_labelled))).round(3).to_dict()
        S["encoder_unavailable"] = b.loc[b.status != "ok", "encoder"].tolist()
        gr = pd.read_csv(pc.out_path("V", "context_weight_grid.csv"))
        S["context_grid_best"] = gr.sort_values("top1_hybrid_pooled", ascending=False).head(3).to_dict("records")
        sel = gr[(gr.w_self_competitor == CONFIG["w_self"]) & (gr.ctx_decay_competitor == CONFIG["ctx_decay"])
                 & (gr.w_self_staples == CONFIG["w_self_focal"])]
        S["context_grid_chosen"] = float(sel.top1_hybrid_pooled.iloc[0]) if len(sel) else None
        base = gr[(gr.w_self_competitor == 1.0) & (gr.w_self_staples == 1.0)]
        S["context_grid_label_only"] = float(base.top1_hybrid_pooled.iloc[0]) if len(base) else None
    except Exception as e:  # noqa
        pc.log(f"      (bake-off summary skipped: {e})")
    S["scope_audit"] = {r: g.set_index("drop_reason")["n"].to_dict() for r, g in
                        nodes[~nodes.in_scope].groupby(["retailer", "drop_reason"]).size().rename("n").reset_index()
                        .groupby("retailer")}
    S["rows_in_files"] = dict(pc.RAW_ROWS)
    S["nodes_after_row_filters"] = {r: int((nodes.retailer == r).sum()) for r in nodes.retailer.unique()}
    S["competitor_shelves_compared"] = int(len(m))
    S["labelled_rows"] = {"total": int(len(gold)), "calibration": int(len(g_cal)), "verification": int(len(ver)),
                          "office_depot_v1": int((g_cal.competitor == "OfficeDepot").sum())}
    ens = m[m.is_leaf].groupby("competitor")["ens"].value_counts(normalize=True).unstack(fill_value=0)
    S["ensemble_status_leaf_share"] = ens.round(4).to_dict("index")
    S["units"] = {"total": int(len(U)), "by_zone": U.label.value_counts().to_dict(),
                  "by_type": U.gap_type.value_counts().to_dict(),
                  "zone_by_type": pd.crosstab(U.label, U.gap_type).to_dict("index"),
                  "enter_concepts": int((U.gap_type == "ENTER").sum()),
                  "enter_k2plus": int(((U.gap_type == "ENTER") & (U.k >= 2) & (U.label != "VERIFY")).sum()),
                  "enter_k3plus": int(((U.gap_type == "ENTER") & (U.k >= 3) & (U.label != "VERIFY")).sum())}
    return S


def top_rows(U, lab, n):
    d = recommendable(U)
    d = d[d.label == lab].sort_values("O", ascending=False).head(n)
    cols = ["unit_id", "gap_type", "name", "family", "competitor_dept", "staples_items", "staples_shelves", "home_path", "nearest_staples_shelf", "k", "carriers", "AAS", "CRS", "GS",
            "PC", "O", "evidence", "p_approved", "p_top_n", "theme", "verify_outcome", "example_paths",
            "competitor_view", "median_depth_gap"]
    return [{k: (None if (isinstance(v, float) and np.isnan(v)) else (round(v, 3) if isinstance(v, float) else v))
             for k, v in r.items()} for r in d[[c for c in cols if c in d]].to_dict("records")]


def zone_summary(U):
    out = {}
    R = recommendable(U)
    for lab in LABELS:
        d = R[R.label == lab]
        out[lab] = {"n": int(len(d)), "enter": int((d.gap_type == "ENTER").sum()),
                    "deepen": int((d.gap_type == "DEEPEN").sum()),
                    "median_AAS": float(d.AAS.median()) if len(d) else None,
                    "median_CRS": float(d.CRS.median()) if len(d) else None,
                    "k3plus": int((d.k >= 3).sum()), "action": pc.ACTION[lab],
                    "all_units_incl_single_competitor": int((U.label == lab).sum())}
    return out


# ---------------------------------------------------------------------------------------
def fig01_data(plt, nodes, totals, cal_V, gold, S):
    g_cal = gold[~gold.source.astype(str).str.contains("verification")]
    rows = []
    for rc in pc.RETAILERS:
        r = rc["name"]
        info = S["retailers"][r]
        counts = {"terminal_only": "Full (leaf pages)", "cumulative": "Full (all pages)",
                  "partial": "Partial (L1-L2 only)", "none": "None - shelves only"}[rc["count_mode"]]
        cv = cal_V[cal_V.competitor == r]
        rows.append([SHORT[r], rc["role"].capitalize(), f"{info['shelves_in_scope']:,}", f"{info['leaf_shelves']:,}",
                     f"L{info['max_depth']}", counts,
                     "-" if r == FOCAL else f"{int((g_cal.competitor == r).sum())}",
                     "-" if r == FOCAL or cv.empty else f"{cv['top1'].iloc[0]:.0%}"])
    cols = ["Retailer", "Role", "Shelves\nin scope", "Leaf\nshelves", "Depth", "Item counts", "Labelled\nrows",
            "Top-1 shelf\nmatch (V)"]
    fig, ax = plt.subplots(figsize=(11, 2.6))
    ax.axis("off")
    t = ax.table(cellText=rows, colLabels=cols, loc="center", cellLoc="center", colLoc="center",
                 colWidths=[0.12, 0.1, 0.1, 0.09, 0.07, 0.2, 0.1, 0.12])
    t.auto_set_font_size(False)
    t.set_fontsize(9.5)
    t.scale(1, 1.7)
    for (i, j), c in t.get_celld().items():
        c.set_edgecolor(INK["grid"])
        if i == 0:
            c.set_facecolor("#f1f0ec")
            c.set_text_props(weight="bold", color=INK["primary"])
        else:
            txt = c.get_text().get_text()
            if j == 5:
                c.set_facecolor({"F": "#e8f1fb", "P": "#fdf3dc", "N": "#fbe7e7"}[txt[0]])
            if j == 0:
                c.set_text_props(weight="bold")
    ax.set_title("Figure 1 - Competitor panel and what each navigation tree gives us", pad=6)
    fig.text(0.02, 0.02, "Shelves in scope = category pages after the common department universe, merchandising pages, "
             "facets and headings are removed. Counts: 'Full' = item counts usable for depth; 'None' = presence and "
             "breadth only.", fontsize=8, color=INK["secondary"])
    fig.savefig(fig_path("fig01_data_availability.png"))
    plt.close(fig)


def fig02_flow(plt, n_labelled):
    from matplotlib.patches import FancyArrowPatch, FancyBboxPatch
    steps = [("1  Load & scope", "6 trees → one table\ncommon department\nuniverse; headings,\nfacets, merch out"),
             ("2  Represent", "label + whatever\npath the retailer\nshows (level-\nagnostic)"),
             ("3  Match", "V: vectors + Qdrant\nG: graph + flooding\ntwo independent\nlenses"),
             ("4  Calibrate\n& verify", f"{n_labelled} labelled rows\nper-competitor\nthresholds +\nverification loop"),
             ("5  Build units", "ENTER: new-category\nconcepts across 5\ncompetitors\nDEEPEN: thin aisles"),
             ("6  Score", "AAS adjacency\nCRS cannibalisation\nPC peers, GS size\nO = opportunity"),
             ("7  Decide", "six zones, rolled\nup to Staples L2\naisles; 500-run\nstress test")]
    fig, ax = plt.subplots(figsize=(14, 3.3))
    ax.set_xlim(0, 7)
    ax.set_ylim(0, 1)
    ax.axis("off")
    for i, (h, b) in enumerate(steps):
        x = i + 0.06
        col = "#e8f1fb" if i in (2, 3) else ("#e7f3e7" if i >= 4 else "#f1f0ec")
        ax.add_patch(FancyBboxPatch((x, 0.1), 0.86, 0.8, boxstyle="round,pad=0.01,rounding_size=0.05",
                                    fc=col, ec=INK["muted"], lw=0.8))
        ax.text(x + 0.43, 0.76, h, ha="center", va="center", fontsize=10, weight="bold", linespacing=1.1)
        ax.text(x + 0.43, 0.37, b, ha="center", va="center", fontsize=8.4, color=INK["secondary"], linespacing=1.3)
        if i < 6:
            ax.add_patch(FancyArrowPatch((x + 0.87, 0.5), (x + 1.05, 0.5), arrowstyle="-|>", mutation_scale=12,
                                         color=INK["muted"]))
    ax.set_title("Figure 2 - End-to-end flow (v2): every step is code in the delivered scripts", pad=4)
    fig.savefig(fig_path("fig02_flow.png"))
    plt.close(fig)


def fig03_quality(plt, cal_V, cal_G):
    comps = list(cal_V.competitor)
    x = np.arange(len(comps))
    fig, axes = plt.subplots(1, 2, figsize=(12, 3.8))
    for ax, key, title in [(axes[0], "top1", "Top-1 shelf match on labelled rows"),
                           (axes[1], "auc", "Carried-vs-missing separation (AUC)")]:
        v = cal_V.set_index("competitor").loc[comps, key].values
        g = cal_G.set_index("competitor").loc[comps, key].values
        ax.bar(x - 0.19, v, 0.36, color="#2a78d6", label="Method V (vectors)")
        ax.bar(x + 0.19, g, 0.36, color="#eb6834", label="Method G (graph)")
        for xi, a, b in zip(x, v, g):
            ax.text(xi - 0.19, a + 0.01, f"{a:.0%}" if key == "top1" else f"{a:.2f}", ha="center", fontsize=8)
            ax.text(xi + 0.19, b + 0.01, f"{b:.0%}" if key == "top1" else f"{b:.2f}", ha="center", fontsize=8)
        ax.set_xticks(x, [SHORT[c] for c in comps])
        ax.set_ylim(0 if key == "top1" else 0.5, 1.02)
        ax.set_title(title, fontsize=11)
        ax.grid(axis="y", color=INK["grid"], lw=0.6)
        ax.set_axisbelow(True)
    axes[0].legend(loc="upper right", fontsize=8.5)
    fig.suptitle("Figure 3 - Matching quality by competitor (calibrated on 695 labelled rows)", x=0.01, ha="left",
                 fontsize=12.5, weight="bold")
    fig.tight_layout()
    fig.savefig(fig_path("fig03_match_quality.png"))
    plt.close(fig)


def fig04_coverage(plt, m, gold, S):
    order = ["carried", "likely", "disputed", "gap_soft", "gap_hard"]
    names = {"carried": "Carried (both methods)", "likely": "Likely carried", "disputed": "Methods disagree",
             "gap_soft": "Possible gap", "gap_hard": "Gap (both methods / confirmed)"}
    lv = m[m.is_leaf]
    sh = lv.groupby("competitor")["ens"].value_counts(normalize=True).unstack(fill_value=0).reindex(columns=order,
                                                                                                    fill_value=0)
    comps = ["OfficeDepot", "WestElm", "Wayfair", "Amazon", "Walmart"]
    sh = sh.loc[comps]
    fig, ax = plt.subplots(figsize=(11, 3.6))
    left = np.zeros(len(comps))
    for st in order:
        ax.barh(range(len(comps)), sh[st].values, left=left, color=STATUS_COLOR[st], label=names[st],
                edgecolor="white", linewidth=1.5, height=0.62)
        left += sh[st].values
    for i, c in enumerate(comps):
        br = S["calibration"][c]["base_rate_carried"]
        ax.text(1.01, i, f"{br:.0%}", va="center", fontsize=9.5, weight="bold")
    ax.text(1.01, len(comps) - 0.35, "labelled\nsample", fontsize=8, color=INK["secondary"], va="bottom")
    ax.set_yticks(range(len(comps)), [SHORT[c] for c in comps])
    ax.invert_yaxis()
    ax.set_xlim(0, 1)
    ax.xaxis.set_major_formatter(plt.FuncFormatter(lambda v, _: f"{v:.0%}"))
    ax.legend(ncol=5, loc="lower center", bbox_to_anchor=(0.5, -0.33), fontsize=8.5)
    ax.set_title("Figure 4 - How much of each competitor's shelf list Staples already covers\n(bars: both methods' "
                 "verdict on every leaf shelf, labelled rows override; right: share carried in a random labelled "
                 "sample)", fontsize=11)
    fig.savefig(fig_path("fig04_staples_coverage.png"))
    plt.close(fig)


def fig05_verification(plt, gold, S):
    ver = gold[gold.source.astype(str).str.contains("verification")]
    t = ver.groupby("competitor").staples_carries.agg(["size", "sum"])
    t["gap"] = t["size"] - t["sum"]
    t = t.sort_values("size")
    fig, ax = plt.subplots(figsize=(10, 3.2))
    y = np.arange(len(t))
    ax.barh(y, t["sum"], color=STATUS_COLOR["carried"], height=0.6, label="Already carried under another name",
            edgecolor="white", linewidth=1.5)
    ax.barh(y, t["gap"], left=t["sum"], color=STATUS_COLOR["gap_hard"], height=0.6, label="Confirmed gap",
            edgecolor="white", linewidth=1.5)
    for yi, (a, b) in enumerate(zip(t["sum"], t["gap"])):
        ax.text(a + b + 0.8, yi, f"{a} carried / {b} gaps", va="center", fontsize=8.5, color=INK["secondary"])
    ax.set_yticks(y, [SHORT[c] for c in t.index])
    ax.set_xlim(0, t["size"].max() * 1.35)
    ax.legend(loc="lower right", fontsize=8.5)
    v = S["verification"]
    ax.set_title(f"Figure 5 - Verification loop: {v['checked']} model-flagged gaps checked - "
                 f"{v['carried_under_other_name']} were already on Staples.com, {v['confirmed_gaps']} are real")
    fig.savefig(fig_path("fig05_verification_loop.png"))
    plt.close(fig)


# ---------------------------------------------------------------------------------------
ZONES = [("1P-CORE GAP", 60, 100, 0, 100), ("REVIEW", 40, 60, 0, 100), ("CURATE", 0, 40, 60, 100),
         ("VERTICAL EXTENSION", 0, 40, 40, 60), ("OFF-BRAND", 0, 40, 0, 40)]


def draw_zones(ax, alpha=0.09, labels=True):
    from matplotlib.patches import Rectangle
    for lab, x0, x1, y0, y1 in ZONES:
        ax.add_patch(Rectangle((x0, y0), x1 - x0, y1 - y0, fc=LABEL_COLOR[lab], alpha=alpha, lw=0, zorder=0))
        if labels:
            right = lab in ("VERTICAL EXTENSION", "OFF-BRAND")
            ax.text(x1 - 1.2 if right else x0 + 1.2, y1 - 1.5, lab, fontsize=9.5, weight="bold", va="top",
                    ha="right" if right else "left", color=INK["primary"], zorder=5, alpha=0.85)
    for x in (40, 60):
        ax.axvline(x, color=INK["muted"], lw=0.8, ls=(0, (3, 3)), zorder=1)
    for y in (40, 60):
        ax.plot([0, 40], [y, y], color=INK["muted"], lw=0.8, ls=(0, (3, 3)), zorder=1)


def fig06_family_matrix(plt, F):
    from adjustText import adjust_text
    fig, ax = plt.subplots(figsize=(12, 7.6))
    draw_zones(ax)
    ax.set_xlim(-2, 100)
    ax.set_ylim(-2, 102)
    size = lambda o: 40 + 330 * np.sqrt(np.clip(o, 0, None) / max(F.O_sum.max(), 1e-9))
    for lab in ["OFF-BRAND", "1P-CORE GAP", "REVIEW", "VERTICAL EXTENSION", "CURATE"]:
        d = F[F.label == lab]
        ax.scatter(d.CRS.clip(0, 99), d.AAS.clip(0, 100), s=size(d.O_sum), c=LABEL_COLOR[lab],
                   alpha=0.35 if lab == "OFF-BRAND" else 0.82, edgecolors="white", linewidths=1.0, zorder=3)
    v = F[F.label == "VERIFY"]
    ax.scatter(v.CRS.clip(0, 99), v.AAS.clip(0, 100), s=size(v.O_sum), facecolors="none",
               edgecolors=LABEL_COLOR["VERIFY"], linewidths=1.3, alpha=0.85, zorder=3)
    texts = []
    for lab, n in (("CURATE", 8), ("VERTICAL EXTENSION", 3), ("REVIEW", 3), ("1P-CORE GAP", 3), ("OFF-BRAND", 3)):
        for _, r in F[F.label == lab].head(n).iterrows():
            nm = r.family_display.replace("Competitor dept: ", "")
            texts.append(ax.text(min(r.CRS, 99), r.AAS, short(nm, 34), fontsize=8.2, zorder=6,
                                 color=INK["primary"] if lab != "OFF-BRAND" else INK["secondary"]))
    adjust_text(texts, ax=ax, arrowprops=dict(arrowstyle="-", color=INK["muted"], lw=0.6),
                expand=(1.15, 1.4), force_text=(0.3, 0.5))
    ax.set_xlabel("Cannibalisation risk to Staples 1P (CRS, 0-100)  →")
    ax.set_ylabel("Adjacency to what Staples already sells (AAS, 0-100)  →")
    ax.text(62, 3, "Dot = one Staples aisle (L1 › L2) with its sub-categories in that zone;\nsize = total opportunity "
            "(ΣO); only sub-categories seen at 2+ competitors. Hollow violet = VERIFY:\nflagged, but already on "
            "Staples.com under another name. OFF-BRAND grouped by competitor department.", fontsize=8.2, color=INK["secondary"], va="bottom", zorder=5,
            bbox=dict(fc="white", ec=INK["grid"], lw=0.6, pad=4))
    n = F.label.value_counts()
    ax.set_title("Figure 6 - The decision matrix at category level: Staples aisles (L2) by zone\n"
                 f"CURATE {n.get('CURATE', 0)} · VERTICAL EXT. {n.get('VERTICAL EXTENSION', 0)} · REVIEW "
                 f"{n.get('REVIEW', 0)} · 1P-CORE GAP {n.get('1P-CORE GAP', 0)} · OFF-BRAND {n.get('OFF-BRAND', 0)} "
                 f"departments · VERIFY {n.get('VERIFY', 0)} aisles", fontsize=12)
    fig.savefig(fig_path("fig06_matrix_aisles.png"))
    plt.close(fig)


def fig07_unit_matrix(plt, U):
    from adjustText import adjust_text
    fig, ax = plt.subplots(figsize=(12, 7.6))
    draw_zones(ax)
    ax.set_xlim(-3, 100)
    ax.set_ylim(-2, 102)
    size = lambda o: 16 + 230 * np.clip(o, 0, 1) ** 1.6
    rj = np.random.default_rng(7)
    jit = lambda c: np.where(c < 1.0, c + rj.uniform(0, 3.2, len(c)), c)   # spread the CRS = 0 pile for legibility
    for lab in ["OFF-BRAND", "1P-CORE GAP", "REVIEW", "VERTICAL EXTENSION", "CURATE"]:
        d = U[U.label == lab]
        for gt, mk in (("ENTER", "o"), ("DEEPEN", "s")):
            e = d[d.gap_type == gt]
            ax.scatter(jit(e.CRS.clip(0, 99).values), e.AAS.clip(0, 100), s=size(e.O), marker=mk, c=LABEL_COLOR[lab],
                       alpha=0.35 if lab == "OFF-BRAND" else 0.8, edgecolors="white", linewidths=0.8, zorder=3)
    v = U[U.label == "VERIFY"]
    ax.scatter(jit(v.CRS.clip(0, 99).values), v.AAS.clip(0, 100), s=size(v.O) * 0.8, marker="o", facecolors="none",
               edgecolors=LABEL_COLOR["VERIFY"], linewidths=1.1, alpha=0.8, zorder=3)
    texts = []
    for lab, n in (("CURATE", 6), ("VERTICAL EXTENSION", 3), ("REVIEW", 3), ("1P-CORE GAP", 3)):
        for _, r in U[U.label == lab].sort_values("O", ascending=False).head(n).iterrows():
            texts.append(ax.text(min(r.CRS, 99), r.AAS, short(unit_display(r), 28), fontsize=8, zorder=6))
    adjust_text(texts, ax=ax, arrowprops=dict(arrowstyle="-", color=INK["muted"], lw=0.6), expand=(1.2, 1.4))
    ax.set_xlabel("Cannibalisation risk to Staples 1P (CRS, 0-100)  →")
    ax.set_ylabel("Adjacency to what Staples already sells (AAS, 0-100)  →")
    from matplotlib.lines import Line2D
    h = [Line2D([], [], marker="o", ls="", color=INK["secondary"], label="New category (ENTER)", ms=7),
         Line2D([], [], marker="s", ls="", color=INK["secondary"], label="Deepen a Staples aisle (DEEPEN)", ms=7),
         Line2D([], [], marker="o", ls="", mfc="none", mec=LABEL_COLOR["VERIFY"], label="VERIFY (already sold)", ms=7)]
    ax.legend(handles=h, loc="lower right", fontsize=8.5, bbox_to_anchor=(1.0, 0.02), frameon=True,
              facecolor="white", edgecolor=INK["grid"])
    n = U.label.value_counts()
    ax.set_title(f"Figure 7 - Detail: all {len(U)} sub-category units (L2-L3) behind Figure 6 (size = opportunity O)\n"
                 f"CURATE {n.get('CURATE', 0)} · VERTICAL EXT. {n.get('VERTICAL EXTENSION', 0)} · REVIEW "
                 f"{n.get('REVIEW', 0)} · 1P-CORE GAP {n.get('1P-CORE GAP', 0)} · OFF-BRAND {n.get('OFF-BRAND', 0)} · "
                 f"VERIFY {n.get('VERIFY', 0)}", fontsize=12)
    fig.savefig(fig_path("fig07_matrix_units.png"))
    plt.close(fig)


def fig_zone(plt, F, U, lab, n=10):
    d = F[F.label == lab].head(n)
    if lab == "VERIFY":
        d = F[F.label == lab].sort_values(["n_units", "O_sum"], ascending=False).head(n)
    if d.empty:
        return
    fig, ax = plt.subplots(figsize=(12, 0.55 * len(d) + 1.4))
    y = np.arange(len(d))
    col = LABEL_COLOR[lab]
    xv = d.n_units.values if lab == "VERIFY" else d.O_sum.values
    ax.barh(y, xv, color=col, height=0.64, alpha=0.9)
    for yi, (_, r) in enumerate(d.iterrows()):
        mix = []
        if r.n_enter:
            mix.append(f"{r.n_enter} new")
        if r.n_deepen:
            mix.append(f"{r.n_deepen} deepen")
        if lab == "VERIFY":
            txt = f"e.g. {short(r.examples, 80)}"
        else:
            txt = (f"{' + '.join(mix)} · AAS {r.AAS:.0f} · CRS {r.CRS:.0f} · peers {r.k_max}/5  |  e.g. "
                   f"{short(r.examples, 62)}")
        ax.text(xv[yi] + xv.max() * 0.012, yi, txt, va="center", fontsize=8.2, color=INK["secondary"])
    ax.set_yticks(y, [short(f.replace("Competitor dept: ", "Competitors' "), 44) for f in d.family_display],
                  fontsize=9)
    ax.invert_yaxis()
    ax.set_xlim(0, xv.max() * 2.35)
    ax.set_xlabel("Number of sub-categories found already on Staples.com" if lab == "VERIFY"
                  else "Total opportunity of the aisle's sub-categories (ΣO)")
    ax.grid(axis="x", color=INK["grid"], lw=0.6)
    ax.set_axisbelow(True)
    tot = int((F.label == lab).sum())
    what = "departments" if lab == "OFF-BRAND" else "aisles"
    nrec = int((recommendable(U).label == lab).sum())
    ax.set_title(f"Figure {FIGNUM[lab]} - {ZONE_TITLE[lab]}  (top {len(d)} of {tot} {what}; "
                 f"{nrec} sub-categories with 2+ competitor evidence)", fontsize=11.2)
    fig.savefig(fig_path(f"fig{FIGNUM[lab]:02d}_zone_{slug(lab)}.png"))
    plt.close(fig)


def fig_zone_slide(plt, F, lab, n=8):
    """Compact version of the zone chart for the deck (same data, larger type, fewer rows)."""
    d = F[F.label == lab].head(n)
    if lab == "VERIFY":
        d = F[F.label == lab].sort_values(["n_units", "O_sum"], ascending=False).head(n)
    if d.empty:
        return
    fig, ax = plt.subplots(figsize=(10, 5.6))
    y = np.arange(len(d))
    xv = d.n_units.values if lab == "VERIFY" else d.O_sum.values
    ax.barh(y, xv, color=LABEL_COLOR[lab], height=0.62, alpha=0.9)
    for yi, (_, r) in enumerate(d.iterrows()):
        ex = short(r.examples, 52)
        sc = (f"{r.n_units} found" if lab == "VERIFY" else f"AAS {r.AAS:.0f} · CRS {r.CRS:.0f} · peers {r.k_max}/5")
        ax.text(xv[yi] + xv.max() * 0.015, yi - 0.13, ex, va="center", fontsize=10, color=INK["primary"])
        ax.text(xv[yi] + xv.max() * 0.015, yi + 0.22, sc, va="center", fontsize=9, color=INK["secondary"])
    ax.set_yticks(y, [two_line(f) for f in d.family], fontsize=10.5)
    ax.invert_yaxis()
    ax.set_xlim(0, xv.max() * 2.3)
    ax.set_xlabel("Sub-categories found on Staples.com" if lab == "VERIFY" else "Total opportunity ΣO", fontsize=10)
    ax.tick_params(axis="x", labelsize=9)
    ax.grid(axis="x", color=INK["grid"], lw=0.6)
    ax.set_axisbelow(True)
    fig.savefig(pc.out_path(OUT, "figures", "slides", f"slide_zone_{slug(lab)}.png"))
    plt.close(fig)


def fig14_themes(plt, U, theme_order, S):
    U = recommendable(U)
    t = pd.crosstab(U.theme, U.label).reindex(index=[x for x in theme_order if x in set(U.theme)],
                                               columns=LABELS, fill_value=0)
    o = U[U.label.isin(["CURATE", "VERTICAL EXTENSION"])].groupby("theme").O.sum().reindex(t.index).fillna(0)
    fig, ax = plt.subplots(figsize=(11.5, 0.5 * len(t) + 1.6))
    M = t.values.astype(float)
    ax.imshow(np.log1p(M), cmap="Blues", aspect="auto", vmin=0, vmax=np.log1p(M.max()) * 1.15)
    for i in range(M.shape[0]):
        for j in range(M.shape[1]):
            if M[i, j]:
                ax.text(j, i, int(M[i, j]), ha="center", va="center", fontsize=9.5,
                        color="white" if np.log1p(M[i, j]) > np.log1p(M.max()) * 0.6 else INK["primary"])
    ax.set_xticks(range(len(LABELS)), [x.replace("VERTICAL EXTENSION", "VERTICAL\nEXTENSION").replace(
        "1P-CORE GAP", "1P-CORE\nGAP") for x in LABELS], fontsize=9)
    ax.set_yticks(range(len(t)), t.index, fontsize=9.5)
    ax.tick_params(length=0)
    for sp in ax.spines.values():
        sp.set_visible(False)
    for i, v in enumerate(o.values):
        ax.text(len(LABELS) - 0.4, i, f"  ΣO {v:.1f}", va="center", fontsize=8.5, color=INK["secondary"])
    ax.set_title("Figure 14 - Strategic themes x zones (sub-categories with 2+ competitor evidence; right: summed "
                 "opportunity of CURATE + VERTICAL EXTENSION)", fontsize=11)
    fig.savefig(fig_path("fig14_themes.png"))
    plt.close(fig)
    S["themes"] = {th: {"by_zone": {k: int(v) for k, v in t.loc[th].items()}, "sum_O_curate_ve": float(o.loc[th])}
                   for th in t.index}


def fig15_peers(plt, U, members, evid):
    d = pd.concat([U[U.label == lab].sort_values("O", ascending=False).head(n)
                   for lab, n in (("CURATE", 12), ("VERTICAL EXTENSION", 5), ("REVIEW", 3))])
    comps = ["OfficeDepot", "WestElm", "Wayfair", "Amazon", "Walmart"]
    fig, ax = plt.subplots(figsize=(10.5, 0.38 * len(d) + 1.6))
    ev = evid.set_index(["staples_path", "competitor"])
    for i, (_, r) in enumerate(d.iterrows()):
        ax.add_patch(plt.Rectangle((-0.95, i - 0.3), 0.18, 0.6, color=LABEL_COLOR[r.label]))
        for j, c in enumerate(comps):
            if r.gap_type == "ENTER":
                mm = members[(members.unit_id == r.unit_id) & (members.competitor == c)]
                if len(mm):
                    own = (mm.label_source != "peer search").any()
                    ax.scatter(j, i, s=90, marker="o", c=INK["primary"] if own else "white",
                               edgecolors=INK["primary"], linewidths=1.2)
            else:
                if (r.example_paths, c) in ev.index:
                    e = ev.loc[(r.example_paths, c)]
                    e = e.iloc[0] if isinstance(e, pd.DataFrame) else e
                    deeper = e["depth_gap"] >= CONFIG["deepen_min"]
                    ax.scatter(j, i, s=85, marker="s", c=INK["primary"] if deeper else "white",
                               edgecolors=INK["primary"], linewidths=1.2)
                    ax.text(j + 0.2, i, f"{np.exp(e['depth_gap']):.1f}x", fontsize=7, va="center",
                            color=INK["secondary"])
    ax.set_xticks(range(len(comps)), [SHORT[c] for c in comps])
    ax.xaxis.tick_top()
    ax.set_yticks(range(len(d)), [short(n.split(" (competitors")[0], 40) for n in d["name"]], fontsize=8.8)
    ax.set_xlim(-1.1, len(comps) - 0.4)
    ax.set_ylim(len(d) - 0.5, -0.5)
    for sp in ax.spines.values():
        sp.set_visible(False)
    ax.tick_params(length=0)
    ax.grid(color=INK["grid"], lw=0.5)
    ax.set_axisbelow(True)
    ax.set_title("Figure 15 - Peer evidence behind the top recommendations\n"
                 "● gap at that competitor (it sells it, Staples does not) · ○ found in that competitor's Qdrant "
                 "collection\n■ competitor ≥2x Staples' relative depth in the aisle · □ carries it, not 2x deeper",
                 fontsize=10.5, pad=24)
    fig.savefig(fig_path("fig15_peer_consensus.png"))
    plt.close(fig)


def fig16_vector_map(plt, U, focal, Vmap):
    from sklearn.decomposition import PCA
    from sklearn.manifold import TSNE
    st = focal[(focal.depth.isin([2, 3])) & focal.uid.isin(Vmap.keys())]
    uu = U[U.centroid_uid.isin(list(Vmap.keys()))].copy()
    X = np.vstack([np.stack([Vmap[u] for u in st.uid]), np.stack([Vmap[u] for u in uu.centroid_uid])])
    Z = PCA(n_components=50, random_state=0).fit_transform(X)
    Y = TSNE(n_components=2, perplexity=35, init="pca", random_state=42, learning_rate="auto").fit_transform(Z)
    ys, yu = Y[:len(st)], Y[len(st):]
    fig, ax = plt.subplots(figsize=(12, 8))
    ax.scatter(ys[:, 0], ys[:, 1], s=9, c="#c9c8c3", alpha=0.7, lw=0, label="Staples aisle / shelf (L2-L3)")
    l1 = st.path.map(lambda p: p[0]).values
    big = pd.Series(l1).value_counts().head(14).index
    for name in big:
        pts = ys[l1 == name]
        cx, cy = np.median(pts, axis=0)
        ax.text(cx, cy, name, fontsize=7.8, color=INK["secondary"], ha="center", style="italic", alpha=0.9)
    for lab in ["OFF-BRAND", "VERIFY", "1P-CORE GAP", "REVIEW", "VERTICAL EXTENSION", "CURATE"]:
        mk = uu.label.values == lab
        if not mk.any():
            continue
        e = uu[mk]
        ax.scatter(yu[mk, 0], yu[mk, 1], s=14 + 120 * e.O.clip(0, 1) ** 1.5,
                   c="none" if lab == "VERIFY" else LABEL_COLOR[lab],
                   edgecolors=LABEL_COLOR[lab] if lab == "VERIFY" else "white", lw=0.9 if lab == "VERIFY" else 0.5,
                   alpha=0.45 if lab == "OFF-BRAND" else 0.85,
                   marker="o", label=lab.title().replace("1P-Core Gap", "1P-core gap"))
    from adjustText import adjust_text
    texts = []
    top = uu[uu.label.isin(["CURATE", "VERTICAL EXTENSION"])].sort_values("O", ascending=False).head(12)
    for i, r in top.iterrows():
        j = uu.index.get_loc(i)
        texts.append(ax.text(yu[j, 0], yu[j, 1], short(r["name"].split(" (competitors")[0], 26), fontsize=8))
    off = uu[uu.label == "OFF-BRAND"].sort_values("O", ascending=False).head(4)
    for i, r in off.iterrows():
        j = uu.index.get_loc(i)
        texts.append(ax.text(yu[j, 0], yu[j, 1], short(r["name"], 22), fontsize=7.5, color=INK["secondary"]))
    adjust_text(texts, ax=ax, arrowprops=dict(arrowstyle="-", color=INK["muted"], lw=0.5))
    ax.set_xticks([])
    ax.set_yticks([])
    for sp in ax.spines.values():
        sp.set_visible(False)
    ax.legend(loc="lower left", fontsize=8.5, markerscale=1.2, ncol=2)
    ax.set_title("Figure 16 - Vector-DB view: the category space (t-SNE of the Qdrant vectors). Grey = Staples' own "
                 "shelves;\ncoloured = recommendation units. CURATE sits on the edge of Staples' territory; "
                 "OFF-BRAND forms islands far from it.", fontsize=11)
    fig.savefig(fig_path("fig16_vector_map.png"))
    plt.close(fig)


def fig17_graph(plt, U):
    """Tripartite graph: recommendation -> the Staples aisles (L2) its shelves attach to in the
    knowledge graph (Method G's nearest shelves) -> Staples department (L1)."""
    from matplotlib.patches import FancyArrowPatch
    d = pd.concat([U[U.label == lab].sort_values("O", ascending=False).head(n)
                   for lab, n in (("CURATE", 9), ("VERTICAL EXTENSION", 3), ("REVIEW", 3))])
    edges = [(r.unit_id, r.aisle, 1) for _, r in d.iterrows()]
    aisles = sorted({a for _, a, _ in edges}, key=lambda a: (a.split(" > ")[0], a))
    l1s = sorted({a.split(" > ")[0] for a in aisles})
    yU = {u: i for i, u in enumerate(d.unit_id)}
    nU, nA, nL = len(d), len(aisles), len(l1s)
    yA = {a: (i + 0.5) * nU / max(nA, 1) - 0.5 for i, a in enumerate(aisles)}
    yL = {l: (i + 0.5) * nU / max(nL, 1) - 0.5 for i, l in enumerate(l1s)}
    fig, ax = plt.subplots(figsize=(13, 0.5 * nU + 1.8))
    lab_of = dict(zip(d.unit_id, d.label))
    for u, a, w in edges:
        ax.add_patch(FancyArrowPatch((0.02, yU[u]), (0.98, yA[a]), connectionstyle="arc3,rad=0.0",
                                     arrowstyle="-", color=LABEL_COLOR[lab_of[u]], lw=0.8 + 0.6 * min(w, 4),
                                     alpha=0.55))
    for a in aisles:
        l1 = a.split(" > ")[0]
        ax.add_patch(FancyArrowPatch((1.02, yA[a]), (1.98, yL[l1]), arrowstyle="-", color=INK["muted"], lw=0.9,
                                     alpha=0.6))
    for _, r in d.iterrows():
        ax.scatter(0, yU[r.unit_id], s=120, c=LABEL_COLOR[r.label], marker="s" if r.gap_type == "DEEPEN" else "o",
                   zorder=3, edgecolors="white")
        ax.text(-0.04, yU[r.unit_id], short(r["name"].split(" (competitors")[0], 32), ha="right", va="center",
                fontsize=8.6)
    for a in aisles:
        ax.scatter(1, yA[a], s=60, c="#c9c8c3", zorder=3, edgecolors=INK["muted"])
        ax.text(1, yA[a] - 0.28, short(a.split(" > ", 1)[-1], 34), ha="center", va="bottom", fontsize=7.6,
                color=INK["secondary"])
    for l in l1s:
        ax.scatter(2, yL[l], s=150, c=INK["secondary"], zorder=3)
        ax.text(2.05, yL[l], l, ha="left", va="center", fontsize=9, weight="bold")
    ax.text(0, -1.3, "Recommendation", ha="center", fontsize=9.5, weight="bold", color=INK["secondary"])
    ax.text(1, -1.3, "Staples aisle it attaches to (graph)", ha="center", fontsize=9.5, weight="bold",
            color=INK["secondary"])
    ax.text(2, -1.3, "Staples department", ha="center", fontsize=9.5, weight="bold", color=INK["secondary"])
    ax.set_xlim(-1.1, 2.9)
    ax.set_ylim(nU - 0.3, -1.8)
    ax.axis("off")
    ax.set_title("Figure 17 - Graph-DB view: where each top recommendation plugs into Staples' existing assortment\n"
                 "(path in the knowledge graph: competitor shelf -CHILD_OF-> its aisle -SAME_AS-> Staples aisle; "
                 "colour = zone)",
                 fontsize=11)
    fig.savefig(fig_path("fig17_graph_dock.png"))
    plt.close(fig)


def fig18_deepen(plt, U, evid):
    d = U[(U.gap_type == "DEEPEN") & U.label.isin(ACTIONABLE)].sort_values("O", ascending=False).head(18)
    comps = ["OfficeDepot", "WestElm", "Wayfair", "Amazon", "Walmart"]
    ev = evid.set_index(["staples_path", "competitor"])
    M = np.full((len(d), len(comps)), np.nan)
    B = np.full((len(d), len(comps)), "", dtype=object)
    for i, p in enumerate(d.example_paths):
        for j, c in enumerate(comps):
            if (p, c) in ev.index:
                e = ev.loc[(p, c)]
                e = e.iloc[0] if isinstance(e, pd.DataFrame) else e
                M[i, j] = np.exp(e["depth_gap"])
                B[i, j] = "i" if e["basis"] == "items" else ""
    fig, ax = plt.subplots(figsize=(11, 0.42 * len(d) + 1.8))
    ax.imshow(np.log(np.clip(np.nan_to_num(M, nan=1.0), 0.25, 16)), cmap="RdBu_r", vmin=-np.log(16), vmax=np.log(16),
              aspect="auto")
    for i in range(M.shape[0]):
        for j in range(M.shape[1]):
            if not np.isnan(M[i, j]):
                ax.text(j, i, f"{M[i, j]:.1f}x{'ⁱ' if B[i, j] else ''}", ha="center", va="center", fontsize=8.5,
                        color="white" if abs(np.log(M[i, j])) > 1.7 else INK["primary"])
            else:
                ax.text(j, i, "–", ha="center", va="center", fontsize=8.5, color=INK["muted"])
    ax.set_xticks(range(len(comps)), [SHORT[c] for c in comps])
    ax.set_yticks(range(len(d)), [f"{short(n.split(' (competitors')[0], 30)}  [{l}]" for n, l in zip(d["name"], d.label)],
                  fontsize=8.6)
    ax.tick_params(length=0)
    for sp in ax.spines.values():
        sp.set_visible(False)
    ax.set_title("Figure 18 - DEEPEN evidence: each competitor's depth in a Staples aisle relative to Staples "
                 "(size-factor adjusted;\n≥2x = deeper; ⁱ = item counts, otherwise number of shelves)", fontsize=11)
    fig.savefig(fig_path("fig18_deepen_evidence.png"))
    plt.close(fig)


def fig19_sensitivity(plt, U, S):
    R = recommendable(U)
    d = R[R.label == "CURATE"].sort_values("O", ascending=False).head(20)
    fig, ax = plt.subplots(figsize=(10.5, 0.36 * len(d) + 1.5))
    y = np.arange(len(d))
    ax.hlines(y, 0, 1, color=INK["grid"], lw=0.8)
    ax.scatter(d.p_approved, y, s=60, c=LABEL_COLOR["CURATE"], label="stays CURATE", zorder=3)
    ax.scatter(d.p_top_n, y, s=60, marker="D", c="white", edgecolors=INK["primary"], label=f"stays in top {CONFIG['top_n']}",
               zorder=3)
    ax.set_yticks(y, [short(n.split(" (competitors")[0], 36) for n in d["name"]], fontsize=8.8)
    ax.invert_yaxis()
    ax.set_xlim(-0.02, 1.02)
    ax.xaxis.set_major_formatter(plt.FuncFormatter(lambda v, _: f"{v:.0%}"))
    ax.axvline(0.8, color=INK["muted"], ls=(0, (3, 3)), lw=0.8)
    ax.legend(loc="lower left", fontsize=8.5, ncol=2)
    ax.set_title("Figure 19 - Stress test: share of 500 runs (weights re-drawn, zone thresholds ±5, penalty 0.5-2x)\n"
                 "in which each top CURATE sub-category (2+ competitors) keeps its zone and its top-20 place", fontsize=11)
    fig.savefig(fig_path("fig19_sensitivity.png"))
    plt.close(fig)
    S["sensitivity"] = {"top20_curate_p_approved_ge_0.8": int((d.p_approved >= 0.8).sum()),
                        "top20_curate_p_top_n_ge_0.8": int((d.p_top_n >= 0.8).sum()), "n": int(len(d))}


def fig20_agreement(plt, U, S):
    from scipy.stats import spearmanr
    e = U[(U.gap_type == "ENTER") & U.AAS_V.notna() & (U.label != "VERIFY")]
    fig, axes = plt.subplots(1, 2, figsize=(11, 4.6))
    out = {}
    for ax, a, b, t in [(axes[0], "AAS_V", "AAS_G", "Adjacency (AAS)"), (axes[1], "CRS_V", "CRS_G", "Cannibalisation (CRS)")]:
        rho = spearmanr(e[a], e[b]).correlation
        out[t] = float(rho)
        ax.scatter(e[a], e[b], s=14, c=[LABEL_COLOR[l] for l in e.label], alpha=0.6, lw=0)
        ax.plot([0, 100], [0, 100], color=INK["muted"], lw=0.8, ls=(0, (3, 3)))
        ax.set_xlim(0, 100)
        ax.set_ylim(0, 100)
        ax.set_xlabel(f"Method V (vectors) {t.split(' ')[-1]}")
        ax.set_ylabel(f"Method G (graph) {t.split(' ')[-1]}")
        ax.set_title(f"{t}: Spearman ρ = {rho:.2f}", fontsize=11)
    agree = float((e.label_V == e.label_G).mean())
    out["same_zone_share"] = agree
    fig.suptitle(f"Figure 20 - Two independent methods, one answer: V and G place {agree:.0%} of new-category units "
                 f"in the same zone (colour = final zone)", x=0.01, ha="left", fontsize=12, weight="bold")
    fig.tight_layout()
    fig.savefig(fig_path("fig20_method_agreement.png"))
    plt.close(fig)
    S["method_agreement"] = out


# ---------------------------------------------------------------------------------------
def workbook(U, F, members, evid, m, nodes, totals, cal_V, cal_G, gold, S):
    fp = pc.out_path(OUT, "Category_Recommendations_v2.xlsx")
    cols = ["rank_in_zone", "label", "gap_type", "name", "family", "theme", "aisle", "competitor_dept", "home_path", "nearest_staples_shelf", "k",
            "carriers", "PC", "GS", "AAS", "CRS", "O", "evidence", "p_approved", "p_top_n", "AAS_V", "AAS_G", "CRS_V",
            "CRS_G", "label_V", "label_G", "competitor_view", "median_depth_gap", "verify_outcome", "action",
            "example_paths", "crs_driver", "unit_id"]
    readme = pd.DataFrame({"item": [
        "What this is", "Unit", "Zones", "AAS", "CRS", "PC", "GS", "O", "evidence", "p_approved / p_top_n",
        "Figures", "Re-run"], "meaning": [
        "Category-level marketplace recommendations for Staples vs Office Depot, West Elm, Wayfair, Amazon, Walmart",
        "ENTER = a category concept Staples does not carry, merged across competitors; DEEPEN = a Staples L2/L3 aisle "
        "where >= 2 competitors are at least 2x deeper (size-factor adjusted)",
        "CURATE (AAS>=60, CRS<40) · VERTICAL EXTENSION (AAS 40-60, CRS<40) · OFF-BRAND (AAS<40, CRS<40) · REVIEW "
        "(CRS 40-60) · 1P-CORE GAP (CRS>=60) · VERIFY (flagged gap that Staples already sells under another name)",
        "Adjacency Affinity Score 0-100: how embedded the category is in the neighbourhood Staples already sells "
        "(mean of Method V vectors and Method G graph)",
        "Cannibalisation Risk Score 0-100: closeness to a Staples CORE 1P shelf x that shelf's coreness",
        "Peer consensus: share of the 5 competitors that carry it (ENTER) or are >= 2x deeper (DEEPEN)",
        "Gap size: percentile of the missing assortment among the carrier's L1-L3 categories (shelves; items too "
        "where counts exist)",
        "Opportunity = [0.35 PC + 0.35 GS + 0.30 AAS/100] x (1 - CRS/100)",
        "A = 3+ competitors and both methods agree on the zone; B = 2+ competitors or methods agree; C = otherwise",
        "Share of 500 stress-test runs in which the unit stays CURATE / stays in the top 20",
        "figures/fig01..fig20 - the numbers on every figure come from this workbook (Figure 6 and the zone "
        "charts: sheet Aisle_Recommendations; Figure 7: All_Units)",
        "python method_v_vector.py && python method_g_graph.py && python run_framework.py"]})
    with pd.ExcelWriter(fp, engine="openpyxl") as xw:
        readme.to_excel(xw, sheet_name="README", index=False)
        zs = pd.DataFrame(S["zones"]).T.reset_index().rename(columns={"index": "zone"})
        zs.to_excel(xw, sheet_name="Zone_Summary", index=False)
        F.round(3).to_excel(xw, sheet_name="Aisle_Recommendations", index=False)
        U.sort_values(["label", "O"], ascending=[True, False])[[c for c in cols if c in U]].round(3).to_excel(
            xw, sheet_name="All_Units", index=False)
        for lab in LABELS:
            U[U.label == lab].sort_values("O", ascending=False)[[c for c in cols if c in U]].round(3).to_excel(
                xw, sheet_name=slug(lab)[:28].upper(), index=False)
        th = pd.DataFrame({k: v["by_zone"] for k, v in S.get("themes", {}).items()}).T
        th.to_excel(xw, sheet_name="Themes")
        members.round(3).to_excel(xw, sheet_name="Concept_Members", index=False)
        evid.round(4).to_excel(xw, sheet_name="Deepen_Evidence", index=False)
        ver = gold[gold.source.astype(str).str.contains("verification")]
        ver.to_excel(xw, sheet_name="Verification_Log", index=False)
        pd.concat([cal_V, cal_G]).round(4).to_excel(xw, sheet_name="Calibration", index=False)
        for f, sh in (("encoder_bakeoff.csv", "Encoder_Bakeoff"), ("context_weight_grid.csv", "Context_Grid")):
            try:
                pd.read_csv(pc.out_path("V", f)).round(4).to_excel(xw, sheet_name=sh, index=False)
            except Exception:  # noqa
                pass
        da = pd.DataFrame(S["retailers"]).T
        da.to_excel(xw, sheet_name="Data_Audit")
        sc = nodes[~nodes.in_scope].groupby(["retailer", "drop_reason"]).size().rename("nodes").reset_index()
        sc.to_excel(xw, sheet_name="Scope_Audit", index=False)
        m[["competitor", "path_str", "depth", "is_leaf", "n_leaves", "subtree_items", "status_V", "status_G", "ens",
           "label_source", "best_staples_V", "best_score_V", "best_staples_G", "best_score_G", "AAS", "CRS",
           "gap_share"]].round(3).to_excel(xw, sheet_name="Node_Ensemble", index=False)
