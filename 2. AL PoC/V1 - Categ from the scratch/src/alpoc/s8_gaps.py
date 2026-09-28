"""S8 METHOD 2 - attribute-level gaps: LSR, PPG, CG, MSG, DFG -> TG; attribute-value gaps (§6.7)."""
from __future__ import annotations

import numpy as np
import pandas as pd
from scipy.stats import wasserstein_distance

from .common import cfg, jeffreys_mean, jsd, load, load_json, log, pct_rank, prob_greater, save, smoothed_dist, split_multi
from .cards import tier2_attrs
from .s5_extract import MULTI_ATTRS


def dist_jsd(a: pd.Series, b: pd.Series) -> float:
    a, b = a.dropna(), b.dropna()
    if len(a) == 0 or len(b) == 0:
        return np.nan
    support = sorted(set(a) | set(b))
    return jsd(smoothed_dist(a.value_counts(), support), smoothed_dist(b.value_counts(), support))


def w1_log(a: pd.Series, b: pd.Series) -> tuple[float, float]:
    a, b = np.log(a.dropna()[a.dropna() > 0]), np.log(b.dropna()[b.dropna() > 0])
    if len(a) < 1 or len(b) < 1:
        return np.nan, np.nan
    return float(wasserstein_distance(a, b)), float(np.sign(np.median(a) - np.median(b)))


def run() -> dict:
    gp = cfg()["gaps"]
    rng = np.random.default_rng(3)
    ft = load("family_table.parquet")
    arches = load("archetypes.parquet")
    members = load("archetype_members.parquet")
    nodes = load("nodes.parquet")
    qx = load_json("qa_extraction.json")["l2"]
    comp = [r for r in ft["retailer"].unique() if r != "staples"][0]
    ft = ft.merge(members, on="family_id", how="left")
    arch_rows, attr_rows, band_rows, node_rows = [], [], [], []
    for _, nd in nodes[nodes["status"].isin(["scored", "thin"])].iterrows():
        nid, l2 = nd["node_id"], nd["l2_key"]
        df = ft[ft["node_id"] == nid]
        S, C = df[df["retailer"] == "staples"], df[df["retailer"] == comp]
        NS, NC = len(S), len(C)
        fails = set(qx.get(l2, {}).get("g2_fail", []))
        msg_attrs = [a for a in ("material_class", "style_family") if a not in fails]
        node_rows.append({"node_id": nid, "n_staples": NS, "n_competitor": NC,
                          "dfi_median_staples": float(S["dfi"].median()) if NS else np.nan,
                          "dfi_median_competitor": float(C["dfi"].median()) if NC else np.nan,
                          "design_forward_share_staples": float((S["dfi"] >= gp["dfi_forward"]).mean()) if NS else np.nan,
                          "design_forward_share_competitor": float((C["dfi"] >= gp["dfi_forward"]).mean()) if NC else np.nan,
                          "price_median_staples": float(S["price"].median()) if NS else np.nan,
                          "price_median_competitor": float(C["price"].median()) if NC else np.nan})
        # --- price-band coverage (node level)
        for b in sorted(df["price_band"].dropna().unique(), key=lambda x: df.loc[df["price_band"] == x, "price"].median()):
            ks, kc = int((S["price_band"] == b).sum()), int((C["price_band"] == b).sum())
            band_rows.append({"node_id": nid, "band": b, "share_staples": ks / max(NS, 1), "share_competitor": kc / max(NC, 1),
                              "n_staples": ks, "n_competitor": kc, "price_mid": float(df.loc[df["price_band"] == b, "price"].median()),
                              "credibility": prob_greater(kc, NC, ks, NS, gp["mc_draws"], rng) if NS and NC else np.nan})
        # --- attribute-value gaps (node level)
        attrs = ["colour_family", "colour_tone", "material_class", "style_family", "size_class"] + \
                [a for a in tier2_attrs(l2) if a in df.columns]
        for a in attrs:
            s_known, c_known = S[a].dropna(), C[a].dropna()
            if len(s_known) == 0 or len(c_known) == 0:
                continue
            j = dist_jsd(s_known, c_known)
            for val in sorted(set(s_known) | set(c_known)):
                ks, kc = int((s_known == val).sum()), int((c_known == val).sum())
                attr_rows.append({"node_id": nid, "attribute": a, "value": val, "kind": "single", "jsd": j,
                                  "share_staples": ks / len(s_known), "share_competitor": kc / len(c_known),
                                  "n_staples": ks, "n_competitor": kc,
                                  "known_staples": len(s_known), "known_competitor": len(c_known),
                                  "credibility": prob_greater(kc, len(c_known), ks, len(s_known), gp["mc_draws"], rng),
                                  "descriptive_only": a in fails})
        for a in MULTI_ATTRS:
            s_tags = S[a].map(split_multi)
            c_tags = C[a].map(split_multi)
            vals = sorted({t for ts in list(s_tags) + list(c_tags) for t in ts})
            if not vals or NS == 0 or NC == 0:
                continue
            fs = pd.Series({t: sum(t in ts for ts in s_tags) for t in vals})
            fc = pd.Series({t: sum(t in ts for ts in c_tags) for t in vals})
            j = jsd(smoothed_dist(fs, vals), smoothed_dist(fc, vals))
            for t in vals:
                attr_rows.append({"node_id": nid, "attribute": a, "value": t, "kind": "multi", "jsd": j,
                                  "share_staples": fs[t] / NS, "share_competitor": fc[t] / NC,
                                  "n_staples": int(fs[t]), "n_competitor": int(fc[t]),
                                  "known_staples": NS, "known_competitor": NC,
                                  "credibility": prob_greater(int(fc[t]), NC, int(fs[t]), NS, gp["mc_draws"], rng),
                                  "descriptive_only": False})
        if nd["status"] != "scored":
            continue
        # --- archetype-level Method-2 components
        for aid, m in df.groupby("archetype_id"):
            ms, mc = m[m["retailer"] == "staples"], m[m["retailer"] == comp]
            ns, nc = len(ms), len(mc)
            within = ns >= gp["min_side_for_within"]
            ref = ms if within else S
            ppg, ppg_dir = w1_log(mc["price"], ref["price"])
            cg = dist_jsd(mc["colour_tone"], ref["colour_tone"])
            msg = np.nanmean([dist_jsd(mc[a], ref[a]) for a in msg_attrs]) if msg_attrs else np.nan
            dfg = float(mc["dfi"].median() - ref["dfi"].median()) if nc and len(ref) else np.nan
            # colourway breadth: colour families the competitor offers here that Staples does not
            st_cols = {c for x in ms["colours_observed"] for c in split_multi(x)} | set(ms["colour_family"].dropna())
            cp_cols = {c for x in mc["colours_observed"] for c in split_multi(x)} | set(mc["colour_family"].dropna())
            arch_rows.append({
                "archetype_id": aid, "node_id": nid, "l2_key": l2,
                "p_staples": ns / NS, "p_competitor": nc / NC,
                "lsr": float(np.log(jeffreys_mean(nc, NC) / jeffreys_mean(ns, NS))),
                "lsr_credibility": prob_greater(nc, NC, ns, NS, gp["mc_draws"], rng),
                "absent": ns == 0 and nc >= 5, "thin": (ns / NS) < 0.25 * (nc / NC),
                "ppg": ppg, "ppg_dir": ppg_dir, "cg": cg, "msg": msg, "dfg": dfg,
                "within_archetype_ref": within,
                "missing_colours": "|".join(sorted(cp_cols - st_cols)),
                "staples_colourways_mean": float(ms["n_colourways"].mean()) if ns else np.nan,
            })
    arch = pd.DataFrame(arch_rows)
    # --- TG: percentiles across all archetypes in the same L2 peer group
    for col in ["lsr", "ppg", "cg", "msg", "dfg"]:
        arch[f"pct_{col}"] = arch.groupby("l2_key")[col].transform(pct_rank).fillna(0.5)
    shrink = np.where(arch["lsr_credibility"] >= gp["credible"], 1.0, gp["noncredible_shrink"])
    w = gp["tg_weights"]
    arch["tg"] = 100 * (w["lsr"] * arch["pct_lsr"] * shrink + w["ppg"] * arch["pct_ppg"] + w["cg"] * arch["pct_cg"]
                        + w["msg"] * arch["pct_msg"] + w["dfg"] * arch["pct_dfg"])
    wo = gp["original_tg_weights"]
    arch["tg_original"] = 100 * (wo["lsr"] * arch["pct_lsr"] * shrink + wo["ppg"] * arch["pct_ppg"]
                                 + wo["cg"] * arch["pct_cg"] + wo["dfg"] * arch["pct_dfg"])
    arch = arches.merge(arch.drop(columns=["node_id", "l2_key"]), on="archetype_id", how="left")
    attrs_df = pd.DataFrame(attr_rows)
    attrs_df["delta"] = attrs_df["share_competitor"] - attrs_df["share_staples"]
    attrs_df["tg_attr"] = attrs_df.groupby("node_id")["delta"].transform(pct_rank) * attrs_df["credibility"]
    save(arch, "archetype_m2.parquet")
    save(attrs_df, "attr_gaps.parquet")
    save(pd.DataFrame(band_rows), "price_bands.parquet")
    save(pd.DataFrame(node_rows), "node_summary.parquet")
    log(f"S8: {len(arch)} archetypes scored, {len(attrs_df)} attribute-value gaps")
    return {"n_archetypes": len(arch), "n_attr_rows": len(attrs_df)}
