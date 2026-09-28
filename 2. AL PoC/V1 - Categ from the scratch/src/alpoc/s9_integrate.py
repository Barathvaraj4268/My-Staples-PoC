"""S9 Integration: VOS and TG kept separate -> shared safety gate -> fused re-rank + agreement tier (§7).
Also the Dirichlet weight-sensitivity analysis for both scores (G7)."""
from __future__ import annotations

import numpy as np
import pandas as pd
from scipy.stats import spearmanr

from .common import cfg, load, log, pct_rank, save, save_json

TIER_ORDER = ["Strong", "Gap-led", "Vector-led", "Weak"]


def _tier(v, t):
    top_v, top_t = v >= 2 / 3, t >= 2 / 3
    return "Strong" if top_v and top_t else "Gap-led" if top_t else "Vector-led" if top_v else "Weak"


def _exclusion_reason(r) -> str:
    if r["eligible"]:
        return ""
    if pd.isna(r.get("safe_share")):
        return "no competitor products could be labelled"
    if r["reject_share"] >= cfg()["integration"]["reject_share_max"]:
        return f"{r['reject_share']:.0%} of its competitor products substitute or undercut a Staples product"
    top = max(["share_OFF-BRAND", "share_EXCLUDE", "share_REVIEW", "share_EDGE"], key=lambda k: r.get(k, 0) or 0)
    return f"only {r['safe_share']:.0%} safe to add; largest group is {top.replace('share_', '')} ({r[top]:.0%})"


def topn_stability(a: pd.DataFrame, score_fn, draws, alpha, n, rng) -> pd.Series:
    """Share of weight draws in which each archetype stays in its node's top-n (eligible only)."""
    base = a.assign(s=score_fn(None))
    base_top = {nid: set(g.nlargest(n, "s")["archetype_id"]) for nid, g in base[base["eligible"]].groupby("node_id")}
    hits = pd.Series(0.0, index=a["archetype_id"])
    for _ in range(draws):
        s = a.assign(s=score_fn(rng))
        for nid, g in s[s["eligible"]].groupby("node_id"):
            for aid in g.nlargest(n, "s")["archetype_id"]:
                hits[aid] += 1
    return hits / draws, base_top


def run() -> dict:
    ic, sc = cfg()["integration"], cfg()["sensitivity"]
    m1 = load("archetype_m1.parquet")
    m2 = load("archetype_m2.parquet")
    keep2 = [c for c in m2.columns if c not in m1.columns or c == "archetype_id"]
    a = m1.merge(m2[keep2], on="archetype_id", how="left")
    a["eligible"] = a["eligible"].fillna(False).astype(bool) & (a["depth"] > 0)
    a["exclusion_reason"] = a.apply(_exclusion_reason, axis=1)
    a.loc[a["depth"] == 0, "exclusion_reason"] = "mixed bucket (not a nameable archetype)"
    # agreement tiers: terciles of each method's rank within the node (all archetypes)
    a["rank_vos_node"] = a.groupby("node_id")["vos"].transform(pct_rank)
    a["rank_tg_node"] = a.groupby("node_id")["tg"].transform(pct_rank)
    a["tier"] = [_tier(v, t) for v, t in zip(a["rank_vos_node"].fillna(0), a["rank_tg_node"].fillna(0))]
    # fused re-rank among eligible archetypes
    el = a["eligible"]
    a["final"] = np.nan
    a.loc[el, "final"] = 100 * (ic["w_vos"] * a[el].groupby("node_id")["vos"].transform(pct_rank)
                                + ic["w_tg"] * a[el].groupby("node_id")["tg"].transform(pct_rank))
    a["final_rank"] = a[el].groupby("node_id")["final"].rank(ascending=False, method="first")
    rk_v = a[el].groupby("node_id")["vos"].rank(ascending=False)
    rk_t = a[el].groupby("node_id")["tg"].rank(ascending=False)
    a.loc[el, "rrf"] = 1 / (ic["rrf_k"] + rk_v) + 1 / (ic["rrf_k"] + rk_t)
    a["shortlisted"] = el & (a["tier"] != "Weak") & (a["final_rank"] <= ic["shortlist_n"])

    # --- sensitivity (Dirichlet around base weights), both scores
    rng = np.random.default_rng(5)
    gw, vw = cfg()["gaps"]["tg_weights"], cfg()["vector"]["vos_weights"]
    shrink = np.where(a["lsr_credibility"].fillna(0) >= cfg()["gaps"]["credible"], 1.0, cfg()["gaps"]["noncredible_shrink"])

    def tg_fn(r):
        w = dict(zip(gw, rng.dirichlet(sc["alpha_scale"] * np.array(list(gw.values()))))) if r is not None else gw
        return 100 * (w["lsr"] * a["pct_lsr"] * shrink + w["ppg"] * a["pct_ppg"] + w["cg"] * a["pct_cg"]
                      + w["msg"] * a["pct_msg"] + w["dfg"] * a["pct_dfg"])

    def vos_fn(r):
        w = dict(zip(vw, rng.dirichlet(sc["alpha_scale"] * np.array(list(vw.values()))))) if r is not None else vw
        return 100 * (w["vw"] * a["pct_vw"] + w["aas"] * a["aas"] / 100 + w["ad"] * a["pct_ad"]) \
            * (1 - a["crs"] / 100) ** cfg()["vector"]["gamma"]

    stab = {}
    for name, fn in (("tg", tg_fn), ("vos", vos_fn)):
        hits, base_top = topn_stability(a, fn, sc["draws"], sc["alpha_scale"], sc["top_n"], rng)
        a[f"stability_{name}"] = a["archetype_id"].map(hits)
        per_node = {nid: float(np.mean([hits[x] for x in top])) for nid, top in base_top.items() if top}
        stab[name] = {"mean_topn_retention": float(np.mean(list(per_node.values()))) if per_node else None,
                      "per_node": per_node}
    # --- per-node agreement
    agree = {}
    for nid, g in a.groupby("node_id"):
        g = g.dropna(subset=["vos", "tg"])
        agree[nid] = float(spearmanr(g["vos"], g["tg"]).statistic) if len(g) >= 4 else None
    save(a, "final_archetypes.parquet")
    qa = {"sensitivity": stab, "spearman_vos_tg": agree,
          "n_eligible": int(el.sum()), "n_shortlisted": int(a["shortlisted"].sum()),
          "tier_counts": a["tier"].value_counts().to_dict()}
    save_json(qa, "qa_integration.json")
    log(f"S9: {qa['n_eligible']} eligible, {qa['n_shortlisted']} shortlisted; stability TG="
        f"{stab['tg']['mean_topn_retention']}, VOS={stab['vos']['mean_topn_retention']}")
    return qa
