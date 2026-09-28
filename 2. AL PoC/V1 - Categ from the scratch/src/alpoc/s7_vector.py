"""S7 METHOD 1 - vector space: VW, AAS, CRS, AD, PPR -> decision labels -> VOS (§6.5-6.6)."""
from __future__ import annotations

import numpy as np
import pandas as pd
from sklearn.linear_model import LogisticRegression
from sklearn.metrics import roc_auc_score
from sklearn.model_selection import cross_val_score

from . import embed
from .cards import full_card, functional_core, neutral_card, neutral_functional_card, tier2_attrs
from .common import cfg, l2_config, load, log, pct_of, pct_rank, save, save_json, split_multi

APPROVE = ["CURATE", "STYLE-EXTENSION", "TRADE-UP"]
HOLD = ["REVIEW", "EDGE"]
REJECT = ["SUBSTITUTE", "UNDERCUT", "OFF-BRAND", "EXCLUDE"]
LABEL_GROUP = {**{l: "Approve" for l in APPROVE}, **{l: "Review / hold" for l in HOLD}, **{l: "Reject" for l in REJECT}}
AESTHETIC = ["colour_family", "colour_tone", "material_class", "style_family"]


def _known(v) -> bool:
    return v is not None and not (isinstance(v, float) and np.isnan(v)) and v != ""


def _t(r: pd.Series, k: str):
    return r["txt_" + k] if ("txt_" + k) in r.index else r.get(k)


def functional_identity(a: pd.Series, b: pd.Series, core: list[str]):
    comp = [(_t(a, k), _t(b, k)) for k in core if k in a and _known(_t(a, k)) and _known(_t(b, k))]
    if not comp:
        return np.nan, 0
    return float(np.mean([x == y for x, y in comp])), len(comp)


def aesthetic_delta(a: pd.Series, b: pd.Series) -> float:
    d = [float(_t(a, k) != _t(b, k)) for k in AESTHETIC if _known(_t(a, k)) and _known(_t(b, k))]
    ta, tb = set(split_multi(a["aesthetic_tags"])), set(split_multi(b["aesthetic_tags"]))
    if ta or tb:
        d.append(1 - len(ta & tb) / len(ta | tb))
    return float(np.mean(d)) if d else 0.0


def calibrate(st: pd.DataFrame, Ef: np.ndarray, core: list[str], n_pairs: int, rng) -> dict:
    """Weak supervision on Staples-only pairs (§6.6): cosine(functional) -> P(substitute)."""
    pos, neg = [], []
    by_node = {n: np.where(st["node_id"].values == n)[0] for n in st["node_id"].unique()}
    nodes = [n for n, ix in by_node.items() if len(ix) >= 2]
    logp = np.log(st["price"].values)
    for _ in range(n_pairs * 4):
        if len(pos) >= n_pairs and len(neg) >= n_pairs:
            break
        n = nodes[rng.integers(len(nodes))]
        i, j = rng.choice(by_node[n], 2, replace=False)
        fi, k = functional_identity(st.iloc[i], st.iloc[j], core)
        cos = float(Ef[i] @ Ef[j])
        if k >= 3 and fi >= 0.9 and abs(logp[i] - logp[j]) <= 0.25 and len(pos) < n_pairs:
            pos.append(cos)
        elif k >= 3 and fi <= 0.34 and len(neg) < n_pairs // 2:
            neg.append(cos)
    if len(nodes) >= 2:
        for _ in range(n_pairs // 2):
            a, b = rng.choice(len(nodes), 2, replace=False)
            neg.append(float(Ef[rng.choice(by_node[nodes[a]])] @ Ef[rng.choice(by_node[nodes[b]])]))
    if len(pos) < 30 or len(neg) < 30:
        return {"method": "fallback-percentile", "n_pos": len(pos), "n_neg": len(neg), "auc": None}
    X = np.array(pos + neg)[:, None]
    y = np.array([1] * len(pos) + [0] * len(neg))
    lr = LogisticRegression(class_weight="balanced").fit(X, y)
    return {"method": "logistic", "coef": float(lr.coef_[0, 0]), "intercept": float(lr.intercept_[0]),
            "n_pos": len(pos), "n_neg": len(neg), "auc": float(roc_auc_score(y, X[:, 0]))}


def cal_prob(cos: np.ndarray, cal: dict, ref: np.ndarray) -> np.ndarray:
    if cal["method"] == "logistic":
        return 1 / (1 + np.exp(-(cal["coef"] * cos + cal["intercept"])))
    return pct_of(cos, ref)


def topk_mean(sims: np.ndarray, k: int) -> np.ndarray:
    k = min(k, sims.shape[1])
    return np.sort(sims, axis=1)[:, -k:].mean(axis=1)


def ctx_score(tags: list[str], shares: dict) -> float:
    if not tags:
        return 0.5
    return float(np.mean([min(1.0, shares.get(t, 0.0) / 0.10) for t in tags]))


def _tags(r) -> list[str]:
    return split_multi(r["use_context"]) + split_multi(r["end_user_segment"])


def label(r, v: dict, T: dict) -> str:
    if r["identical"]:
        return "EXCLUDE"
    if r["aas"] < T["aas_low"]:
        return "OFF-BRAND"
    if r["crs"] >= v["crs_high"]:
        ppr = r["ppr"] if not np.isnan(r["ppr"]) else 1.0
        if ppr < v["ppr_undercut"]:
            return "UNDERCUT"
        if ppr >= v["ppr_tradeup"] and (r["d_dfi"] >= v["dfi_delta_tradeup"] or r["material_upgrade"]):
            return "TRADE-UP"
        if r["ad"] >= v["ad_threshold"]:
            return "STYLE-EXTENSION"
        return "SUBSTITUTE"
    if r["crs"] >= v["crs_low"]:
        return "REVIEW"
    return "CURATE" if r["aas"] >= T["aas_high"] else "EDGE"


def run() -> dict:
    v = cfg()["vector"]
    rng = np.random.default_rng(11)
    ft = load("family_table.parquet").reset_index(drop=True)
    nodes = load("nodes.parquet")
    members = load("archetype_members.parquet")
    arches = load("archetypes.parquet")
    comp = [r for r in ft["retailer"].unique() if r != "staples"][0]
    scored = set(nodes.loc[nodes["status"] == "scored", "node_id"])
    cand_rows, qa = [], {}
    for l2, g in ft.groupby("l2_key"):
        g = g.reset_index(drop=True)
        t2, core = tier2_attrs(l2), functional_core(l2)
        tiers = l2_config(l2).get("material_tiers") or {}
        Ef = embed.encode([neutral_card(r, t2) for _, r in g.iterrows()])
        Eu = embed.encode([neutral_functional_card(r, core) for _, r in g.iterrows()])
        Et = embed.encode([full_card(r, t2) for _, r in g.iterrows()])      # title-bearing card: diagnostic only
        st_mask = (g["retailer"] == "staples").values
        # AAS measures fit in FUNCTION and use (functional view); the look is rewarded separately through AD
        S_all = Eu[st_mask]
        # AAS reference: Staples' own semantic affinity to the rest of Staples in the L2 (leave-one-out)
        ss = S_all @ S_all.T
        np.fill_diagonal(ss, -1)
        a_sem_st = topk_mean(ss, v["aas_top"])
        st_tags = [_tags(r) for _, r in g[st_mask].iterrows()]
        shares = pd.Series([t for ts in st_tags for t in set(ts)]).value_counts() / max(1, st_mask.sum())
        shares = shares.to_dict()
        aas_st = 100 * (0.6 * pct_of(a_sem_st, a_sem_st) + 0.4 * np.array([ctx_score(t, shares) for t in st_tags]))
        T = {"aas_low": float(np.percentile(aas_st, v["aas_low_pct"])),
             "aas_high": float(np.percentile(aas_st, v["aas_high_pct"]))}
        st_scored = g[st_mask & g["node_id"].isin(scored).values].reset_index()
        cal = calibrate(st_scored, Eu[st_scored["index"].values], core, v["calib_pairs"], rng)
        # reference for the percentile fallback: Staples nearest-neighbour functional similarity within node
        ref = []
        for n, ix in st_scored.groupby("node_id")["index"]:
            if len(ix) > 1:
                m = Eu[ix.values] @ Eu[ix.values].T
                np.fill_diagonal(m, -1)
                ref += list(topk_mean(m, v["crs_top"]))
        qa[l2] = {"thresholds": T, "calibration": cal, "staples_aas_median": float(np.median(aas_st))}

        for nid in sorted(scored & set(g["node_id"])):
            si = np.where((g["node_id"] == nid).values & st_mask)[0]
            ci = np.where((g["node_id"] == nid).values & ~st_mask)[0]
            if len(si) == 0 or len(ci) == 0:
                continue
            # Vector Whitespace: balanced k-NN share of Staples around each competitor family
            P = np.concatenate([si, ci])
            w = np.where(np.isin(P, si), len(ci) / len(si), 1.0)
            simsP = Ef[ci] @ Ef[P].T
            simsP[np.arange(len(ci)), len(si) + np.arange(len(ci))] = -9      # exclude self
            k = min(v["k_whitespace"], len(P) - 1)
            kth = -np.partition(-simsP, k - 1, axis=1)[:, k - 1]
            inn = simsP >= (kth[:, None] - 1e-6)                 # k nearest, keeping every tie at the k-th value
            is_s = np.isin(P, si)[None, :]
            lsc = (inn * w * is_s).sum(1) / (inn * w).sum(1)
            vw = np.clip(1 - 2 * lsc, 0, 1)
            # AAS
            a_sem = topk_mean(Eu[ci] @ S_all.T, v["aas_top"])
            a_ctx = np.array([ctx_score(_tags(g.iloc[i]), shares) for i in ci])
            aas = 100 * (0.6 * pct_of(a_sem, a_sem_st) + 0.4 * a_ctx)
            # CRS on the functional view (same node), nearest Staples family
            fs = Eu[ci] @ Eu[si].T
            s_max = topk_mean(fs, v["crs_top"])
            p_sub = cal_prob(s_max, cal, np.array(ref))
            full_t = Et[ci] @ Et[si].T                                  # title-bearing card: identity check only
            for r_i, i in enumerate(ci):
                c_row = g.iloc[i]
                # functional peers = every Staples family within 0.02 of the best functional match (ties are common)
                order = np.argsort(-fs[r_i])
                peers = si[order[fs[r_i, order] >= fs[r_i, order[0]] - 0.02][:15]]
                if len(peers) < v["crs_top"]:
                    peers = si[order[: v["crs_top"]]]
                # nearest = the best-looking peer (smallest aesthetic delta): conservative for STYLE-EXTENSION
                ads = [aesthetic_delta(c_row, g.iloc[j]) for j in peers]
                j = peers[int(np.argmin(ads))]
                s_row = g.iloc[j]
                fi, n_fi = functional_identity(c_row, s_row, core)
                crs = 100 * (0.6 * p_sub[r_i] + 0.4 * fi) if n_fi else 100 * p_sub[r_i]
                nb_prices = g.iloc[peers]["price"].dropna()
                ppr = c_row["price"] / nb_prices.median() if len(nb_prices) and c_row["price"] > 0 else np.nan
                jj = int(np.argmax(full_t[r_i]))
                identical = bool(full_t[r_i, jj] >= v["identical_cos"] and
                                 abs(np.log(c_row["price"] / g.iloc[si[jj]]["price"])) <= v["identical_log_price"])
                cand_rows.append({
                    "family_id": c_row["family_id"], "node_id": nid, "l2_key": l2,
                    "vw": float(vw[r_i]), "aas": float(aas[r_i]), "a_sem_pct": float(pct_of([a_sem[r_i]], a_sem_st)[0]),
                    "a_ctx": float(a_ctx[r_i]), "s_max_p": float(p_sub[r_i]), "fi": fi, "n_fi": n_fi,
                    "crs": float(crs), "ppr": float(ppr) if not np.isnan(ppr) else np.nan,
                    "ad": float(min(ads)), "n_peers": len(peers), "d_dfi": float(c_row["dfi"] - s_row["dfi"]),
                    "material_upgrade": bool(tiers.get(c_row["material_class"], 0) > tiers.get(s_row["material_class"], 99)),
                    "nearest_staples": s_row["family_id"], "identical": identical,
                    "identical_staples": g.iloc[si[jj]]["family_id"] if identical else None,
                })
            # G5: how well can the retailer be predicted from the card? (lower = less source style)
            if len(si) >= 30 and len(ci) >= 30:
                y = np.r_[np.zeros(len(si)), np.ones(len(ci))]
                for view, E in (("neutral", Ef), ("title_card", Et)):
                    auc = cross_val_score(LogisticRegression(max_iter=1000, class_weight="balanced"),
                                          E[np.r_[si, ci]], y, cv=5, scoring="roc_auc").mean()
                    qa[l2].setdefault("g5_" + view, []).append(float(auc))
        log(f"S7 {l2}: calibration {cal['method']} AUC={cal.get('auc')}, T_AAS={T}")
    cand = pd.DataFrame(cand_rows)
    Tmap = {l2: q["thresholds"] for l2, q in qa.items()}
    cand["label"] = [label(r, v, Tmap[r["l2_key"]]) for _, r in cand.iterrows()]
    cand["label_group"] = cand["label"].map(LABEL_GROUP)
    cand["data_error"] = (cand["aas"] < cand["l2_key"].map(lambda l: Tmap[l]["aas_low"])) & (cand["crs"] >= v["crs_high"])
    cand = cand.merge(members, on="family_id", how="left")
    save(cand, "candidates.parquet")

    # --- archetype roll-up + VOS
    agg = cand.groupby("archetype_id").agg(
        vw=("vw", "mean"), aas=("aas", "median"), ad=("ad", "median"), crs=("crs", "median"),
        ppr=("ppr", "median"), n_labelled=("label", "size"),
        safe_share=("label", lambda s: s.isin(APPROVE).mean()),
        reject_share=("label", lambda s: s.isin(["SUBSTITUTE", "UNDERCUT"]).mean()),
        **{f"share_{l}": ("label", (lambda l: (lambda s: (s == l).mean()))(l)) for l in APPROVE + HOLD + REJECT},
    ).reset_index()
    ic = cfg()["integration"]
    agg["eligible"] = (agg["safe_share"] >= ic["safe_share_min"]) & (agg["reject_share"] < ic["reject_share_max"])
    agg = arches.merge(agg, on="archetype_id", how="left")
    w = v["vos_weights"]
    agg["pct_vw"] = agg.groupby("l2_key")["vw"].transform(pct_rank)
    agg["pct_ad"] = agg.groupby("l2_key")["ad"].transform(pct_rank)
    agg["vos"] = 100 * (w["vw"] * agg["pct_vw"] + w["aas"] * agg["aas"] / 100 + w["ad"] * agg["pct_ad"]) \
        * (1 - agg["crs"] / 100) ** v["gamma"]
    save(agg, "archetype_m1.parquet")
    for l2 in qa:
        qa[l2]["g5"] = {v: (float(np.mean(qa[l2].pop("g5_" + v))) if qa[l2].get("g5_" + v) else None)
                        for v in ("neutral", "title_card")}
    qa["label_counts"] = cand["label"].value_counts().to_dict()
    qa["data_error_share"] = float(cand["data_error"].mean())
    save_json(qa, "qa_vector.json")
    log(f"S7 labels: {qa['label_counts']}")
    return qa
