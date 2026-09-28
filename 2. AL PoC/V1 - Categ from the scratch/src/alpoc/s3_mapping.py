"""S3 Product-level mapping: competitor family -> Staples leaf (-> pseudo-L5 node) | NONE (§5.4).

Cascade: page prior (crosswalk) -> k-NN classifier trained on Staples families labelled with their own
leaf -> NONE when the product sits outside every candidate leaf. LLM adjudication of low-confidence
cases is the planned 3rd stage (needs an API key); low-confidence rows are flagged instead.
"""
from __future__ import annotations

import numpy as np
import pandas as pd

from . import embed
from .common import cfg, load, load_yaml, log, save, save_json


def mapping_card(r) -> str:
    words = " ".join(str(r["desc"]).split()[: cfg()["extraction"]["desc_words_in_card"]])
    return f"{r['title_clean']}. {words}".strip()


def _leaf_lookup(st: pd.DataFrame) -> dict:
    return st.drop_duplicates("leaf")[["leaf", "leaf_path"]].set_index("leaf")["leaf_path"].to_dict()


def page_candidates(comp: pd.DataFrame, st: pd.DataFrame, crosswalk: dict, model: str) -> tuple[dict, list]:
    """page -> list of candidate leaf paths. Unknown pages get top-5 leaves by name similarity (provisional)."""
    lookup = _leaf_lookup(st)
    pages = sorted({p for ps in comp["pages"] for p in ps.split("|") if p})
    known = crosswalk.get("pages") or {}
    out, auto = {}, []
    leaf_paths = sorted(st["leaf_path"].unique())
    missing = [p for p in pages if p not in known]
    if missing:
        le = embed.encode(leaf_paths, model)
        pe = embed.encode(missing, model)
        for p, v in zip(missing, pe):
            out[p] = [leaf_paths[i] for i in np.argsort(-(le @ v))[:5]]
            auto.append(p)
    for p in pages:
        if p in known:
            out[p] = [lookup[n] for n in known[p] if n in lookup]
    return out, auto


def classify(W: np.ndarray, S: np.ndarray, s_leaf: np.ndarray, cands: list[tuple], k: int, top: int):
    """Returns per competitor row: (pred leaf, prob, margin, affinity of pred)."""
    res = [None] * len(W)
    groups: dict = {}
    for i, c in enumerate(cands):
        groups.setdefault(c, []).append(i)
    for cset, idx in groups.items():
        if not cset:
            for i in idx:
                res[i] = (None, 0.0, 0.0, np.nan)
            continue
        mask = np.isin(s_leaf, cset)
        Sm, lm = S[mask], s_leaf[mask]
        sims = W[idx] @ Sm.T
        kk = min(k, sims.shape[1])
        nn = np.argpartition(-sims, kk - 1, axis=1)[:, :kk]
        leaves = list(cset)
        for r, i in enumerate(idx):
            votes = {l: 0.0 for l in leaves}
            for j in nn[r]:
                votes[lm[j]] += max(sims[r, j], 0)
            tot = sum(votes.values()) or 1.0
            probs = sorted(((v / tot, l) for l, v in votes.items()), reverse=True)
            p1, pred = probs[0]
            p2 = probs[1][0] if len(probs) > 1 else 0.0
            ls = sims[r, lm == pred]
            aff = float(np.sort(ls)[-min(top, len(ls)):].mean())
            res[i] = (pred, float(p1), float(p1 - p2), aff)
    return res


def staples_affinity_loo(S: np.ndarray, s_leaf: np.ndarray, top: int) -> np.ndarray:
    """Each Staples family's mean top-`top` cosine to its own leaf (excluding itself)."""
    out = np.full(len(S), np.nan)
    for leaf in np.unique(s_leaf):
        m = np.where(s_leaf == leaf)[0]
        if len(m) < 2:
            continue
        sims = S[m] @ S[m].T
        np.fill_diagonal(sims, -1)
        t = min(top, len(m) - 1)
        out[m] = np.sort(sims, axis=1)[:, -t:].mean(axis=1)
    return out


def knn_leaf_accuracy(S: np.ndarray, s_leaf: np.ndarray, k: int) -> float:
    sims = S @ S.T
    np.fill_diagonal(sims, -1)
    nn = np.argpartition(-sims, k - 1, axis=1)[:, :k]
    correct = 0
    for i in range(len(S)):
        votes: dict = {}
        for j in nn[i]:
            votes[s_leaf[j]] = votes.get(s_leaf[j], 0) + sims[i, j]
        correct += max(votes, key=votes.get) == s_leaf[i]
    return correct / len(S)


def run() -> dict:
    c = cfg()["mapping"]
    fam = load("families.parquet")
    st = fam[fam["retailer"] == "staples"].reset_index(drop=True)
    retailers = load_yaml("retailers.yaml")
    comp_names = [n for n, a in retailers.items() if a["role"] == "competitor"]
    comp = fam[fam["retailer"].isin(comp_names)].reset_index(drop=True)
    crosswalk = {}
    for n in comp_names:
        cw = load_yaml(retailers[n]["crosswalk"])
        crosswalk.setdefault("pages", {}).update(cw.get("pages") or {})

    # --- candidate leaves per product (union over the pages it was found on)
    default_model = cfg()["encoder"]["default"]
    page_cands, auto_pages = page_candidates(comp, st, crosswalk, default_model)
    comp["cands"] = [tuple(sorted({l for p in ps.split("|") for l in page_cands.get(p, [])})) for ps in comp["pages"]]
    universe = sorted({l for cs in comp["cands"] for l in cs})
    st_u = st[st["leaf_path"].isin(universe)].reset_index(drop=True)
    silver = comp[comp["pages"].map(lambda ps: len({l for p in ps.split("|") for l in page_cands.get(p, [])}) == 1)
                  & comp["cands"].map(len).eq(1)]

    # --- encoder bake-off on the mapping task
    # NONE threshold is calibrated ACROSS sources (Staples-vs-Staples similarity is systematically higher):
    #   positives = products from single-candidate pages (known in scope), affinity to that leaf
    #   negatives = products from pages with no Staples shelf (known out of scope), best affinity over all leaves
    # tau = Youden-J cut. Metrics: NONE balanced accuracy + Staples leave-one-out k-NN leaf accuracy.
    models = cfg()["encoder"]["candidates"] if cfg()["encoder"]["bakeoff"] else [default_model]
    s_cards = [mapping_card(r) for _, r in st_u.iterrows()]
    c_cards = [mapping_card(r) for _, r in comp.iterrows()]
    negatives = comp[comp["cands"].map(len) == 0]
    bake = []
    for m in models:
        S = embed.encode(s_cards, m)
        Wv = embed.encode(c_cards, m)
        leaves_u = st_u["leaf_path"].values
        pos = np.array([r[3] for r in classify(Wv[silver.index.values], S, leaves_u, list(silver["cands"]), c["k"], c["affinity_top"])])
        neg = np.array([r[3] for r in classify(Wv[negatives.index.values], S, leaves_u,
                                              [tuple(universe)] * len(negatives), c["k"], c["affinity_top"])])
        if len(pos) >= 30 and len(neg) >= 30:
            grid = np.quantile(np.concatenate([pos, neg]), np.linspace(0.01, 0.99, 197))
            j = [(np.mean(pos >= g) + np.mean(neg < g) - 1, g) for g in grid]
            tau = float(max(j)[1])
            tau_src = "youden (cross-source)"
        else:
            tau = float(np.nanpercentile(staples_affinity_loo(S, leaves_u, c["affinity_top"]), c["none_percentile"]))
            tau_src = "staples percentile (fallback)"
        recall_in = float(np.mean(pos >= tau)) if len(pos) else np.nan
        reject_out = float(np.mean(neg < tau)) if len(neg) else np.nan
        cv = knn_leaf_accuracy(S, leaves_u, c["k"])
        bake.append({"model": m, "tau_none": tau, "tau_source": tau_src, "in_scope_recall": recall_in,
                     "out_scope_rejection": reject_out, "none_balanced_acc": (recall_in + reject_out) / 2,
                     "n_pos": int(len(pos)), "n_neg": int(len(neg)), "staples_loo_knn": cv})
        log(f"S3 bake-off {m}: tau={tau:.3f} recall={recall_in:.3f} reject={reject_out:.3f} staples-LOO={cv:.3f}")
    best = max(bake, key=lambda b: (b["none_balanced_acc"] + b["staples_loo_knn"]) / 2)
    model = best["model"]
    embed.set_active_model(model)
    save_json({"model": model, "bakeoff": bake}, "encoder.json")

    # --- classify every competitor family with the chosen encoder
    S = embed.encode(s_cards, model)
    Wv = embed.encode(c_cards, model)
    tau = best["tau_none"]
    res = classify(Wv, S, st_u["leaf_path"].values, list(comp["cands"]), c["k"], c["affinity_top"])
    comp["leaf_pred"] = [r[0] for r in res]
    comp["prob"] = [r[1] for r in res]
    comp["margin"] = [r[2] for r in res]
    comp["affinity"] = [r[3] for r in res]
    comp["status"] = np.where(comp["cands"].map(len) == 0, "none_page",
                              np.where(comp["affinity"] >= tau, "mapped", "none_far"))
    comp.loc[comp["status"] != "mapped", "leaf_pred"] = None
    comp["low_conf"] = (comp["status"] == "mapped") & (comp["prob"] < c["low_conf_prob"])

    # --- pseudo-L5 assignment inside split leaves (k-NN over that leaf's Staples families)
    sa = load("staples_node_initial.parquet").set_index("family_id")["node_id"]
    st_u["node_id"] = st_u["family_id"].map(sa)
    comp["node_id"] = comp["leaf_pred"]
    split_leaves = st_u.loc[st_u["node_id"] != st_u["leaf_path"], "leaf_path"].unique()
    for leaf in split_leaves:
        ci = comp.index[comp["leaf_pred"] == leaf].values
        si = st_u.index[st_u["leaf_path"] == leaf].values
        if len(ci) == 0:
            continue
        sims = Wv[ci] @ S[si].T
        kk = min(c["k"], len(si))
        nn = np.argpartition(-sims, kk - 1, axis=1)[:, :kk]
        labels = st_u.loc[si, "node_id"].values
        for r, i in enumerate(ci):
            votes: dict = {}
            for j in nn[r]:
                votes[labels[j]] = votes.get(labels[j], 0) + sims[r, j]
            comp.at[i, "node_id"] = max(votes, key=votes.get)

    # --- nearest Staples L2 for backlog items (product card vs every Staples leaf path name)
    none_idx = comp.index[comp["status"] != "mapped"].values
    comp["nearest_staples_l2"] = None
    if len(none_idx):
        leaves = st.drop_duplicates("leaf_path")[["leaf_path", "l2_key"]].reset_index(drop=True)
        L = embed.encode(leaves["leaf_path"].tolist(), model)
        nn = np.argmax(Wv[none_idx] @ L.T, axis=1)
        comp.loc[none_idx, "nearest_staples_l2"] = leaves["l2_key"].values[nn]

    comp["cands"] = comp["cands"].map(lambda t: "|".join(t))
    cols = ["family_id", "retailer", "pages", "cands", "leaf_pred", "node_id", "prob", "margin", "affinity",
            "status", "low_conf", "nearest_staples_l2"]
    save(comp[cols], "mapping.parquet")
    qa = {
        "encoder": model, "bakeoff": bake, "tau_none": tau, "auto_candidate_pages": auto_pages,
        "status_counts": comp["status"].value_counts().to_dict(),
        "low_conf_share_of_mapped": float(comp.loc[comp["status"] == "mapped", "low_conf"].mean()),
        "in_scope_recall": best["in_scope_recall"], "out_scope_rejection": best["out_scope_rejection"],
        "none_balanced_acc": best["none_balanced_acc"], "n_pos": best["n_pos"], "n_neg": best["n_neg"],
        "staples_loo_knn": best["staples_loo_knn"], "tau_source": best["tau_source"],
        "page_candidates": page_cands,
    }
    save_json(qa, "qa_mapping.json")
    log(f"S3 mapping: {qa['status_counts']}, low-conf share {qa['low_conf_share_of_mapped']:.2f}")
    return qa
