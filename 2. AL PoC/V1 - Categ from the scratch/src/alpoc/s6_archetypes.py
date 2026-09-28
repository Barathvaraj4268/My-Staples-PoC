"""S6 Archetype construction: facet grid first, clusters as validation (§6.3)."""
from __future__ import annotations

import json
import warnings

import numpy as np
import pandas as pd
from sklearn.cluster import HDBSCAN
from sklearn.metrics import adjusted_mutual_info_score, adjusted_rand_score

from . import embed
from .cards import neutral_card, tier2_attrs
from .common import cfg, l2_config, load, load_json, log, money, nice_breaks, save, save_json

UNIVERSAL_GRID = ["material_class", "colour_tone", "style_family", "price_band"]
LABELS = {
    "arm_type": {"armless": "armless", "fixed": "with arms", "adjustable": "adjustable-arm"},
    "base_type": {"casters": "on casters", "legs": "on legs", "sled": "sled-base", "pedestal": "pedestal"},
    "back_height": {"high": "high-back", "mid": "mid-back", "low": "low-back"},
    "shape": {"l_shaped": "L-shaped", "u_shaped": "U-shaped", "corner": "corner", "bow": "bow-front", "rectangular": "rectangular"},
    "sit_stand": {"electric": "electric sit-stand", "manual": "manual sit-stand"},
    "desk_type": {"computer": "computer", "workstation": "workstation", "executive": "executive", "writing": "writing",
                  "standing": "standing", "gaming": "gaming", "reception": "reception", "credenza": "credenza",
                  "secretary": "secretary", "vanity": "vanity", "ladder": "ladder"},
}
NAME_ORDER = ["style_family", "material_class", "arm_type", "base_type", "back_height", "shape", "sit_stand",
              "desk_type", "width_band"]


def price_bands(prices: pd.Series, n: int) -> tuple[list[float], list[str]]:
    qs = prices.dropna().quantile([i / n for i in range(1, n)]).tolist()
    br = nice_breaks(qs)
    labels, lo = [], 0
    for b in br:
        labels.append(f"{money(lo)}–{money(b)}" if lo else f"under {money(b)}")
        lo = b
    labels.append(f"{money(lo)}+")
    return br, labels


def assign_band(p, br, labels):
    if p is None or np.isnan(p):
        return None
    return labels[int(np.searchsorted(br, p, side="right"))]


def _norm_entropy(s: pd.Series) -> float:
    vc = s.dropna().value_counts()
    if len(vc) < 2:
        return 0.0
    p = vc / vc.sum()
    return float(-(p * np.log(p)).sum() / np.log(len(p)))


def choose_facets(df: pd.DataFrame, candidates: list[str], c: dict, comp: str):
    scores = []
    for f in candidates:
        if f not in df.columns:
            continue
        unk = {r: df.loc[df["retailer"] == r, f].isna().mean() for r in ("staples", comp)}
        if max(unk.values()) > c["max_unknown_share"]:
            scores.append({"facet": f, "score": 0.0, "unknown": unk, "eligible": False})
            continue
        score = _norm_entropy(df[f]) * (1 - df[f].isna().mean())
        scores.append({"facet": f, "score": score, "unknown": unk, "eligible": score > 0.05})
    chosen = [s["facet"] for s in sorted(scores, key=lambda s: -s["score"]) if s["eligible"]][: c["max_facets"]]
    return chosen, scores


def grid_assign(df: pd.DataFrame, facets: list[str], c: dict, comp: str) -> pd.Series:
    arche = pd.Series(index=df.index, dtype=object)
    remaining = df.index
    for d in range(len(facets), 0, -1):
        sub = df.loc[remaining, facets[:d]]
        ok = sub.notna().all(axis=1)
        sub = sub[ok]
        if sub.empty:
            continue
        key = sub.astype(str).agg("||".join, axis=1)
        is_comp = df.loc[sub.index, "retailer"] == comp
        stats = pd.DataFrame({"key": key, "comp": is_comp}).groupby("key").agg(n=("comp", "size"), nc=("comp", "sum"))
        good = stats[(stats["n"] >= c["min_pooled"]) & (stats["nc"] >= c["min_competitor"])].index
        hit = key[key.isin(good)]
        arche.loc[hit.index] = [f"{d}::{k}" for k in hit]
        remaining = remaining.difference(hit.index)
    arche.loc[remaining] = "0::mixed"
    return arche


def archetype_name(facets: list[str], values: list[str]) -> str:
    fv = dict(zip(facets, values))
    words = []
    for f in NAME_ORDER:
        if f in fv:
            v = fv[f]
            words.append(LABELS.get(f, {}).get(v, v.replace("_", " ")))
    head = " ".join(words).strip().capitalize() or "Any style"
    tail = [fv[f].replace("-", " ") for f in ("colour_tone",) if f in fv] + [fv[f] for f in ("price_band",) if f in fv]
    return head + (" — " + ", ".join(tail) if tail else "")


def validate(E: np.ndarray, labels: np.ndarray, c: dict, rng: np.random.Generator) -> dict:
    n = len(E)
    if n < 60 or len(set(labels)) < 2:
        return {"ari": None, "ami": None, "stability": None, "n_clusters": None, "noise_share": None}
    import umap
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        Z = umap.UMAP(n_neighbors=15, min_dist=0.0, n_components=c["umap_components"], random_state=42).fit_transform(E)
    mcs = max(8, n // 30)
    hl = HDBSCAN(min_cluster_size=mcs, cluster_selection_method="leaf").fit_predict(Z)
    keep = hl >= 0
    ari = adjusted_rand_score(labels[keep], hl[keep]) if keep.sum() > 10 else np.nan
    ami = adjusted_mutual_info_score(labels[keep], hl[keep]) if keep.sum() > 10 else np.nan
    boots = []
    for _ in range(c["bootstrap"]):
        idx = rng.choice(n, int(0.8 * n), replace=False)
        hb = HDBSCAN(min_cluster_size=mcs, cluster_selection_method="leaf").fit_predict(Z[idx])
        k = hb >= 0
        if k.sum() > 10:
            boots.append(adjusted_rand_score(labels[idx][k], hb[k]))
    return {"ari": float(ari), "ami": float(ami), "stability": float(np.mean(boots)) if boots else None,
            "n_clusters": int(hl.max() + 1), "noise_share": float((~keep).mean())}


def run() -> dict:
    c = cfg()["archetypes"]
    att = load("attributes.parquet")
    fam = load("families.parquet")[["family_id", "title", "title_clean", "brand", "url", "house_brand",
                                    "brand_on_staples", "colours_observed", "n_colourways"]]
    att = att.merge(fam, on="family_id", how="left")
    nodes = load("nodes.parquet")
    qx = load_json("qa_extraction.json")["l2"]
    comp = [r for r in att["retailer"].unique() if r != "staples"][0]
    rng = np.random.default_rng(7)
    members, arches, qa = [], [], {}
    att["price_band"] = None
    for _, nd in nodes[nodes["status"].isin(["scored", "thin"])].iterrows():
        idx = att.index[att["node_id"] == nd["node_id"]]
        df = att.loc[idx]
        br, labels = price_bands(df["price"], c["price_bands"])
        att.loc[idx, "price_band"] = df["price"].map(lambda p: assign_band(p, br, labels))
        df = att.loc[idx]
        if nd["status"] != "scored":
            continue
        l2 = nd["l2_key"]
        fails = set(qx.get(l2, {}).get("g2_fail", []))
        cands = [f for f in (l2_config(l2).get("grid_candidates") or UNIVERSAL_GRID) if f not in fails]
        facets, fscores = choose_facets(df, cands, c, comp)
        arche = grid_assign(df, facets, c, comp)
        for key in arche.unique():
            depth, vals = key.split("::", 1)
            depth = int(depth)
            fs = facets[:depth]
            vs = vals.split("||") if depth else []
            aid = f"{nd['node_id']} :: " + (" · ".join(f"{f}={v}" for f, v in zip(fs, vs)) if depth else "mixed")
            m = arche[arche == key].index
            arches.append({"archetype_id": aid, "node_id": nd["node_id"], "l2_key": l2, "depth": depth,
                           "facets": json.dumps(dict(zip(fs, vs))),
                           "name": archetype_name(fs, vs) if depth else "Mixed / not classifiable",
                           "n_staples": int((df.loc[m, "retailer"] == "staples").sum()),
                           "n_competitor": int((df.loc[m, "retailer"] == comp).sum())})
            members += [(f, aid) for f in df.loc[m, "family_id"]]
        # validation against unsupervised structure on full-card embeddings
        t2 = tier2_attrs(l2)
        E = embed.encode([neutral_card(r, t2) for _, r in df.iterrows()])
        lab = arche.reindex(df.index).values
        v = validate(E, lab, c, rng)
        qa[nd["node_id"]] = {"facets": facets, "facet_scores": fscores, "price_breaks": br,
                             "n_archetypes": int(arche.nunique()), "mixed_share": float((arche == "0::mixed").mean()),
                             **v}
        log(f"S6 {nd['node_id']}: {arche.nunique()} archetypes on {facets}; ARI={v['ari']}")
    save(att, "family_table.parquet")          # attributes + family info + node price band
    save(pd.DataFrame(arches), "archetypes.parquet")
    save(pd.DataFrame(members, columns=["family_id", "archetype_id"]), "archetype_members.parquet")
    save_json(qa, "qa_archetypes.json")
    return qa
