#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
=====================================================================================
 STAPLES MARKETPLACE PoC  -  CATEGORY-LEVEL RECOMMENDATION  (v2, six retailers)
 METHOD V : VECTOR-FIRST  (path-aware embeddings + Qdrant vector DB, retrieve-then-rerank)
=====================================================================================

Plain-language summary
----------------------
Every category page on every retailer's site is a "shelf label". Method V turns every
label into a vector - using the label itself plus whatever parent labels the retailer
shows above it (one level for Wayfair rugs, four for Walmart) - stores the vectors in
Qdrant, and for every competitor shelf asks: "which Staples shelf is closest, and is it
close enough to be the same shelf?"

 1. LOAD      six trees -> one node table (poc_common.load_all): scope, headings, counts
 2. EMBED     label + decaying ancestor context (level-agnostic)            [section 2]
 3. INDEX     one Qdrant collection per retailer (persisted in outputs/V/qdrant_db)
 4. MATCH     top-50 Staples shelves by vector, re-scored with wording (TF-IDF)
 5. CALIBRATE per competitor on labelled rows (+ ablation negatives): tau, gap screen
 6. SCORE     per competitor shelf: status (matched / likely / gap), AAS (adjacency to the
              Staples neighbourhood) and CRS (closeness to a 1P CORE shelf)
 7. WRITE     outputs/V/node_scores_V.csv (+ calibration, encoder bake-off, vectors)

run_framework.py turns these node scores (and Method G's) into recommendations.

Run
---
 Colab : upload poc_common.py, this file, the six .xlsx trees and gold_matches_v2.csv to
         /content, then
            !pip install -q qdrant-client spacy networkx openpyxl
            !python -m spacy download en_core_web_lg
            !python method_v_vector.py
 Local : python method_v_vector.py   (data in ../data, outputs in ../outputs by default;
         override with env vars POC_DATA_DIR / POC_OUT_DIR)
"""
# %% [0] INSTALL (uncomment in Colab) ---------------------------------------------------
# !pip install -q pandas numpy scipy scikit-learn openpyxl matplotlib qdrant-client spacy
# !python -m spacy download en_core_web_lg
# optional encoders:  !pip install -q sentence-transformers     (encoder="st", needs HuggingFace)
#                     !pip install -q wordllama                 (encoder="wordllama")

# %% [1] IMPORTS ------------------------------------------------------------------------
import math
import os
import shutil
import time
from collections import defaultdict

import numpy as np
import pandas as pd

import poc_common as pc
from poc_common import CONFIG, FOCAL, log

OUT = "V"


# %% [2] ENCODERS -----------------------------------------------------------------------
class Encoder:
    """Pluggable label encoder. Downstream thresholds are re-calibrated for whichever is used."""

    def __init__(self, kind, corpus):
        self.kind = kind
        getattr(self, f"_init_{kind}")(corpus)

    # spaCy en_core_web_lg: 300-d word vectors, IDF-weighted mean (offline once installed)
    def _init_spacy(self, corpus):
        import spacy
        self.nlp = spacy.load(CONFIG["spacy_model"], exclude=["tok2vec", "tagger", "parser", "ner",
                                                              "lemmatizer", "attribute_ruler", "senter"])
        if self.nlp.vocab.vectors.shape[0] < 50000:
            raise RuntimeError("model has too few vectors - install en_core_web_lg")
        df = defaultdict(int)
        for t in corpus:
            for w in set(pc.tokens(t)):
                df[w] += 1
        n = len(corpus)
        self.idf = {w: math.log((1 + n) / (1 + c)) + 1 for w, c in df.items()}
        self.dim = self.nlp.vocab.vectors.shape[1]

    def _enc_spacy(self, texts):
        out = np.zeros((len(texts), self.dim), dtype=np.float32)
        for i, t in enumerate(texts):
            acc, wsum = np.zeros(self.dim, dtype=np.float32), 0.0
            for w in pc.tokens(t):
                lex = self.nlp.vocab[w]
                if lex.has_vector:
                    wt = self.idf.get(w, 1.0)
                    acc += wt * lex.vector
                    wsum += wt
            out[i] = acc / wsum if wsum else 0
        return out

    # sentence-transformers (bge-small / MiniLM): the recommended encoder where HuggingFace is reachable
    def _init_st(self, corpus):
        from sentence_transformers import SentenceTransformer
        self.model = SentenceTransformer(CONFIG["st_model"])

    def _enc_st(self, texts):
        return self.model.encode(list(texts), batch_size=128, normalize_embeddings=True,
                                 show_progress_bar=False).astype(np.float32)

    # WordLlama (weights ship inside the pip wheel - no download)
    def _init_wordllama(self, corpus):
        import pathlib
        import wordllama
        from wordllama import WordLlama
        pkg = pathlib.Path(wordllama.__file__).parent
        cache = pathlib.Path.home() / ".cache" / "wordllama" / "tokenizers"
        cache.mkdir(parents=True, exist_ok=True)
        src = pkg / "tokenizers" / "l2_supercat_tokenizer_config.json"
        if src.exists() and not (cache / src.name).exists():
            shutil.copy(src, cache / src.name)       # the wheel's own file; works around a path bug
        self.wl = WordLlama.load(disable_download=True)

    def _enc_wordllama(self, texts):
        return np.asarray(self.wl.embed(list(texts), norm=True), dtype=np.float32)

    # TF-IDF character n-grams + SVD: no downloads at all (weakest; last resort)
    def _init_tfidf(self, corpus):
        from sklearn.decomposition import TruncatedSVD
        from sklearn.feature_extraction.text import TfidfVectorizer
        self.vec = TfidfVectorizer(analyzer="char_wb", ngram_range=(3, 5), sublinear_tf=True)
        X = self.vec.fit_transform([t.lower() for t in corpus])
        self.svd = TruncatedSVD(n_components=min(256, X.shape[1] - 1), random_state=0).fit(X)

    def _enc_tfidf(self, texts):
        return self.svd.transform(self.vec.transform([t.lower() for t in texts])).astype(np.float32)

    def encode(self, texts):
        E = getattr(self, f"_enc_{self.kind}")(texts)
        n = np.linalg.norm(E, axis=1, keepdims=True)
        n[n == 0] = 1
        return E / n


def node_vectors(nodes, enc, w_self=None, decay=None, cache=None):
    """Level-agnostic node vector: own label + ancestors with geometrically decaying weight.
         v = w_self * e(label) + (1 - w_self) * sum_k decay^(k-1) e(ancestor_k) / sum_k decay^(k-1)
    A Wayfair rug (2 levels) and a Walmart L4 shelf (4 levels) are built the same way, from
    whatever context their retailer shows; navigation headings were already skipped."""
    w_self = CONFIG["w_self"] if w_self is None else w_self
    decay = CONFIG["ctx_decay"] if decay is None else decay
    chains = [[ne] + ch[1:] for ne, ch in zip(nodes["name_expanded"], nodes["sem_chain"])]
    if cache is None:
        uniq = sorted({t for ch in chains for t in ch})
        E = enc.encode(uniq)
        cache = {t: E[i] for i, t in enumerate(uniq)}
    dim = len(next(iter(cache.values())))
    V = np.zeros((len(chains), dim), dtype=np.float32)
    for i, ch in enumerate(chains):
        v = cache[ch[0]].copy()
        if len(ch) > 1:
            ws = np.array([decay ** k for k in range(len(ch) - 1)])
            ws = (1 - w_self) * ws / ws.sum()
            v = w_self * v + sum(w * cache[t] for w, t in zip(ws, ch[1:]))
        V[i] = v
    V /= np.clip(np.linalg.norm(V, axis=1, keepdims=True), 1e-9, None)
    return V, cache


# %% [3] VECTOR DB ----------------------------------------------------------------------
class VectorDB:
    """Qdrant (local mode, persisted to disk) with ONE COLLECTION PER RETAILER.
    Hub-and-spoke queries always target one retailer, so per-retailer collections avoid
    payload filtering (which local mode evaluates in Python and is ~40x slower here) and
    mirror how a Qdrant server / cloud deployment would shard the SKU-level index later.
    Falls back to NumPy brute force (identical results) if qdrant-client is missing."""

    def __init__(self, path, dim):
        self.dim, self.backend, self.store = dim, "qdrant", {}
        try:
            from qdrant_client import QdrantClient
            if os.path.exists(path):
                shutil.rmtree(path)
            self.client = QdrantClient(path=path)
        except Exception as e:  # noqa
            log(f"  qdrant unavailable ({e}); using numpy")
            self.backend = "numpy"

    def add(self, retailer, vectors, payloads):
        self.store[retailer] = (vectors, payloads)
        if self.backend == "qdrant":
            from qdrant_client.models import Distance, VectorParams
            name = retailer.lower()
            self.client.create_collection(name, vectors_config=VectorParams(size=self.dim, distance=Distance.COSINE))
            self.client.upload_collection(name, vectors=vectors, ids=list(range(len(vectors))),
                                          payload=payloads, batch_size=512)

    def search(self, retailer, Q, k):
        """(indices into that retailer's vectors, cosine scores), each shaped (len(Q), k)."""
        n = len(self.store[retailer][0])
        k = min(k, n)
        if self.backend == "qdrant":
            from qdrant_client.models import QueryRequest
            I = np.zeros((len(Q), k), dtype=int)
            S = np.zeros((len(Q), k), dtype=np.float32)
            for s in range(0, len(Q), 1000):
                reqs = [QueryRequest(query=q.tolist(), limit=k) for q in Q[s:s + 1000]]
                for j, res in enumerate(self.client.query_batch_points(retailer.lower(), requests=reqs)):
                    I[s + j] = [p.id for p in res.points]
                    S[s + j] = [p.score for p in res.points]
            return I, S
        M = Q @ self.store[retailer][0].T
        I = np.argsort(-M, axis=1)[:, :k]
        return I, np.take_along_axis(M, I, axis=1)


# %% [4] MATCHING -----------------------------------------------------------------------
def match_to_staples(cn, focal, Qv, focal_V, lex, focal_L, db=None, k=None):
    """Stage 1: top-k Staples shelves by vector.  Stage 2: hybrid rerank
       score = w_sem * cosine + w_lex * wording.  Returns sorted (I, S, S_sem, S_lex)."""
    k = k or CONFIG["k_retrieve"]
    if db is not None:
        I, S_sem = db.search(FOCAL, Qv, k)
    else:
        M = Qv @ focal_V.T
        I = np.argsort(-M, axis=1)[:, :k]
        S_sem = np.take_along_axis(M, I, axis=1)
    CL = lex.transform(cn["name_expanded"])
    S_lex = np.zeros_like(S_sem)
    for i in range(len(cn)):
        S_lex[i] = (focal_L[I[i]] @ CL[i].T).toarray().ravel()
    S = CONFIG["w_sem"] * S_sem + CONFIG["w_lex"] * S_lex
    order = np.argsort(-S, axis=1)
    return tuple(np.take_along_axis(a, order, axis=1) for a in (I, S, S_sem, S_lex)) + (CL,)


def top1_on_gold(gold, cn_paths, best_paths):
    bp = dict(zip(cn_paths, best_paths))
    ok = n = 0
    for _, r in gold[(gold.staples_carries == 1) & gold.gold_staples_path.notna()].iterrows():
        pred = bp.get(r["comp_path"])
        if pred is None:
            continue
        n += 1
        alts = str(r["gold_staples_path"]).split("|")
        ok += any(pred == a or pred.startswith(a + " > ") or a.startswith(pred + " > ") for a in alts)
    return ok / max(n, 1), n


# %% [5] MODEL SELECTION: encoder bake-off + context-weight grid ---------------------------
def bakeoff(nodes, focal, lex, focal_L, gold):
    """Top-1 shelf accuracy on the labelled rows for every encoder that loads, and for a small
    grid of context weights. Evidence for the model choice, recorded in encoder_bakeoff.csv."""
    labelled = gold[(gold.staples_carries == 1) & gold.gold_staples_path.notna()]
    rows, grid = [], []
    un = nodes[nodes.unit_ok]
    corpus = sorted(set(un["name_expanded"]) | {t for ch in un["sem_chain"] for t in ch})
    comp_nodes = nodes[nodes.retailer.isin(labelled.competitor.unique()) & nodes.unit_ok]
    comp_nodes = comp_nodes[comp_nodes.path_str.isin(labelled.comp_path)]
    sub = pd.concat([focal.drop(columns=["index"], errors="ignore"), comp_nodes], ignore_index=True)
    nf = len(focal)
    for kind in CONFIG["encoder_candidates"]:
        t0 = time.time()
        try:
            enc = Encoder(kind, corpus)
        except Exception as e:  # noqa
            rows.append({"encoder": kind, "status": f"unavailable ({type(e).__name__})"})
            log(f"      {kind:<10} unavailable ({type(e).__name__}: {str(e)[:70]})")
            continue
        Vs, cache = node_vectors(sub, enc)
        Vs[:nf], _ = node_vectors(sub.iloc[:nf], enc, CONFIG["w_self_focal"], CONFIG["ctx_decay_focal"], cache)
        fV = Vs[:nf]
        for comp, g in labelled.groupby("competitor"):
            idx = np.where(sub["retailer"].values[nf:] == comp)[0] + nf
            if not len(idx):
                continue
            cn = sub.iloc[idx].reset_index(drop=True)
            Qv = Vs[idx]
            I, S, S_sem, S_lex, _ = match_to_staples(cn, focal, Qv, fV, lex, focal_L)
            acc_h, n = top1_on_gold(g, cn["path_str"], focal["path_str"].values[I[:, 0]])
            acc_s, _ = top1_on_gold(g, cn["path_str"], focal["path_str"].values[np.argmax(Qv @ fV.T, axis=1)])
            rows.append({"encoder": kind, "status": "ok", "competitor": comp, "n_labelled": n,
                         "top1_vector_only": acc_s, "top1_hybrid": acc_h, "seconds": round(time.time() - t0, 1)})
            log(f"      {kind:<10} {comp:<12} top-1 hybrid {acc_h:.1%} | vector only {acc_s:.1%} (n={n})")
        if kind == CONFIG["encoder"] or (CONFIG["encoder"] == "auto" and kind == "spacy"):
            for ws_c in (0.3, 0.4, 0.55, 1.0):
                for dc in (0.35, 0.5):
                    Vc, _ = node_vectors(sub, enc, ws_c, dc, cache)
                    for ws_f in (0.55, 0.7, 0.85, 1.0):
                        fVg, _ = node_vectors(sub.iloc[:nf], enc, ws_f, 0.5, cache)
                        accs = []
                        for comp, g in labelled.groupby("competitor"):
                            idx = np.where(sub["retailer"].values[nf:] == comp)[0] + nf
                            if not len(idx):
                                continue
                            cn = sub.iloc[idx].reset_index(drop=True)
                            I, *_ = match_to_staples(cn, focal, Vc[idx], fVg, lex, focal_L)
                            accs.append(top1_on_gold(g, cn["path_str"], focal["path_str"].values[I[:, 0]]))
                        tot = sum(a * n for a, n in accs) / max(sum(n for _, n in accs), 1)
                        grid.append({"encoder": kind, "w_self_competitor": ws_c, "ctx_decay_competitor": dc,
                                     "w_self_staples": ws_f, "top1_hybrid_pooled": tot})
    return pd.DataFrame(rows), pd.DataFrame(grid)


# %% [6] MAIN ---------------------------------------------------------------------------
def main():
    t_start = time.time()
    log("\n[1] Loading the six navigation trees")
    nodes, xedges, totals, files = pc.load_all()
    gold = pc.load_gold(for_calibration=True)
    focal = nodes[(nodes.retailer == FOCAL) & nodes.unit_ok].reset_index()
    focal["coreness"] = pc.staples_coreness(focal)
    lex = pc.Lexical(sorted(set(nodes.loc[nodes.in_scope, "name_expanded"])))
    focal_L = lex.transform(focal["name_expanded"])

    log("\n[2] Encoder bake-off and context-weight grid (top-1 shelf accuracy on labelled rows)")
    bake, grid = bakeoff(nodes, focal, lex, focal_L, gold)
    bake.to_csv(pc.out_path(OUT, "encoder_bakeoff.csv"), index=False)
    grid.to_csv(pc.out_path(OUT, "context_weight_grid.csv"), index=False)
    kind = CONFIG["encoder"]
    if kind == "auto":
        ok = bake[bake.status == "ok"].groupby("encoder").apply(
            lambda d: np.average(d.top1_hybrid, weights=d.n_labelled))
        kind = ok.idxmax()
    log(f"      encoder in use: {kind}")

    log("\n[3] Embedding every in-scope node (label + decaying ancestor context)")
    work = nodes[nodes.unit_ok].reset_index(drop=False)
    corpus = sorted(set(work["name_expanded"]) | {t for ch in work["sem_chain"] for t in ch})
    enc = Encoder(kind, corpus)
    V, cache = node_vectors(work, enc)                                        # competitor weighting
    is_f = work["retailer"].eq(FOCAL).values
    Vf, _ = node_vectors(work[is_f], enc, CONFIG["w_self_focal"], CONFIG["ctx_decay_focal"], cache)
    V[is_f] = Vf                                                              # Staples weighting
    np.savez_compressed(pc.out_path(OUT, "vectors_V.npz"), uid=work["uid"].values, V=V)
    log(f"      {len(work):,} vectors, dim {V.shape[1]}")

    log("\n[4] Loading the Qdrant vector DB (one collection per retailer)")
    db = VectorDB(pc.out_path(OUT, "qdrant_db"), V.shape[1])
    pos_of = {u: i for i, u in enumerate(work["uid"])}
    focal_V = V[[pos_of[u] for u in focal["uid"]]]
    for r in [FOCAL] + pc.COMPETITORS:
        sl = work[work.retailer == r]
        db.add(r, V[sl.index.values], [{"uid": u, "path": p, "depth": int(d), "has_count": bool(c)}
                                       for u, p, d, c in zip(sl["uid"], sl["path_str"], sl["depth"], sl["count_known"])])
    log(f"      backend: {db.backend}; collections: {', '.join(r.lower() for r in [FOCAL] + pc.COMPETITORS)}")

    log("\n[5] Retrieve-then-rerank each competitor against the Staples spine + calibrate")
    all_scores, cal_rows, pooled = [], [], []
    calibs = {}
    for comp in pc.COMPETITORS:
        cn = work[work.retailer == comp].reset_index(drop=True)
        Qv = V[[pos_of[u] for u in cn["uid"]]]
        t0 = time.time()
        I, S, S_sem, S_lex, CL = match_to_staples(cn, focal, Qv, focal_V, lex, focal_L, db)
        cn["best_score"], cn["best_sem"], cn["best_lex"] = S[:, 0], S_sem[:, 0], S_lex[:, 0]
        cn["best_staples"] = focal["path_str"].values[I[:, 0]]
        cn["best_staples_uid"] = focal["uid"].values[I[:, 0]]
        cn["alt_staples"] = [" | ".join(focal["path_str"].values[I[i, 1:4]]) for i in range(len(cn))]
        cn["A_raw"] = S[:, :CONFIG["k_adjacency"]].mean(axis=1)
        gc = gold[gold.competitor == comp]
        if len(gc) >= CONFIG["min_gold_rows"]:
            row_of = dict(zip(cn["path_str"], range(len(cn))))
            abl = {}
            for _, r in gc[(gc.staples_carries == 1) & gc.gold_staples_path.notna()].iterrows():
                i = row_of.get(r["comp_path"])
                if i is None:
                    continue
                mask = pc.ablation_mask(focal, r["gold_staples_path"], r["comp_path"])
                sem = focal_V[mask] @ Qv[i]
                lx = (focal_L[mask] @ CL[i].T).toarray().ravel()
                abl[r["comp_path"]] = float((CONFIG["w_sem"] * sem + CONFIG["w_lex"] * lx).max())
            cal, cal_df = pc.calibrate(gc, dict(zip(cn["path_str"], cn["best_score"])), abl,
                                       dict(zip(cn["path_str"], cn["best_staples"])))
            cal["source"] = "own labels"
            cal_df["competitor"] = comp
            pooled.append(cal_df)
        else:
            cal = None
        calibs[comp] = cal
        all_scores.append((comp, cn, I, S))
        log(f"      {comp:<12} {len(cn):>5} shelves matched in {time.time() - t0:5.1f}s"
            + (f" | tau {cal['tau']:.3f}, gap screen {cal['s_enter']:.3f}, AUC {cal['auc']:.3f}, "
               f"top-1 {cal['top1']:.1%} (n={cal['n_top1']})" if cal else " | no labels yet"))

    # pooled calibration for competitors without enough labels (same scorer, same spine)
    pool = None
    if pooled:
        pdf = pd.concat(pooled)
        y, s = pdf["y"].values, pdf["score"].values
        cands = [(float(t), CONFIG["false_gap_cost"] * (s[y == 1] >= t).mean() - (s[y == 0] >= t).mean())
                 for t in np.unique(np.round(s, 4))]
        tau_p = max(cands, key=lambda z: z[1])[0]
        pos = pdf[(pdf.kind == "gold") & (pdf.y == 1)]["score"]
        pool = {"tau": tau_p, "s_enter": min(float(pos.quantile(CONFIG["gap_screen_false_alarm"])), tau_p),
                "neg_median": float(pdf.loc[pdf.y == 0, "score"].median()), "source": "pooled labels"}

    scored = []
    for comp, cn, I, S in all_scores:
        cal = calibs[comp] or pc.fallback_calibration(cn["best_score"].values, pool)
        calibs[comp] = cal
        tau, s_enter, lo = cal["tau"], cal["s_enter"], cal["neg_median"]
        cn["status"] = pc.status_from_scores(cn["best_score"].values, tau, s_enter)
        m = CONFIG["multi_match_margin"]
        cn["match_set"] = [" | ".join(focal["uid"].values[I[i, j]] for j in range(I.shape[1])
                                      if S[i, j] >= tau and S[i, j] >= S[i, 0] - m) for i in range(len(cn))]
        # CRS: 0 = a typical non-match, 100 = a certain match to a fully core 1P shelf (max-pool)
        f = np.clip((S - lo) / max(tau - lo, 1e-6), 0, 1)
        core = focal["coreness"].values[I]
        cn["crs_raw"] = (f * core).max(axis=1)
        cn["crs_driver"] = [focal["path_str"].values[I[i, np.argmax(f[i] * core[i])]] for i in range(len(cn))]
        cn["CRS_leaf"] = 100 * cn["crs_raw"]
        scored.append((comp, cn, cal))
    # adjacency on one pooled scale (all competitors' leaves), then roll up each competitor tree
    allc = pd.concat([x[1] for x in scored])
    lf = allc["is_leaf"].values.astype(bool)
    scale = pc.aas_scaler(allc["A_raw"].values[lf], allc["status"].values[lf] == "matched")
    for comp, cn, cal in scored:
        cn["AAS_leaf"] = scale(cn["A_raw"].values)
        cn = pc.rollup_competitor(cn)
        cn["competitor"] = comp
        cn["method"] = "V"
        keep = ["competitor", "method", "uid", "path_str", "sem_path", "depth", "is_leaf", "n_leaves",
                "subtree_items", "count_known", "best_staples", "best_staples_uid", "best_score", "best_sem",
                "best_lex", "alt_staples", "status", "match_set", "A_raw", "AAS_leaf", "CRS_leaf", "coverage",
                "coverage_leaf", "coverage_items", "AAS", "CRS", "crs_driver", "n_gap_leaves", "gap_leaves_example"]
        all_scores_df = cn[keep]
        all_scores_df.to_csv(pc.out_path(OUT, f"node_scores_V_{comp}.csv"), index=False)
        cal_rows.append({"competitor": comp, "method": "V", "encoder": kind, **{k: v for k, v in cal.items()}})
        st = cn["status"].value_counts(normalize=True)
        log(f"      {comp:<12} matched {st.get('matched', 0):.0%} | likely {st.get('likely', 0):.0%} | "
            f"gap {st.get('gap', 0):.0%}   ({cal.get('source', cal.get('note', ''))})")
    pd.concat([pd.read_csv(pc.out_path(OUT, f"node_scores_V_{c}.csv")) for c in pc.COMPETITORS]).to_csv(
        pc.out_path(OUT, "node_scores_V.csv"), index=False)
    pd.DataFrame(cal_rows).to_csv(pc.out_path(OUT, "calibration_V.csv"), index=False)
    if pooled:
        pd.concat(pooled).to_csv(pc.out_path(OUT, "calibration_rows_V.csv"), index=False)
    pc.save_json({"encoder": kind, "vector_backend": db.backend, "n_vectors": int(len(work)),
                  "dim": int(V.shape[1]), "totals": totals, "seconds": round(time.time() - t_start, 1)},
                 OUT, "run_info_V.json")
    log(f"\nDone in {time.time() - t_start:.0f}s. Outputs in {pc.out_path(OUT, '')}")


if __name__ == "__main__":
    main()
