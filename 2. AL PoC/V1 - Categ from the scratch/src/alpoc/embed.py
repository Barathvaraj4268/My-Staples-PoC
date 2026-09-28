"""Pluggable text encoder with an on-disk cache (re-runs never re-encode the same text)."""
from __future__ import annotations

import os
import warnings

import numpy as np
import pandas as pd

from .common import cfg, hash_text, log, path, slug

os.environ.setdefault("HF_HUB_DISABLE_SYMLINKS_WARNING", "1")
warnings.filterwarnings("ignore", message=".*symlinks.*")

_MODELS: dict = {}


def _model(name: str):
    if name not in _MODELS:
        from sentence_transformers import SentenceTransformer
        log(f"loading encoder {name}")
        _MODELS[name] = SentenceTransformer(name, device="cpu")
    return _MODELS[name]


def _cache_file(name: str):
    return path("cache_dir") / f"emb_{slug(name)}.parquet"


def encode(texts: list[str], model: str | None = None) -> np.ndarray:
    """L2-normalised embeddings for texts (float32), cached by text hash."""
    model = model or active_model()
    cf = _cache_file(model)
    cache = pd.read_parquet(cf) if cf.exists() else pd.DataFrame(columns=["h", "v"])
    known = dict(zip(cache["h"], cache["v"]))
    hashes = [hash_text(t) for t in texts]
    todo = sorted({h: t for h, t in zip(hashes, texts) if h not in known}.items())
    if todo:
        log(f"encoding {len(todo):,} new texts with {model}")
        vecs = _model(model).encode([t for _, t in todo], batch_size=cfg()["encoder"]["batch_size"],
                                    normalize_embeddings=True, show_progress_bar=False)
        new = pd.DataFrame({"h": [h for h, _ in todo], "v": [v.astype(np.float32).tolist() for v in vecs]})
        cache = pd.concat([cache, new], ignore_index=True)
        cache.to_parquet(cf, index=False)
        known.update(zip(new["h"], new["v"]))
    return np.asarray([known[h] for h in hashes], dtype=np.float32)


_ACTIVE = {"name": None}


def set_active_model(name: str) -> None:
    _ACTIVE["name"] = name


def active_model() -> str:
    if _ACTIVE["name"] is None:
        from .common import interim, load_json
        if interim("encoder.json").exists():          # chosen by the S3 bake-off
            _ACTIVE["name"] = load_json("encoder.json")["model"]
    return _ACTIVE["name"] or cfg()["encoder"]["default"]
