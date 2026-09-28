"""Shared utilities: paths, config, IO, logging, statistics."""
from __future__ import annotations

import hashlib
import json
import re
import time
from functools import lru_cache
from pathlib import Path

import numpy as np
import pandas as pd
import yaml

ROOT = Path(__file__).resolve().parents[2]
CONFIG_DIR = ROOT / "config"


# ----------------------------------------------------------------------------- config
@lru_cache(maxsize=None)
def cfg() -> dict:
    return yaml.safe_load((CONFIG_DIR / "pipeline.yaml").read_text(encoding="utf-8"))


@lru_cache(maxsize=None)
def load_yaml(rel: str) -> dict:
    p = CONFIG_DIR / rel
    return yaml.safe_load(p.read_text(encoding="utf-8")) if p.exists() else {}


def slug(s: str) -> str:
    return re.sub(r"[^a-z0-9]+", "_", s.lower()).strip("_")


def l2_config(l2_key: str) -> dict:
    """Per-L2 generated artefact; {} (provisional, universal-only) when none exists."""
    name = "l2/" + "__".join(slug(p) for p in l2_key.split(" > ")) + ".yaml"
    return load_yaml(name)


# ----------------------------------------------------------------------------- paths & IO
def path(key: str) -> Path:
    p = ROOT / cfg()["paths"][key]
    p.mkdir(parents=True, exist_ok=True)
    return p


def interim(name: str) -> Path:
    return path("interim_dir") / name


def out(*parts: str) -> Path:
    p = path("out_dir").joinpath(*parts)
    p.parent.mkdir(parents=True, exist_ok=True)
    return p


def save(df: pd.DataFrame, name: str) -> None:
    df.to_parquet(interim(name), index=False)


def load(name: str) -> pd.DataFrame:
    return pd.read_parquet(interim(name))


def save_json(obj, name: str) -> None:
    interim(name).write_text(json.dumps(obj, indent=2, default=_json_default), encoding="utf-8")


def load_json(name: str):
    return json.loads(interim(name).read_text(encoding="utf-8"))


def _json_default(o):
    if isinstance(o, (np.integer,)):
        return int(o)
    if isinstance(o, (np.floating,)):
        return None if np.isnan(o) else float(o)
    if isinstance(o, np.ndarray):
        return o.tolist()
    return str(o)


def hash_text(s: str) -> str:
    return hashlib.md5(s.encode("utf-8")).hexdigest()


_T0 = time.time()


def log(msg: str) -> None:
    print(f"[{time.time() - _T0:7.1f}s] {msg}", flush=True)


# ----------------------------------------------------------------------------- statistics
def pct_rank(x: pd.Series) -> pd.Series:
    """Percentile rank in [0, 1]; NaN stays NaN; a single value maps to 0.5."""
    x = pd.Series(x, dtype=float)
    n = x.notna().sum()
    if n <= 1:
        return x.where(x.isna(), 0.5)
    return (x.rank(method="average") - 1) / (n - 1)


def pct_of(values: np.ndarray, reference: np.ndarray) -> np.ndarray:
    """Percentile (0-1) of each value within a reference distribution."""
    ref = np.sort(np.asarray(reference, dtype=float))
    ref = ref[~np.isnan(ref)]
    if len(ref) == 0:
        return np.full(len(values), 0.5)
    return np.searchsorted(ref, values, side="right") / len(ref)


def jeffreys_mean(k, n):
    return (np.asarray(k, float) + 0.5) / (np.asarray(n, float) + 1.0)


def prob_greater(k1, n1, k2, n2, draws: int, rng: np.random.Generator) -> float:
    """Pr(p1 > p2) under independent Beta(k+0.5, n-k+0.5) posteriors (Jeffreys prior)."""
    a = rng.beta(k1 + 0.5, n1 - k1 + 0.5, draws)
    b = rng.beta(k2 + 0.5, n2 - k2 + 0.5, draws)
    return float((a > b).mean())


def smoothed_dist(counts: pd.Series, support: list, alpha: float = 0.5) -> np.ndarray:
    c = np.array([counts.get(v, 0) for v in support], dtype=float) + alpha
    return c / c.sum()


def jsd(p: np.ndarray, q: np.ndarray) -> float:
    """Jensen-Shannon divergence, base 2, in [0, 1]."""
    m = 0.5 * (p + q)
    def kl(a, b):
        mask = a > 0
        return float(np.sum(a[mask] * np.log2(a[mask] / b[mask])))
    return 0.5 * kl(p, m) + 0.5 * kl(q, m)


def split_multi(v) -> list:
    if v is None or (isinstance(v, float) and np.isnan(v)) or v == "":
        return []
    return [t for t in str(v).split("|") if t]


def nice_breaks(qs: list[float]) -> list[float]:
    """Round price quantile cut points to merchant-friendly numbers."""
    steps = [5, 10, 15, 20, 25, 30, 40, 50, 60, 75, 100, 125, 150, 175, 200, 250, 300, 350, 400,
             450, 500, 600, 700, 750, 800, 900, 1000, 1250, 1500, 1750, 2000, 2500, 3000, 4000, 5000]
    out_ = []
    for q in qs:
        b = min(steps, key=lambda s: abs(s - q))
        if not out_ or b > out_[-1]:
            out_.append(b)
    return out_


def money(x: float) -> str:
    return f"${x:,.0f}"
