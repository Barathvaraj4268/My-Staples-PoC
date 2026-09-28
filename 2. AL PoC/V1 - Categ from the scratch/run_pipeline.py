"""Run the Staples AL PoC pipeline end to end, or from a given stage.

    python run_pipeline.py              # all stages
    python run_pipeline.py --from s6    # re-run from archetypes onward (earlier outputs are reused)
    python run_pipeline.py --only s11   # rebuild the report only
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent / "src"))

from alpoc import (s0_ingest, s1_families, s2_nodes, s3_mapping, s5_extract, s6_archetypes,  # noqa: E402
                   s7_vector, s8_gaps, s9_integrate, s10_skus, s11_report)
from alpoc.common import log  # noqa: E402

STAGES = [
    ("s0", "ingest & clean", s0_ingest.run),
    ("s1", "families & dedup", s1_families.run),
    ("s2", "analysis nodes (define)", s2_nodes.define),
    ("s3", "product-level mapping", s3_mapping.run),
    ("s2f", "analysis nodes (finalize)", s2_nodes.finalize),
    ("s5", "attribute extraction + validation", s5_extract.run),
    ("s6", "archetypes", s6_archetypes.run),
    ("s7", "method 1: vector view (VOS)", s7_vector.run),
    ("s8", "method 2: attribute gaps (TG)", s8_gaps.run),
    ("s9", "integration: gate + fused re-rank", s9_integrate.run),
    ("s10", "SKU exemplars & sellers", s10_skus.run),
    ("s11", "figures + HTML report", s11_report.run),
]


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--from", dest="start", default="s0")
    ap.add_argument("--only", default=None)
    a = ap.parse_args()
    keys = [k for k, _, _ in STAGES]
    todo = [a.only] if a.only else keys[keys.index(a.start):]
    for k, name, fn in STAGES:
        if k in todo:
            log(f"=== {k}: {name}")
            fn()
    log("done")


if __name__ == "__main__":
    main()
