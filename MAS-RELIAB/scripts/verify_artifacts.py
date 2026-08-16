#!/usr/bin/env python3
"""Fail when expected reproducibility artifacts are absent or internally inconsistent."""
from pathlib import Path
import json
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
required = [
    ROOT / "data/tasks/tasks_v1.jsonl",
    ROOT / "results/raw/episodes.csv",
    ROOT / "results/raw/events.jsonl",
    ROOT / "results/raw/reproducibility_manifest.json",
    ROOT / "results/analysis/statistical_tests.csv",
    ROOT / "results/analysis/hypothesis_decisions.csv",
    ROOT / "results/figures/figure_06_tradeoff.png",
    ROOT / "results/tables/table_08_tradeoff.csv",
]
missing = [str(path.relative_to(ROOT)) for path in required if not path.exists() or path.stat().st_size == 0]
if missing:
    raise SystemExit(f"Missing/empty artifacts: {missing}")
episodes = pd.read_csv(ROOT / "results/raw/episodes.csv")
counts = json.loads((ROOT / "results/raw/run_counts.json").read_text())
assert len(episodes) == counts["episode_count"]
assert set(episodes.topology.unique()) == {"single", "sequential", "parallel", "hierarchical"}
assert episodes.seed.notna().all()
assert episodes[episodes.injection_scheduled].injection_consumed.mean() > 0.8
print(f"Verified {len(required)} artifacts and {len(episodes)} episode records")
