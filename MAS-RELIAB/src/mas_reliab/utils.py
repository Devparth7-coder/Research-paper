"""I/O, hashing, seeding, and exact-state helpers."""
from __future__ import annotations

import hashlib
import json
import platform
import random
import subprocess
import sys
from pathlib import Path
from typing import Any, Iterable

import numpy as np
import yaml


def load_yaml(path: str | Path) -> dict[str, Any]:
    with open(path, encoding="utf-8") as handle:
        return yaml.safe_load(handle)


def ensure_dirs(root: Path) -> None:
    for relative in ["data/tasks", "results/raw", "results/tables", "results/figures", "results/analysis"]:
        (root / relative).mkdir(parents=True, exist_ok=True)


def stable_seed(*parts: Any) -> int:
    text = "|".join(str(part) for part in parts).encode("utf-8")
    return int.from_bytes(hashlib.sha256(text).digest()[:8], "big") % (2**32 - 1)


def canonical(value: Any) -> str:
    return json.dumps(value, sort_keys=True, ensure_ascii=False, separators=(",", ":"))


def exact_equal(a: Any, b: Any) -> bool:
    return canonical(a) == canonical(b)


def write_jsonl(path: Path, rows: Iterable[dict[str, Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as handle:
        for row in rows:
            handle.write(json.dumps(row, ensure_ascii=False, sort_keys=True) + "\n")


def read_jsonl(path: Path) -> list[dict[str, Any]]:
    with open(path, encoding="utf-8") as handle:
        return [json.loads(line) for line in handle if line.strip()]


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with open(path, "rb") as handle:
        for block in iter(lambda: handle.read(65536), b""):
            digest.update(block)
    return digest.hexdigest()


def environment_manifest(root: Path, config_path: Path, data_path: Path) -> dict[str, Any]:
    try:
        git_commit = subprocess.check_output(
            ["git", "rev-parse", "HEAD"], cwd=root, stderr=subprocess.DEVNULL, text=True
        ).strip()
    except (subprocess.CalledProcessError, FileNotFoundError):
        git_commit = None
    packages = {}
    for name in ["numpy", "pandas", "scipy", "matplotlib", "seaborn", "networkx", "PyYAML"]:
        try:
            from importlib.metadata import version
            packages[name] = version(name)
        except Exception:
            packages[name] = "unavailable"
    return {
        "study_type": "offline stochastic synthetic-agent simulation pilot",
        "external_model_api": False,
        "python": sys.version,
        "platform": platform.platform(),
        "machine": platform.machine(),
        "processor": platform.processor() or "not reported by platform",
        "logical_cpu_count": __import__("os").cpu_count(),
        "packages": packages,
        "git_commit": git_commit,
        "config_file": str(config_path.relative_to(root)),
        "config_sha256": sha256(config_path),
        "dataset_file": str(data_path.relative_to(root)),
        "dataset_sha256": sha256(data_path),
        "randomness": "SHA-256-derived per-episode NumPy seeds; no hosted model sampling",
        "token_accounting": "not applicable; simulated work units are reported",
        "latency": "deterministic simulated milliseconds, not wall-clock/API latency",
    }


def rng(seed: int) -> np.random.Generator:
    random.seed(seed)
    return np.random.default_rng(seed)
