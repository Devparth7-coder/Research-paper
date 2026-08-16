"""Command-line interface for dataset generation, execution, and analysis."""
from __future__ import annotations

import argparse
import json
from pathlib import Path

import pandas as pd

from .analysis import analyze
from .dataset import generate_tasks
from .experiments import run_all
from .models import Task
from .utils import ensure_dirs, environment_manifest, load_yaml, read_jsonl, sha256


def _root() -> Path:
    return Path(__file__).resolve().parents[2]


def _config(root: Path, path: str) -> tuple[dict, Path]:
    config_path = (root / path).resolve() if not Path(path).is_absolute() else Path(path)
    return load_yaml(config_path), config_path


def _load_tasks(path: Path) -> list[Task]:
    return [Task.from_dict(row) for row in read_jsonl(path)]


def _write_manifest(root: Path, config_path: Path, data_path: Path) -> None:
    manifest = environment_manifest(root, config_path, data_path)
    tracked = [
        *sorted((root / "src").rglob("*.py")),
        *sorted((root / "configs").glob("*.yaml")),
        data_path,
    ]
    manifest["artifact_hashes"] = {
        str(path.relative_to(root)): sha256(path) for path in tracked if path.exists()
    }
    output = root / "results" / "raw" / "reproducibility_manifest.json"
    output.write_text(json.dumps(manifest, indent=2), encoding="utf-8")


def generate_command(config_path: str) -> None:
    root = _root(); config, cfg = _config(root, config_path); ensure_dirs(root)
    path = root / "data" / "tasks" / "tasks_v1.jsonl"
    tasks = generate_tasks(config, path)
    pd.DataFrame([task.to_dict() for task in tasks]).to_csv(root / "data" / "tasks" / "tasks_v1.csv", index=False)
    _write_manifest(root, cfg, path)
    print(f"Generated {len(tasks)} tasks at {path}")


def run_command(config_path: str) -> None:
    root = _root(); config, cfg = _config(root, config_path); ensure_dirs(root)
    data_path = root / "data" / "tasks" / "tasks_v1.jsonl"
    if not data_path.exists():
        generate_tasks(config, data_path)
    tasks = _load_tasks(data_path)
    episodes, attribution = run_all(tasks, config, root)
    _write_manifest(root, cfg, data_path)
    print(f"Executed {len(episodes)} episodes and {len(attribution)} attribution-view records")


def analyze_command(config_path: str) -> None:
    root = _root(); config, cfg = _config(root, config_path); ensure_dirs(root)
    data_path = root / "data" / "tasks" / "tasks_v1.jsonl"
    tasks = _load_tasks(data_path)
    raw = root / config["outputs"]["root"] / "raw"
    episodes = pd.read_csv(raw / "episodes.csv")
    attribution = pd.read_csv(raw / "attribution.csv")
    tables = analyze(episodes, attribution, pd.DataFrame([t.to_dict() for t in tasks]), config, root)
    _write_manifest(root, cfg, data_path)
    print(f"Wrote {len(tables)} measured tables, statistical outputs, and figures")


def all_command(config_path: str) -> None:
    root = _root(); config, cfg = _config(root, config_path); ensure_dirs(root)
    data_path = root / "data" / "tasks" / "tasks_v1.jsonl"
    tasks = generate_tasks(config, data_path)
    pd.DataFrame([task.to_dict() for task in tasks]).to_csv(root / "data" / "tasks" / "tasks_v1.csv", index=False)
    episodes, attribution = run_all(tasks, config, root)
    analyze(episodes, attribution, pd.DataFrame([t.to_dict() for t in tasks]), config, root)
    _write_manifest(root, cfg, data_path)
    print(f"Complete: {len(tasks)} tasks, {len(episodes)} episodes, measured analysis written")


def main() -> None:
    parser = argparse.ArgumentParser(prog="mas-reliab", description="MAS-RELIAB offline simulation pilot")
    parser.add_argument("command", choices=["generate-data", "run", "analyze", "all"])
    parser.add_argument("--config", default="configs/pilot.yaml")
    args = parser.parse_args()
    {"generate-data": generate_command, "run": run_command, "analyze": analyze_command, "all": all_command}[args.command](args.config)


if __name__ == "__main__":
    main()
