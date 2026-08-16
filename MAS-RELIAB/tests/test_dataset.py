from pathlib import Path

from mas_reliab.dataset import generate_tasks
from mas_reliab.utils import load_yaml, read_jsonl


def test_dataset_is_deterministic(tmp_path: Path):
    config = load_yaml(Path(__file__).parents[1] / "configs" / "pilot.yaml")
    a, b = tmp_path / "a.jsonl", tmp_path / "b.jsonl"
    first = generate_tasks(config, a)
    second = generate_tasks(config, b)
    assert a.read_bytes() == b.read_bytes()
    assert len(first) == len(second) == 150
    assert all(task.gold is not None for task in first)


def test_split_counts(tmp_path: Path):
    config = load_yaml(Path(__file__).parents[1] / "configs" / "pilot.yaml")
    path = tmp_path / "tasks.jsonl"
    generate_tasks(config, path)
    rows = read_jsonl(path)
    assert sum(row["split"] == "evaluation" for row in rows) == 90
