"""Deterministic generation of objectively checkable synthetic tasks."""
from __future__ import annotations

from collections import Counter
from pathlib import Path
from typing import Any

import numpy as np

from .models import Task
from .utils import stable_seed, write_jsonl

FAMILIES = [
    "arithmetic_pipeline",
    "shard_aggregation",
    "evidence_synthesis",
    "constraint_decision",
    "stateful_workflow",
]


def combine_values(values: list[Any], operation: str) -> Any:
    """Associative state combination used for both gold labels and execution."""
    if operation == "sum":
        return int(sum(int(v) for v in values))
    if operation == "vector_sum":
        width = len(values[0]) if values else 2
        return [int(sum(v[i] for v in values)) for i in range(width)]
    if operation == "bitwise_or":
        result = 0
        for value in values:
            result |= int(value)
        return result
    if operation == "dict_sum":
        keys = sorted({key for value in values for key in value})
        return {key: int(sum(value.get(key, 0) for value in values)) for key in keys}
    if operation == "concat":
        return [item for value in values for item in value]
    raise ValueError(f"Unknown operation: {operation}")


def _make_task(index: int, family: str, split: str, master_seed: int) -> Task:
    local = np.random.default_rng(stable_seed(master_seed, "task", index, family))
    difficulty = 1 + (index // len(FAMILIES)) % 3
    n = 4 + difficulty
    if family == "arithmetic_pipeline":
        partials = [int(x) for x in local.integers(-20, 31, size=n)]
        operation = "sum"
        payload = {"terms": partials, "instruction": "sum all signed terms"}
    elif family == "shard_aggregation":
        partials = [int(x) for x in local.integers(5, 101, size=n)]
        operation = "sum"
        payload = {"shard_counts": partials, "instruction": "aggregate all shard counts"}
    elif family == "evidence_synthesis":
        partials = [[int(a), int(b)] for a, b in local.integers(0, 6, size=(n, 2))]
        operation = "vector_sum"
        payload = {"evidence_vectors": partials, "labels": ["support", "against"]}
    elif family == "constraint_decision":
        partials = [int(1 << int(x)) for x in local.integers(0, 8, size=n)]
        operation = "bitwise_or"
        payload = {"constraint_masks": partials, "instruction": "return the union bitmask"}
    elif family == "stateful_workflow":
        keys = ["queued", "processed", "flagged"]
        partials = []
        for _ in range(n):
            partials.append({key: int(local.integers(0, 5)) for key in keys})
        operation = "dict_sum"
        payload = {"state_deltas": partials, "instruction": "apply all state deltas"}
    else:
        raise ValueError(family)
    gold = combine_values(partials, operation)
    task_id = f"{split}-{index:04d}-{family[:3]}"
    return Task(task_id, family, difficulty, payload, partials, operation, gold, split)


def generate_tasks(config: dict, output_path: Path) -> list[Task]:
    spec = config["dataset"]
    master_seed = int(config["study"]["master_seed"])
    counts = spec["split_counts"]
    splits = [name for name, count in counts.items() for _ in range(int(count))]
    n_tasks = int(spec["n_tasks"])
    if len(splits) != n_tasks:
        raise ValueError("dataset.n_tasks must equal the sum of split_counts")
    tasks = []
    family_counter = Counter()
    for index, split in enumerate(splits):
        family = spec["families"][index % len(spec["families"])]
        tasks.append(_make_task(index, family, split, master_seed))
        family_counter[family] += 1
    write_jsonl(output_path, [task.to_dict() for task in tasks])
    return tasks
