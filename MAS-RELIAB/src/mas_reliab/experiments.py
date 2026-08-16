"""Six-experiment pilot orchestration and trace-derived attribution records."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pandas as pd

from .engine import run_episode
from .models import EpisodeResult, FaultSpec, Task
from .topologies import build_plan, target_for_position
from .utils import stable_seed, write_jsonl


def _task_subset(tasks: list[Task], config: dict[str, Any]) -> list[Task]:
    evaluation = [task for task in tasks if task.split == "evaluation"]
    limit = int(config["dataset"]["measured_task_limit"])
    if len(evaluation) < limit:
        raise ValueError(f"Requested {limit} measured tasks but only {len(evaluation)} evaluation tasks exist")
    return evaluation[:limit]


def _fault(task: Task, topology: str, fault_type: str, position: str, seed: int, severity: int = 1) -> FaultSpec:
    plan = build_plan(topology, len(task.partials))
    target = target_for_position(plan, position)
    return FaultSpec(
        fault_id=f"fault::{task.task_id}::{topology}::{fault_type}::sev{severity}::{position}::{seed}",
        fault_type=fault_type,
        position=position,
        target_node=target,
        severity=severity,
        seed=seed,
    )


def _append_result(results: list[EpisodeResult], result: EpisodeResult) -> None:
    results.append(result)


def _run_e1(tasks: list[Task], config: dict, results: list[EpisodeResult]) -> None:
    sim = config["simulation"]
    master = config["study"]["master_seed"]
    for task in tasks:
        for repeat in range(int(sim["baseline_repeats"])):
            block = f"E1::{task.task_id}::{repeat}"
            for topology in sim["topologies"]:
                seed = stable_seed(master, block, topology)
                result = run_episode(
                    task, topology, seed, sim, "E1_baseline_reliability",
                    f"{block}::{topology}", block, condition="baseline"
                )
                _append_result(results, result)


def _run_e2(tasks: list[Task], config: dict, results: list[EpisodeResult]) -> None:
    sim = config["simulation"]
    master = config["study"]["master_seed"]
    for task in tasks:
        for fault_type in sim["fault_types"]:
            for severity in sim["severities"]:
                for position in sim["positions"]:
                    for repeat in range(int(sim["fault_repeats"])):
                        block = f"E2::{task.task_id}::{fault_type}::sev{severity}::{position}::{repeat}"
                        for topology in sim["topologies"]:
                            seed = stable_seed(master, block, topology)
                            fault = _fault(task, topology, fault_type, position, seed, int(severity))
                            result = run_episode(
                                task, topology, seed, sim, "E2_fault_propagation",
                                f"{block}::{topology}", block, condition="fault",
                                fault=fault, verification="none", recovery="none"
                            )
                            _append_result(results, result)


def _run_e4(tasks: list[Task], config: dict, results: list[EpisodeResult]) -> None:
    sim, exp = config["simulation"], config["experiments"]
    master = config["study"]["master_seed"]
    strategies = [("none", "none"), ("final", "retry"), ("local", "retry"), ("both", "retry")]
    for task in tasks:
        for fault_type in exp["verification_fault_types"]:
            for repeat in range(int(exp["verification_repeats"])):
                block = f"E4::{task.task_id}::{fault_type}::{repeat}"
                for topology in sim["topologies"]:
                    for verification, recovery in strategies:
                        seed = stable_seed(master, block, topology, verification)
                        position = "early"
                        fault = _fault(task, topology, fault_type, position, seed)
                        condition = f"verify={verification};recovery={recovery}"
                        result = run_episode(
                            task, topology, seed, sim, "E4_verification_ablation",
                            f"{block}::{topology}::{verification}", block,
                            condition=condition, fault=fault,
                            verification=verification, recovery=recovery
                        )
                        _append_result(results, result)


def _run_e5(tasks: list[Task], config: dict, results: list[EpisodeResult]) -> None:
    sim, exp = config["simulation"], config["experiments"]
    master = config["study"]["master_seed"]
    strategies = ["none", "retry", "isolate", "redundant"]
    for task in tasks:
        for repeat in range(int(exp["recovery_repeats"])):
            block = f"E5::{task.task_id}::{repeat}"
            for topology in sim["topologies"]:
                for recovery in strategies:
                    seed = stable_seed(master, block, topology, recovery)
                    fault = _fault(
                        task, topology, exp["recovery_fault_type"], exp["recovery_position"], seed
                    )
                    result = run_episode(
                        task, topology, seed, sim, "E5_recovery",
                        f"{block}::{topology}::{recovery}", block,
                        condition=f"recovery={recovery}", fault=fault,
                        verification="both", recovery=recovery
                    )
                    _append_result(results, result)


def _rank_candidates(result: EpisodeResult, observability: str) -> list[str]:
    nodes = [event.node_id for event in result.trace]
    if observability == "full_trace":
        observed = result.trace
    elif observability == "sparse_trace":
        observed = [event for event in result.trace if event.event_index % 2 == 0 or not event.parents]
    elif observability == "output_only":
        observed = [result.trace[-1]]
    else:
        raise ValueError(observability)
    anomalous = [event.node_id for event in observed if not event.locally_valid]
    # Evidence-ordered candidates; stable hashed tie-breaking avoids privileged fault labels.
    remainder = [node for node in nodes if node not in anomalous]
    remainder.sort(key=lambda node: stable_seed(result.episode_id, observability, node))
    return anomalous + remainder


def derive_attribution(results: list[EpisodeResult], max_cases: int) -> pd.DataFrame:
    eligible = [
        result for result in results
        if result.experiment == "E2_fault_propagation" and result.injection_consumed
    ]
    # Deterministic hash ordering distributes a capped sample across tasks, faults, and topologies.
    cases = sorted(eligible, key=lambda result: stable_seed("attribution-sample", result.episode_id))[:max_cases]
    rows = []
    for result in cases:
        for observability in ["output_only", "sparse_trace", "full_trace"]:
            ranking = _rank_candidates(result, observability)
            rank = ranking.index(result.origin_node) + 1 if result.origin_node in ranking else len(ranking) + 1
            rows.append({
                "experiment": "E3_attribution_observability",
                "episode_id": result.episode_id,
                "task_id": result.task_id,
                "topology": result.topology,
                "fault_type": result.fault_type,
                "fault_severity": result.fault_severity,
                "fault_position": result.fault_position,
                "observability": observability,
                "origin_node": result.origin_node,
                "predicted_origin": ranking[0],
                "candidate_count": len(ranking),
                "rank": rank,
                "top1_correct": int(rank == 1),
                "reciprocal_rank": 1.0 / rank,
            })
    return pd.DataFrame(rows)


def run_all(tasks: list[Task], config: dict[str, Any], root: Path) -> tuple[pd.DataFrame, pd.DataFrame]:
    selected = _task_subset(tasks, config)
    results: list[EpisodeResult] = []
    flags = config["experiments"]
    if flags.get("run_e1", True):
        _run_e1(selected, config, results)
    if flags.get("run_e2", True):
        _run_e2(selected, config, results)
    # E3 is computed from E2 traces rather than rerunning identical episodes.
    if flags.get("run_e4", True):
        _run_e4(selected, config, results)
    if flags.get("run_e5", True):
        _run_e5(selected, config, results)

    raw_dir = root / config["outputs"]["root"] / "raw"
    write_jsonl(raw_dir / "episodes.jsonl", [result.flat_dict(False) for result in results])
    trace_rows = []
    for result in results:
        for event in result.trace:
            row = event.to_dict()
            row.update({
                "episode_id": result.episode_id,
                "experiment": result.experiment,
                "condition": result.condition,
                "task_id": result.task_id,
                "topology": result.topology,
                "seed": result.seed,
                "fault_type": result.fault_type,
                "fault_severity": result.fault_severity,
                "fault_position": result.fault_position,
                "origin_node": result.origin_node,
            })
            trace_rows.append(row)
    write_jsonl(raw_dir / "events.jsonl", trace_rows)
    episodes = pd.DataFrame([result.flat_dict(False) for result in results])
    episodes.to_csv(raw_dir / "episodes.csv", index=False)
    attribution = derive_attribution(results, int(flags["attribution_max_cases"]))
    attribution.to_csv(raw_dir / "attribution.csv", index=False)
    with open(raw_dir / "run_counts.json", "w", encoding="utf-8") as handle:
        json.dump({
            "episode_count": len(episodes),
            "event_count": len(trace_rows),
            "attribution_records": len(attribution),
            "measured_tasks": len(selected),
        }, handle, indent=2)
    return episodes, attribution
