"""Stochastic, fault-sensitive synthetic multi-agent execution engine."""
from __future__ import annotations

from typing import Any

import networkx as nx
import numpy as np

from .dataset import combine_values
from .models import Artifact, EpisodeResult, Event, FaultSpec, Task
from .topologies import NodeSpec, build_plan, graph_for
from .utils import exact_equal


def _mutate(value: Any, fault_type: str, local: np.random.Generator, severity: int = 1) -> Any:
    """Apply a guaranteed state-changing corruption when possible."""
    if isinstance(value, bool):
        return not value
    if isinstance(value, int):
        if fault_type == "timeout":
            return None
        if fault_type == "truncation":
            return int(value / (10 ** severity))
        if fault_type == "referential_drift":
            result = value
            for _ in range(severity):
                result ^= 1 << int(local.integers(0, 8))
            return result
        return value + severity * int(local.choice([-7, -3, 3, 7]))
    if isinstance(value, list):
        if fault_type == "timeout":
            return []
        if fault_type == "truncation":
            return value[:-severity] if len(value) >= severity else []
        result = list(value)
        if fault_type == "referential_drift" and len(result) > 1:
            result[0], result[-1] = result[-1], result[0]
            if result == value:
                result[0] = int(result[0]) + 1
        elif result:
            result[0] = int(result[0]) + severity * int(local.choice([-3, 3]))
        else:
            result = [1]
        return result
    if isinstance(value, dict):
        if fault_type == "timeout":
            return {}
        result = dict(value)
        keys = sorted(result)
        if fault_type == "truncation" and keys:
            for key in keys[-severity:]:
                result.pop(key, None)
        elif fault_type == "referential_drift" and len(keys) > 1:
            result[keys[0]], result[keys[1]] = result[keys[1]], result[keys[0]]
            if result == value:
                result[keys[0]] = int(result[keys[0]]) + 1
        elif keys:
            result[keys[0]] = int(result[keys[0]]) + severity * int(local.choice([-3, 3]))
        else:
            result = {"corrupt": 1}
        return result
    return {"corrupted": True, "original_type": type(value).__name__}


def _node_expected(task: Task, node: NodeSpec, leaf_sets: dict[str, tuple[int, ...]]) -> Any:
    indices = list(node.leaf_indices)
    for parent in node.parents:
        indices.extend(leaf_sets[parent])
    leaf_sets[node.node_id] = tuple(indices)
    return combine_values([task.partials[i] for i in indices], task.combine)


def _node_actual(task: Task, node: NodeSpec, outputs: dict[str, Artifact]) -> Artifact:
    values: list[Any] = []
    lineages: set[str] = set()
    for parent in node.parents:
        values.append(outputs[parent].data)
        lineages.update(outputs[parent].active_lineages)
    values.extend(task.partials[i] for i in node.leaf_indices)
    # Timeout-like None values are consumable but yield a deliberately invalid state.
    if any(value is None for value in values):
        data: Any = None
    else:
        try:
            data = combine_values(values, task.combine)
        except (TypeError, ValueError, KeyError, IndexError, AttributeError):
            data = None
    return Artifact(data, lineages)


def _error_multiplier(topology: str, role: str) -> float:
    table = {
        "single": {"solo": 0.90},
        "sequential": {"producer": 0.95, "handoff": 1.25},
        "parallel": {"worker": 1.05, "aggregator": 1.20},
        "hierarchical": {"worker": 1.00, "manager": 1.10, "root": 0.95},
    }
    return table[topology][role]


def run_episode(
    task: Task,
    topology: str,
    seed: int,
    simulation: dict[str, Any],
    experiment: str,
    episode_id: str,
    block_id: str,
    condition: str = "baseline",
    fault: FaultSpec | None = None,
    verification: str = "none",
    recovery: str = "none",
) -> EpisodeResult:
    local = np.random.default_rng(seed)
    plan = build_plan(topology, len(task.partials))
    graph = graph_for(plan)
    sink = [node for node in graph if graph.out_degree(node) == 0][0]
    outputs: dict[str, Artifact] = {}
    expected: dict[str, Any] = {}
    leaf_sets: dict[str, tuple[int, ...]] = {}
    finish_times: dict[str, float] = {}
    events: list[Event] = []
    injection_applied = False
    injection_consumed = False
    origin_node = fault.target_node if fault else None

    for index, node in enumerate(plan):
        expected_value = _node_expected(task, node, leaf_sets)
        expected[node.node_id] = expected_value
        artifact = _node_actual(task, node, outputs)
        input_lineages = set(artifact.active_lineages)
        input_data = [outputs[p].data for p in node.parents] + [task.partials[i] for i in node.leaf_indices]
        intrinsic_error = False
        fault_scheduled = bool(fault and node.node_id == fault.target_node)
        fault_applied_here = False
        fault_consumed_here = False

        error_probability = (
            float(simulation["base_error_rate"])
            + (task.difficulty - 1) * float(simulation["difficulty_error_step"])
        ) * _error_multiplier(topology, node.role)
        if local.random() < min(error_probability, 0.8):
            artifact.data = _mutate(artifact.data, "wrong_value", local)
            artifact.active_lineages.add(f"intrinsic::{node.node_id}")
            intrinsic_error = True

        if fault_scheduled:
            fault_applied_here = local.random() >= float(simulation["injection_application_failure_rate"])
            injection_applied = injection_applied or fault_applied_here
            if fault_applied_here:
                fault_consumed_here = local.random() >= float(simulation["injection_nonconsumption_rate"])
                injection_consumed = injection_consumed or fault_consumed_here
                if fault_consumed_here:
                    artifact.data = _mutate(artifact.data, fault.fault_type, local, fault.severity)
                    artifact.active_lineages.add(fault.fault_id)

        invalid = not exact_equal(artifact.data, expected_value)
        verify_here = verification in {"local", "both"} or (
            verification in {"final", "both"} and node.node_id == sink
        )
        detected = False
        if verify_here:
            if invalid:
                sensitivity = (
                    float(simulation["final_detector_sensitivity"])
                    if node.node_id == sink and verification == "final"
                    else float(simulation["local_detector_sensitivity"])
                )
                detected = local.random() < sensitivity
            else:
                detected = local.random() < float(simulation["false_positive_rate"])

        corrected = False
        recovery_overhead = 0.0
        if detected and recovery != "none":
            success_probability = {
                "retry": float(simulation["correction_success"]),
                "isolate": max(0.0, float(simulation["correction_success"]) - 0.12),
                "redundant": min(0.995, float(simulation["correction_success"]) + 0.075),
            }.get(recovery, 0.0)
            recovery_overhead = {"retry": 0.75, "isolate": 0.45, "redundant": 1.60}.get(recovery, 0.0)
            if local.random() < success_probability:
                artifact.data = expected_value
                artifact.active_lineages.clear()
                corrected = True

        # A lineage ceases to be active if later algebra masks the corruption.
        if exact_equal(artifact.data, expected_value):
            artifact.active_lineages.clear()
        outputs[node.node_id] = artifact
        base_work = 1.0 + 0.12 * len(input_data)
        work = base_work + recovery_overhead
        duration = float(local.lognormal(mean=2.45, sigma=0.18)) * (1 + recovery_overhead)
        start = max((finish_times[parent] for parent in node.parents), default=0.0)
        finish_times[node.node_id] = start + duration
        injected_id = fault.fault_id if fault else ""
        events.append(
            Event(
                event_index=index,
                node_id=node.node_id,
                agent_id=node.agent_id,
                role=node.role,
                parents=list(node.parents),
                stage="execute",
                input_data=input_data,
                output_data=artifact.data,
                expected_data=expected_value,
                locally_valid=exact_equal(artifact.data, expected_value),
                active_lineages_in=sorted(input_lineages),
                active_lineages_out=sorted(artifact.active_lineages),
                intrinsic_error=intrinsic_error,
                fault_scheduled=fault_scheduled,
                fault_applied=fault_applied_here,
                fault_consumed=fault_consumed_here,
                detected=detected,
                corrected=corrected,
                affected_by_injected_fault=bool(injected_id and injected_id in artifact.active_lineages),
                work_units=work,
                latency_ms=duration,
            )
        )

    final = outputs[sink].data
    task_success = exact_equal(final, task.gold)
    descendants = nx.descendants(graph, origin_node) if origin_node in graph else set()
    affected_nodes = {
        event.node_id for event in events
        if event.node_id in descendants and event.affected_by_injected_fault
    }
    reachable = len(descendants)
    affected = len(affected_nodes)
    epr = affected / reachable if reachable else 0.0
    daf = float(bool(fault and fault.fault_id in outputs[sink].active_lineages and not task_success))
    depths = [nx.shortest_path_length(graph, origin_node, node) for node in affected_nodes] if affected_nodes else []
    injected_id = fault.fault_id if fault else ""
    detected_any = any(
        event.detected
        and (event.fault_scheduled or injected_id in event.active_lineages_in or injected_id in event.active_lineages_out)
        for event in events
    )
    recovered = bool(fault and injection_consumed and task_success and any(event.corrected for event in events))
    return EpisodeResult(
        experiment=experiment,
        condition=condition,
        episode_id=episode_id,
        block_id=block_id,
        task_id=task.task_id,
        task_family=task.family,
        difficulty=task.difficulty,
        topology=topology,
        seed=seed,
        fault_type=fault.fault_type if fault else None,
        fault_severity=fault.severity if fault else None,
        fault_position=fault.position if fault else None,
        fault_target=fault.target_node if fault else None,
        verification=verification,
        recovery=recovery,
        injection_scheduled=fault is not None,
        injection_applied=injection_applied,
        injection_consumed=injection_consumed,
        task_success=task_success,
        final_output=final,
        gold_output=task.gold,
        epr=epr,
        reachable_downstream=reachable,
        affected_downstream=affected,
        daf=daf,
        amplification=float(affected),
        propagation_depth=max(depths, default=0),
        detected=detected_any,
        recovered=recovered,
        work_units=float(sum(event.work_units for event in events)),
        latency_ms=float(finish_times[sink]),
        event_count=len(events),
        origin_node=origin_node,
        trace=events,
    )
