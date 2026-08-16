"""Typed records for MAS-RELIAB's offline pilot."""
from __future__ import annotations

from dataclasses import dataclass, asdict, field
from typing import Any


@dataclass(frozen=True)
class Task:
    task_id: str
    family: str
    difficulty: int
    payload: dict[str, Any]
    partials: list[Any]
    combine: str
    gold: Any
    split: str
    version: str = "1.0"

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, value: dict[str, Any]) -> "Task":
        return cls(**value)


@dataclass
class Artifact:
    data: Any
    active_lineages: set[str] = field(default_factory=set)

    def clone(self) -> "Artifact":
        import copy
        return Artifact(copy.deepcopy(self.data), set(self.active_lineages))


@dataclass
class FaultSpec:
    fault_id: str
    fault_type: str
    position: str
    target_node: str
    severity: int = 1
    seed: int = 0


@dataclass
class Event:
    event_index: int
    node_id: str
    agent_id: str
    role: str
    parents: list[str]
    stage: str
    input_data: Any
    output_data: Any
    expected_data: Any
    locally_valid: bool
    active_lineages_in: list[str]
    active_lineages_out: list[str]
    intrinsic_error: bool
    fault_scheduled: bool
    fault_applied: bool
    fault_consumed: bool
    detected: bool
    corrected: bool
    affected_by_injected_fault: bool
    work_units: float
    latency_ms: float

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass
class EpisodeResult:
    experiment: str
    condition: str
    episode_id: str
    block_id: str
    task_id: str
    task_family: str
    difficulty: int
    topology: str
    seed: int
    fault_type: str | None
    fault_severity: int | None
    fault_position: str | None
    fault_target: str | None
    verification: str
    recovery: str
    injection_scheduled: bool
    injection_applied: bool
    injection_consumed: bool
    task_success: bool
    final_output: Any
    gold_output: Any
    epr: float
    reachable_downstream: int
    affected_downstream: int
    daf: float
    amplification: float
    propagation_depth: int
    detected: bool
    recovered: bool
    work_units: float
    latency_ms: float
    event_count: int
    origin_node: str | None
    trace: list[Event] = field(default_factory=list)

    def flat_dict(self, include_trace: bool = False) -> dict[str, Any]:
        d = asdict(self)
        if not include_trace:
            d.pop("trace", None)
        return d
