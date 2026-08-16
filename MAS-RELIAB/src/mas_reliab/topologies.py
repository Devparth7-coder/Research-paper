"""Executable topology plans for four coordination structures."""
from __future__ import annotations

from dataclasses import dataclass

import networkx as nx


@dataclass(frozen=True)
class NodeSpec:
    node_id: str
    agent_id: str
    role: str
    parents: tuple[str, ...]
    leaf_indices: tuple[int, ...] = ()


def _chunks(n: int, count: int = 4) -> list[tuple[int, ...]]:
    import numpy as np
    return [tuple(int(x) for x in chunk) for chunk in np.array_split(range(n), count)]


def build_plan(topology: str, n_partials: int) -> list[NodeSpec]:
    chunks = _chunks(n_partials)
    if topology in {"single", "sequential"}:
        nodes = []
        for i, chunk in enumerate(chunks):
            agent = "agent-0" if topology == "single" else f"agent-{i}"
            parent = () if i == 0 else (f"step-{i-1}",)
            role = "solo" if topology == "single" else ("producer" if i == 0 else "handoff")
            nodes.append(NodeSpec(f"step-{i}", agent, role, parent, chunk))
        return nodes
    if topology == "parallel":
        workers = [NodeSpec(f"worker-{i}", f"agent-{i}", "worker", (), chunk) for i, chunk in enumerate(chunks)]
        return workers + [NodeSpec("aggregate", "agent-4", "aggregator", tuple(n.node_id for n in workers))]
    if topology == "hierarchical":
        workers = [NodeSpec(f"worker-{i}", f"agent-{i}", "worker", (), chunk) for i, chunk in enumerate(chunks)]
        managers = [
            NodeSpec("manager-0", "agent-4", "manager", ("worker-0", "worker-1")),
            NodeSpec("manager-1", "agent-5", "manager", ("worker-2", "worker-3")),
        ]
        return workers + managers + [NodeSpec("root", "agent-6", "root", ("manager-0", "manager-1"))]
    raise ValueError(f"Unknown topology: {topology}")


def graph_for(plan: list[NodeSpec]) -> nx.DiGraph:
    graph = nx.DiGraph()
    for node in plan:
        graph.add_node(node.node_id)
        for parent in node.parents:
            graph.add_edge(parent, node.node_id)
    return graph


def target_for_position(plan: list[NodeSpec], position: str) -> str:
    graph = graph_for(plan)
    sinks = [node for node in graph if graph.out_degree(node) == 0]
    sink = sorted(sinks)[0]
    generations = list(nx.topological_generations(graph))
    if position == "early":
        return sorted(generations[0])[0]
    if position == "late":
        return sink
    if len(generations) == 2:
        first = sorted(generations[0])
        return first[len(first) // 2]
    middle_generation = generations[len(generations) // 2]
    return sorted(middle_generation)[0]
