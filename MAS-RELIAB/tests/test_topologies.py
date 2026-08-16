import networkx as nx
import pytest

from mas_reliab.topologies import build_plan, graph_for, target_for_position


@pytest.mark.parametrize("topology", ["single", "sequential", "parallel", "hierarchical"])
def test_topology_is_dag_with_one_sink(topology):
    plan = build_plan(topology, 7)
    graph = graph_for(plan)
    assert nx.is_directed_acyclic_graph(graph)
    assert len([node for node in graph if graph.out_degree(node) == 0]) == 1
    assert {target_for_position(plan, p) for p in ["early", "middle", "late"]} <= set(graph.nodes)


def test_single_reuses_one_agent_and_sequential_does_not():
    single = build_plan("single", 6)
    sequential = build_plan("sequential", 6)
    assert len({node.agent_id for node in single}) == 1
    assert len({node.agent_id for node in sequential}) == 4
