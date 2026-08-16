from pathlib import Path

import pytest

from mas_reliab.dataset import generate_tasks
from mas_reliab.engine import run_episode
from mas_reliab.models import FaultSpec
from mas_reliab.topologies import build_plan, target_for_position
from mas_reliab.utils import load_yaml


@pytest.fixture
def setup(tmp_path):
    config = load_yaml(Path(__file__).parents[1] / "configs" / "pilot.yaml")
    config["simulation"]["base_error_rate"] = 0
    config["simulation"]["difficulty_error_step"] = 0
    config["simulation"]["injection_application_failure_rate"] = 0
    config["simulation"]["injection_nonconsumption_rate"] = 0
    tasks = generate_tasks(config, tmp_path / "tasks.jsonl")
    return tasks[0], config["simulation"]


@pytest.mark.parametrize("topology", ["single", "sequential", "parallel", "hierarchical"])
def test_clean_execution_matches_exact_gold(setup, topology):
    task, simulation = setup
    result = run_episode(task, topology, 1, simulation, "test", "e", "b")
    assert result.task_success
    assert result.final_output == result.gold_output


@pytest.mark.parametrize("fault_type", ["wrong_value", "truncation", "referential_drift", "timeout"])
def test_fault_is_applied_consumed_and_recorded(setup, fault_type):
    task, simulation = setup
    plan = build_plan("hierarchical", len(task.partials))
    target = target_for_position(plan, "early")
    fault = FaultSpec("known-fault", fault_type, "early", target)
    result = run_episode(task, "hierarchical", 2, simulation, "test", "e", "b", fault=fault)
    assert result.injection_applied and result.injection_consumed
    assert any(event.fault_consumed for event in result.trace)
    assert result.origin_node == target


def test_local_recovery_can_remove_lineage(setup):
    task, simulation = setup
    simulation["local_detector_sensitivity"] = 1
    simulation["correction_success"] = 1
    plan = build_plan("sequential", len(task.partials))
    target = target_for_position(plan, "early")
    fault = FaultSpec("known-fault", "wrong_value", "early", target)
    result = run_episode(task, "sequential", 3, simulation, "test", "e", "b",
                         fault=fault, verification="local", recovery="retry")
    assert result.detected and result.recovered and result.task_success
    assert result.daf == 0
