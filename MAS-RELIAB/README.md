# MAS-RELIAB

**A reproducible offline simulation pilot for multi-agent reliability, failure propagation, attribution, mitigation, and reliability–cost trade-offs.**

Author: **Dev Parth**  
Release: **v0.2.0 simulation pilot**

> **Scope:** The measured results in this repository come from stochastic synthetic agents, not hosted or local LLMs. They validate the implementation and configured causal mechanisms only. They are not evidence about deployed multi-agent systems.

## What was executed

The fixed pilot configuration generated 150 exact-state tasks in five families and measured the first 60 evaluation tasks. It executed:

- four topologies: single-agent, sequential, parallel, and hierarchical;
- four fault operators: wrong value, truncation, referential drift, and timeout;
- severity levels 1–2 and early/middle/late functional positions;
- output-only, sparse-trace, and full-trace attribution views;
- no, final, local, and combined verification;
- retry, isolate, and redundant recovery; and
- six analysis experiments plus the verification/recovery component ablation.

The recorded release contains **16,560 episodes**, **82,800 trace events**, and **1,800 attribution-view records**. E2 scheduled 11,520 fault episodes; 96.84% were applied and 94.97% were consumed. Propagation analyses use consumed interventions.

## Selected measured results

Individual metrics are reported before any composite score; MAS-RELIAB does not claim scientifically justified weights for a single reliability index.

| Topology | Baseline task success (95% task-clustered CI) | Pass-all across 5 repeats | Mean work units | Simulated latency (ms) |
|---|---:|---:|---:|---:|
| Single | 0.870 [0.833, 0.903] | 0.483 | 5.08 | 47.21 |
| Sequential | 0.860 [0.813, 0.903] | 0.533 | 5.08 | 46.86 |
| Parallel | 0.800 [0.757, 0.843] | 0.333 | 6.20 | 25.75 |
| Hierarchical | 0.780 [0.727, 0.833] | 0.333 | 8.44 | 38.05 |

Attribution top-1 accuracy was 0.302 for output-only, 0.803 for sparse traces, and 0.892 for full traces. In the paired full-trace versus output-only contrast, the difference was 0.590 (95% bootstrap CI [0.543, 0.637], Holm-adjusted p < 0.001; paired sign effect 0.839).

Local verification plus retry reduced EPR by 0.619 relative to final-only verification plus retry (local minus final; 95% CI [-0.657, -0.582], Holm-adjusted p < 0.001) while adding 0.505 mean work units. This is an oracle-verification upper bound inside the simulator.

### Hypothesis decisions

| Hypothesis | Decision |
|---|---|
| H1: earlier faults produce greater normalized impact and depth | **Mixed support** — EPR and depth increased, but early DAF was slightly lower than late DAF |
| H2: full traces improve attribution | **Supported** |
| H3: local verification contains cascades better | **Supported** |
| H4: isolate/redundant recovery beats retry with additional cost | **Not supported** after multiplicity correction |
| H5: topology × location interaction follows the proposed pattern | **Not supported** |
| H6: multidimensional profile establishes a significant masked weakness/rank reversal | **Not supported** by the direct paired DAF contrast |

See `results/analysis/statistical_tests.csv` for confidence intervals, paired tests, effect sizes, and Holm-adjusted p-values. Null and contradictory findings are retained.

## Reproduce

Python 3.10+ is required.

```bash
python -m venv .venv
source .venv/bin/activate
# Exact measured environment (Python 3.13.14):
pip install -r requirements-lock.txt
# Or install the package with supported dependency ranges:
pip install -e .[dev]
mas-reliab all --config configs/pilot.yaml
pytest
python scripts/verify_artifacts.py
```

Or run:

```bash
bash scripts/run_pilot.sh
```

Subcommands are `generate-data`, `run`, `analyze`, and `all`. Per-episode seeds are derived by SHA-256 from the fixed master seed. `results/raw/reproducibility_manifest.json` records configuration, dataset, source hashes, package versions, platform, and hardware fields.

## Repository map

```text
configs/pilot.yaml                 fixed pilot configuration
data/tasks/                        generated exact-state task dataset
src/mas_reliab/                    simulator, topology, experiment, and analysis code
results/raw/                       episodes, events, attribution records, manifests
results/tables/                    measured CSV tables with confidence intervals
results/figures/                   six empirical PNG figures
results/analysis/                  statistical tests, RQ answers, hypothesis decisions
scripts/run_pilot.sh               full rerun entry point
tests/                             deterministic unit and integration tests
docs/                              data card, experiment registry, fault protocol, review
requirements-lock.txt              exact packages used for the measured pilot
CHANGELOG.md / RELEASE_CHECKLIST.md release history and publication gates
.github/workflows/ci.yml           GitHub Actions test workflow
```

## Metric semantics

- **Task success:** canonical exact equality with the generated end state.
- **Pass-all:** fraction of tasks successful in all five baseline repeats.
- **EPR:** affected downstream nodes / nodes reachable downstream from the injected node.
- **DAF:** whether the injected lineage reaches an incorrect final state.
- **Amplification:** count of downstream nodes carrying the active injected lineage.
- **Attribution:** top-1 origin accuracy and reciprocal rank under controlled evidence views.
- **Recovery:** consumed fault episode ending successfully after a recorded correction.
- **Cost:** simulated work units; **latency is simulated**, not measured wall-clock/API time.

## Important limitations

The simulator uses associative task algebra, oracle intermediate states, researcher-selected error/detector/recovery parameters, functional positions that are not structurally identical across topologies, and synthetic work/latency. There are no tokens because no language model is called. Read `docs/METHODOLOGY_REVIEW.md` before citing the measurements.

## Citation and license

Citation metadata is in `CITATION.cff`. Code and generated pilot artifacts are released under the MIT License. Replace the placeholder repository URL before publication.
