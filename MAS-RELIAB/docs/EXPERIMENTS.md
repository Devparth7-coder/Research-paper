# Experiment registry

All experiments use the fixed evaluation task subset and per-episode SHA-256-derived seeds. `configs/pilot.yaml` is the source of truth.

| ID | Question | Design | Primary outputs |
|---|---|---|---|
| E1 | Baseline reliability beyond one success score | 60 tasks × 4 topologies × 5 repeats | task success, pass-all, exact-output consistency, work, simulated latency |
| E2 | Fault type, severity, and position | 60 × 4 topologies × 4 faults × 2 severities × 3 positions × 2 repeats | application/consumption, EPR, DAF, amplification, depth, task success |
| E3 | Observability-conditioned attribution | Deterministic 600-case sample from consumed E2 faults × 3 evidence views | top-1 origin accuracy, reciprocal rank |
| E4 | Verification ablation | 60 × 4 topologies × 2 fault types × 4 verification/recovery configurations | EPR, DAF, success, detection, recovery, work |
| E5 | Recovery strategy | 60 × 4 topologies × 4 recovery modes × 2 repeats | recovery, success, EPR, DAF, work, simulated latency |
| E6 | Reliability–cost trade-off and topology synthesis | Operating points derived from E4–E5, plus E2 topology contrasts | non-dominated interpretation, topology/location contrasts |

E3 and E6 are trace-/summary-derived experiments and do not duplicate episode execution. The component analysis is embedded in E4 and E5.

## Pairing and blocking units

The inferential unit is an aggregated matched cell, not an individual event. Replicates within a cell are averaged by the analysis pivot before a paired difference is calculated.

| Experiment / contrast | Matched-cell key | Treatment varied inside cell |
|---|---|---|
| E1 topology summaries | Task clusters for uncertainty; five repeats summarize each task–topology | Topology is descriptive in E1; H6 pairs by task/fault/severity/position in E2 |
| E2 H1 position | task × topology × fault type × severity | early versus late position; the two repeats are averaged |
| E2 severity exploration | task × topology × fault type × position | severity 2 versus 1; repeats are averaged |
| E3 H2 observability | source fault episode ID | full, sparse, or output-only views of the identical episode |
| E4 H3 verification | task × topology × fault type | verification/recovery condition; two repeats are averaged |
| E5 H4 recovery | task × topology | recovery strategy; two repeats are averaged |
| E6 H5 topology/location | task × fault type × severity | position or topology, depending on the contrast |

E4 and E5 are blocked on task and design factors but do **not** replay an identical random-number trajectory across treatments: the treatment label contributes to the deterministic seed. They are matched-design contrasts, not synchronized counterfactual replays.

### E2 control-arm limitation

E2 has no separately executed matched clean/null-intervention arm. Its EPR and DAF variables describe the consumed intervention lineage and final corruption, respectively; they are not paired causal differences from a clean counterfactual episode. E1 is a baseline topology experiment but is not a seed-synchronized E2 control. A confirmatory replication should execute null and fault conditions with independent named random streams synchronized for all non-intervention stochastic draws.

## Inference

Means use 95% task-clustered bootstrap confidence intervals (2,000 replicates). Blocked paired contrasts use an exact McNemar/binomial sign test for binary differences and Wilcoxon signed-rank tests otherwise. Reported matched effect size is discordant-pair direction for binary outcomes or matched rank-biserial correlation for continuous/ordinal outcomes. Holm adjustment is applied across the reported contrast family.

No p-value is interpreted as deployment evidence. Practical interpretation must also inspect effect direction, confidence interval, work units, and simulator limitations.
