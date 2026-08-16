# Methodological quality review

## Scope judgment

This repository is a credible **implementation and causal-mechanism pilot inside a declared simulator**. It is not evidence about real LLM agents, production orchestration, or human-agent teams. The strongest defensible contribution is an auditable experimental methodology linking controlled intervention, trace lineage, propagation, attribution, mitigation, and cost. Claims must remain conditional on the configuration and synthetic task algebra.

## Strengths

1. **Objectively checkable endpoints.** Every task has a generated oracle state and canonical exact-state evaluation; no subjective LLM judge is used.
2. **Intervention integrity.** Scheduling, application, consumption, propagation, detection, and correction are distinct recorded events. Propagation denominators exclude unconsumed interventions.
3. **Trace-level causal bookkeeping.** Unique fault lineages traverse an explicit DAG and terminate on correction or algebraic masking.
4. **Topology coverage.** Single-agent, sequential, parallel, and hierarchical execution plans use the same task labels and configuration.
5. **Reproducibility.** Task, episode, and analysis artifacts include seeds, configuration, software environment, hashes, and rerun scripts.
6. **Inference discipline.** Tables report individual outcomes with task-clustered bootstrap intervals. Planned contrasts use paired exact/sign or Wilcoxon tests, effect sizes, and Holm correction.
7. **Negative findings retained.** H1 is mixed; H4, H5, and H6 are not supported. These outcomes are not rewritten as confirmations.

## High-priority weaknesses and required remedies

### 1. External validity is absent
The agents are stochastic state transformers, not language models. There are no prompts, natural-language ambiguities, learned policies, tool APIs, context windows, or model-specific failure modes. The pilot validates code paths and simulator-conditional mechanisms only.

**Remedy:** replicate with at least two hosted or local model families, frozen prompts, real tool sandboxes, multiple task domains, and model-by-topology interactions. Preserve this simulator as a preregistered calibration layer rather than pooling the two evidence types.

### 2. Oracle verification is unusually strong
Local and final verification compare states with exact simulator ground truth. Real deployments rarely possess an oracle at intermediate steps. Consequently, the large benefit of local verification is an upper-bound result under privileged information.

**Remedy:** test fallible executable constraints, schema checks, learned verifiers, and blinded human review; calibrate sensitivity/specificity on held-out data and include verifier cost and correlated errors.

### 3. Simulator parameters encode substantive outcomes
Intrinsic error rates, detector sensitivity, correction probabilities, topology multipliers, and application/consumption rates are researcher-set. A different parameterization can change rankings and hypothesis decisions.

**Remedy:** preregister parameter ranges, conduct global sensitivity analysis, use factorial or Latin-hypercube sampling, and report which conclusions survive across plausible regimes. Estimate parameters from empirical agent traces before making deployment claims.

### 4. Fault operators are narrow and partly algebra-specific
Wrong values, truncation, referential drift, and timeout capture useful atomic mechanisms, but they do not cover strategic deception, prompt injection, tool misuse, memory poisoning, policy violations, coordinated faults, or persistent environment corruption. Severity level does not create an independent effect for categorical timeout.

**Remedy:** validate each operator against empirical trace examples, add multi-fault and persistent faults, calibrate severity by measured state distance, and analyze operator validity separately from propagation.

### 5. Functional positions are not structurally equivalent
“Early,” “middle,” and “late” resolve differently in chains, parallel graphs, and trees. Late injection has no downstream reachable nodes, so EPR is structurally zero even when final failure is certain. This explains the observed H1 tension: early injections increased EPR and depth while late injections had slightly greater DAF.

**Remedy:** match nodes by centrality/reachable-set size, model reachability as a covariate, report topology-specific contrasts, and avoid interpreting raw position effects as intrinsic transmissibility.

### 6. Experimental pairing is incomplete
Contrasts are blocked by task and design cell, but stochastic streams are not fully counterfactually coupled across treatment arms. Injection-specific random draws can shift subsequent randomness.

**Remedy:** allocate independent RNG streams to intrinsic errors, intervention protocol, detection, recovery, and latency; run null-intervention counterfactuals with identical stream states; store pair IDs explicitly.

### 7. Attribution is optimistic and internally defined
The full-trace heuristic uses event-level oracle validity, making it much stronger than realistic debugging evidence. Only 600 consumed fault episodes are sampled, although the deterministic sample covers all 60 measured tasks.

**Remedy:** hide oracle validity from attribution methods, expose realistic logs only, add blinded human and learned baselines, report calibration and localization distance, and evaluate missing/misleading traces.

### 8. Effective sample size and generalization are limited
The measured pilot uses 60 evaluation tasks and many repeated episodes. Episode counts are large, but task diversity—not episode count—is the main generalization bottleneck. Bootstrap intervals cluster by task, yet only five algebraic task families exist.

**Remedy:** increase the number of independently generated task templates and domains, conduct leave-family-out analysis, and use hierarchical models with task-family and task random effects.

### 9. Cost and latency are simulated
Work units are defined by node and intervention overhead. Latency is sampled from a configured lognormal distribution and aggregated on each DAG critical path. Neither is wall-clock latency, token use, energy, nor monetary cost.

**Remedy:** in real-agent replication, record wall-clock timing, retries, tool calls, input/output tokens, model pricing snapshots, hardware, and queueing effects. Keep simulated and observed cost columns separate.

### 10. Multiplicity and decision-rule sensitivity remain
Holm correction controls the reported contrast family, but the study still contains many summaries and design choices. H6’s rank-discordance criterion has low power with four topologies, and the pilot did not support it after the direct paired DAF contrast.

**Remedy:** preregister primary contrasts and minimal effect sizes before external replication; separate confirmatory and exploratory analyses; report all exclusions and analysis changes.

## Secondary engineering weaknesses

- No repository Git commit was available in the generated manifest (`git_commit: null`); release runs should occur from a tagged commit.
- CI validates code and deterministic dataset generation but intentionally does not execute the full 16,560-episode pilot on every push.
- The repository has no container image or lockfile. Add a pinned environment and archived release artifact.
- Office files are structurally validated but cannot be visually rendered through LibreOffice in this environment.
- The related-work comparison must be rechecked cell by cell against final cited versions before submission; publication dates and versions are moving targets.

## Research-quality verdict

**Fit for:** reproducibility demonstration, simulator-method pilot, software artifact, hypothesis refinement, and a transparent viva.

**Not fit for:** claims of real-world MAS reliability, comparative claims about LLM vendors or frameworks, production safety certification, or universal topology/mitigation recommendations.

A strong next study would preregister the simulator-informed contrasts, run real model agents in a sandbox, replace oracle-dependent attribution/verification with observable evidence, and report cross-model and cross-domain heterogeneity.
