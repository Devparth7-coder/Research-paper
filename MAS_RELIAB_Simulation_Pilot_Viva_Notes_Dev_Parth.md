# MAS-RELIAB simulation-pilot viva notes

Target talk: 14–16 minutes plus questions.

## Slide 1: Title
Open with the evidence boundary: this is the measured simulation pilot, not an LLM benchmark.

## Slide 2: Evidence boundary
State the scale and what zero model calls means. Treat this as integrity, not an apology.

## Slide 3: Problem
Use one example: final failure alone cannot reveal whether the injection failed, propagated, or was detected.

## Slide 4: Contribution
Avoid priority language. Emphasize integration and auditability.

## Slide 5: RQs
Explain that hypotheses were frozen directional claims; null findings remain informative.

## Slide 6: Simulator
Explain exact-state algebra and lineage termination. Mention stochastic intrinsic errors.

## Slide 7: Topologies
Clarify that graph size differs and therefore cost/critical path must be shown.

## Slide 8: Fault protocol
Explain scheduled/applied/consumed denominators. Propagation uses consumed faults.

## Slide 9: Experiments
E3 and E6 are derived analyses, so episode totals do not double-count them.

## Slide 10: E1
Point out that single leads TSR but sequential leads pass-all. H6 still failed its direct inferential criterion.

## Slide 11: E2
This is the key counterintuitive result: early creates more cascade opportunity; late directly breaks final state.

## Slide 12: E3
State the +0.590 paired effect, then immediately disclose oracle-validity optimism.

## Slide 13: E4
State containment gain and cost together. Call it an upper bound.

## Slide 14: E5
Do not claim redundancy wins; its raw recovery gain failed Holm correction.

## Slide 15: E6
Explain the Pareto idea without claiming a universal optimum or composite score.

## Slide 16: Decisions
Read H1, H4, H5, H6 honestly. Stress that the protocol survives rejected hypotheses.

## Slide 17: Weaknesses
Prioritize external validity, oracle verification, parameter sensitivity, and position inequivalence.

## Slide 18: Artifact
Mention raw trace count, hashes, tests, CI, and missing Git commit as an openly recorded weakness.

## Slide 19: Next study
Propose real models, tools, fallible verifiers, synchronized counterfactual streams, observed costs.

## Slide 20: Conclusion
Finish with the integrated methodology and evidence boundary.

## Slide 21: Backup
Use result → limitation → next step for every challenge.

## Short answers

**Why no LLMs?** No credentials or local weights were available; simulation was chosen explicitly to produce reproducible causal-mechanism evidence without fabricating model results.

**Why is H1 mixed?** Early faults had more reachable descendants, so EPR and depth rose. Late faults acted directly on the sink, so DAF was slightly higher despite zero downstream nodes.

**Is full-trace attribution realistic?** No. It contains oracle validity and is an upper bound. A real replication must hide oracle labels and expose only realistic logs.

**Why no composite score?** Weights encode stakeholder utility and can hide rank reversals. Component metrics are primary; any utility model must be declared separately.

**Main contribution?** A unified, executable, auditable methodology connecting intervention integrity, propagation, attribution, mitigation, and reliability–cost analysis.