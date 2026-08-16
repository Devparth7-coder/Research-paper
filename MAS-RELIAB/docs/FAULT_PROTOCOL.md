# Controlled fault protocol

Each scheduled fault has a unique ID, type, functional position, concrete node target, severity, and episode seed. E2 crosses severity levels 1 and 2; for value and truncation operators, level 2 increases perturbation magnitude, while timeout remains a categorical missing-state intervention. The four operators are:

- **wrong_value:** numerical or state-value perturbation;
- **truncation:** removes a component or reduces a scalar;
- **referential_drift:** changes a bit/reference-like coordinate or swaps state fields;
- **timeout:** returns a missing/empty state.

Positions are resolved against each topology: **early** is an entry/worker node, **middle** is a chain midpoint, manager, or representative parallel worker, and **late** is the terminal/aggregation node. This is a functional-position comparison, not an assertion that all graphs have equal depth.

For each intervention the engine separately records `injection_scheduled`, `injection_applied`, and `injection_consumed`. Propagation analyses use consumed interventions. A stable per-episode seed controls intrinsic errors, protocol application, consumption, detector outcomes, recovery, and simulated latency.

A lineage label follows a consumed corruption through downstream state transformations. It is removed when correction succeeds or the state algebra masks the difference. EPR is affected downstream nodes divided by reachable downstream nodes. DAF is one only when the injected lineage reaches an incorrect final state.

The verifier compares against simulator oracle state. That is an experimental control, not a realistic claim that deployed systems possess perfect ground truth.
