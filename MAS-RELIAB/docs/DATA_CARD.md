# MAS-RELIAB synthetic task data card

## Purpose
`tasks_v1` is a deterministic, objectively scored dataset for an offline simulation pilot of failure propagation and recovery. It is not a corpus of human or model interactions.

## Composition
The generator creates 150 tasks across five families: signed arithmetic pipelines, shard aggregation, evidence-vector synthesis, constraint-bitmask union, and state-delta workflows. Splits are 30 development, 30 pilot, and 90 evaluation tasks. The configured measured pilot uses the first 60 evaluation tasks.

## Labels and evaluation
Every label is computed by an associative state operator (`sum`, `vector_sum`, `bitwise_or`, or `dict_sum`). Evaluation uses canonical exact end-state equality. No subjective judge or text similarity is used.

## Generation and versioning
Run `mas-reliab generate-data --config configs/pilot.yaml`. Every task is generated from a SHA-256-derived seed rooted in master seed `20260816`. The run manifest stores dataset and configuration hashes.

## Intended and out-of-scope use
The data supports implementation testing and controlled simulator experiments. It must not be used to claim performance of real LLM agents, humans, deployed orchestration systems, natural-language understanding, or real tool use.

## Known limitations
Tasks use associative state combination and have clean oracle states. They are deliberately narrower and more verifiable than open-ended agent work. Their family labels increase surface diversity but do not reproduce linguistic ambiguity or real API semantics.
