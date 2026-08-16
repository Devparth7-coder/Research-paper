# Release checklist

## Completed

- [x] Pilot configuration frozen in `configs/pilot.yaml`.
- [x] Dataset, raw traces, tables, statistical tests, and figures regenerated.
- [x] E2 severity/task/fault/position/topology coverage inspected.
- [x] E3 attribution records cover all 60 measured tasks and all three evidence views.
- [x] `scripts/verify_artifacts.py` passes.
- [x] Test suite passes: 18 tests.
- [x] Full repeat-run SHA-256 comparison passes for `data/tasks/` and `results/`.
- [x] Exact executed Python dependency versions captured in `requirements-lock.txt` and the reproducibility manifest.
- [x] Simulation-only evidence boundary appears in README, paper, and deck.

## Required before public/tagged release

- [ ] Create the public repository and replace `https://github.com/USERNAME/MAS-RELIAB` in `CITATION.cff`.
- [ ] Commit the release, rerun from a clean checkout, and confirm `git_commit` is populated.
- [ ] Create and sign/tag `v0.2.0`; attach paper and viva artifacts.
- [ ] Render the DOCX/PPTX with LibreOffice or Microsoft Office and perform page-by-page visual QA.
- [ ] If adding a related-work comparison table, verify every cell against the full cited publication.
- [ ] Archive the tagged release and record its DOI in `CITATION.cff`.

Do not describe this release as evidence about real LLM agents or deployed multi-agent systems.
