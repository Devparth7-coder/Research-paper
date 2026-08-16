# MAS-RELIAB deliverable QA report

Date: 2026-08-16  
Author: Dev Parth

## Execution QA

- Full pilot rerun completed: 150 tasks, 60 measured tasks, 16,560 episodes, and 82,800 events.
- `python scripts/verify_artifacts.py`: passed; eight required artifact classes and all 16,560 episode records verified.
- `PYTHONPATH=src pytest -q`: passed; 18 tests.
- Severity/task/fault coverage: E2 contains severity levels 1–2, four fault types, three positions, four topologies, and all 60 measured tasks.
- Attribution coverage: all 60 measured tasks and output-only, sparse-trace, and full-trace views.
- E2 integrity: 11,520 scheduled, 11,156 applied, and 10,940 consumed interventions.
- A complete second execution produced identical SHA-256 checksums for every file under `data/tasks/` and `results/`: determinism check passed.
- The final full run produced no Seaborn deprecation warning after figure-code cleanup.

## Repository QA

- README, changelog, release checklist, exact environment lock, citation metadata, license, tests, CI workflow, experiment registry, data card, fault protocol, and methodology review are present.
- `docs/EXPERIMENTS.md` explicitly records E1–E6 blocking/pairing units, repeat aggregation, unsynchronized E4/E5 treatment streams, and the absence of a matched clean/null E2 arm.
- `CITATION.cff` intentionally retains a placeholder repository URL; public release is gated on replacing it.
- The reproducibility manifest correctly reports `git_commit: null`; a clean committed/tagged rerun remains a release requirement.

## Office-file structural QA

### Paper

File: `MAS_RELIAB_Simulation_Pilot_Paper_Dev_Parth.docx`

- Valid OOXML ZIP package; no corrupt member detected.
- 129 paragraphs, eight tables, nine embedded figures, approximately 3,086 words.
- Required evidence-scope wording, execution total, mixed H1 decision, unsupported H4 decision, and E2 control-arm limitation are present.
- Complete measured Method, Results, Discussion, threats-to-validity, Conclusion, references, and appendices are present.

### Viva deck

File: `MAS_RELIAB_Simulation_Pilot_Viva_Deck_Dev_Parth.pptx`

- Valid OOXML ZIP package; no corrupt member detected.
- 21 slides.
- Programmatic geometry check found no shape outside slide boundaries.
- All six experiments, H1–H6 decisions, evidence boundary, methodological verdict, reproducibility, and next-study plan are covered.
- Companion presenter notes are in `MAS_RELIAB_Simulation_Pilot_Viva_Notes_Dev_Parth.md`.

## Visual-rendering limitation

LibreOffice is not installed in this environment, so the DOCX and PPTX could not be rendered to PDF or inspected page by page as raster images. Structural OOXML and geometry checks passed, but a final release should still be opened in Microsoft Office or LibreOffice to inspect pagination, table wrapping, font substitution, figure legibility, and animation-free slide display. This unresolved visual QA gate is recorded in `MAS-RELIAB/RELEASE_CHECKLIST.md`.
