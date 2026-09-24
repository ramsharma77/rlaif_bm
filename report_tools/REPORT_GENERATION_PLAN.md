# Before/After Report Plan (v1 -> v3)

## Objective
Generate client-ready before/after docs for this repo using:
- Before = `trace-v1`
- After = `trace-v3`
- Per-section summary (`The problem / The change / The effect`)
- Per-section code-change table (`Setting / Before / After`)
- Clickable GitHub permalink evidence

## Repo
- Local: `C:/projects/proposal/rlaif-harness-traces`
- GitHub: `https://github.com/ramsharma77/rlaif_bm`

## Tooling
- Generator: `report_tools/generate_trace_versions_docx.py`
- Config: `report_tools/trace_versions_docx_config.rlaif_before_after.json`
- Executive wrapper: `report_tools/build_rlaif_before_after_exec_variant.py`

## Run
1. `python report_tools/generate_trace_versions_docx.py --config report_tools/trace_versions_docx_config.rlaif_before_after.json`
2. `python report_tools/build_rlaif_before_after_exec_variant.py`

## Outputs
- `output/rlaif_before_after.docx`
- `output/rlaif_before_after_exec.docx`
- `output/trace_docx_artifacts_rlaif_before_after/`

## Validation Checklist
- Config refs are exactly `trace-v1` and `trace-v3`.
- DOCX opens successfully.
- Hyperlinks resolve to `github.com/ramsharma77/rlaif_bm/blob/<40-char-sha>/...`.
- Each section includes summary + code-change table.
