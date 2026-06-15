# Telecom PR Automation + BOM Comparison Context

## Existing Production Engine

RAN PR Automation already works and must not be broken.

Main entry point:

* src/run_pipeline.py

Pipeline:

* simple_normalize.py
* simple_calculation.py
* simple_pr_generator.py
* simple_ecc_export.py

FastAPI calls:

run_pipeline(selected_project)

Project selection comes from HTML dropdown.

## New BOM Comparison Tool

Two workflows:

### Workflow 1

PR Generated BOM vs Uploaded BOM

Inputs:

* output/simple_calculated.json (existing PR output)
* input/BOM_COMPARE_C.xlsx

Generated files:

* output/simple_normalized_bom_c.json
* output/simple_calculated_bom_c.json

Comparison:

compare_pr_vs_bom()

Output:

* output/bom_revision_report.xlsx
* output/bom_revision_report.json

### Workflow 2

BOM A vs BOM B

Inputs:

* input/BOM_COMPARE_A.xlsx
* input/BOM_COMPARE_B.xlsx

Generated files:

* output/simple_normalized_bom_a.json
* output/simple_calculated_bom_a.json
* output/simple_normalized_bom_b.json
* output/simple_calculated_bom_b.json

Comparison:

compare_bom_vs_bom()

Output:

* output/bom_revision_report.xlsx
* output/bom_revision_report.json

## Important Architecture Rules

1. Do not modify run_pipeline.py behavior.
2. Do not modify existing PR generation outputs.
3. Do not overwrite:

   * simple_normalized.json
   * simple_calculated.json
4. Use dedicated A/B/C output files.
5. Prefer pipeline-style orchestration matching run_pipeline.py.
6. Web UI uses FastAPI.
7. Upload endpoints already exist:

   * /upload-bom-compare-a
   * /upload-bom-compare-b
   * /upload-bom-compare-c
8. Existing comparison engine works and should not be rewritten.
9. Minimize risk to production PR automation workflow.

