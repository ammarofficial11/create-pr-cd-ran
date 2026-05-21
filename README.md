````markdown
# Telecom PR Automation Engine

## Overview

This project automates the full telecom PR workflow:

- BOM normalization
- telecom equipment categorization
- semantic quantity calculation
- PR line item generation
- ECC PR template export

The system replaces:
- manual Excel formulas
- repetitive copy-paste workflows
- VBA-heavy logic
- fragile hardcoded mappings

with:
- config-driven rule engines
- semantic equipment mapping
- reusable calculation pipelines
- automated ECC export generation

---

# Final Pipeline Architecture

```text
BOM.xlsx
EPMS.xlsx
ReferenceSubcon&Region.xlsx
        ↓
simple_normalize.py
        ↓
simple_normalized.json
        ↓
simple_calculation.py
        ↓
simple_calculated.json
        ↓
simple_pr_generator.py
        ↓
simple_pr_output.json
        ↓
simple_ecc_export.py
        ↓
ECC_PR_Output.xlsx
````

---

# Folder Structure

```text
Telecom-PR-Automation/
│
├── config/
│   ├── MainConfig.xlsx
│   ├── ReferenceSubcon&Region.xlsx
│   └── ecc_template.xls
│
├── input/
│   ├── BOM.xlsx
│   └── EPMS.xlsx
│
├── output/
│   ├── simple_normalized.json
│   ├── simple_calculated.json
│   ├── simple_pr_output.json
│   └── ECC_PR_Output.xlsx
│
├── src/
│   ├── simple_normalize.py
│   ├── simple_calculation.py
│   ├── simple_pr_generator.py
│   ├── simple_ecc_export.py
│   └── run_pipeline.py
│
└── README.md
```

---

# Core Input Files

## 1. BOM.xlsx

Main telecom BOM source file.

Contains:

* site information
* DU codes
* equipment quantities
* raw telecom equipment columns

### Important Structure

* Header row starts at Excel Row 3
* Site metadata columns:

  * Site Code
  * Site Name
  * Region
  * DU Code
* Equipment quantities exist as dynamic equipment columns

---

## 2. EPMS.xlsx

Used for:

* subcontractor lookup
* site deployment metadata

### Important Structure

* Header row starts at Excel Row 4
* Required columns:

  * customer site code
  * du code
  * SubCon - TI

The automation engine uses:

```text
(customer site code + du code)
```

to lookup:

```text
SubCon - TI
```

---

## 3. ReferenceSubcon&Region.xlsx

Static lookup reference file.

Used for:

* Region → Purchasing Area
* Official Subcontractor → Contract Number
* Subcontractor normalization

### Important Areas

## Region Mapping

| Region   | Purchasing Area             |
| -------- | --------------------------- |
| Central  | Malaysia_Central Region     |
| Northern | Malaysia_South North Region |

---

## Subcontractor Mapping

| Official Subcontractor | Contract Number    |
| ---------------------- | ------------------ |
| Datasco                | S1MY2024071001WBF1 |
| CCSMY                  | S1MY2024071004WBF1 |

---

# MainConfig.xlsx Sheets

---

# 1. Equipment_Normalization

Purpose:
Normalize raw BOM equipment names into semantic ItemKeys.

Example:

| Item List                            | ItemKey      |
| ------------------------------------ | ------------ |
| Rectifier-ZTE Rectifier Module 3000W | Power Module |
| GPS Unit Huawei                      | GPS          |
| RF Antenna 4 Port                    | Antenna      |

This layer removes:

* vendor naming differences
* BOM inconsistencies
* raw equipment complexity

---

# 2. Calculation_Rules

Purpose:
Perform telecom semantic calculations.

Example logic:

```text
If Cabinet exists:
    return 0
Else:
    calculate based on rectifier count
```

Supported calculation types:

| CalcType                   | Description                            |
| -------------------------- | -------------------------------------- |
| SUM                        | Sum quantities directly                |
| DIVIDE_BY_ITEMKEY          | Divide one ItemKey quantity by another |
| CONDITIONAL_SUM_BY_ITEMKEY | Conditional aggregation                |
| CONDITIONAL_DIVIDE         | Conditional division                   |
| RANGE                      | Threshold/range logic                  |

---

# 3. MainRuleTable

Purpose:
Convert semantic quantities into ECC PR line items.

Example:

| ItemKey | LineItemText               |
| ------- | -------------------------- |
| RRU     | RRU installation/expansion |
| Antenna | RF Antenna installation    |
| GPS     | GPS installation per Site  |

Contains:

* MaterialCode
* PR_Mode
* MinQty
* MaxQty
* ECC SOW descriptions

---

# Pipeline Scripts

---

# 1. simple_normalize.py

Purpose:
Convert raw BOM equipment into semantic ItemKeys.

Also enriches metadata using:

* EPMS.xlsx
* ReferenceSubcon&Region.xlsx

## Output

```text
output/simple_normalized.json
```

## Responsibilities

### Equipment Normalization

```text
Raw Equipment → ItemKey
```

### Metadata Enrichment

Adds:

* site_code
* site_name
* region
* du_code
* subcon_ti
* contract_number
* purchasing_area

### Subcontractor Logic

Flow:

```text
EPMS SubCon - TI
    ↓
Normalize against official subcon list
    ↓
Retrieve contract number
```

### Matching Logic

Uses:

```text
(site_code + du_code)
```

to lookup EPMS records.

---

# 2. simple_calculation.py

Purpose:
Apply telecom quantity calculation rules.

## Input

```text
simple_normalized.json
```

## Output

```text
simple_calculated.json
```

## Responsibilities

* semantic quantity calculations
* conditional telecom logic
* reusable engineering calculations

---

# 3. simple_pr_generator.py

Purpose:
Generate ECC PR line items.

## Input

```text
simple_calculated.json
```

## Output

```text
simple_pr_output.json
```

## Responsibilities

Generate:

* MaterialCode
* LineItemText
* PR_Mode
* FinalPRQty

based on:

* ItemKey
* quantity ranges
* MainRuleTable rules

---

# 4. simple_ecc_export.py

Purpose:
Generate final ECC Excel output.

## Input

```text
simple_pr_output.json
```

## Output

```text
output/ECC_PR_Output.xlsx
```

## ECC Fields Generated

| ECC Column         | Source          |
| ------------------ | --------------- |
| SN.                | Auto numbering  |
| Purchasing Area    | purchasing_area |
| Region             | region          |
| Site ID            | site_code       |
| Site Name          | site_name       |
| Delivery Unit Code | du_code         |
| Logical Site       | blank           |
| Contract Number    | contract_number |
| Subcontractor      | subcon_ti       |
| PBOM Code          | MaterialCode    |
| SOW                | LineItemText    |
| Unit               | blank           |
| Quantity           | FinalPRQty      |
| Remarks            | blank           |

---

# Full Execution Workflow

## Step 1 — Normalize BOM

```bash
python src/simple_normalize.py
```

Generates:

```text
output/simple_normalized.json
```

---

## Step 2 — Run Calculations

```bash
python src/simple_calculation.py
```

Generates:

```text
output/simple_calculated.json
```

---

## Step 3 — Generate PR Items

```bash
python src/simple_pr_generator.py
```

Generates:

```text
output/simple_pr_output.json
```

---

## Step 4 — Export ECC File

```bash
python src/simple_ecc_export.py
```

Generates:

```text
output/ECC_PR_Output.xlsx
```

---

# One Click Pipeline

Run entire workflow:

```bash
python src/run_pipeline.py
```

Pipeline sequence:

```text
simple_normalize.py
    ↓
simple_calculation.py
    ↓
simple_pr_generator.py
    ↓
simple_ecc_export.py
```

---

# Design Philosophy

The engine is built around:

## Semantic Engineering Logic

Instead of:

* hardcoded BOM columns
* Excel formulas
* vendor-specific naming

the system uses:

* ItemKeys
* semantic categories
* reusable telecom logic

---

# Example Semantic Flow

## Raw BOM

```text
Rectifier-ZTE Rectifier Module 3000W
```

## Normalized

```text
Power Module
```

## Calculation

```text
Rectifier Installation Qty
```

## PR Rule

```text
Generate Rectifier Installation SOW
```

## ECC Export

```text
350000597821
DCPD or PDU Installation
```

---

# Current Technology Stack

## Current Implementation

* Python
* Pandas
* Excel Config Engine
* JSON Pipeline
* VS Code

---

# Current Completed Features

## Completed

* BOM normalization
* semantic ItemKey mapping
* EPMS subcontractor lookup
* subcontractor normalization
* contract number mapping
* purchasing area mapping
* calculation rule engine
* PR generation engine
* ECC export generation
* one-click pipeline execution

---

# Future Improvements

## Phase 2

* GUI frontend
* drag-and-drop BOM upload
* validation dashboard
* duplicate detection
* missing rule detection

---

## Phase 3

* AI-assisted rule suggestion
* automatic ItemKey recommendation
* PR anomaly detection
* telecom engineering knowledge base
* auto rule generation
* BOM comparison engine

---

# Important Notes

## BOM Header

BOM header row:

```text
Excel Row 3
```

Code:

```python
header=2
```

---

## EPMS Header

EPMS header row:

```text
Excel Row 4
```

Code:

```python
header=3
```

---

# Final Output

Final deliverable:

```text
output/ECC_PR_Output.xlsx
```

Ready for:

* ECC upload
* procurement processing
* telecom PR submission

```
```
#   c r e a t e - p r - c d - r a n  
 #   c r e a t e - p r - c d - r a n  
 #   c r e a t e - p r - c d - r a n  
 