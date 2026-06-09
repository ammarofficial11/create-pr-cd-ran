# Telecom PR Automation Engine

## Overview

Telecom PR Automation Engine is a Python-based automation platform designed to generate ECC Purchase Requisition (PR) outputs from telecom BOM and EPMS data.

The system replaces manual Excel processing, repetitive copy-paste activities, and complex spreadsheet formulas with a configurable, rule-driven automation pipeline.

### Key Benefits

* Automated BOM normalization
* Automated subcontractor lookup
* Automated contract number mapping
* Automated purchasing area mapping
* Telecom-specific quantity calculations
* ECC-compatible PR generation
* General Item automation support
* Web-based user interface
* Reduced manual effort and human error

---

## Current Features

### Core Processing

* BOM normalization
* Equipment semantic mapping
* Calculation rule engine
* PR generation engine
* ECC Excel export

### Data Enrichment

* EPMS subcontractor lookup
* Subcontractor normalization
* Contract number mapping
* Purchasing area mapping

### Web Application

* Upload BOM
* Upload EPMS
* Dynamic project selection
* Generate PR from browser
* Download ECC output
* Download ECC output with General Items
* Job status tracking
* Generation audit trail

---

## Technology Stack

### Backend

* Python 3.x
* FastAPI
* Pandas
* OpenPyXL

### Frontend

* HTML
* CSS
* JavaScript

### Data Format

* Excel
* JSON

---

## High Level Workflow

```text
BOM.xlsx
EPMS.xlsx
ReferenceSubcon&Region.xlsx
          |
          v
  Normalize Equipment
          |
          v
  Apply Calculation Rules
          |
          v
  Generate PR Lines
          |
          v
  Export ECC Template
          |
          v
ECC_PR_Output.xlsx
```

---

## Web Application Workflow

```text
1. Upload BOM
2. Upload EPMS
3. Select Project
4. Generate PR
5. Wait for Processing
6. Download ECC Output
```

---

## Project Structure

```text
Telecom-PR-Automation
│
├── api
│   ├── __init__.py
│   └── app.py
│
├── config
│   ├── MainConfig.xlsx
│   ├── ReferenceSubcon&Region.xlsx
│   ├── GENERAL ITEM FOR ALL DU PROJECT Overall.xlsx
│   └── ecc_template.xls
│
├── input
│   ├── BOM.xlsx
│   └── EPMS.xlsx
│
├── output
│   ├── simple_normalized.json
│   ├── simple_calculated.json
│   ├── simple_pr_output.json
│   ├── general_pr_output.json
│   ├── ECC_PR_Output.xlsx
│   └── ECC_PR_Output_With_GeneralItems.xlsx
│
├── src
│   ├── simple_normalize.py
│   ├── simple_calculation.py
│   ├── simple_pr_generator.py
│   ├── simple_ecc_export.py
│   └── run_pipeline.py
│
├── web
│   └── index.html
│
└── README.md
```

---

## Processing Pipeline

### Step 1 - Normalize BOM

Script:

```bash
python src/simple_normalize.py
```

Responsibilities:

* Equipment normalization
* Site metadata extraction
* EPMS lookup
* Subcontractor lookup
* Contract number lookup
* Purchasing area lookup

Output:

```text
output/simple_normalized.json
```

---

### Step 2 - Apply Calculation Rules

Script:

```bash
python src/simple_calculation.py
```

Responsibilities:

* Semantic quantity calculations
* Engineering rule processing
* Conditional logic evaluation

Output:

```text
output/simple_calculated.json
```

---

### Step 3 - Generate PR Lines

Script:

```bash
python src/simple_pr_generator.py
```

Responsibilities:

* ECC material selection
* PR quantity generation
* General Item processing
* ECC line item preparation

Output:

```text
output/simple_pr_output.json
output/general_pr_output.json
```

---

### Step 4 - Export ECC Output

Script:

```bash
python src/simple_ecc_export.py
```

Responsibilities:

* ECC row creation
* Excel export generation

Output:

```text
output/ECC_PR_Output.xlsx
output/ECC_PR_Output_With_GeneralItems.xlsx
```

---

## One Click Execution

Run the full pipeline:

```bash
python src/run_pipeline.py
```

Pipeline order:

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

## Input Files

### BOM.xlsx

Primary engineering source file.

Contains:

* Site Code
* Site Name
* Region
* DU Code
* Equipment quantities

Header row:

```python
header=2
```

---

### EPMS.xlsx

Deployment reference source.

Used for:

* Subcontractor lookup
* Deployment metadata

Header row:

```python
header=3
```

Key lookup:

```text
(site_code + du_code)
```

---

### ReferenceSubcon&Region.xlsx

Reference data source.

Used for:

* Region → Purchasing Area mapping
* Official Subcontractor mapping
* Contract Number mapping

---

## General Item Automation

The system supports project-based General Item generation.

Configuration source:

```text
config/GENERAL ITEM FOR ALL DU PROJECT Overall.xlsx
```

Capabilities:

* Region-based item selection
* Project-based item selection
* Optional item logic
* Transportation model support
* Automatic PR line generation

Output:

```text
output/general_pr_output.json
output/ECC_PR_Output_With_GeneralItems.xlsx
```

---

## REST API

### Get Projects

```http
GET /projects
```

Returns available General Item projects.

---

### Upload BOM

```http
POST /upload-bom
```

---

### Upload EPMS

```http
POST /upload-epms
```

---

### Generate PR

```http
POST /generate-pr
```

Example:

```json
{
  "project": "CD consolidation 2023 (Swap/ Modernize)"
}
```

---

### Check Job Status

```http
GET /job-status
```

---

### Download ECC Output

```http
GET /download-pr
```

---

### Download ECC Output With General Items

```http
GET /download-pr-general
```

---

### Open Web UI

```http
GET /ui
```

---

## Current Status

### Completed

* BOM normalization
* Equipment semantic mapping
* EPMS integration
* Subcontractor normalization
* Contract number mapping
* Purchasing area mapping
* Calculation rule engine
* PR generation engine
* General Item automation
* ECC export generation
* FastAPI backend
* Web user interface
* Job tracking
* Output timestamp tracking

---

## Planned Roadmap

### Phase 3.1

* UI refresh state recovery
* Persistent status restoration

### Phase 3.2

* Site filtering
* Single-site PR generation
* Multi-site PR generation

### Phase 4

* Job storage system
* Historical outputs
* Download previous jobs

### Phase 5

* AI Assistant
* PR explanation engine
* Missing material analysis
* Quantity troubleshooting

---

## Author

Developed as an internal telecom engineering automation platform for RAN Purchase Requisition generation and ECC export preparation.
