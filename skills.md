````markdown
# Telecom PR Automation Agent Skills

# Purpose

The AI agent assists with:

- BOM normalization
- telecom equipment categorization
- semantic quantity calculation
- PR line item generation
- ECC export generation
- subcontractor normalization
- contract mapping
- mismatch detection
- rule analysis
- validation checking

---

# Current Production Pipeline

The current production workflow is:

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
```

All processing stages are intentionally separated to:
- simplify debugging
- isolate business logic
- reduce coupling
- simplify AI-assisted maintenance
- simplify future GUI integration
- simplify future API integration

---

# Domain Knowledge

The AI agent understands telecom infrastructure equipment categories.

## Equipment Categories

Examples:
- RRU
- Antenna
- AAU
- BBU
- BBU Cards
- Cabinet
- Power Module
- Battery
- Feeder
- Feeder Connector
- GPS
- DCPD
- ODCPD
- Combiner
- Bracket
- Post
- Circuit Breaker
- Security Bar
- PadPower

---

# Semantic Architecture

The AI agent must understand the following execution flow:

```text
Raw BOM Equipment
        ↓
Equipment_Normalization
        ↓
Semantic ItemKeys
        ↓
Calculation_Rules
        ↓
Calculated Quantities
        ↓
MainRuleTable
        ↓
PR Line Items
        ↓
ECC Template Export
```

---

# Static Lookup Tables

Purpose:
Provide static configuration and output templates that support the engine without containing rule logic.

Examples:
- ECC_Template.xlsx
- ReferenceSubcon&Region.xlsx

Important:
- These files are NOT part of normalization logic
- These files are NOT part of calculation logic
- These files provide static lookup values only
- These files provide output formatting only

---

# EPMS Integration

Purpose:
Retrieve subcontractor deployment information.

## EPMS Lookup Key

```text
(site_code + du_code)
```

## Required EPMS Columns

- customer site code
- du code
- SubCon - TI

## EPMS Logic Flow

```text
site_code + du_code
        ↓
lookup EPMS
        ↓
retrieve SubCon - TI
        ↓
normalize subcontractor
        ↓
retrieve contract number
        ↓
enrich metadata
```

---

# Metadata Enrichment Layer

The normalization phase also enriches metadata.

## Added Metadata

- site_code
- site_name
- region
- du_code
- subcon_ti
- contract_number
- purchasing_area

Purpose:
Allow downstream stages to reuse metadata consistently.

Used by:
- calculation engine
- PR generator
- ECC export

---

# JSON Contract Philosophy

Each pipeline stage exports structured JSON.

Purpose:
- decouple pipeline stages
- simplify debugging
- simplify AI inspection
- simplify future APIs
- simplify future GUI integration
- create stable interfaces between stages

## JSON Pipeline

```text
simple_normalized.json
        ↓
simple_calculated.json
        ↓
simple_pr_output.json
```

Each JSON output acts as a stable interface contract between stages.

---

# Equipment_Normalization Layer

Purpose:
Normalize raw BOM equipment names into semantic ItemKeys.

Example:

```text
Rectifier-Huawei Rectifier Module 3000W
→ Power Module
```

Important Rules:
- Never depend on raw BOM names directly
- Always normalize into semantic ItemKeys
- Vendor naming differences must be absorbed during normalization
- BOM column positions must not affect logic
- ItemKeys are the semantic source of truth

---

# Calculation_Rules Layer

Purpose:
Execute semantic telecom calculation logic.

The AI agent must interpret calculation behavior based on CalcType.

---

# Supported Calculation Types

## SUM

Meaning:
Aggregate all quantities under same ItemKey.

Example:

```text
Sum all ItemKey = RRU
```

---

## CONDITIONAL_SUM_BY_ITEMKEY

Meaning:
Conditionally aggregate semantic quantities.

Example:

```text
If Cabinet > 0:
    return 0
Else:
    sum all PowerModule quantities
```

Another Example:

```text
If BBU Cards > 0:
    return 0
Else:
    sum all BBU quantities
```

---

## CONDITIONAL_DIVIDE

Meaning:
Divide semantic quantity only when condition passes.

Example:

```text
If Feeder 7/8 > 0:
    connector qty / 2
Else:
    0
```

---

## DIVIDE_BY_ITEMKEY

Meaning:
Divide one semantic quantity by another.

Example:

```text
Feeder 1/2
/
FeederConnector 1/2
```

If result >= 20:
- round up

Else:
- return 0

---

# MainRuleTable Layer

Purpose:
Convert calculated semantic quantities into:
- PR line items
- PBOM quantities
- installation quantities
- ECC upload rows

This layer contains:
- procurement rules
- PR business logic
- line item mapping
- ECC field mapping

Important:
- MainRuleTable must NOT contain calculation formulas
- Calculation logic belongs ONLY in Calculation_Rules
- Export scripts must NOT contain telecom business logic

---

# ECC Export Layer

Purpose:
Generate final customer ECC upload format.

The export layer should:
- remain formatting-focused
- avoid telecom business calculations
- avoid procurement logic
- consume finalized PR outputs only

ECC export should behave as a presentation/output layer only.

---

# Validation Philosophy

The system must validate:
- unmatched ItemKeys
- missing subcontractors
- missing contract numbers
- duplicate sites
- invalid quantities
- empty PR outputs
- invalid EPMS mappings
- missing normalization mappings

Validation should occur BEFORE ECC export.

---

# AI Agent Responsibilities

---

# 1. Normalize BOM

Convert raw BOM equipment names into semantic ItemKeys.

Example:

```text
Rectifier-ZTE Rectifier Module 3000W
→ Power Module
```

---

# 2. Execute Calculation Rules

Apply semantic telecom calculations.

Examples:
- SUM
- CONDITIONAL_SUM_BY_ITEMKEY
- CONDITIONAL_DIVIDE
- DIVIDE_BY_ITEMKEY

---

# 3. Generate PR Quantities

Use calculated semantic quantities to generate:
- PBOM quantities
- installation quantities
- PR line items

---

# 4. Generate ECC Export

Generate customer ECC upload format using:
- PR outputs
- metadata enrichment
- subcon normalization
- contract mappings

---

# 5. Detect Mismatch

Compare:
- generated PR
- existing PR
- BOM quantities
- EPMS data

Detect:
- missing PR lines
- incorrect quantities
- mismatched ItemKeys
- dependency conflicts
- quantity inconsistencies
- missing subcon
- invalid contracts

---

# Important Architectural Rules

# NEVER

- Depend on Excel column positions
- Depend on raw BOM names directly
- Use hardcoded formulas
- Use worksheet-specific logic
- Mix normalization with calculation logic
- Mix calculation logic with PR business rules
- Embed telecom business logic directly into export scripts
- Couple pipeline stages tightly

---

# ALWAYS

- Use semantic ItemKeys
- Use config-driven rules
- Use Equipment_Normalization layer
- Use Calculation_Rules layer
- Use MainRuleTable for PR generation
- Separate normalization from calculation
- Separate calculation from procurement logic
- Export structured JSON between stages
- Preserve metadata consistency across stages

---

# AI Agent Thinking Model

The AI agent must think semantically.

Incorrect:

```text
N2 + R2 + U2 + V2
```

Correct:

```text
Sum all ItemKey = Power Module
```

Incorrect:

```text
AF2:AP2
```

Correct:

```text
Sum all ItemKey = BBU
```

Incorrect:

```text
Check Excel column 73
```

Correct:

```text
Check semantic ItemKey = GPS
```

---

# Future AI Features

Potential future capabilities:
- auto-suggest ItemKeys
- auto-detect new equipment
- auto-generate calculation rules
- detect duplicated rules
- detect missing normalization
- validate PR completeness
- explain calculation reasoning
- compare PR revisions
- detect procurement anomalies
- AI-assisted telecom rule discovery
- BOM comparison intelligence
- ECC validation assistant

---

# Development Goal

Build a config-driven telecom procurement automation engine using:
- semantic equipment normalization
- reusable calculation patterns
- PR business rules
- AI-assisted validation
- Python execution engine
- JSON pipeline architecture
- VS Code AI workflow
- future AI-assisted engineering automation
````
