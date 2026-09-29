# Power Query Architecture

## Overview

The NTPC CBCMP PO Analytics System uses a layered Power Query architecture to transform multiple structured source datasets into a validated Unified Master, automated QA layer, and executive reporting layer.

The public portfolio version uses sanitized synthetic source data while preserving the architecture and transformation approach of the original analytics workflow.

The architecture is designed around:

- Source configuration
- Source ingestion
- Data cleaning
- PO validation
- Source consolidation
- Department consolidation
- Controlled manual inputs
- Unified Master construction
- Automated quality assurance
- Dashboard reporting

---

## Architecture Flow

```text
Source Configuration
        │
        ▼
   Source Files
        │
        ├── ERP / Legacy PO Data
        ├── CCP Status Data
        ├── Pending Contract Closing Data
        └── Department-wise Data
                │
                ▼
       Raw / Staging Queries
                │
                ▼
       Cleaning & Validation
                │
                ▼
       Source Consolidation
                │
                ▼
      Department Consolidation
                │
                ▼
        Manual Control Layer
        ├── Additions
        ├── Overrides
        └── Exclusions
                │
                ▼
          Unified Master
                │
        ┌───────┴────────┐
        ▼                ▼
 Automated QA      Executive Dashboard
 Checks
```

---

## Source Configuration

The workbook uses a dedicated `PQ_Settings` area to store configurable source information.

This allows the Power Query workflow to reference source-folder and source-file settings without embedding individual file paths throughout every query.

The portfolio source configuration represents:

- Legacy PO source file
- CCP status source file
- Pending contract-closing source file
- Portfolio source folder

This approach makes the workbook easier to move between environments and simplifies source-file replacement.

---

## Reusable Functions

### `fn_GetSetting`

`fn_GetSetting` retrieves configuration values from the `PQ_Settings` table.

It is used to obtain settings such as:

- Source Folder
- Legacy PO File
- CCP File
- Pending File

This centralizes source configuration.

### `fn_ValidatePO`

`fn_ValidatePO` provides a reusable PO validation rule across source datasets.

A valid PO is defined as:

> A trimmed value containing exactly 10 numeric digits.

The function rejects blanks, headers, notes, descriptive text, alphabetic values, punctuation, and other non-PO records.

Using a reusable function ensures that PO validation is applied consistently rather than being duplicated with different logic in individual queries.

---

## Source Ingestion Layer

The source ingestion layer reads the configured source files and exposes their raw structures to downstream queries.

### Legacy / ERP

- `q_Legacy_Raw`
- `q_Legacy_Cleaned`

`q_Legacy_Raw` reads the Legacy PO source and preserves the source structure.

`q_Legacy_Cleaned` applies text cleaning and the reusable PO validation function.

### CCP

- `q_CCP_Raw`
- `q_CCP_Cleaned`
- `q_CCP_Staging`

The CCP workflow combines the relevant CCP source structures and applies PO cleaning and validation.

`q_CCP_Staging` provides a standardized staging layer for downstream consolidation.

### Pending Contract Closing

- `q_Pending_Raw`
- `q_Pending_Cleaned`

The Pending source is loaded through the configured source file, cleaned, and validated before being incorporated into the consolidated Master.

### Department Data

- `q_Department_Cleaned`

Department data is consolidated into a standardized structure before being used by the Unified Master.

---

## Cleaning & Validation Layer

The cleaning layer standardizes source values before reconciliation.

Typical transformations include:

- Trimming whitespace
- Removing non-printing characters
- Converting PO identifiers to a consistent text representation
- Applying the reusable PO validation rule
- Standardizing column structures
- Handling source-specific layouts
- Preserving source information required for traceability

Invalid PO records are excluded from the valid PO populations rather than being artificially converted into valid identifiers.

---

## Manual Control Layer

Manual intervention is separated from automated source transformation.

The workbook contains three controlled Excel input tables:

- `Manual_PO_Additions`
- `Manual_PO_Overrides`
- `Manual_PO_Exclusions`

These tables allow controlled exceptions without directly editing Power Query output.

### Manual Additions

`q_Manual_Additions` processes records entered into `Manual_PO_Additions`.

The table follows the Unified Master field structure so that manually added records can enter the same downstream schema.

### Manual Overrides

`q_Manual_Overrides` processes field-level overrides.

The override design supports multiple fields for the same PO without creating duplicate Master records.

For example, different fields for the same PO can be overridden independently while the Unified Master remains one row per PO.

### Manual Exclusions

`q_Manual_Exclusions` processes controlled exclusions.

Only active exclusions are applied to the consolidated Master population.

---

## Raw Manual-Control Queries

The workbook also contains raw query layers for the manual-control tables:

- `q_Manual_Additions_Raw`
- `q_Manual_Overrides_Raw`
- `q_Manual_Exclusions_Raw`

These provide a separation between the Excel input tables and the cleaned/transformed manual-control queries.

This follows the same layered principle used for source data ingestion.

---

## Unified Master Construction

### `q_Unified_Master`

`q_Unified_Master` is the central consolidated output of the Power Query workflow.

It brings together the validated source populations and controlled manual inputs into a single active PO Master.

The Master is designed to:

- Maintain one active row per PO
- Combine information from multiple source populations
- Preserve source traceability
- Apply controlled exclusions
- Apply field-level manual overrides
- Incorporate manual additions
- Identify data conflicts
- Preserve validation status
- Carry pending-contract information where available

The current public portfolio contains:

**25 unique valid POs.**

---

## Master Output Structure

The final Unified Master contains the following fields:

1. PO No.
2. Legacy PO Number
3. Department
4. PO Desc.
5. PO Value
6. Creation Dt
7. Expiry Dt
8. PO Vendor
9. Vendor Description
10. EIC
11. PO_Status
12. Status
13. Present CCP Status
14. Status of CCP/CCC
15. Final Dev. Status
16. Closed Date
17. Remarks
18. Data Source(s)
19. Data Conflict
20. Validation Status
21. Added Date
22. Pending Ordering Plant
23. Pending Delivery Plant
24. Pending Status
25. Pending Remarks

---

## Data Conflict & Exception Handling

The Unified Master does not silently overwrite source-level differences.

Where relevant, source differences are retained through fields such as:

- `Data Conflict`
- `Validation Status`
- `Data Source(s)`

This allows downstream users to distinguish between consolidated information and records requiring review.

Manual overrides provide a controlled mechanism when an authorized correction or clarification is required.

---

## Automated QA Layer

### `q_QA_Checks`

`q_QA_Checks` provides the automated quality-control layer.

The current framework contains 14 checks covering:

1. ERP Unique Legacy POs
2. CCP Unique POs
3. Pending Unique POs
4. Department Unique POs
5. Master Unique POs
6. Duplicate Master POs
7. ERP Relationship Review
8. CCP Missing from Master
9. Pending Missing from Master
10. Status Conflicts
11. Date Conflicts / Review
12. Alignment Reviews
13. Traceability Issues
14. Non-PO Records in Master

The QA layer is designed to refresh automatically with the underlying Power Query data.

---

## Reporting Layer

The consolidated Master and QA outputs feed the Excel reporting layer.

### Executive Dashboard

The `Executive_Dashboard` provides:

- Total PO count
- Closed PO count
- Not Closed PO count
- Total PO value
- Status conflict count
- Department-wise PO breakdown
- PO status visualization

The dashboard is formula-driven from the consolidated Master and is designed to update after a Power Query refresh.

---

## Refresh Workflow

The intended workflow is:

```text
1. Update / replace source files
            │
            ▼
2. Refresh Power Query
            │
            ▼
3. Review QA Checks
            │
            ▼
4. Review manual exceptions
            │
            ▼
5. Review Unified Master
            │
            ▼
6. Review Executive Dashboard
```

Manual changes should be made through the designated control tables rather than directly inside Power Query output tables.

---

## Design Principles

The architecture follows several important design principles:

### 1. Separate Raw Data from Transformation Logic

Raw and staging queries provide a controlled starting point for downstream transformations.

### 2. Reuse Validation Logic

The `fn_ValidatePO` function ensures consistent PO validation across different datasets.

### 3. Separate Automated Processing from Manual Intervention

Manual additions, overrides, and exclusions are maintained through dedicated control tables rather than direct edits to query outputs.

### 4. Preserve Traceability

Source information and data-quality indicators are retained so that consolidated records can be reviewed when necessary.

### 5. Avoid Unsupported Inference

Relationships between different PO identifiers are not automatically assumed to be equivalent without supporting evidence.

### 6. Maintain One Active Master Record per PO

The Unified Master is designed to prevent duplicate active records for the same PO, including when multiple manual overrides apply to the same PO.

### 7. Make the Workflow Refreshable

The architecture is designed so that source data can be replaced or updated and the downstream Master, QA checks, and dashboard can be refreshed without manually rebuilding the transformation workflow.

---

## Portfolio Data

The public portfolio uses synthetic source files representing the multi-source architecture.

The synthetic data is intended to demonstrate:

- Source ingestion
- Data cleaning
- PO validation
- Multi-source consolidation
- Exception management
- QA checks
- Dashboard reporting

No confidential operational records are required for the public workflow.

---

## Technology

The core solution uses:

- Microsoft Excel
- Power Query
- Power Query M

The core data-processing workflow does not require:

- Python
- VBA
- Office Scripts
- External databases

The project therefore demonstrates how a structured multi-source analytics workflow can be implemented using Excel and Power Query alone.