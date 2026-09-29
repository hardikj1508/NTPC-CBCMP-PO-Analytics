# NTPC CBCMP Procurement & Contract Closing Analytics System

An end-to-end Excel + Power Query analytics solution designed to consolidate, clean, validate, reconcile, and analyze Purchase Order (PO) and contract-closing data.

The project demonstrates how Excel and Power Query can transform multiple structured source datasets into a unified PO master, automated data-quality checks, controlled exception handling, and an executive dashboard.

> Portfolio version: The public workbook uses sanitized synthetic data. It is designed to demonstrate the architecture, transformation logic, analytical approach, and reporting workflow without exposing confidential operational information.

---

## Project Objectives

- Consolidate PO data from multiple source datasets
- Clean and standardize PO identifiers and source fields
- Validate records using a defined PO validation rule
- Create a unified PO master database
- Reconcile data across ERP/Legacy, CCP, pending-contract, and department sources
- Detect status conflicts and other data-quality issues
- Provide controlled handling of manual additions, overrides, and exclusions
- Build an executive-level dashboard for contract-closing analysis
- Create a refreshable workflow using Power Query
- Demonstrate a structured Excel-based analytics architecture suitable for repeatable reporting

---

## Solution Architecture

The workflow follows this structure:

**Source Files → Raw Queries → Cleaning & Validation → Source Consolidation → Manual Controls → Unified Master → QA Checks → Dashboard**

### Main Data Sources

The portfolio architecture represents the following source categories:

1. ERP / Legacy PO List
2. CCP Status Data
3. Pending Contract Closing Data
4. Department-wise PO Data
5. Controlled manual additions, overrides, and exclusions

The public repository uses synthetic source files representing these data structures.

---

## Power Query Components

The workbook contains dedicated Power Query components for:

- Source configuration
- PO validation
- Raw data ingestion
- Data cleaning
- CCP staging and consolidation
- Pending-contract processing
- Department-wise consolidation
- Manual additions
- Manual overrides
- Manual exclusions
- Unified master construction
- Automated QA checks

### Core Queries

- `fn_GetSetting`
- `fn_ValidatePO`
- `q_Legacy_Raw`
- `q_Legacy_Cleaned`
- `q_CCP_Raw`
- `q_CCP_Cleaned`
- `q_CCP_Staging`
- `q_Pending_Raw`
- `q_Pending_Cleaned`
- `q_Department_Cleaned`
- `q_Manual_Additions`
- `q_Manual_Overrides`
- `q_Manual_Exclusions`
- `q_Manual_Additions_Raw`
- `q_Manual_Overrides_Raw`
- `q_Manual_Exclusions_Raw`
- `q_Unified_Master`
- `q_QA_Checks`

---

## PO Validation Rule

A valid PO is defined as:

> A trimmed value containing exactly 10 numeric digits.

Invalid records include:

- Blank values
- Headers
- Notes or comments
- Descriptive text
- Values containing letters
- Values containing punctuation
- Other non-PO records

The validation rule is implemented through the reusable Power Query function:

`fn_ValidatePO`

This allows the same validation logic to be applied consistently across different source datasets.

---

## Unified Master Database

The current portfolio version contains:

**25 unique valid POs**

The `q_Unified_Master` query creates the active consolidated PO population from the available synthetic source datasets and controlled manual inputs.

The master is designed to maintain source traceability while supporting controlled exception handling.

### Key Fields

- PO No.
- Legacy PO Number
- Department
- PO Desc.
- PO Value
- Creation Dt
- Expiry Dt
- PO Vendor
- Vendor Description
- EIC
- PO_Status
- Status
- Present CCP Status
- Status of CCP/CCC
- Final Dev. Status
- Closed Date
- Remarks
- Data Source(s)
- Data Conflict
- Validation Status
- Added Date
- Pending Ordering Plant
- Pending Delivery Plant
- Pending Status
- Pending Remarks

---

## Automated Data Quality Checks

The workbook contains an automated **14-item QA framework**.

Current checks include:

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

### Current Portfolio QA Results

| QA Check | Result | Status |
|---|---:|---|
| ERP Unique Legacy POs | 10 | PASS |
| CCP Unique POs | 15 | PASS |
| Pending Unique POs | 10 | PASS |
| Department Unique POs | — | PASS |
| Master Unique POs | 25 | PASS |
| Duplicate Master POs | 0 | PASS |
| ERP Relationship Review | 10 | REVIEW |
| CCP Missing from Master | 0 | PASS |
| Pending Missing from Master | 0 | PASS |
| Status Conflicts | 0 | PASS |
| Date Conflicts / Review | — | REVIEW |
| Alignment Reviews | — | REVIEW |
| Traceability Issues | — | REVIEW |
| Non-PO Records in Master | 0 | PASS |

The QA framework is designed to update dynamically when the underlying Power Query data is refreshed.

`REVIEW` statuses represent checks that require additional relationship-level, date-level, structural, or traceability validation rather than automatically indicating a data error.

---

## Exception Management

Instead of manually editing Power Query output tables, the workbook uses dedicated Excel control tables for:

- Manual additions
- Manual overrides
- Manual exclusions

These tables provide a controlled and traceable mechanism for handling legitimate exceptions without modifying the underlying query logic.

### Manual Additions

`Manual_PO_Additions` provides a structured input area for adding legitimate records that are not available in the automated source datasets.

The portfolio version keeps this table available for demonstration but does not require an active manual addition.

### Manual Overrides

`Manual_PO_Overrides` supports field-level overrides for existing PO records.

The portfolio demonstrates multiple overrides for the same PO without creating duplicate Master records.

### Manual Exclusions

`Manual_PO_Exclusions` provides a controlled mechanism for excluding records from the active Master population.

The public example uses synthetic data only.

---

## Executive Dashboard

The `Executive_Dashboard` provides a high-level view of the consolidated synthetic PO universe.

### Current Portfolio KPIs

- **Total POs:** 25
- **Not Closed:** 19
- **Closed:** 6
- **Total PO Value:** ₹20,365,000
- **Status Conflicts:** 0

### Visualizations

- PO Status Breakdown
- Department-wise PO Breakdown

The dashboard is connected to the consolidated Master and is designed to update after a Power Query refresh.

The departmental summary reconciles with the overall Master population.

---

## Data Reconciliation

The project demonstrates reconciliation across multiple source universes.

### Current Portfolio Source Counts

| Dataset | Unique Valid POs |
|---|---:|
| ERP / Legacy | 10 |
| CCP | 15 |
| Pending Closing | 10 |
| Department | — |
| Unified Master | 25 |

The reconciliation framework distinguishes between:

- Direct matches
- Missing records
- Status conflicts
- Historical relationships
- Records requiring review

The public version uses synthetic source populations, so the counts above describe the portfolio dataset rather than operational NTPC data.

---

## Historical & Relationship Tracking

The architecture separates active Master data from relationship and exception logic where appropriate.

Legacy-to-SAP/ERP relationships are not automatically assumed to be equivalent simply because two PO identifiers appear in different source systems.

Unverified relationships can instead be classified for review rather than being treated as confirmed relationships.

This approach is intended to reduce false matches during multi-source reconciliation.

---

## Workbook Structure

The final portfolio workbook contains:

| Sheet | Purpose |
|---|---|
| `Navigation_Hub` | Workbook navigation and workflow guidance |
| `Executive_Dashboard` | Executive KPIs and analytical visualizations |
| `q_Unified_Master` | Active consolidated PO master |
| `q_QA_Checks` | Automated 14-item data-quality framework |
| `Manual_PO_Additions` | Controlled manual record additions |
| `Manual_PO_Overrides` | Controlled field-level overrides |
| `Manual_PO_Exclusions` | Controlled record exclusions |
| `PQ_Settings` | Source configuration and refresh settings |

The public workbook does not require the historical operational output sheets used during development.

---

## Technology Stack

### Primary Tools

- Microsoft Excel
- Power Query
- Power Query M

### Supporting Concepts

- Data cleaning
- Data validation
- Data reconciliation
- Exception handling
- Quality assurance
- Dashboarding
- Data transformation
- Source traceability
- Master-data management

No Python, VBA, or Office Scripts are required for the core data-processing workflow.

---

## Refresh Workflow

The intended workflow is:

1. Update or replace the synthetic/source files
2. Open the workbook
3. Refresh Power Query using **Data → Refresh All**
4. Review automated QA checks
5. Review exceptions requiring attention
6. Review the Unified Master
7. Review the Executive Dashboard

Manual changes should be made through the designated control tables rather than directly inside Power Query output tables.

---

## Source Configuration

The portfolio workbook uses a configurable source-folder approach rather than hard-coding individual source paths throughout the transformation logic.

`PQ_Settings` contains the configuration required by the Power Query workflow, including the source folder and source filenames.

This makes the workbook easier to move between environments and simplifies source-file replacement.

---

## Key Outcomes

The project demonstrates the ability to:

- Build a multi-source data pipeline using Power Query
- Design reusable validation functions
- Consolidate heterogeneous Excel datasets
- Create a controlled master-data workflow
- Implement automated data-quality checks
- Handle exceptions without breaking query logic
- Reconcile multiple PO populations
- Build an executive reporting layer
- Separate automated transformation from controlled manual intervention
- Design a refreshable Excel analytics solution
- Structure an analytics project for portfolio and GitHub presentation

---

## Project Context

This project was developed around a procurement and contract-closing analytics workflow involving multiple PO sources and department-level datasets.

The public portfolio version uses synthetic data and is intentionally separated from confidential operational records.

The objective of the repository is to demonstrate:

- Data-engineering logic
- Power Query transformation design
- Data-quality methodology
- Reconciliation techniques
- Exception-management workflow
- Analytical reporting
- Dashboard design

without exposing confidential operational information.

---

## Portfolio Note

This repository is a portfolio representation of an **Excel + Power Query data analytics project**.

The public version uses **sanitized or synthetic data** rather than confidential operational records.

The synthetic dataset is intended to demonstrate the workflow and architecture; its PO numbers, vendors, plants, values, and other operational attributes are fictional and should not be interpreted as real operational records.

---

## Repository Structure

```text
ntpc_cbcmp_github/
│
├── workbook/
│   └── NTPC_CBCMP_PO_Analytics_Portfolio.xlsx
│
├── portfolio_data/
│   ├── Legacy_PO_Source.xlsx
│   ├── CCP_Status_Source.xlsx
│   └── Pending_Closing_Source.xlsx
│
├── docs/
│   ├── architecture.md
│   └── data-dictionary.md
│
├── screenshots/
│
├── README.md
└── .gitignore
```

---

## Disclaimer

This repository is a **portfolio demonstration project**.

It is not an official NTPC reporting system and should not be used for operational decision-making.

The public dataset is synthetic and is included solely to demonstrate the data transformation, validation, reconciliation, exception-management, and reporting techniques used in the project.