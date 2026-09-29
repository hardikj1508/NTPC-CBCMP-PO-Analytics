# Data Dictionary

## Purpose

This document describes the main fields used in the consolidated PO analytics model.

The public portfolio version uses sanitized or synthetic operational values. The field definitions describe the structure and analytical role of the current Unified Master and its supporting control and QA layers.

---

## Unified Master Fields

| Field | Description | Role |
|---|---|---|
| `PO No.` | Primary PO identifier used for the active master | Key |
| `Legacy PO Number` | Legacy/reference PO number where available | Relationship |
| `Department` | Department associated with the PO where available | Classification |
| `PO Desc.` | Purchase order description | Description |
| `PO Value` | Purchase order monetary value | Financial |
| `Creation Dt` | PO creation date | Date |
| `Expiry Dt` | PO expiry date | Date |
| `PO Vendor` | Vendor identifier/name available in source data | Vendor |
| `Vendor Description` | Vendor description available in source data | Vendor |
| `EIC` | EIC field available in source data | Reference |
| `PO_Status` | Final/processed PO status used in the analytical workflow | Status |
| `Status` | Consolidated PO status | Status |
| `Present CCP Status` | Present CCP status associated with the PO | Status |
| `Status of CCP/CCC` | CCP/CCC status information | Status |
| `Final Dev. Status` | Final development/status information available from source data | Status |
| `Closed Date` | Date associated with closure | Date |
| `Remarks` | Record-level remarks | Notes |
| `Data Source(s)` | Source dataset(s) contributing information to the record | Traceability |
| `Data Conflict` | Indicator of detected conflicting information | Data Quality |
| `Validation Status` | Result/status of validation processing | Data Quality |
| `Added Date` | Date associated with addition to the active workflow | Audit |
| `Pending Ordering Plant` | Ordering plant information from pending-contract data | Pending Contract |
| `Pending Delivery Plant` | Delivery plant information from pending-contract data | Pending Contract |
| `Pending Status` | Pending-contract status | Pending Contract |
| `Pending Remarks` | Remarks associated with pending-contract processing | Pending Contract |

---

## Manual Control Tables

The workbook uses dedicated manual-control tables instead of directly editing Power Query output.

### Manual Additions

`Manual_PO_Additions` provides a structured input area for approved PO records that need to enter the active workflow.

The table follows the Unified Master structure and contains fields for:

- PO identifier
- Legacy PO Number
- Department
- PO description
- PO value
- Creation and expiry dates
- Vendor information
- EIC
- PO and CCP/CCC status information
- Development status
- Closed Date
- Remarks
- Data Source(s)
- Validation Status
- Added Date
- Pending ordering plant
- Pending delivery plant
- Pending status
- Pending remarks

The portfolio version keeps this control table available for future use and demonstration.

### Manual Overrides

`Manual_PO_Overrides` is used for controlled field-level corrections or overrides.

Typical control information includes:

- PO identifier
- Field to override
- Manual value
- Reason
- Source/Reference
- Updated By
- Update Date
- Active flag

Multiple active overrides can be applied to different fields of the same PO while the Unified Master remains one row per PO.

### Manual Exclusions

`Manual_PO_Exclusions` is used to identify approved records that should be excluded from the active workflow.

Typical control information includes:

- PO identifier
- Reason
- Source/Reference
- Excluded By
- Exclusion Date
- Active flag

---

## QA Fields

The automated QA layer evaluates the consolidated model using checks such as:

- Unique PO counts
- Duplicate PO detection
- Missing-from-master checks
- Status conflicts
- Date conflicts
- Alignment reviews
- Traceability issues
- Non-PO record detection

Each QA item contains:

- Check description
- Actual result
- QA status
- Notes

Typical QA statuses include:

- `PASS`
- `REVIEW`

`REVIEW` indicates that a check requires further validation or supporting evidence rather than automatically indicating an error.

---

## PO Validation

The project uses a standardized PO validation rule.

A valid PO is:

> A trimmed value containing exactly 10 numeric digits.

Examples of invalid records include:

- Blank values
- Headers
- Notes
- Descriptive text
- Values containing letters
- Values containing punctuation
- Other non-PO records

Validation is implemented through the reusable Power Query function:

`fn_ValidatePO`

The same validation principle is applied across the relevant source datasets.

---

## Status Interpretation

The Unified Master contains multiple status-related fields because the source systems provide different types of status information.

These fields should not automatically be treated as interchangeable.

Examples include:

- `Present CCP Status`
- `Status`
- `Status of CCP/CCC`
- `Pending Status`
- `PO_Status`
- `Final Dev. Status`
- `Data Conflict`

The consolidated workflow retains these fields so that source-level information, processed status, and reconciliation outcomes remain distinguishable.

---

## Relationship Fields

### `Legacy PO Number`

Used to retain a legacy/reference PO relationship where such information is available.

Relationships should be based on documented source evidence or controlled mapping.

Unverified relationships should remain identifiable as requiring review rather than being inferred automatically.

The workflow does not construct or invent legacy-to-current PO relationships.

---

## Data Quality Principles

The model follows these principles:

1. Do not infer missing PO identifiers.
2. Do not construct PO numbers.
3. Do not treat a blank field as a confirmed status.
4. Do not overwrite source information without controlled exception handling.
5. Keep reconciliation issues visible.
6. Preserve source traceability where available.
7. Separate validated relationships from relationships requiring review.
8. Maintain one active Master record per PO.
9. Apply manual corrections through controlled input tables rather than directly editing Power Query output.

---

## Public Portfolio Protection

The public repository should not expose actual:

- PO numbers
- Vendor names
- Vendor codes
- Employee information
- Confidential descriptions
- Internal documents
- Operational correspondence
- Local file paths

The public version uses synthetic source data and synthetic examples.

Only the structure, methodology, transformation logic, analytical approach, and sanitized examples should be published.