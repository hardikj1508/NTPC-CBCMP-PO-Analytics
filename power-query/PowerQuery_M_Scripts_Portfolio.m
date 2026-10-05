// ============================================================
// NTPC CBCMP PO ANALYTICS SYSTEM
// Power Query M Scripts - Portfolio Version
// ============================================================
//
// Purpose:
// Reference copy of the Power Query M logic used by the
// sanitized synthetic portfolio workbook.
//
// The sections below represent individual Power Query
// functions and queries. They are maintained separately
// inside the Excel workbook.
//
// Portfolio data only.
// ============================================================
// ============================================================
// 01. FUNCTION — fn_GetSetting
// ============================================================
(settingName as text) as text =>
let
    Source = Excel.CurrentWorkbook(){[Name="Table_PQSettings"]}[Content],
    FilteredRow = Table.SelectRows(Source, each ([Setting] = settingName)),
    Value = FilteredRow{0}[Value]
in
    Value
// ============================================================
// 02. FUNCTION — fn_ValidatePO
// ============================================================
(poValue as any) as logical =>
let
    cleanText = try Text.Trim(Text.Clean(Text.From(poValue))) otherwise "",
    isValid =
        Text.Length(cleanText) = 10
        and Text.Select(cleanText, {"0".."9"}) = cleanText
in
    isValid
// ============================================================
// 03. QUERY — q_Legacy_Raw
// ============================================================
let
    FolderPath = fn_GetSetting("Source Folder"),
    FileName = fn_GetSetting("Legacy PO File"),
    Source = Excel.Workbook(File.Contents(FolderPath & FileName), null),
    Sheet1 = Source{[Item="Legacy PO created after OCT 25",Kind="Sheet"]}[Data],
    Header1 = Table.PromoteHeaders(Sheet1, [PromoteAllScalars=true]),
    Sheet2 = Source{[Item="PO OCT 25 till date",Kind="Sheet"]}[Data],
    Header2 = Table.PromoteHeaders(Sheet2, [PromoteAllScalars=true]),
    Combined = Table.Combine({Header1, Header2}),
    RemoveEmpty = Table.SelectRows(Combined, each Record.Field(_, "Legacy PO No.") <> null)
in
    RemoveEmpty
// ============================================================
// 04. QUERY — q_Legacy_Cleaned
// ============================================================
let
    Source = q_Legacy_Raw,
    CleanPO = Table.TransformColumns(
        Source,
        {
            {
                "Legacy PO No.",
                each Text.Trim(Text.Clean(Text.From(_))),
                type text
            }
        }
    ),
    FilterValid = Table.SelectRows(
        CleanPO,
        each fn_ValidatePO(Record.Field(_, "Legacy PO No."))
    )
in
    FilterValid
// ============================================================
// 05. QUERY — q_CCP_Raw
// ============================================================
let
    FolderPath = fn_GetSetting("Source Folder"),
    FileName = fn_GetSetting("CCP File"),
    Source = Excel.Workbook(File.Contents(FolderPath & FileName), null, true),
    NotClosed = Source{[Item="Not Closed Master",Kind="Sheet"]}[Data],
    NotClosedHdr = Table.PromoteHeaders(NotClosed, [PromoteAllScalars=true]),
    NotClosedStatus = Table.AddColumn(NotClosedHdr, "PO_Status", each "Not Closed"),
    Closed = Source{[Item="Closed Master",Kind="Sheet"]}[Data],
    ClosedHdr = Table.PromoteHeaders(Closed, [PromoteAllScalars=true]),
    ClosedStatus = Table.AddColumn(ClosedHdr, "PO_Status", each "Closed"),
    Combined = Table.Combine({NotClosedStatus, ClosedStatus})
in
    Combined
// ============================================================
// 06. QUERY — q_CCP_Cleaned
// ============================================================
let
    Source = q_CCP_Raw,
    CleanPO =
        Table.TransformColumns(
            Source,
            {
                {
                    "PO No.",
                    each
                        try
                            Text.Trim(
                                Text.Clean(
                                    Text.From(_)
                                )
                            )
                        otherwise null,
                    type text
                }
            }
        ),
    FilterValid =
        Table.SelectRows(
            CleanPO,
            each fn_ValidatePO([#"PO No."])
        )
in
    FilterValid
// ============================================================
// 07. QUERY — q_CCP_Staging
// ============================================================
let
    FolderPath = fn_GetSetting("Source Folder"),
    FileName = fn_GetSetting("CCP File"),
    Source = Excel.Workbook(File.Contents(FolderPath & FileName), null, true),
    NotClosed = Source{[Item="Not Closed Master",Kind="Sheet"]}[Data],
    NotClosedHdr = Table.PromoteHeaders(NotClosed, [PromoteAllScalars=true]),
    NotClosedStatus = Table.AddColumn(NotClosedHdr, "PO_Status", each "Not Closed"),
    Closed = Source{[Item="Closed Master",Kind="Sheet"]}[Data],
    ClosedHdr = Table.PromoteHeaders(Closed, [PromoteAllScalars=true]),
    ClosedStatus = Table.AddColumn(ClosedHdr, "PO_Status", each "Closed"),
    Combined = Table.Combine({NotClosedStatus, ClosedStatus}),
    CleanPO = Table.TransformColumns(Combined, {{"PO No.", each Text.Trim(Text.Clean(Text.From(_))), type text}}),
    FilterValid = Table.SelectRows(CleanPO, each fn_ValidatePO(Record.Field(_, "PO No.")))
in
    FilterValid
// ============================================================
// 08. QUERY — q_Pending_Raw
// ============================================================
let
    FolderPath = fn_GetSetting("Source Folder"),
    FileName = fn_GetSetting("Pending File"),
    Source = Excel.Workbook(File.Contents(FolderPath & FileName), null),
    DataSheet = Source{[Item="Data",Kind="Sheet"]}[Data],
    Header = Table.PromoteHeaders(DataSheet, [PromoteAllScalars=true])
in
    Header
// ============================================================
// 09. QUERY — q_Pending_Cleaned
// ============================================================
let
    Source = q_Pending_Raw,
    CleanPO = Table.TransformColumns(
        Source,
        {
            {
                "PO No.",
                each Text.Trim(Text.Clean(Text.From(_))),
                type text
            }
        }
    ),
    FilterValid = Table.SelectRows(
        CleanPO,
        each fn_ValidatePO(Record.Field(_, "PO No."))
    )
in
    FilterValid
// ============================================================
// 10. QUERY — q_Department_Cleaned
// ============================================================
let
    FolderPath = fn_GetSetting("Source Folder"),
    FileName = fn_GetSetting("CCP File"),
    Source =
        Excel.Workbook(
            File.Contents(FolderPath & FileName),
            null,
            true
        ),
    DepartmentSheets = {
        "LA R&R",
        "HR",
        "CSR",
        "Mining",
        "FQA",
        "P&S",
        "UPL Cases",
        "IT",
        "Dispatch",
        "EED",
        "C&M",
        "Medical",
        "Civil"
    },
    SelectedSheets =
        Table.SelectRows(
            Source,
            each
                [Kind] = "Sheet"
                and List.Contains(
                    DepartmentSheets,
                    [Item]
                )
        ),
    ProcessSheet =
        (SheetName as text, SheetData as table) as table =>
        let
            Prepared =
                if SheetName = "UPL Cases" then
                    Table.Skip(SheetData, 1)
                else
                    SheetData,
            Promoted =
                Table.PromoteHeaders(
                    Prepared,
                    [PromoteAllScalars = true]
                ),
            HasDepartment =
                Table.HasColumns(
                    Promoted,
                    "Department"
                ),
            DepartmentPrepared =
                if HasDepartment then
                    Table.TransformColumns(
                        Promoted,
                        {
                            {
                                "Department",
                                each
                                    try
                                        Text.Trim(
                                            Text.Clean(
                                                Text.From(_)
                                            )
                                        )
                                    otherwise
                                        null,
                                type text
                            }
                        }
                    )
                else
                    Table.AddColumn(
                        Promoted,
                        "Department",
                        each
                            if SheetName = "UPL Cases" then
                                "UPL"
                            else
                                SheetName,
                        type text
                    ),
            Standardized =
                if SheetName = "UPL Cases" then
                    Table.RenameColumns(
                        DepartmentPrepared,
                        {
                            {
                                "Status as on 03.07.2025",
                                "Present CCP Status"
                            }
                        },
                        MissingField.Ignore
                    )
                else
                    DepartmentPrepared,
            KeepColumns =
                Table.SelectColumns(
                    Standardized,
                    {
                        "Department",
                        "PO No.",
                        "PO Desc.",
                        "PO Value",
                        "Creation Dt",
                        "Expiry Dt",
                        "PO Vendor",
                        "Vendor Description",
                        "EIC",
                        "Status of CCP/CCC",
                        "Status",
                        "Closed Date",
                        "Remarks",
                        "Final Dev. Status",
                        "Present CCP Status",
                        "PO Closed (Enter 1 if yes)"
                    },
                    MissingField.UseNull
                ),
            CreationRaw =
                Table.DuplicateColumn(
                    KeepColumns,
                    "Creation Dt",
                    "Creation Dt Raw"
                ),
            ExpiryRaw =
                Table.DuplicateColumn(
                    CreationRaw,
                    "Expiry Dt",
                    "Expiry Dt Raw"
                ),
            SafeCreationDate =
                Table.TransformColumns(
                    ExpiryRaw,
                    {
                        {
                            "Creation Dt",
                            each
                                let
                                    v = _,
                                    t =
                                        try
                                            Text.Trim(
                                                Text.From(v)
                                            )
                                        otherwise
                                            ""
                                in
                                    if Value.Is(v, type date) then
                                        v
                                    else if Value.Is(v, type datetime) then
                                        Date.From(v)
                                    else if Text.Contains(t, "-") then
                                        try
                                            Date.FromText(
                                                t,
                                                [Culture = "en-GB"]
                                            )
                                        otherwise
                                            null
                                    else if Text.Contains(t, "/") then
                                        try
                                            Date.FromText(
                                                t,
                                                [Culture = "en-US"]
                                            )
                                        otherwise
                                            null
                                    else
                                        null,
                            type date
                        }
                    }
                ),
            SafeExpiryDate =
                Table.TransformColumns(
                    SafeCreationDate,
                    {
                        {
                            "Expiry Dt",
                            each
                                let
                                    v = _,
                                    t =
                                        try
                                            Text.Trim(
                                                Text.From(v)
                                            )
                                        otherwise
                                            ""
                                in
                                    if Value.Is(v, type date) then
                                        v
                                    else if Value.Is(v, type datetime) then
                                        Date.From(v)
                                    else if Text.Contains(t, "-") then
                                        try
                                            Date.FromText(
                                                t,
                                                [Culture = "en-GB"]
                                            )
                                        otherwise
                                            null
                                    else if Text.Contains(t, "/") then
                                        try
                                            Date.FromText(
                                                t,
                                                [Culture = "en-US"]
                                            )
                                        otherwise
                                            null
                                    else
                                        null,
                            type date
                        }
                    }
                ),
            CleanPO =
                Table.TransformColumns(
                    SafeExpiryDate,
                    {
                        {
                            "PO No.",
                            each
                                try
                                    Text.Trim(
                                        Text.Clean(
                                            Text.From(_)
                                        )
                                    )
                                otherwise
                                    null,
                            type text
                        }
                    }
                ),
            FilterValid =
                Table.SelectRows(
                    CleanPO,
                    each
                        fn_ValidatePO(
                            [#"PO No."]
                        )
                ),
            RemoveBlankDepartments =
                Table.TransformColumns(
                    FilterValid,
                    {
                        {
                            "Department",
                            each
                                if _ = null then
                                    null
                                else if
                                    Text.Trim(
                                        Text.From(_)
                                    ) = "" then
                                    null
                                else
                                    Text.Trim(
                                        Text.From(_)
                                    ),
                            type text
                        }
                    }
                )
        in
            RemoveBlankDepartments,
    ProcessedTables =
        List.Transform(
            Table.ToRecords(
                SelectedSheets
            ),
            each
                ProcessSheet(
                    [Item],
                    [Data]
                )
        ),
    EmptyDepartmentTable =
        #table(
            {
                "Department",
                "PO No.",
                "PO Desc.",
                "PO Value",
                "Creation Dt",
                "Expiry Dt",
                "PO Vendor",
                "Vendor Description",
                "EIC",
                "Status of CCP/CCC",
                "Status",
                "Closed Date",
                "Remarks",
                "Final Dev. Status",
                "Present CCP Status",
                "PO Closed (Enter 1 if yes)",
                "Creation Dt Raw",
                "Expiry Dt Raw"
            },
            {}
        ),
    Combined =
        if List.Count(ProcessedTables) = 0 then
            EmptyDepartmentTable
        else
            Table.Combine(
                ProcessedTables
            ),
    ReplaceDateErrors =
        Table.ReplaceErrorValues(
            Combined,
            {
                {
                    "Creation Dt",
                    null
                },
                {
                    "Expiry Dt",
                    null
                }
            }
        )
in
    ReplaceDateErrors
// ============================================================
// 11. QUERY — q_Manual_Additions_Raw
// ============================================================
let
    Source = Excel.CurrentWorkbook(){[Name="Table_ManualAdditions"]}[Content]
in
    Source
// ============================================================
// 12. QUERY — q_Manual_Additions
// ============================================================
let
    Source = q_Manual_Additions_Raw,
    CleanPO = Table.TransformColumns(
        Source,
        {{"PO No.", each Text.Trim(Text.Clean(Text.From(_))), type text}}
    ),
    FilterValid = Table.SelectRows(
        CleanPO,
        each [Validation Status] = "Valid New PO"
    )
in
    FilterValid
// ============================================================
// 13. QUERY — q_Manual_Overrides_Raw
// ============================================================
let
    Source = Excel.CurrentWorkbook(){[Name="Table_ManualOverrides"]}[Content]
in
    Source
// ============================================================
// 14. QUERY — q_Manual_Overrides
// ============================================================
let
    Source = q_Manual_Overrides_Raw,
    CleanPO = Table.TransformColumns(Source, {{"PO No.", each Text.Trim(Text.Clean(Text.From(_))), type text}}),
    FilterActive = Table.SelectRows(CleanPO, each [Active] = "Yes")
in
    FilterActive
// ============================================================
// 15. QUERY — q_Manual_Exclusions_Raw
// ============================================================
let
    Source = Excel.CurrentWorkbook(){[Name="Table_ManualExclusions"]}[Content]
in
    Source
// ============================================================
// 16. QUERY — q_Manual_Exclusions
// ============================================================
let
    Source = q_Manual_Exclusions_Raw,
    CleanPO = Table.TransformColumns(Source, {{"PO No.", each Text.Trim(Text.Clean(Text.From(_))), type text}}),
    FilterActive = Table.SelectRows(CleanPO, each [Active] = "Yes")
in
    FilterActive
// ============================================================
// 17. QUERY — q_Unified_Master
// ============================================================
let
    CCP = q_CCP_Cleaned,
    Pending = q_Pending_Cleaned,
    ManualAdd = q_Manual_Additions,
    // =========================================================
    // 1. STANDARDIZE MANUAL ADDITIONS
    // =========================================================
    ManualAdd_Standardized =
        Table.RenameColumns(
            ManualAdd,
            {
                {"Pending Ordering Plant", "Manual Pending Ordering Plant"},
                {"Pending Delivery Plant", "Manual Pending Delivery Plant"},
                {"Pending Status", "Manual Pending Status"},
                {"Pending Remarks", "Manual Pending Remarks"}
            },
            MissingField.Ignore
        ),
    // =========================================================
    // 1A. STANDARD MASTER COLUMNS FOR MANUAL ADDITIONS
    // Missing fields are deliberately created as null.
    // =========================================================
    ManualAdd_Final =
        Table.SelectColumns(
            ManualAdd_Standardized,
            {
                "PO No.",
                "Department",
                "Final Dev. Status",
                "Present CCP Status",
                "PO Vendor",
                "PO Desc.",
                "Creation Dt",
                "Expiry Dt",
                "PO Value",
                "Vendor Description",
                "EIC",
                "Status",
                "Status of CCP/CCC",
                "Closed Date",
                "Remarks",
                "Legacy PO Number",
                "PO_Status",
                "Data Source(s)",
                "Added Date",
                "Validation Status",
                "Manual Pending Ordering Plant",
                "Manual Pending Delivery Plant",
                "Manual Pending Status",
                "Manual Pending Remarks"
            },
            MissingField.UseNull
        ),
    // =========================================================
    // 2. COMBINE CCP + MANUAL ADDITIONS
    // =========================================================
    SourceBase =
        Table.Combine(
            {
                CCP,
                ManualAdd_Final
            }
        ),
    // =========================================================
    // 3. MERGE CCP + PENDING
    // =========================================================
    MergeCCP_Pending =
        Table.NestedJoin(
            SourceBase,
            {"PO No."},
            Pending,
            {"PO No."},
            "PendingData",
            JoinKind.FullOuter
        ),
    // =========================================================
    // 4. EXPAND PENDING FIELDS
    // =========================================================
    ExpandedPending =
        Table.ExpandTableColumn(
            MergeCCP_Pending,
            "PendingData",
            {
                "PO No.",
                "Ordering Plant",
                "Delivery Plant",
                "PO Desc.",
                "Creation Dt",
                "Expiry Dt",
                "PO Value",
                "PO Vendor",
                "Vendor Description",
                "EIC",
                "Status",
                "Status of CCP/CCC",
                "Closed Date",
                "Remarks",
                "Legacy PO Number",
                "Remarks.1"
            },
            {
                "Pending PO No.",
                "Pending Ordering Plant",
                "Pending Delivery Plant",
                "Pending PO Desc.",
                "Pending Creation Dt",
                "Pending Expiry Dt",
                "Pending PO Value",
                "Pending PO Vendor",
                "Pending Vendor Description",
                "Pending EIC",
                "Pending Status",
                "Pending Status of CCP/CCC",
                "Pending Closed Date",
                "Pending Remarks",
                "Pending Legacy PO Number",
                "Pending Remarks.1"
            }
        ),
    // =========================================================
    // 5. MERGE MANUAL PENDING VALUES WITH PENDING VALUES
    // PENDING SOURCE HAS PRIORITY
    // MANUAL ADDITION VALUE USED IF PENDING IS BLANK
    // =========================================================
    FillPendingOrderingPlant =
        Table.AddColumn(
            ExpandedPending,
            "Pending Ordering Plant_Final",
            each
                if [#"Pending Ordering Plant"] <> null
                    and Text.Trim(Text.From([#"Pending Ordering Plant"])) <> ""
                then
                    [#"Pending Ordering Plant"]
                else
                    [#"Manual Pending Ordering Plant"],
            type any
        ),
    RemoveOldPendingOrderingPlant =
        Table.RemoveColumns(
            FillPendingOrderingPlant,
            {"Pending Ordering Plant"}
        ),
    RenamePendingOrderingPlant =
        Table.RenameColumns(
            RemoveOldPendingOrderingPlant,
            {
                {"Pending Ordering Plant_Final", "Pending Ordering Plant"}
            }
        ),
    FillPendingDeliveryPlant =
        Table.AddColumn(
            RenamePendingOrderingPlant,
            "Pending Delivery Plant_Final",
            each
                if [#"Pending Delivery Plant"] <> null
                    and Text.Trim(Text.From([#"Pending Delivery Plant"])) <> ""
                then
                    [#"Pending Delivery Plant"]
                else
                    [#"Manual Pending Delivery Plant"],
            type any
        ),
    RemoveOldPendingDeliveryPlant =
        Table.RemoveColumns(
            FillPendingDeliveryPlant,
            {"Pending Delivery Plant"}
        ),
    RenamePendingDeliveryPlant =
        Table.RenameColumns(
            RemoveOldPendingDeliveryPlant,
            {
                {"Pending Delivery Plant_Final", "Pending Delivery Plant"}
            }
        ),
    FillPendingStatus =
        Table.AddColumn(
            RenamePendingDeliveryPlant,
            "Pending Status_Final",
            each
                if [#"Pending Status"] <> null
                    and Text.Trim(Text.From([#"Pending Status"])) <> ""
                then
                    [#"Pending Status"]
                else
                    [#"Manual Pending Status"],
            type text
        ),
    RemoveOldPendingStatus =
        Table.RemoveColumns(
            FillPendingStatus,
            {"Pending Status"}
        ),
    RenamePendingStatus =
        Table.RenameColumns(
            RemoveOldPendingStatus,
            {
                {"Pending Status_Final", "Pending Status"}
            }
        ),
    FillPendingRemarks =
        Table.AddColumn(
            RenamePendingStatus,
            "Pending Remarks_Final",
            each
                if [#"Pending Remarks"] <> null
                    and Text.Trim(Text.From([#"Pending Remarks"])) <> ""
                then
                    [#"Pending Remarks"]
                else
                    [#"Manual Pending Remarks"],
            type text
        ),
    RemoveOldPendingRemarks =
        Table.RemoveColumns(
            FillPendingRemarks,
            {"Pending Remarks"}
        ),
    RenamePendingRemarks =
        Table.RenameColumns(
            RemoveOldPendingRemarks,
            {
                {"Pending Remarks_Final", "Pending Remarks"}
            }
        ),
    // =========================================================
    // 6. CREATE FINAL PO NUMBER
    // =========================================================
    AddFinalPO =
        Table.AddColumn(
            RenamePendingRemarks,
            "Final PO No.",
            each
                if [#"PO No."] <> null
                then
                    [#"PO No."]
                else
                    [#"Pending PO No."],
            type text
        ),
    RemoveOldPO =
        Table.RemoveColumns(
            AddFinalPO,
            {
                "PO No.",
                "Pending PO No."
            }
        ),
    RenameFinalPO =
        Table.RenameColumns(
            RemoveOldPO,
            {
                {"Final PO No.", "PO No."}
            }
        ),
    // =========================================================
    // 7. FILL MISSING PO DESCRIPTION
    // =========================================================
    FillPODesc =
        Table.AddColumn(
            RenameFinalPO,
            "PO Desc_Final",
            each
                if [#"PO Desc."] <> null
                    and Text.Trim(Text.From([#"PO Desc."])) <> ""
                then
                    [#"PO Desc."]
                else
                    [#"Pending PO Desc."],
            type text
        ),
    RemoveOldPODesc =
        Table.RemoveColumns(
            FillPODesc,
            {"PO Desc."}
        ),
    RenamePODesc =
        Table.RenameColumns(
            RemoveOldPODesc,
            {
                {"PO Desc_Final", "PO Desc."}
            }
        ),
    // =========================================================
    // 8. FILL MISSING CREATION DATE
    // =========================================================
    FillCreationDt =
        Table.AddColumn(
            RenamePODesc,
            "Creation Dt_Final",
            each
                if [#"Creation Dt"] <> null
                then
                    [#"Creation Dt"]
                else
                    [#"Pending Creation Dt"],
            type any
        ),
    RemoveOldCreationDt =
        Table.RemoveColumns(
            FillCreationDt,
            {"Creation Dt"}
        ),
    RenameCreationDt =
        Table.RenameColumns(
            RemoveOldCreationDt,
            {
                {"Creation Dt_Final", "Creation Dt"}
            }
        ),
    // =========================================================
    // 9. FILL MISSING EXPIRY DATE
    // =========================================================
    FillExpiryDt =
        Table.AddColumn(
            RenameCreationDt,
            "Expiry Dt_Final",
            each
                if [#"Expiry Dt"] <> null
                then
                    [#"Expiry Dt"]
                else
                    [#"Pending Expiry Dt"],
            type any
        ),
    RemoveOldExpiryDt =
        Table.RemoveColumns(
            FillExpiryDt,
            {"Expiry Dt"}
        ),
    RenameExpiryDt =
        Table.RenameColumns(
            RemoveOldExpiryDt,
            {
                {"Expiry Dt_Final", "Expiry Dt"}
            }
        ),
    // =========================================================
    // 10. FILL MISSING PO VALUE
    // =========================================================
    FillPOValue =
        Table.AddColumn(
            RenameExpiryDt,
            "PO Value_Final",
            each
                if [#"PO Value"] <> null
                then
                    [#"PO Value"]
                else
                    [#"Pending PO Value"],
            type any
        ),
    RemoveOldPOValue =
        Table.RemoveColumns(
            FillPOValue,
            {"PO Value"}
        ),
    RenamePOValue =
        Table.RenameColumns(
            RemoveOldPOValue,
            {
                {"PO Value_Final", "PO Value"}
            }
        ),
    // =========================================================
    // 11. FILL MISSING PO VENDOR
    // =========================================================
    FillPOVendor =
        Table.AddColumn(
            RenamePOValue,
            "PO Vendor_Final",
            each
                if [#"PO Vendor"] <> null
                    and Text.Trim(Text.From([#"PO Vendor"])) <> ""
                then
                    [#"PO Vendor"]
                else
                    [#"Pending PO Vendor"],
            type text
        ),
    RemoveOldPOVendor =
        Table.RemoveColumns(
            FillPOVendor,
            {"PO Vendor"}
        ),
    RenamePOVendor =
        Table.RenameColumns(
            RemoveOldPOVendor,
            {
                {"PO Vendor_Final", "PO Vendor"}
            }
        ),
    // =========================================================
    // 12. FILL MISSING VENDOR DESCRIPTION
    // =========================================================
    FillVendorDescription =
        Table.AddColumn(
            RenamePOVendor,
            "Vendor Description_Final",
            each
                if [#"Vendor Description"] <> null
                    and Text.Trim(Text.From([#"Vendor Description"])) <> ""
                then
                    [#"Vendor Description"]
                else
                    [#"Pending Vendor Description"],
            type text
        ),
    RemoveOldVendorDescription =
        Table.RemoveColumns(
            FillVendorDescription,
            {"Vendor Description"}
        ),
    RenameVendorDescription =
        Table.RenameColumns(
            RemoveOldVendorDescription,
            {
                {"Vendor Description_Final", "Vendor Description"}
            }
        ),
    // =========================================================
    // 13. FILL MISSING EIC
    // =========================================================
    FillEIC =
        Table.AddColumn(
            RenameVendorDescription,
            "EIC_Final",
            each
                if [EIC] <> null
                    and Text.Trim(Text.From([EIC])) <> ""
                then
                    [EIC]
                else
                    [#"Pending EIC"],
            type text
        ),
    RemoveOldEIC =
        Table.RemoveColumns(
            FillEIC,
            {"EIC"}
        ),
    RenameEIC =
        Table.RenameColumns(
            RemoveOldEIC,
            {
                {"EIC_Final", "EIC"}
            }
        ),
    // =========================================================
    // 14. FILL MISSING MASTER STATUS
    // =========================================================
    FillMasterStatus =
        Table.AddColumn(
            RenameEIC,
            "Status_Final",
            each
                if [Status] <> null
                    and Text.Trim(Text.From([Status])) <> ""
                then
                    [Status]
                else
                    [#"Pending Status"],
            type text
        ),
    RemoveOldMasterStatus =
        Table.RemoveColumns(
            FillMasterStatus,
            {"Status"}
        ),
    RenameMasterStatus =
        Table.RenameColumns(
            RemoveOldMasterStatus,
            {
                {"Status_Final", "Status"}
            }
        ),
    // =========================================================
    // 15. FILL MISSING CCP/CCC STATUS
    // =========================================================
    FillCCPStatus =
        Table.AddColumn(
            RenameMasterStatus,
            "Status of CCP/CCC_Final",
            each
                if [#"Status of CCP/CCC"] <> null
                    and Text.Trim(Text.From([#"Status of CCP/CCC"])) <> ""
                then
                    [#"Status of CCP/CCC"]
                else
                    [#"Pending Status of CCP/CCC"],
            type text
        ),
    RemoveOldCCPStatus =
        Table.RemoveColumns(
            FillCCPStatus,
            {"Status of CCP/CCC"}
        ),
    RenameCCPStatus =
        Table.RenameColumns(
            RemoveOldCCPStatus,
            {
                {"Status of CCP/CCC_Final", "Status of CCP/CCC"}
            }
        ),
    // =========================================================
    // 16. FILL MISSING CLOSED DATE
    // =========================================================
    FillClosedDate =
        Table.AddColumn(
            RenameCCPStatus,
            "Closed Date_Final",
            each
                if [#"Closed Date"] <> null
                then
                    [#"Closed Date"]
                else
                    [#"Pending Closed Date"],
            type any
        ),
    RemoveOldClosedDate =
        Table.RemoveColumns(
            FillClosedDate,
            {"Closed Date"}
        ),
    RenameClosedDate =
        Table.RenameColumns(
            RemoveOldClosedDate,
            {
                {"Closed Date_Final", "Closed Date"}
            }
        ),
    // =========================================================
    // 17. FILL MISSING REMARKS
    // =========================================================
    FillRemarks =
        Table.AddColumn(
            RenameClosedDate,
            "Remarks_Final",
            each
                if [Remarks] <> null
                    and Text.Trim(Text.From([Remarks])) <> ""
                then
                    [Remarks]
                else
                    [#"Pending Remarks"],
            type text
        ),
    RemoveOldRemarks =
        Table.RemoveColumns(
            FillRemarks,
            {"Remarks"}
        ),
    RenameRemarks =
        Table.RenameColumns(
            RemoveOldRemarks,
            {
                {"Remarks_Final", "Remarks"}
            }
        ),
    // =========================================================
    // 18. FILL MISSING LEGACY PO NUMBER
    // =========================================================
    FillLegacyPO =
        Table.AddColumn(
            RenameRemarks,
            "Legacy PO Number_Final",
            each
                if [#"Legacy PO Number"] <> null
                    and Text.Trim(Text.From([#"Legacy PO Number"])) <> ""
                then
                    [#"Legacy PO Number"]
                else
                    [#"Pending Legacy PO Number"],
            type text
        ),
    RemoveOldLegacyPO =
        Table.RemoveColumns(
            FillLegacyPO,
            {"Legacy PO Number"}
        ),
    RenameLegacyPO =
        Table.RenameColumns(
            RemoveOldLegacyPO,
            {
                {"Legacy PO Number_Final", "Legacy PO Number"}
            }
        ),
    // =========================================================
    // 19. FINAL STATUS
    // CCP STATUS HAS PRIORITY
    // PENDING STATUS USED ONLY WHEN CCP STATUS IS MISSING
    // =========================================================
    AddFinalStatus =
        Table.AddColumn(
            RenameLegacyPO,
            "PO_Status_Final",
            each
                if [#"PO_Status"] <> null
                    and Text.Trim(Text.From([#"PO_Status"])) <> ""
                then
                    [#"PO_Status"]
                else
                    [#"Pending Status"],
            type text
        ),
    RemoveOldStatus =
        Table.RemoveColumns(
            AddFinalStatus,
            {"PO_Status"}
        ),
    RenameFinalStatus =
        Table.RenameColumns(
            RemoveOldStatus,
            {
                {"PO_Status_Final", "PO_Status"}
            }
        ),
    // =========================================================
    // 20. DEDUPLICATE MASTER BY PO NUMBER
    // =========================================================
    Deduplicated =
        let
            Columns =
                Table.ColumnNames(RenameFinalStatus),
            OtherColumns =
                List.RemoveItems(
                    Columns,
                    {"PO No.", "PO_Status"}
                ),
            Aggregations =
                List.Combine(
                    {
                        {
                            {
                                "PO_Status",
                                each
                                    let
                                        Values =
                                            List.RemoveNulls(
                                                List.Transform(
                                                    Table.Column(_, "PO_Status"),
                                                    each
                                                        try
                                                            Text.Trim(Text.From(_))
                                                        otherwise
                                                            null
                                                )
                                            )
                                    in
                                        if List.Count(Values) > 0
                                        then
                                            Values{0}
                                        else
                                            null,
                                type text
                            }
                        },
                        {
                            {
                                "StatusVariants",
                                each
                                    List.Distinct(
                                        List.RemoveNulls(
                                            List.Transform(
                                                Table.Column(_, "PO_Status"),
                                                each
                                                    try
                                                        Text.Trim(Text.From(_))
                                                    otherwise
                                                        null
                                            )
                                        )
                                    ),
                                type list
                            }
                        },
                        List.Transform(
                            OtherColumns,
                            (c) =>
                                {
                                    c,
                                    each
                                        let
                                            Values =
                                                List.RemoveNulls(
                                                    Table.Column(_, c)
                                                )
                                        in
                                            if List.Count(Values) > 0
                                            then
                                                Values{0}
                                            else
                                                null,
                                    type any
                                }
                        )
                    }
                ),
            Grouped =
                Table.Group(
                    RenameFinalStatus,
                    {"PO No."},
                    Aggregations
                ),
            AddDataConflict =
                Table.AddColumn(
                    Grouped,
                    "Data Conflict",
                    each
                        if List.Count([StatusVariants]) > 1
                        then
                            "Yes – Status Conflict"
                        else
                            null,
                    type text
                ),
            RemoveStatusVariants =
                Table.RemoveColumns(
                    AddDataConflict,
                    {"StatusVariants"}
                )
        in
            RemoveStatusVariants,
    // =========================================================
    // 21. APPLY MANUAL EXCLUSIONS
    // =========================================================
    Exclusions =
        Table.SelectRows(
            q_Manual_Exclusions,
            each
                Text.Upper(
                    Text.Trim(
                        Text.From([Active])
                    )
                ) = "YES"
        ),
    ExcludedPOs =
        Table.Column(
            Exclusions,
            "PO No."
        ),
    FilteredExclusions =
        Table.SelectRows(
            Deduplicated,
            each
                not List.Contains(
                    ExcludedPOs,
                    Record.Field(_, "PO No.")
                )
        ),
    // =========================================================
    // 22. APPLY MANUAL OVERRIDES
    //
    // Long-form override table:
    // one row = one PO + one field to override.
    //
    // Multiple active overrides for the same PO are converted
    // into one record, preventing duplicate Master rows.
    // =========================================================
    ActiveOverrides =
        Table.SelectRows(
            q_Manual_Overrides,
            each
                Text.Upper(
                    Text.Trim(
                        Text.From([Active])
                    )
                ) = "YES"
        ),
    // Map the user-facing "PO Status" label to the actual
    // Master column name "PO_Status".
    NormalizeOverrideField =
        Table.AddColumn(
            ActiveOverrides,
            "Normalized Field",
            each
                let
                    FieldName =
                        Text.Trim(
                            Text.From([#"Field to Override"])
                        )
                in
                    if FieldName = "PO Status"
                    then
                        "PO_Status"
                    else
                        FieldName,
            type text
        ),
    // Fields that are permitted to be manually overridden.
    AllowedOverrideFields =
        {
            "Department",
            "PO Desc.",
            "PO Value",
            "Creation Dt",
            "Expiry Dt",
            "PO Vendor",
            "Vendor Description",
            "EIC",
            "PO_Status",
            "Status",
            "Present CCP Status",
            "Status of CCP/CCC",
            "Final Dev. Status",
            "Closed Date",
            "Remarks",
            "Pending Ordering Plant",
            "Pending Delivery Plant",
            "Pending Status",
            "Pending Remarks"
        },
    ValidOverrides =
        Table.SelectRows(
            NormalizeOverrideField,
            each
                List.Contains(
                    AllowedOverrideFields,
                    [#"Normalized Field"]
                )
        ),
    // Sort so the most recently updated active override
    // appears first for each PO + field combination.
    SortedOverrides =
        Table.Sort(
            ValidOverrides,
            {
                {"PO No.", Order.Ascending},
                {"Normalized Field", Order.Ascending},
                {"Update Date", Order.Descending}
            }
        ),
    // Group by PO + field and retain the first record.
    // This prevents duplicate Master rows.
    LatestOverrides =
        Table.Group(
            SortedOverrides,
            {
                "PO No.",
                "Normalized Field"
            },
            {
                {
                    "Override Value",
                    each
                        let
                            Values =
                                List.RemoveNulls(
                                    [#"Manual Value"]
                                )
                        in
                            if List.Count(Values) > 0
                            then
                                Values{0}
                            else
                                null,
                    type any
                }
            }
        ),
    // Convert the long-form overrides for each PO into
    // a single record.
    OverrideRecords =
        Table.Group(
            LatestOverrides,
            {"PO No."},
            {
                {
                    "OverrideRecord",
                    each
                        let
                            Names =
                                [#"Normalized Field"],
                            Values =
                                [#"Override Value"],
                            RecordPairs =
                                List.Zip(
                                    {
                                        Names,
                                        Values
                                    }
                                )
                        in
                            Record.FromList(
                                Values,
                                Names
                            ),
                    type record
                }
            }
        ),
    // Merge one override record onto each Master PO.
    OverrideMerge =
        Table.NestedJoin(
            FilteredExclusions,
            {"PO No."},
            OverrideRecords,
            {"PO No."},
            "OverrideData",
            JoinKind.LeftOuter
        ),
    ExpandOverrideRecord =
        Table.ExpandTableColumn(
            OverrideMerge,
            "OverrideData",
            {"OverrideRecord"},
            {"OverrideRecord"}
        ),
    // Apply each permitted override independently.
    // Record.FieldOrDefault safely returns null when a PO
    // has no override for that particular field.
    ApplyOverrides =
        List.Accumulate(
            AllowedOverrideFields,
            ExpandOverrideRecord,
            (CurrentTable, FieldName) =>
                Table.AddColumn(
                    CurrentTable,
                    FieldName & "_Final",
                    each
                        let
                            OverrideRecord =
                                try
                                    [OverrideRecord]
                                otherwise
                                    null,
                            OverrideValue =
                                if OverrideRecord = null
                                then
                                    null
                                else
                                    Record.FieldOrDefault(
                                        OverrideRecord,
                                        FieldName,
                                        null
                                    ),
                            CurrentValue =
                                try
                                    Record.Field(
                                        _,
                                        FieldName
                                    )
                                otherwise
                                    null
                        in
                            if OverrideValue <> null
                                and Text.Trim(
                                    Text.From(OverrideValue)
                                ) <> ""
                            then
                                OverrideValue
                            else
                                CurrentValue,
                    type any
                )
        ),
    // Remove the original versions of the overrideable fields.
    RemoveOriginalOverrideFields =
        Table.RemoveColumns(
            ApplyOverrides,
            AllowedOverrideFields,
            MissingField.Ignore
        ),
    // Remove the temporary OverrideRecord.
    RemoveOverrideRecord =
        Table.RemoveColumns(
            RemoveOriginalOverrideFields,
            {"OverrideRecord"},
            MissingField.Ignore
        ),
    // Rename all generated *_Final columns back to their
    // original Master field names.
    RenameFinalOverrideFields =
        Table.RenameColumns(
            RemoveOverrideRecord,
            List.Transform(
                AllowedOverrideFields,
                each
                    {
                        _ & "_Final",
                        _
                    }
            ),
            MissingField.Ignore
        ),
    // =========================================================
    // 23. REMOVE TEMPORARY HELPER COLUMNS
    // =========================================================
    RemovePendingHelperColumns =
        Table.RemoveColumns(
            RenameFinalOverrideFields,
            {
                "PO Closed (Enter 1 if yes)",
                "Pending PO Desc.",
                "Pending Creation Dt",
                "Pending Expiry Dt",
                "Pending PO Value",
                "Pending PO Vendor",
                "Pending Vendor Description",
                "Pending EIC",
                "Pending Status of CCP/CCC",
                "Pending Closed Date",
                "Pending Legacy PO Number",
                "Pending Remarks.1",
                "Manual Pending Ordering Plant",
                "Manual Pending Delivery Plant",
                "Manual Pending Status",
                "Manual Pending Remarks"
            },
            MissingField.Ignore
        ),
    // =========================================================
    // 24. FINAL MASTER COLUMN ORDER
    // =========================================================
    FinalColumns =
        {
            "PO No.",
            "Legacy PO Number",
            "Department",
            "PO Desc.",
            "PO Value",
            "Creation Dt",
            "Expiry Dt",
            "PO Vendor",
            "Vendor Description",
            "EIC",
            "PO_Status",
            "Status",
            "Present CCP Status",
            "Status of CCP/CCC",
            "Final Dev. Status",
            "Closed Date",
            "Remarks",
            "Data Source(s)",
            "Data Conflict",
            "Validation Status",
            "Added Date",
            "Pending Ordering Plant",
            "Pending Delivery Plant",
            "Pending Status",
            "Pending Remarks"
        },
    FinalMaster =
        Table.SelectColumns(
            RemovePendingHelperColumns,
            FinalColumns,
            MissingField.UseNull
        )
in
    FinalMaster
// ============================================================
// 18. QUERY — q_QA_Checks
// ============================================================
let
    // =========================================================
    // SOURCE QUERIES
    // =========================================================
    Master = q_Unified_Master,
    CCP = q_CCP_Cleaned,
    Pending = q_Pending_Cleaned,
    Legacy = q_Legacy_Cleaned,
    Department = q_Department_Cleaned,
    // =========================================================
    // UNIQUE PO LISTS
    // =========================================================
    MasterPOs =
        List.Distinct(
            List.RemoveNulls(
                Table.Column(
                    Master,
                    "PO No."
                )
            )
        ),
    CCPPOs =
        List.Distinct(
            List.RemoveNulls(
                Table.Column(
                    CCP,
                    "PO No."
                )
            )
        ),
    PendingPOs =
        List.Distinct(
            List.RemoveNulls(
                Table.Column(
                    Pending,
                    "PO No."
                )
            )
        ),
    LegacyPOs =
        List.Distinct(
            List.RemoveNulls(
                Table.Column(
                    Legacy,
                    "Legacy PO No."
                )
            )
        ),
    DepartmentPOs =
        List.Distinct(
            List.RemoveNulls(
                Table.Column(
                    Department,
                    "PO No."
                )
            )
        ),
    // =========================================================
    // COUNTS
    // =========================================================
    ERPUnique =
        List.Count(
            LegacyPOs
        ),
    CCPUnique =
        List.Count(
            CCPPOs
        ),
    PendingUnique =
        List.Count(
            PendingPOs
        ),
    DepartmentUnique =
        List.Count(
            DepartmentPOs
        ),
    MasterUnique =
        List.Count(
            MasterPOs
        ),
    MasterRows =
        Table.RowCount(
            Master
        ),
    DuplicateMaster =
        MasterRows - MasterUnique,
    // =========================================================
    // SOURCE → MASTER COVERAGE
    // =========================================================
    CCPMissing =
        List.Count(
            List.Difference(
                CCPPOs,
                MasterPOs
            )
        ),
    PendingMissing =
        List.Count(
            List.Difference(
                PendingPOs,
                MasterPOs
            )
        ),
    // =========================================================
    // STATUS CONFLICTS
    // =========================================================
    StatusConflictRows =
        Table.SelectRows(
            Master,
            each
                [Data Conflict] = "Yes – Status Conflict"
        ),
    StatusConflicts =
        Table.RowCount(
            StatusConflictRows
        ),
    // =========================================================
    // VALID PO CHECK
    // =========================================================
    NonPORecords =
        List.Count(
            List.Select(
                MasterPOs,
                each
                    not fn_ValidatePO(_)
            )
        ),
    // =========================================================
    // QA TABLE
    // =========================================================
    QA =
        #table(
            {
                "QA Check Description",
                "Actual Result",
                "QA Status",
                "Notes"
            },
            {
                {
                    "1. ERP Unique Legacy POs",
                    ERPUnique,
                    "PASS",
                    "Distinct valid Legacy PO numbers"
                },
                {
                    "2. CCP Unique POs",
                    CCPUnique,
                    "PASS",
                    "Distinct valid CCP PO numbers"
                },
                {
                    "3. Pending Unique POs",
                    PendingUnique,
                    "PASS",
                    "Distinct valid Pending PO numbers"
                },
                {
                    "4. Department Unique POs",
                    DepartmentUnique,
                    "PASS",
                    "Distinct valid Department PO numbers"
                },
                {
                    "5. Master Unique POs",
                    MasterUnique,
                    if MasterUnique > 0
                    then "PASS"
                    else "FAIL",
                    "Current unique PO population in Unified Master"
                },
                {
                    "6. Duplicate Master POs",
                    DuplicateMaster,
                    if DuplicateMaster = 0
                    then "PASS"
                    else "FAIL",
                    "Unified Master must contain one row per PO"
                },
                {
                    "7. ERP Relationship Review",
                    ERPUnique,
                    "REVIEW",
                    "Legacy-to-SAP/ERP relationships require explicit evidence; direct PO equality is not assumed"
                },
                {
                    "8. CCP Missing from Master",
                    CCPMissing,
                    if CCPMissing = 0
                    then "PASS"
                    else "FAIL",
                    "Current CCP PO not found in current Unified Master"
                },
                {
                    "9. Pending Missing from Master",
                    PendingMissing,
                    if PendingMissing = 0
                    then "PASS"
                    else "FAIL",
                    "Current Pending PO not found in current Unified Master"
                },
                {
                    "10. Status Conflicts",
                    StatusConflicts,
                    if StatusConflicts > 0
                    then "REVIEW"
                    else "PASS",
                    "Known source-level status conflicts are retained and flagged"
                },
                {
                    "11. Date Conflicts / Review",
                    null,
                    "REVIEW",
                    "Dates are preserved from source; detailed cross-source date comparison requires separate review"
                },
                {
                    "12. Alignment Reviews",
                    null,
                    "REVIEW",
                    "Department source layouts require structural validation"
                },
                {
                    "13. Traceability Issues",
                    null,
                    "REVIEW",
                    "Source-to-record traceability requires relationship-level validation"
                },
                {
                    "14. Non-PO Records in Master",
                    NonPORecords,
                    if NonPORecords = 0
                    then "PASS"
                    else "FAIL",
                    "Unified Master must contain only valid 10-digit PO numbers"
                }
            }
        )
in
    QA
// ============================================================
// END OF POWER QUERY PORTFOLIO SCRIPTS
// ============================================================