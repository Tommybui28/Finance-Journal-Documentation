    # DB Journal Power Query – Developer Handover (Block-by-Block Annotation)

    This document is a **developer handover guide** for the **`DB Journal`** Power Query.

    It explains the query **block by block**, with each section showing:

    - **Code block** – the relevant M code
    - **What it does** – plain-English explanation
    - **Why it matters** – technical / business significance
    - **Risks / gotchas** – things that may break, confuse, or create mismatches

    ---

    ## 1. Purpose of this query

    The **`DB Journal`** query takes source Daily Banking data from Excel, enriches it using lookup queries, and outputs a **Business Central import-ready journal**.

    The query does all of the following in one flow:

    1. Loads the workbook source table
    2. Sets data types
    3. Cleans text data
    4. Builds posting/accounting fields
    5. Derives document numbers and dimensions
    6. Joins lookup queries
    7. Creates balancing bank lines
    8. Shapes the final output layout for import

    ---

    ## 2. High-level query flow

    ```text
    Table1 source
    → Type conversion
    → Text cleanup
    → Base accounting fields
    → Document number logic
    → Dimension extraction
    → Fund / Department / VAT lookup joins
    → Journal base rows
    → Balancing bank rows
    → Combine original + balancing rows
    → Final importer layout
    ```

    ---

    ## 3. Block-by-block annotation

    ## Block 1 – Load source table

    ### Code

    ```m
    Source = Excel.CurrentWorkbook(){[Name="Table1"]}[Content]
    ```

    ### What it does

    Loads data from the Excel table named **`Table1`** in the current workbook.

    ### Why it matters

    This is the root data source for the whole query. Every downstream step depends on the structure and contents of `Table1`.

    ### Risks / gotchas

    - If the Excel table is renamed, refresh will fail.
    - If columns are added/removed in the table, later hardcoded column references may break.
    - Because this is workbook-based, upstream refresh issues may appear here even if the logic below is correct.

    ---

    ## Block 2 – Set data types

    ### Code

    ```m
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Gf_Type", type text}, {"Gf_Date", type date}, {"Gf_Gift_subtype", type text}, {"Gf_Amount", type number}, {"Gf_Pay_method", type text}, {"Gf_Batch_Number", Int64.Type}, {"Gf_Reference", Int64.Type}, {"Gf_DateAdded", type date}, {"Gf_Reference_number", type text}, {"Gf_Gift_code", type text}, {"Gf_NL_post_date", type date}, {"Gf_Import_ID", type text}, {"Gf_Gift_ID", type text}, {"Gf_NL_post_status", type text}, {"Gf_Markting_Sorc_Cod", type text}, {"Gf_Agency", type text}, {"Gf_AttrCat_1_01_Description", type text}, {"Gf_AttrCat_2_01_Description", type text}, {"Gf_AttrCat_3_01_Description", type text}, {"Gf_AttrCat_4_01_Description", Int64.Type}, {"Gf_CnBio_ID", Int64.Type}, {"Gf_CnBio_Name", type text}, {"Gf_Fnds_1_01_Description", type text}, {"Gf_Fnds_1_01_Fn_Restricted", type text}, {"Gf_Fnds_1_01_Fn_Fund_category", type text}, {"Gf_Fnds_1_01_Fn_Fund_type", type text}, {"Gf_Apls_1_01_Appeal_ID", type text}, {"Gf_Apls_1_01_Description", type text}, {"Gf_Apls_1_01_Ap_Appeal_category", type text}, {"Gf_Apls_1_01_ApAtrCat_1_01_Description", Int64.Type}, {"Gf_Apls_1_01_ApAtrCat_2_01_Description", type text}})
    ```

    ### What it does

    Explicitly assigns data types to the source columns.

    ### Why it matters

    - Makes joins and transformations predictable
    - Prevents accidental text/number/date issues later
    - Ensures Business Central fields are generated from correctly typed inputs

    ### Risks/gotchas

    - `Int64.Type` on fields like `Gf_Reference`, `Gf_AttrCat_4_01_Description`, `Gf_CnBio_ID`, and `Gf_Apls_1_01_ApAtrCat_1_01_Description` assumes those values are always numeric.
    - If any of these fields contain blanks, mixed formatting, or non-numeric values, refresh may fail.
    - If IDs should preserve leading zeros, typing them as numbers may lose format fidelity.

    ---

    ## Block 3 – Trim text fields

    ### Code

    ```m
    #"Trimmed Text" = Table.TransformColumns(#"Changed Type",{{"Gf_Type", Text.Trim, type text}, {"Gf_Gift_subtype", Text.Trim, type text}, {"Gf_Pay_method", Text.Trim, type text}, {"Gf_Reference_number", Text.Trim, type text}, {"Gf_Gift_code", Text.Trim, type text}, {"Gf_Import_ID", Text.Trim, type text}, {"Gf_Gift_ID", Text.Trim, type text}, {"Gf_NL_post_status", Text.Trim, type text}, {"Gf_Markting_Sorc_Cod", Text.Trim, type text}, {"Gf_Agency", Text.Trim, type text}, {"Gf_AttrCat_1_01_Description", Text.Trim, type text}, {"Gf_AttrCat_2_01_Description", Text.Trim, type text}, {"Gf_AttrCat_3_01_Description", Text.Trim, type text}, {"Gf_CnBio_Name", Text.Trim, type text}, {"Gf_Fnds_1_01_Description", Text.Trim, type text}, {"Gf_Fnds_1_01_Fn_Restricted", Text.Trim, type text}, {"Gf_Fnds_1_01_Fn_Fund_category", Text.Trim, type text}, {"Gf_Fnds_1_01_Fn_Fund_type", Text.Trim, type text}, {"Gf_Apls_1_01_Ap_Appeal_category", Text.Trim, type text}, {"Gf_Apls_1_01_Description", Text.Trim, type text}, {"Gf_Apls_1_01_Appeal_ID", Text.Trim, type text}, {"Gf_Apls_1_01_ApAtrCat_2_01_Description", Text.Trim, type text}})
    ```

    ### What it does

    Removes leading and trailing spaces from selected text columns.

    ### Why it matters

    This improves matching reliability in lookups and helps avoid invisible data-quality issues.

    ### Risks / gotchas

    - Good cleanup step, but it does **not** standardise case.
    - `Cash`, `cash`, and `CASH` would still be different values during joins.

    ---

    ## Block 4 – Clean text fields

    ### Code

    ```m
    #"Cleaned Text" = Table.TransformColumns(#"Trimmed Text",{{"Gf_Type", Text.Clean, type text}, {"Gf_Gift_subtype", Text.Clean, type text}, {"Gf_Pay_method", Text.Clean, type text}, {"Gf_Reference_number", Text.Clean, type text}, {"Gf_Gift_code", Text.Clean, type text}, {"Gf_Import_ID", Text.Clean, type text}, {"Gf_Gift_ID", Text.Clean, type text}, {"Gf_NL_post_status", Text.Clean, type text}, {"Gf_Markting_Sorc_Cod", Text.Clean, type text}, {"Gf_Agency", Text.Clean, type text}, {"Gf_AttrCat_1_01_Description", Text.Clean, type text}, {"Gf_AttrCat_2_01_Description", Text.Clean, type text}, {"Gf_AttrCat_3_01_Description", Text.Clean, type text}, {"Gf_CnBio_Name", Text.Clean, type text}, {"Gf_Fnds_1_01_Description", Text.Clean, type text}, {"Gf_Fnds_1_01_Fn_Restricted", Text.Clean, type text}, {"Gf_Fnds_1_01_Fn_Fund_category", Text.Clean, type text}, {"Gf_Fnds_1_01_Fn_Fund_type", Text.Clean, type text}, {"Gf_Apls_1_01_Ap_Appeal_category", Text.Clean, type text}, {"Gf_Apls_1_01_Description", Text.Clean, type text}, {"Gf_Apls_1_01_Appeal_ID", Text.Clean, type text}, {"Gf_Apls_1_01_ApAtrCat_2_01_Description", Text.Clean, type text}})
    ```

    ### What it does

    Removes non-printable characters from the same set of text columns.

    ### Why it matters

    Useful when source files contain hidden characters from exports or copy/paste operations.

    ### Risks / gotchas

    - Good for hygiene, but still does not guarantee lookup keys are fully normalised.
    - If case inconsistency is an issue, add a `Text.Upper()` or `Text.Lower()` step before merging.

    ---

    ## Block 5 – Rename source date and create signed amount

    ### Code

    ```m
    #"Renamed Columns" = Table.RenameColumns(#"Cleaned Text",{{"Gf_Date", "Posting Date"}}),
    #"Inserted Multiplication" = Table.AddColumn(#"Renamed Columns", "Amount", each [Gf_Amount] * -1, type number)
    ```

    ### What it does

    - Renames `Gf_Date` to `Posting Date`
    - Creates a new `Amount` column by negating `Gf_Amount`

    ### Why it matters

    This establishes the accounting date and the signed amount used in the journal.

    ### Risks/gotchas

    - `Posting Date` is now anchored to `Gf_Date`, not `Gf_NL_post_date`.
    - `Amount = Gf_Amount * -1` is a critical accounting assumption.
    - If source data sometimes comes in negative already, the sign could be flipped incorrectly.

    ---

    ## Block 6 – Create Document Date

    ### Code

    ```m
    #"Duplicated Column" = Table.AddColumn(#"Reordered Columns", "Document Date", each [Posting Date], type date)
    ```

    ### What it does

    Creates `Document Date` by copying `Posting Date`.

    ### Why it matters

    Business Central commonly expects both posting and document dates. Here they are intentionally the same.

    ### Risks / gotchas

    - If the business later needs separate document date logic, this is the block to revisit.

    ---

    ## Block 7 – Reorder columns (early query housekeeping)

    ### Code

    ```m
    #"Reordered Columns" = ...
    #"Reordered Columns1" = ...
    #"Reordered Columns2" = ...
    ```

    ### What it does

    Moves columns into a preferred order.

    ### Why it matters

    Mostly visual / structural housekeeping created by the Power Query UI.

    ### Risks / gotchas

    - These steps do not change the values.
    - They make the query longer and harder to read.
    - In a refactor, many of these can be removed and replaced with one final reorder near the end.

    ---

    ## Block 8 – Document No. logic using Cat4 vs Batch Number

    ### Code

    ```m
    Cat4List = Table.Column(#"Reordered Columns2", "Gf_AttrCat_4_01_Description"),
    HasAnyCat4 = List.Count(List.Select(Cat4List, each _ <> null)) > 0,

    #"Added Document No." =
        Table.AddColumn(
            #"Reordered Columns2",
            "Document No.",
            each
                if HasAnyCat4 and [Gf_AttrCat_4_01_Description] <> null
                then [Gf_AttrCat_4_01_Description]
                else [Gf_Batch_Number],
            Int64.Type
        ),

    #"Removed Cat4 Column" =
        Table.RemoveColumns(#"Added Document No.", {"Gf_AttrCat_4_01_Description"})
    ```

    ### What it does

    - Looks at the whole dataset to see whether **any** row has a `Gf_AttrCat_4_01_Description`
    - Creates `Document No.` using:
    - `Gf_AttrCat_4_01_Description` if available on the row **and** at least one Cat4 exists somewhere in the dataset
    - otherwise `Gf_Batch_Number`
    - Removes the original Cat4 column afterward

    ### Why it matters

    `Document No.` is a core Business Central field and likely also drives traceability.

    ### Risks / gotchas

    - `HasAnyCat4` is a **table-level rule**, not a row-level rule.
    - This means behaviour in one row can depend on data existing in other rows.
    - Simpler and safer row-level logic would normally be:

    ```m
    if [Gf_AttrCat_4_01_Description] <> null then [Gf_AttrCat_4_01_Description] else [Gf_Batch_Number]
    ```

    ---

    ## Block 9 – External Document No.

    ### Code

    ```m
    #"Duplicated Column1" = Table.AddColumn(#"Removed Cat4 Column", "External Document No.", each [#"Document No."], type number)
    ```

    ### What it does

    Creates `External Document No.` as a copy of `Document No.`.

    ### Why it matters

    This preserves the same key in both BC-style fields, at least at this stage.

    ### Risks / gotchas

    - The column is later rebuilt again after more transformations.
    - Because it is typed as a number, any non-numeric document number logic in future would break.

    ---

    ## Block 10 – Remove no-longer-needed source columns

    ### Code

    ```m
    #"Removed Columns4" = Table.RemoveColumns(...)
    #"Removed Columns4" / #"Removed Columns" / #"Removed Columns1" / #"Removed Columns2" etc.
    ```

    ### What it does

    Strips out original source columns that are no longer needed for final journal output.

    ### Why it matters

    This keeps the shaped dataset leaner and more focused on BC-required fields.

    ### Risks / gotchas

    - Removing source columns too early can make troubleshooting harder later.
    - For supportability, it can be useful to maintain a separate debug query that keeps more of the original source fields.

    ---

    ## Block 11 – Original Index creation using hardcoded date replacement

    ### Code

    ```m
    #"Inserted Replaced Text" = Table.AddColumn(#"Reordered Columns5", "Index", each Text.Replace(Text.From([Posting Date], "en-GB"), "/09/2025", "09251000"), type text)
    ```

    ### What it does

    Creates an `Index` value by converting `Posting Date` to text and replacing the substring `/09/2025` with `09251000`.

    ### Why it matters

    This appears to be building a format required by the downstream importer or journal load process.

    ### Risks / gotchas

    - This is highly brittle and period-specific.
    - It assumes the date string will contain `/09/2025`.
    - For any other month/year, the output may be wrong or inconsistent.
    - This is one of the strongest candidates for refactoring.

    ### Safer long-term pattern

    ```m
    Date.ToText([Posting Date], "ddMMyy", "en-GB") & "1000"
    ```

    ---

    ## Block 12 – Derive Project Code and Appeal Code

    ### Code

    ```m
    #"Inserted Text Before Delimiter" = Table.AddColumn(#"Reordered Columns6", "Project Code", each Text.BeforeDelimiter([Gf_Apls_1_01_Ap_Appeal_category], " "), type text),
    #"Inserted Text Range" = Table.AddColumn(#"Inserted Text Before Delimiter", "Appeal Code", each Text.Middle([Gf_Apls_1_01_Appeal_ID], 2, 4), type text)
    ```

    ### What it does

    - `Project Code` = text before the first space in `Gf_Apls_1_01_Ap_Appeal_category`
    - `Appeal Code` = 4 characters extracted from `Gf_Apls_1_01_Appeal_ID`, starting at position 2

    ### Why it matters

    These are dimension values used in the final BC journal.

    ### Risks / gotchas

    - Both steps rely on **strict source formatting assumptions**.
    - If `Appeal ID` changes structure, `Appeal Code` can silently become wrong.
    - If `Appeal category` does not contain a space, `Text.BeforeDelimiter` may fail depending on source consistency.

    ---

    ## Block 13 – Fund lookup join

    ### Code

    ```m
    #"Merged Queries" = Table.NestedJoin(#"Inserted Text Range", {"Gf_Fnds_1_01_Fn_Fund_category"}, #"Fund-PQ", {"RE Data Export Name"}, "Fund-PQ", JoinKind.LeftOuter),
    #"Expanded Fund-PQ" = Table.ExpandTableColumn(#"Merged Queries", "Fund-PQ", {"BC Fund"}, {"Fund-PQ.BC Fund"}),
    #"Removed Columns6" = Table.RemoveColumns(#"Expanded Fund-PQ",{"Gf_Fnds_1_01_Fn_Fund_type", "Gf_Fnds_1_01_Fn_Fund_category", "Gf_Fnds_1_01_Description"}),
    #"Renamed Columns2" = Table.RenameColumns(#"Removed Columns6",{{"Fund-PQ.BC Fund", "Fund Code"}})
    ```

    ### What it does

    Uses `Fund-PQ` to map the RE fund category into the Business Central `Fund Code`.

    ### Why it matters

    This is a core dimension mapping. If it fails, the journal may be invalid or incomplete.

    ### Risks / gotchas

    - Any mismatch between source value and lookup key results in null `Fund Code`.
    - Duplicate keys in `Fund-PQ` can duplicate rows after expansion.
    - Because trimming/cleaning was done but case was not normalised, case differences may still fail to match.

    ---

    ## Block 14 – Derive Year Code from Appeal ID

    ### Code

    ```m
    #"Inserted Text Range1" = Table.AddColumn(#"Renamed Columns2", "Text Range", each Text.Middle([Gf_Apls_1_01_Appeal_ID], 6, 2), type text),
    #"Added Conditional Column" = Table.AddColumn(#"Inserted Text Range1", "Year Code", each if [Text Range] = "00" then "NA" else [Text Range], type text),
    #"Removed Columns7" = Table.RemoveColumns(#"Added Conditional Column",{"Text Range"})
    ```

    ### What it does

    Extracts a 2-character year code from the appeal ID.
    If that value is `00`, it replaces it with `NA`.

    ### Why it matters

    Creates another required dimension for the final journal.

    ### Risks / gotchas

    - Again depends on fixed character positions inside `Gf_Apls_1_01_Appeal_ID`.
    - If the source ID format changes, this logic may not error — it may just return the wrong answer.

    ---

    ## Block 15 – Add Document Type = PAYMENT

    ### Code

    ```m
    #"Added Custom" = Table.AddColumn(#"Removed Columns7", "Document Type", each "PAYMENT")
    ```

    ### What it does

    Adds a constant `Document Type` value of `PAYMENT` to all rows.

    ### Why it matters

    At this point in the logic, the journal appears intended to represent payment entries.

    ### Risks / gotchas

    - This value does **not survive to the final output**.
    - Later in the query, `Document Type` is removed and replaced with a blank/null-based field.
    - That means this block is misleading unless you know the later shaping logic.

    ---

    ## Block 16 – Department lookup join

    ### Code

    ```m
    #"Merged Queries1" = Table.NestedJoin(#"Added Custom", {"Gf_Apls_1_01_ApAtrCat_2_01_Description"}, #"DEP-PQ", {"RE data export name"}, "DEP-PQ", JoinKind.LeftOuter),
    #"Expanded DEP-PQ" = Table.ExpandTableColumn(#"Merged Queries1", "DEP-PQ", {"BC Department Code"}, {"DEP-PQ.BC Department Code"}),
    #"Renamed Columns3" = Table.RenameColumns(#"Expanded DEP-PQ",{{"DEP-PQ.BC Department Code", "Department Code"}})
    ```

    ### What it does

    Uses `DEP-PQ` to map a RE appeal attribute description to the Business Central `Department Code`.

    ### Why it matters

    Department is another key dimension for posting and reporting.

    ### Risks / gotchas

    - Null matches create blank department codes.
    - Duplicate keys in `DEP-PQ` could duplicate rows.
    - If the appeal attribute values are not clean and standardised, join failures may occur.

    ---

    ## Block 17 – VAT lookup join repurposed as Account No.

    ### Code

    ```m
    #"Merged Queries2" = Table.NestedJoin(#"Renamed Columns3", {"Gf_Gift_subtype"}, #"VAT-PQ", {"Gift Subtype"}, "VAT-PQ", JoinKind.LeftOuter),
    #"Expanded VAT-PQ" = Table.ExpandTableColumn(#"Merged Queries2", "VAT-PQ", {"BC Destination Code"}, {"VAT-PQ.BC Destination Code"}),
    #"Renamed Columns4" = Table.RenameColumns(#"Expanded VAT-PQ",{{"VAT-PQ.BC Destination Code", "Account No."}})
    ```

    ### What it does

    Maps `Gf_Gift_subtype` to `BC Destination Code` using `VAT-PQ`, then renames that result to `Account No.`.

    ### Why it matters

    This is the logic that decides which main account the original journal line posts to.

    ### Risks / gotchas

    - The lookup name suggests VAT, but the result is used as the actual `Account No.`.
    - That can be confusing for future maintainers.
    - Missing or incorrect mappings here directly affect accounting output.

    ---

    ## Block 18 – Remove intermediate columns after joins

    ### Code

    ```m
    #"Removed Columns" = Table.RemoveColumns(#"Renamed Columns4",{"Gf_Apls_1_01_ApAtrCat_2_01_Description", "Gf_Apls_1_01_ApAtrCat_1_01_Description", "Gf_Apls_1_01_Ap_Appeal_category"})
    ```

    ### What it does

    Removes no-longer-needed source columns after lookup/dimension derivation.

    ### Why it matters

    Leaves a cleaner journal shape.

    ### Risks / gotchas

    - If you later need to troubleshoot why project/department/account mapping is wrong, these source fields are no longer available in the final shaping steps.

    ---

    ## Block 19 – Build BC account type fields and description

    ### Code

    ```m
    #"Added Custom1" = Table.AddColumn(#"Removed Columns", "Bal. Account No.", each "null"),
    #"Added Custom2" = Table.AddColumn(#"Added Custom1", "Balance Account Type", each "G/L ACCOUNT"),
    #"Renamed Columns5" = Table.RenameColumns(#"Added Custom2",{{"Balance Account Type", "Bal. Account Type"}}),
    #"Added Custom3" = Table.AddColumn(#"Renamed Columns5", "Account Type", each "G/L ACCOUNT"),
    #"Inserted Merged Column" = Table.AddColumn(#"Added Custom3", "Merged", each Text.Combine({[Gf_AttrCat_2_01_Description], " - ", Text.From([#"Document No."], "en-GB")}), type text),
    #"Renamed Columns6" = Table.RenameColumns(#"Inserted Merged Column",{{"Merged", "Description"}}),
    #"Added Custom4" = Table.AddColumn(#"Removed Columns1", "Major Appeal Code", each "NA")
    ```

    ### What it does

    - Creates placeholder balancing account fields
    - Sets `Account Type = G/L ACCOUNT`
    - Builds `Description` as `[AttrCat2 Description] - [Document No.]`
    - Hardcodes `Major Appeal Code = NA`

    ### Why it matters

    This is where the query starts looking like a BC journal row structure.

    ### Risks / gotchas

    - `Bal. Account No.` is initially the literal text `"null"`, not a real null.
    - That is corrected later, but it is still a potential point of confusion.
    - `Major Appeal Code` is static, so there is no dynamic logic behind it.

    ---

    ## Block 20 – Shape the pre-balancing journal base

    ### Code

    ```m
    #"Removed Columns1" = Table.RemoveColumns(...),
    #"Added Custom5" = Table.AddColumn(#"Removed Columns1", "Custom", each "BANK ACCOUNT"),
    #"Removed Columns2" = Table.RemoveColumns(#"Added Custom5",{"Bal. Account Type"}),
    #"Reordered Columns7" = Table.ReorderColumns(...),
    #"Renamed Columns7" = Table.RenameColumns(#"Reordered Columns7",{{"Custom", "Bal. Account Type"}})
    ```

    ### What it does

    Further shapes the row structure and sets `Bal. Account Type` to `BANK ACCOUNT` via an intermediate column rename.

    ### Why it matters

    This is the final prepared dataset before the balancing row logic starts.

    ### Risks / gotchas

    - The query uses intermediate columns like `Custom`, then renames them.
    - Functionally fine, but not very readable.
    - At this point, the row is still just the original side of the journal.

    ---

    ## Block 21 – Parameter/configuration section

    ### Code

    ```m
    AmountColName        = "Amount",
    AccountTypeColName   = "Account Type",
    AccountNoColName     = "Account No.",
    DescriptionColName   = "Description",
    GFBioNameColName     = "Gf_CnBio_Name",
    PostingDateColName   = "Posting Date",
    IndexColName         = "Index",

    MainAccountNumber    = "53472586 MAIN",
    MainAccountType      = "Bank Account",
    DescriptionSeparator = " - ",
    ```

    ### What it does

    Defines reusable configuration values used by the balancing logic.

    ### Why it matters

    Makes later transformations easier to read and simplifies future changes.

    ### Risks / gotchas

    - `MainAccountNumber = "53472586 MAIN"` is hardcoded.
    - If the bank account changes, this is the key place to update.
    - `GFBioNameColName` and `PostingDateColName` are defined but not always used explicitly in dynamic ways.

    ---

    ## Block 22 – Set balancing start point

    ### Code

    ```m
    Start = #"Renamed Columns7"
    ```

    ### What it does

    Creates a named handoff point from the journal base rows into the balancing logic.

    ### Why it matters

    This is the effective start of the **double-entry generation** portion of the query.

    ### Risks / gotchas

    - None directly — this is a useful handover alias.

    ---

    ## Block 23 – Build helper description for balancing rows

    ### Code

    ```m
    AddedNewDescription =
        Table.AddColumn(
            Start,
            "NewDescription",
            each
                let
                    d = Text.Trim([Description]),
                    g = Text.Trim([Gf_CnBio_Name]),
                    parts = List.RemoveNulls({ if d <> "" then d else null, if g <> "" then g else null })
                in
                    if List.Count(parts) = 0 then "" else Text.Combine(parts, DescriptionSeparator),
            type text
        )
    ```

    ### What it does

    Creates a helper field called `NewDescription` by combining:
    - the existing `Description`
    - the donor name `Gf_CnBio_Name`

    Blank values are skipped.

    ### Why it matters

    This gives the balancing bank rows a richer description, improving audit traceability.

    ### Risks / gotchas

    - If donor names are sensitive or inconsistent, this affects balancing descriptions directly.
    - If either input is null rather than blank, current logic handles that reasonably well through the list filtering.

    ---

    ## Block 24 – Generate balancing base rows

    ### Code

    ```m
    BalancingBase =
        Table.TransformColumns(
            AddedNewDescription,
            {
                {AmountColName,      each if _ = null then null else -_, type number},
                {AccountTypeColName, each MainAccountType,                type text},
                {AccountNoColName,   each MainAccountNumber,              type text}
            }
        )
    ```

    ### What it does

    Creates a transformed version of the rows where:
    - `Amount` sign is reversed
    - `Account Type` becomes `Bank Account`
    - `Account No.` becomes `53472586 MAIN`

    ### Why it matters

    This is the **core double-entry logic**.
    It creates the balancing bank side of the journal.

    ### Risks / gotchas

    - If the original `Amount` sign convention is wrong, this balancing logic will amplify the issue.
    - If `MainAccountNumber` changes, all balancing lines will be wrong until updated.

    ---

    ## Block 25 – Replace balancing description with enriched version

    ### Code

    ```m
    BalancingRows0 =
        Table.RenameColumns(
            Table.RemoveColumns(BalancingBase, {DescriptionColName}),
            {{"NewDescription", DescriptionColName}}
        )
    ```

    ### What it does

    Removes the original `Description` from the balancing rows and replaces it with `NewDescription`.

    ### Why it matters

    Ensures balancing rows include donor-name-enhanced descriptions.

    ### Risks / gotchas

    - If `NewDescription` creation is ever changed or removed, this block must be updated too.

    ---

    ## Block 26 – Ensure balancing columns exist

    ### Code

    ```m
    BalancingRowsEnsure =
        let
            t1 = if Table.HasColumns(BalancingRows0, "Bal. Account Type")
                then BalancingRows0
                else Table.AddColumn(BalancingRows0, "Bal. Account Type", each null, type text),
            t2 = if Table.HasColumns(t1, "Bal. Account No.")
                then t1
                else Table.AddColumn(t1, "Bal. Account No.", each null, type text)
        in
            t2
    ```

    ### What it does

    Defensively ensures both balancing-related columns exist before proceeding.

    ### Why it matters

    Makes the query more resilient if earlier steps are edited later.

    ### Risks / gotchas

    - Good defensive coding pattern.
    - If someone restructures the earlier query and removes those columns, this prevents a hard failure here.

    ---

    ## Block 27 – Force balancing fields to true nulls

    ### Code

    ```m
    BalancingRowsSetBalBlank =
        Table.TransformColumns(
            BalancingRowsEnsure,
            {
                {"Bal. Account Type", each null, type any},
                {"Bal. Account No.",  each null, type any}
            }
        )
    ```

    ### What it does

    Overwrites both balancing fields with real null values.

    ### Why it matters

    Your comment suggests this is required for the **TES importer**.
    That means the final import format expects those columns to be blank.

    ### Risks / gotchas

    - If someone removes this step, text placeholders like `"null"` may leak into the output.
    - This is one of the rows where integration rules override what may look like standard BC journal behaviour.

    ---

    ## Block 28 – Rebuild balancing index

    ### Code

    ```m
    BalancingRowsWithIndex =
        let
            prep = if Table.HasColumns(BalancingRowsSetBalBlank, IndexColName)
                then Table.RemoveColumns(BalancingRowsSetBalBlank, {IndexColName})
                else BalancingRowsSetBalBlank
        in
            Table.AddColumn(
                prep,
                IndexColName,
                each Date.ToText([Posting Date], "ddMMyy", "en-GB") & "5000",
                type text
            )
    ```

    ### What it does

    Recreates the `Index` on balancing rows using:

    ```text
    ddMMyy + 5000
    ```

    ### Why it matters

    Likely creates a distinct import ordering/key pattern for the bank side of the journal.

    ### Risks / gotchas

    - This logic is much clearer than the earlier original index logic.
    - The fact that originals and balancing rows use different index-generation styles is important and should be preserved intentionally if the importer depends on it.

    ---

    ## Block 29 – Cleanup original rows before combining

    ### Code

    ```m
    OriginalRowsRaw = Table.RemoveColumns(AddedNewDescription, {"NewDescription"}),

    OriginalRowsEnsure =
        let
            o1 = if Table.HasColumns(OriginalRowsRaw, "Bal. Account Type")
                then OriginalRowsRaw
                else Table.AddColumn(OriginalRowsRaw, "Bal. Account Type", each null, type text),
            o2 = if Table.HasColumns(o1, "Bal. Account No.")
                then o1
                else Table.AddColumn(o1, "Bal. Account No.", each null, type text)
        in
            o2,

    OriginalRowsBalBlank =
        Table.TransformColumns(
            OriginalRowsEnsure,
            {
                {"Bal. Account Type", each null, type any},
                {"Bal. Account No.",  each null, type any}
            }
        )
    ```

    ### What it does

    - Removes the helper description from the original row set
    - Ensures balancing columns exist
    - Forces balancing columns to blank/null on original rows too

    ### Why it matters

    The final import is designed so that **both original and balancing rows have blank Bal. fields**.

    ### Risks / gotchas

    - This is slightly counterintuitive if someone expects balancing columns to carry the counter-entry info.
    - In this query, balancing is represented by duplicated rows, not by filling `Bal. Account` fields.

    ---

    ## Block 30 – Combine original rows and balancing rows

    ### Code

    ```m
    Output = Table.Combine({OriginalRowsBalBlank, BalancingRowsWithIndex})
    ```

    ### What it does

    Stacks the original journal rows together with the balancing bank rows.

    ### Why it matters

    This is the moment the output becomes a full **double-entry journal**.

    ### Risks / gotchas

    - If row counts double unexpectedly, remember that this is intentional here.
    - Any issue upstream is now duplicated into both row sets.

    ---

    ## Block 31 – Remove zero / null amount rows

    ### Code

    ```m
    Output_NoZero = Table.SelectRows(Output, each [Amount] <> 0 and [Amount] <> null)
    ```

    ### What it does

    Removes rows where `Amount` is zero or null.

    ### Why it matters

    Prevents invalid or pointless journal lines from reaching the final import.

    ### Risks / gotchas

    - If a valid business row is genuinely meant to be zero-valued (rare), it would be removed.
    - In most journal scenarios this is the right thing to do.

    ---

    ## Block 32 – Final cleanup and layout shaping

    ### Code

    ```m
    #"Removed Columns3" = Table.RemoveColumns(Output_NoZero,{"Gf_CnBio_Name"}),
    #"Reordered Columns8" = Table.ReorderColumns(...),
    #"Removed Columns5" = Table.RemoveColumns(#"Reordered Columns8",{"External Document No."}),
    #"Reordered Columns9" = Table.ReorderColumns(...)
    ```

    ### What it does

    Starts the final output shaping:
    - removes donor name after it has already been used in balancing descriptions
    - reorders fields into BC/import layout
    - removes and later rebuilds `External Document No.`

    ### Why it matters

    Creates the final output structure the importer expects.

    ### Risks / gotchas

    - Lots of column-order manipulation here makes maintenance harder.
    - Removing and re-adding fields can be confusing for future developers.

    ---

    ## Block 33 – Replace Document Type with blank/null field

    ### Code

    ```m
    #"Added Custom6" = Table.AddColumn(#"Reordered Columns9", "replaceCustom", each null),
    #"Reordered Columns10" = Table.ReorderColumns(#"Added Custom6",{"Index", "Posting Date", "Document Date", "Document No.", "Fund Code", "Department Code", "Project Code", "Year Code", "Appeal Code", "Bal. Account No.", "Document Type", "replaceCustom", "Account No.", "Account Type", "Description", "Bal. Account Type", "Amount", "Major Appeal Code"}),
    #"Removed Columns8" = Table.RemoveColumns(#"Reordered Columns10",{"Document Type"}),
    #"Renamed Columns8" = Table.RenameColumns(#"Removed Columns8",{{"replaceCustom", "Document Type"}})
    ```

    ### What it does

    - Creates a blank/null helper column called `replaceCustom`
    - Removes the existing `Document Type`
    - Renames the blank helper into `Document Type`

    ### Why it matters

    This means the **final `Document Type` is blank**, even though earlier it was set to `PAYMENT`.

    ### Risks / gotchas

    - Very easy place for developer confusion.
    - If someone reviews only the earlier steps, they may assume output `Document Type = PAYMENT`, which is not true.
    - If the importer actually needs blanks, this is intentional and should be documented clearly.

    ---

    ## Block 34 – Add Line No.

    ### Code

    ```m
    #"Added Index" = Table.AddIndexColumn(#"Renamed Columns8", "Index.1", 0, 1, Int64.Type),
    #"Reordered Columns11" = Table.ReorderColumns(...),
    #"Renamed Columns9" = Table.RenameColumns(#"Reordered Columns11",{{"Index.1", "Line No."}}),
    #"Removed Columns9" = Table.RemoveColumns(#"Renamed Columns9",{"Index"})
    ```

    ### What it does

    - Adds a sequential row index starting at 0
    - Renames it to `Line No.`
    - Removes the previous `Index` column

    ### Why it matters

    `Line No.` is likely required by the target import format.

    ### Risks / gotchas

    - The original `Index` values are discarded here.
    - If the import process depends on the old `Index`, then removing it would matter — but based on current output design, `Line No.` appears to replace it.

    ---

    ## Block 35 – Recreate External Document No.

    ### Code

    ```m
    #"Duplicated Column2" = Table.AddColumn(#"Reordered Columns12", "External Document No.", each [#"Document No."], type number)
    ```

    ### What it does

    Recreates `External Document No.` from `Document No.` after the final reshaping.

    ### Why it matters

    Ensures the field is present in the final output layout.

    ### Risks / gotchas

    - Again typed as number, so it assumes document numbers remain numeric.

    ---

    ## Block 36 – Add Break / temporary shaping column

    ### Code

    ```m
    #"Added Custom7" = Table.AddColumn(#"Removed Columns10", "Break", each "null"),
    #"Reordered Columns14" = Table.ReorderColumns(...),
    #"Removed Columns11" = Table.RemoveColumns(#"Reordered Columns15",{"Break"})
    ```

    ### What it does

    Creates a temporary `Break` column, reorders around it, then removes it.

    ### Why it matters

    This is likely just UI-generated layout manipulation.

    ### Risks / gotchas

    - No business logic value.
    - Good candidate to simplify/remove in any future refactor.

    ---

    ## Block 37 – Create Currency from final Document Type

    ### Code

    ```m
    #"Duplicated Column3" = Table.DuplicateColumn(#"Reordered Columns14", "Document Type", "Document Type - Copy"),
    #"Reordered Columns15" = Table.ReorderColumns(...),
    #"Renamed Columns1" = Table.RenameColumns(#"Removed Columns11",{{"Document Type - Copy", "Currency"}})
    ```

    ### What it does

    Duplicates final `Document Type` into a new column and renames it to `Currency`.

    ### Why it matters

    Creates a final `Currency` column for the export layout.

    ### Risks / gotchas

    - Since final `Document Type` was replaced with null/blank, `Currency` will also be blank/null.
    - If the importer expects blank currency, this is okay.
    - If the importer expects a currency code like `GBP`, this is a bug waiting to happen.

    ---

    ## Block 38 – Final output

    ### Code

    ```m
    in
        #"Renamed Columns1"
    ```

    ### What it does

    Returns the fully transformed journal table.

    ### Why it matters

    This is the final query output loaded into the workbook / destination.

    ### Risks / gotchas

    - Whatever state `Renamed Columns1` is in becomes the export/output dataset.
    - Any confusion earlier in the script must be resolved before this point because this is what users actually consume.

    ---

    # 4. Final output fields (functional view)

    The final output is effectively shaped into these importer-facing fields:

    - `Posting Date`
    - `Amount`
    - `Document Date`
    - `Document No.`
    - `External Document No.`
    - `Line No.`
    - `Currency`
    - `Project Code`
    - `Appeal Code`
    - `Fund Code`
    - `Year Code`
    - `Document Type`
    - `Department Code`
    - `Account No.`
    - `Bal. Account No.`
    - `Bal. Account Type`
    - `Account Type`
    - `Description`
    - `Major Appeal Code`

    ---

    # 5. Most important places to check when debugging

    ## 1. Amount sign

    ```m
    Amount = [Gf_Amount] * -1
    ```

    If totals are off, this is one of the first places to validate.

    ## 2. Document No. generation

    ```m
    if HasAnyCat4 and [Gf_AttrCat_4_01_Description] <> null then ... else ...
    ```

    This can create unexpected behaviour in mixed datasets.

    ## 3. Lookup joins

    The three critical mappings are:

    - `Fund-PQ` → `Fund Code`
    - `DEP-PQ` → `Department Code`
    - `VAT-PQ` → `Account No.`

    ## 4. Balancing rows

    If the journal does not net to zero, check:

    - sign logic
    - balancing account number
    - row duplication logic

    ## 5. Final blanking of `Document Type` and `Currency`

    The output is blank/null even though earlier steps seem to populate these fields.

    ---

    # 6. Recommended future improvements

    ## Improvement 1 – Replace hardcoded original Index logic

    Current:

    ```m
    Text.Replace(Text.From([Posting Date], "en-GB"), "/09/2025", "09251000")
    ```

    Suggested:

    ```m
    Date.ToText([Posting Date], "ddMMyy", "en-GB") & "1000"
    ```

    ---

    ## Improvement 2 – Normalise lookup keys further

    Before merges, consider:

    ```m
    Text.Upper(Text.Trim(Text.Clean(_)))
    ```

    on both source keys and lookup keys.

    ---

    ## Improvement 3 – Simplify column reordering

    Many `ReorderColumns` steps can probably be removed and replaced with one final reorder.

    ---

    ## Improvement 4 – Clarify final Document Type / Currency design

    If blanks are intentional, add a clear comment in the code.
    If not intentional, fix the final shaping logic.

    ---

    ## Improvement 5 – Add exception queries

    Recommended checks:

    - unmapped Fund Code
    - unmapped Department Code
    - unmapped Account No.
    - duplicate keys in lookup tables
    - final total of original + balancing rows = 0

    ---

    # 7. Short developer summary

    > The `DB Journal` query reads Daily Banking source data from `Table1`, cleans and enriches it using lookup queries, derives Business Central journal dimensions and account fields, then creates balancing bank rows to produce a final import-ready double-entry journal.

    ---

    # 8. Suggested handover note for future maintainers

    If you need to modify this query, review these areas first:

    1. **Source schema and data types**
    2. **Lookup query structures (`Fund-PQ`, `DEP-PQ`, `VAT-PQ`)**
    3. **Document No. generation logic**
    4. **Balancing row creation (`53472586 MAIN`)**
    5. **Final blank `Document Type` / `Currency` behaviour**

    These are the parts most likely to affect output correctness.
