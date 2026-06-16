# Addition Journal Power Query – Developer Handover (Block-by-Block Annotation)

This document is a **developer handover guide** for the **Addition journal** Power Query.

It explains the query **block by block**, with each section showing:

- **Code block** – the relevant M code
- **What it does** – plain-English explanation
- **Why it matters** – technical / business significance
- **Risks / gotchas** – things that may break, confuse, or create mismatches

---

# 1. Purpose of this query

The **Addition** journal query builds a **Business Central import-ready journal** for **Braintree disbursement additions**, including:

1. source gift rows from `Table16`
2. enrichment from Braintree disbursement data
3. lookup-driven account / department / fund mapping
4. grouped income journal lines
5. proportional fee allocation lines
6. orphan fee handling for fee dates with no same-day income
7. append of `Bank Journal`
8. final Business Central output shaping

At a high level, the query produces a journal that combines:

- **income lines**
- **fee lines**
- **fallback/orphan fee lines**
- **bank journal rows**

---

# 2. High-level query flow

```text
Table16 source
→ Date normalisation helper
→ Braintree match
→ Text cleanup
→ Posting/document/amount fields
→ VAT / Department / Fund lookups
→ Dimension extraction
→ Description creation
→ Group income rows
→ Calculate daily fee allocation
→ Build fee rows
→ Build orphan fee rows
→ Combine all output rows
→ Add BC-required columns
→ Append Bank Journal
→ Final document-number reconciliation
→ Add line number
→ Filter zero rows
```

---

# 3. Block-by-block annotation

## Block 1 – Reusable date conversion helper

### Code

```m
FxToDate = (v as any) as nullable date =>
if v = null then
    null
else if Value.Is(v, type date) then
    v
else if Value.Is(v, type datetime) or Value.Is(v, type datetimezone) then
    Date.From(v)
else
    Date.FromText(Text.From(v), [Culture = "en-GB"])
```

### What it does
Creates a helper function called `FxToDate` that converts different date-like inputs into a proper `date` value.

It handles:
- `null`
- already-typed dates
- datetimes / datetimezones
- text values parsed using `en-GB`

### Why it matters
This is one of the strongest parts of the query from a reliability point of view. It standardises how dates are parsed and avoids repeated custom date logic throughout the rest of the script.

### Risks / gotchas
- If a text value is not valid under `en-GB`, refresh will fail.
- If source values are inconsistent (for example mixed locale text), this helper may still error.
- If source contains Excel serial numbers as numbers, this helper does **not** explicitly convert them.

---

## Block 2 – Load source table

### Code

```m
Source = Excel.CurrentWorkbook(){[Name="Table16"]}[Content]
```

### What it does
Loads the Excel table named **`Table16`** from the current workbook.

### Why it matters
This is the root data source for the Addition journal. All downstream logic depends on the structure and quality of this table.

### Risks / gotchas
- If `Table16` is renamed, refresh fails.
- If source column names change, many downstream steps will break because the query uses hardcoded column names.

---

## Block 3 – Initial date normalisation on source columns

### Code

```m
#"Changed Type3" = Table.TransformColumns(
Source,
{
    {"Gift Date", each FxToDate(_), type date},
    {"NL post date", each FxToDate(_), type date}
}
)
```

### What it does
Applies the `FxToDate` function to:
- `Gift Date`
- `NL post date`

### Why it matters
This ensures those key source date fields are consistently typed before joins and downstream logic.

### Risks / gotchas
- Only these two fields are normalised here; `Braintree` disbursement date is handled later.
- `NL post date` is typed but later removed, so if someone assumes it drives posting logic, that is not true in the final output.

---

## Block 4 – Join to Braintree transaction data

### Code

```m
#"Merged Queries" = Table.NestedJoin(#"Changed Type3", {"Gift ID"}, #"Braintree PQ", {"Transaction ID"}, "Braintree PQ", JoinKind.LeftOuter)
```

### What it does
Joins the source table to **`Braintree PQ`** using:
- source key: `Gift ID`
- lookup key: `Transaction ID`

### Why it matters
This is the bridge between the gift data and the Braintree disbursement information. It is what allows the query to derive a posting date from Braintree.

### Risks / gotchas
- If `Gift ID` does not reliably equal `Transaction ID`, posting dates will go missing.
- If `Braintree PQ` contains duplicate `Transaction ID` values, source rows can duplicate after expansion.
- Because this is a Left Outer join, unmatched records are retained but with null Braintree values.

---

## Block 5 – Clean Gift subtype

### Code

```m
#"Cleaned Gift Subtype" = Table.TransformColumns(
#"Merged Queries",
{{"Gift subtype", each Text.Upper(Text.Trim(Text.Clean(_))), type text}}
)
```

### What it does
Normalises `Gift subtype` using:
- `Text.Clean`
- `Text.Trim`
- `Text.Upper`

### Why it matters
This is excellent defensive preparation for lookup joins later, especially for VAT/account mapping.

### Risks / gotchas
- This standardisation is only applied to `Gift subtype`, not all other lookup keys.
- If the lookup table `VAT PQ[Gift Subtype]` is not standardised the same way, mismatches can still happen.

---

## Block 6 – Expand and type Braintree disbursement date

### Code

```m
#"Expanded Braintree PQ" = Table.ExpandTableColumn(#"Cleaned Gift Subtype", "Braintree PQ", {"Disbursement Date"}, {"Braintree PQ.Disbursement Date"}),

#"Typed Braintree Date" = Table.TransformColumns(
#"Expanded Braintree PQ",
{{"Braintree PQ.Disbursement Date", each FxToDate(_), type date}}
)
```

### What it does
- Expands `Disbursement Date` from the Braintree join
- Converts it to a proper date using `FxToDate`

### Why it matters
This field later becomes the **Posting Date**, so this is one of the most important enrichment steps in the whole query.

### Risks / gotchas
- If Braintree has no matching row, posting date will become null later.
- If disbursement dates are malformed, `FxToDate` can fail.

---

## Block 7 – Reorder and retype main dates

### Code

```m
#"Reordered Columns" = Table.ReorderColumns(...),

#"Changed Type" = Table.TransformColumns(
#"Reordered Columns",
{{"Gift Date", each FxToDate(_), type date}}
)
```

### What it does
- Reorders the source columns
- Reapplies date typing to `Gift Date`

### Why it matters
Mostly housekeeping, but it ensures `Gift Date` is still safely typed before being renamed to document date.

### Risks / gotchas
- The extra date retyping is harmless but slightly redundant.
- Many reorder steps add complexity without adding logic value.

---

## Block 8 – Rename posting/document/amount fields

### Code

```m
#"Renamed Columns" = Table.RenameColumns(#"Changed Type", {{"Braintree PQ.Disbursement Date", "Posting Date"}, {"Gift Date", "Document Date"}, {"Gift Amount", "Amount"}}),

#"Multiplied Column" = Table.TransformColumns(#"Renamed Columns", {{"Amount", each _ * -1, type number}})
```

### What it does
- `Braintree PQ.Disbursement Date` becomes `Posting Date`
- `Gift Date` becomes `Document Date`
- `Gift Amount` becomes `Amount`
- `Amount` is multiplied by `-1`

### Why it matters
This locks in the journal’s accounting and dating basis:
- **Posting Date** = Braintree disbursement date
- **Document Date** = original gift date
- **Amount** = negative source amount (for journal sign convention)

### Risks / gotchas
- If Braintree dates are missing, posting date is null.
- The negative sign assumption is critical. If source gift amounts are already negative, amounts will flip incorrectly.
- This is one of the first places to inspect if totals look reversed.

---

## Block 9 – VAT lookup to derive Account No.

### Code

```m
#"Merged Queries1" = Table.NestedJoin(#"Multiplied Column", {"Gift subtype"}, #"VAT PQ", {"Gift Subtype"}, "VAT PQ", JoinKind.LeftOuter),
#"Expanded VAT PQ" = Table.ExpandTableColumn(#"Merged Queries1", "VAT PQ", {"BC Destination Code"}, {"VAT PQ.BC Destination Code"}),
#"Renamed Columns1" = Table.RenameColumns(#"Expanded VAT PQ", {{"VAT PQ.BC Destination Code", "Account No."}})
```

### What it does
Uses `VAT PQ` to map `Gift subtype` to a Business Central `Account No.`.

### Why it matters
This is the first core account mapping for the income side of the journal.

### Risks / gotchas
- This mapping is critical. Missing lookup matches create null `Account No.`.
- The lookup name suggests VAT, but it is effectively being used as the income account map.
- Both source and lookup should be normalised the same way to avoid silent mismatches.

---

## Block 10 – Department lookup

### Code

```m
#"Merged Queries2" = Table.NestedJoin(#"Renamed Columns1", {"Department Code Description"}, #"DEP PQ", {"RE data export name"}, "DEP PQ", JoinKind.LeftOuter),
#"Expanded DEP PQ" = Table.ExpandTableColumn(#"Merged Queries2", "DEP PQ", {"Code"}, {"DEP PQ.Code"}),
#"Renamed Columns2" = Table.RenameColumns(#"Expanded DEP PQ", {{"DEP PQ.Code", "Department Code"}})
```

### What it does
Maps source department text into `Department Code` using `DEP PQ`.

### Why it matters
`Department Code` is a key reporting and posting dimension in the final Business Central journal.

### Risks / gotchas
- Unmatched department text produces null department codes.
- If `DEP PQ` contains duplicate keys, rows can duplicate on expand.

---

## Block 11 – Remove source-only columns and derive Project Code

### Code

```m
#"Removed Columns" = Table.RemoveColumns(#"Renamed Columns2", {"Department Code Description", "Appeal VAT Code Description"}),

#"Inserted Text Before Delimiter" = Table.AddColumn(#"Removed Columns", "Project Code", each Text.BeforeDelimiter([Income Ref], " "), type text),
#"Removed Columns1" = Table.RemoveColumns(#"Inserted Text Before Delimiter", {"Income Ref"})
```

### What it does
- Drops intermediate source fields no longer needed
- Derives `Project Code` from the first token in `Income Ref`
- Removes `Income Ref` afterwards

### Why it matters
`Project Code` becomes a final BC dimension.

### Risks / gotchas
- `Text.BeforeDelimiter([Income Ref], " ")` assumes that `Income Ref` always contains a space.
- If `Income Ref` is blank or differently structured, this can error or produce wrong project codes.

---

## Block 12 – Fund lookup

### Code

```m
#"Merged Queries3" = Table.NestedJoin(#"Removed Columns1", {"Expenditure Ref"}, #"FUND PQ", {"RE Data Export Name"}, "FUND PQ", JoinKind.LeftOuter),
#"Expanded FUND PQ" = Table.ExpandTableColumn(#"Merged Queries3", "FUND PQ", {"Fund"}, {"FUND PQ.Fund"})
```

### What it does
Uses `FUND PQ` to derive a fund value from `Expenditure Ref`.

### Why it matters
This contributes to the journal’s `Fund Code` / fund classification later.

### Risks / gotchas
- The query later renames `FUND PQ.Fund` to `Fund Code`, but then later also renames `Fund ID` to `Fund Code` after output shaping.
- This means there is a naming collision / override risk that is important to understand.

---

## Block 13 – Remove noise fields, derive Appeal and Year codes

### Code

```m
#"Removed Columns2" = Table.RemoveColumns(#"Expanded FUND PQ", {"Expenditure Ref", "Funder Ref", "Name", "Constituent ID", "Gift Aid Received Description", "Gift Aid Collection Fees Description", "Paying In Slip Number Description", "Postage & Packaging Description", "Agency", "Marketing Source Code", "NL post status", "Gift code", "Reference number", "Batch Number", "Gf_Receipt_amount"}),

#"Inserted Text Range" = Table.AddColumn(#"Removed Columns2", "Appeal Code", each Text.Middle([Appeal ID], 2, 4), type text),
#"Inserted Text Range1" = Table.AddColumn(#"Inserted Text Range", "Year Code", each Text.Middle([Appeal ID], 6, 2), type text),

#"Removed Columns3" = Table.RemoveColumns(#"Inserted Text Range1", {"Appeal ID"}),
#"Renamed Columns3" = Table.RenameColumns(#"Removed Columns3", {{"FUND PQ.Fund", "Fund Code"}}),
#"Removed Columns4" = Table.RemoveColumns(#"Renamed Columns3", {"NL post date"})
```

### What it does
- Removes many source-only descriptive columns
- Derives `Appeal Code` from characters 2–5 of `Appeal ID`
- Derives `Year Code` from characters 6–7 of `Appeal ID`
- Renames the fund lookup output to `Fund Code`
- Removes `NL post date`

### Why it matters
This step establishes the core dimensions used in grouping and final output.

### Risks / gotchas
- `Appeal ID` parsing assumes the ID format is fixed.
- Wrong/short `Appeal ID` values can silently create bad codes.
- Removing `NL post date` confirms that field is **not** used in final posting logic.

---

## Block 14 – Build Description from source descriptions + posting date

### Code

```m
#"Inserted Merged Column" = Table.AddColumn(
#"Removed Columns4",
"Merged",
each Text.Combine(
    List.RemoveNulls({
        [Income Source Description],
        [Data Collection Source Description],
        if [Posting Date] <> null then "Income - " & Date.ToText([Posting Date], "dd/MM/yyyy", "en-GB") else null
    }),
    " "
),
type text
),

#"Renamed Columns4" = Table.RenameColumns(#"Inserted Merged Column", {{"Merged", "Description"}})
```

### What it does
Builds `Description` by combining:
- `Income Source Description`
- `Data Collection Source Description`
- a generated label like `Income - dd/MM/yyyy`

### Why it matters
Creates a more descriptive journal narrative for the income rows.

### Risks / gotchas
- This description is later overwritten after output shaping.
- So although this step is useful, it does **not** survive as the final `Description` in the final output.

---

## Block 15 – Remove additional source columns

### Code

```m
#"Removed Columns5" = Table.RemoveColumns(#"Renamed Columns4", {"Braintree Description", "Reference", "Gift Import ID", "Income Source Description", "Data Collection Source Description", "Gift ID", "Donation Collection Fees Description"})
```

### What it does
Removes source/descriptive columns no longer needed before grouping.

### Why it matters
Makes the grouped dataset cleaner.

### Risks / gotchas
- Once these are removed, tracing line-level output back to raw input becomes harder.

---

## Block 16 – Group source rows into income buckets

### Code

```m
#"Grouped Rows" = Table.Group(
#"Removed Columns5",
{"Pay method", "Gift subtype", "Appeal Code", "Project Code", "Fund ID", "Department Code", "Posting Date", "Document Date", "Year Code"},
{{"All", each _, type table [Gift Type=text, Gift subtype=text, Posting Date=nullable date, Document Date=nullable date, Amount=number, Pay method=text, Fund ID=text, #"Account No."=nullable number, Department Code=number, Project Code=text, Fund Code=nullable text, Appeal Code=text, Year Code=text, Description=text]}}
),

#"Added Amount" = Table.AddColumn(#"Grouped Rows", "Amount", each List.Sum([All][Amount]), type number)
```

### What it does
Groups source rows by a set of journal dimensions and then calculates the summed total `Amount` for each group.

### Why it matters
This is the main **income-line aggregation step**. It collapses individual source transactions into grouped journal lines.

### Risks / gotchas
- Grouping uses `Fund ID`, not `Fund Code`, even though a `Fund Code` already exists from the fund lookup.
- That suggests the grouping logic and final output naming may not fully align.
- Any dimension omitted from the group keys can cause lines to merge unexpectedly.

---

## Block 17 – Reapply VAT/account mapping for grouped income

### Code

```m
#"Merged VAT For Income" = Table.NestedJoin(#"Added Amount", {"Gift subtype"}, #"VAT PQ", {"Gift Subtype"}, "VAT PQ", JoinKind.LeftOuter),
#"Expanded VAT For Income" = Table.ExpandTableColumn(#"Merged VAT For Income", "VAT PQ", {"BC Destination Code"}, {"VAT PQ.BC Destination Code"}),
#"Renamed Account" = Table.RenameColumns(#"Expanded VAT For Income", {{"VAT PQ.BC Destination Code", "Account No."}})
```

### What it does
After grouping, it reapplies the VAT/account mapping to make sure grouped income rows still have `Account No.`.

### Why it matters
This keeps the grouped income output self-contained and posting-ready.

### Risks / gotchas
- There is now account mapping logic both before grouping and after grouping.
- Functionally okay, but it can confuse maintainers because the earlier mapping is effectively intermediate and this later mapping is the one that matters for grouped income output.

---

## Block 18 – Prepare Braintree fee source table

### Code

```m
FeesTyped = Table.TransformColumns(
#"PP Fees PQ",
{
    {"disbursement_date", each FxToDate(_), type date},
    {"discount_GBP", each Number.From(_), type number},
    {"per_transaction_fees_GBP", each Number.From(_), type number},
    {"cross_border_fees_GBP", each Number.From(_), type number}
}
),

FeesNullToZero = Table.ReplaceValue(
FeesTyped, null, 0, Replacer.ReplaceValue,
{"discount_GBP","per_transaction_fees_GBP","cross_border_fees_GBP"}
),

FeesAddTotal = Table.AddColumn(
FeesNullToZero,
"FeeTotal",
each [discount_GBP] + [per_transaction_fees_GBP] + [cross_border_fees_GBP],
type number
),

DailyFees = Table.Group(
FeesAddTotal,
{"disbursement_date"},
{{"TargetFee", each Number.Round(-List.Sum([FeeTotal]), 2), type number}}
)
```

### What it does
- Loads and types the Braintree fee data from `PP Fees PQ`
- Converts null fee values to 0
- Calculates total fee per row
- Groups fees by `disbursement_date`
- Produces a daily fee target amount `TargetFee`

### Why it matters
This creates the **daily fee pool** that will be proportionally allocated across grouped income lines for the same day.

### Risks / gotchas
- The negative sign in `TargetFee = -List.Sum([FeeTotal])` is intentional and critical.
- If fee source values are already signed differently than expected, fee postings will be wrong.
- Converting nulls to zero is sensible, but it masks whether fee source values were actually missing.

---

## Block 19 – Prepare daily income totals for fee allocation

### Code

```m
#"Add AbsIncome" = Table.AddColumn(#"Renamed Account", "AbsIncome", each Number.Abs([Amount]), type number),

DailyIncome = Table.Group(
#"Add AbsIncome",
{"Posting Date"},
{{"DailyAbsIncome", each List.Sum([AbsIncome]), type number}}
)
```

### What it does
- Creates `AbsIncome` = absolute value of each grouped income amount
- Sums absolute income by posting date

### Why it matters
This is the denominator used to proportionally allocate fee amounts across each day’s grouped income lines.

### Risks / gotchas
- Using absolute values avoids sign problems but means the fee allocation is driven by magnitude only.
- If the day contains mixed positive/negative income corrections, allocation may not behave as expected.

---

## Block 20 – Merge daily income and daily fee targets into grouped income rows

### Code

```m
#"Merge Daily Income" = Table.NestedJoin(
#"Add AbsIncome", {"Posting Date"},
DailyIncome, {"Posting Date"},
"DI", JoinKind.LeftOuter
),

#"Expand Daily Income" = Table.ExpandTableColumn(#"Merge Daily Income", "DI", {"DailyAbsIncome"}, {"DailyAbsIncome"}),

#"Merge Daily Fees" = Table.NestedJoin(
#"Expand Daily Income", {"Posting Date"},
DailyFees, {"disbursement_date"},
"DF", JoinKind.LeftOuter
),

#"Expand Daily Fees" = Table.ExpandTableColumn(#"Merge Daily Fees", "DF", {"TargetFee"}, {"TargetFee"})
```

### What it does
Adds to each grouped income row:
- the day’s total absolute income
- the day’s total target fee

### Why it matters
This prepares each grouped income row to receive its proportional share of the day’s fee total.

### Risks / gotchas
- If a posting date has fees but no matching income, `TargetFee` exists but grouped rows do not. This is handled later by the orphan fee logic.
- If posting date is null, joins may behave unexpectedly or leave values blank.

---

## Block 21 – Calculate fee allocation share and rounded fee amount

### Code

```m
#"Add Share" = Table.AddColumn(
#"Expand Daily Fees",
"ShareOfDay",
each if [TargetFee] = null or [DailyAbsIncome] = null or [DailyAbsIncome] = 0 then null else [AbsIncome] / [DailyAbsIncome],
type number
),

#"Add Fee Unrounded" = Table.AddColumn(
#"Add Share",
"Fee_Unrounded",
each if [ShareOfDay] = null then null else [TargetFee] * [ShareOfDay],
type number
),

#"Add Fee Rounded" = Table.AddColumn(
#"Add Fee Unrounded",
"Fee_Rounded",
each if [Fee_Unrounded] = null then null else Number.Round([Fee_Unrounded], 2),
type number
)
```

### What it does
For each grouped income row:
- calculates the row’s share of the day’s income
- applies that share to the day’s target fee
- rounds the calculated fee line to 2 dp

### Why it matters
This is the **proportional fee allocation engine**.

### Risks / gotchas
- Rounding introduces day-level small differences that must be corrected later.
- If `DailyAbsIncome = 0`, no fee is allocated.
- If source income aggregation is wrong, fee allocation will also be wrong.

---

## Block 22 – Create fee candidate rows and assign fee account

### Code

```m
FeeCandidates0 = Table.SelectRows(#"Add Fee Rounded", each [Fee_Rounded] <> null and [Fee_Rounded] <> 0),

FeeLinesPrep = Table.RemoveColumns(FeeCandidates0, {"Amount"}),
FeeLinesAddAmount = Table.AddColumn(FeeLinesPrep, "Amount", each [Fee_Rounded], type number),
FeeLinesSetAcct = Table.TransformColumns(FeeLinesAddAmount, {{"Account No.", each 25100, Int64.Type}})
```

### What it does
- Keeps only rows with a non-zero rounded fee allocation
- Replaces `Amount` with the fee amount
- Forces `Account No.` to **25100** for fee rows

### Why it matters
This is where fee rows become their own journal line type.

### Risks / gotchas
- `25100` is hardcoded. If the fee account changes, this must be updated.
- The logic assumes all allocated fees go to one fixed code.

---

## Block 23 – Correct fee rounding drift at daily level

### Code

```m
FeeLinesGroupedByDay = Table.Group(
FeeLinesSetAcct,
{"Posting Date"},
{{"AllFeeLines", each _, type table}}
),

FeeLinesAdjustedTables = Table.TransformColumns(
FeeLinesGroupedByDay,
{"AllFeeLines", (t as table) as table =>
    let
        tgtList = List.RemoveNulls(t[TargetFee]),
        tgt = if List.Count(tgtList)=0 then null else tgtList{0},
        t0 = t,
        sumAmt = List.Sum(t0[Amount]),
        diff = if tgt = null then 0 else Number.Round(tgt - sumAmt, 2),
        tIdx = Table.AddIndexColumn(t0, "Idx", 0, 1, Int64.Type),
        tAdj = Table.AddColumn(
            tIdx,
            "AmountAdj",
            each if [Idx] = 0 then [Amount] + diff else [Amount],
            type number
        ),
        tDrop = Table.RemoveColumns(tAdj, {"Amount","Idx"}),
        tOut = Table.RenameColumns(tDrop, {{"AmountAdj","Amount"}})
    in
        tOut
}
),

FeeLinesAdjusted = Table.ExpandTableColumn(
FeeLinesAdjustedTables,
"AllFeeLines",
List.RemoveItems(Table.ColumnNames(FeeLinesSetAcct), {"Posting Date"})
)
```

### What it does
For each posting date:
- compares the sum of rounded fee lines to the day’s target fee
- calculates the difference caused by rounding
- adds that difference to the **first fee row only**

### Why it matters
This is a solid mechanism to ensure:

```text
sum(adjusted daily fee lines) = target daily fee
```

### Risks / gotchas
- The first row of the day takes the rounding adjustment, which is fine mathematically but can look arbitrary to users.
- If ordering within a day changes, a different row will absorb the rounding correction.

---

## Block 24 – Separate final income and fee row sets

### Code

```m
IncomeLines = Table.RemoveColumns(
#"Renamed Account",
{"Gift subtype", "Pay method", "All"}
),

FeeLinesFinal = Table.RemoveColumns(
FeeLinesAdjusted,
{"Gift subtype", "Pay method", "All", "AbsIncome", "DailyAbsIncome", "ShareOfDay", "Fee_Unrounded", "Fee_Rounded"}
)
```

### What it does
Creates two cleaner datasets:
- `IncomeLines`
- `FeeLinesFinal`

### Why it matters
This separates the main line types before building orphan fees and before final combine.

### Risks / gotchas
- `IncomeLines` comes from `Renamed Account` (post-group / post-second VAT merge), which is correct.
- If someone changes upstream naming or grouping, these line sets may no longer align structurally.

---

## Block 25 – Build orphan fee dates (fees without same-day income)

### Code

```m
OrphanFeeDates = Table.NestedJoin(
DailyFees,
{"disbursement_date"},
DailyIncome,
{"Posting Date"},
"IncomeMatch",
JoinKind.LeftAnti
),

OrphanFeeDatesOnly = Table.SelectColumns(OrphanFeeDates, {"disbursement_date", "TargetFee"})
```

### What it does
Finds fee dates that exist in `DailyFees` but **do not** exist in `DailyIncome`.

### Why it matters
Without this logic, fee-only days would be lost from the journal.

### Risks / gotchas
- This assumes orphan fee days should still be posted by copying a template fee row, which is reasonable but introduces synthetic behaviour.

---

## Block 26 – Build orphan fee rows using nearest previous fee line as template

### Code

```m
FeeLinesForTemplate = Table.Sort(FeeLinesFinal, {{"Posting Date", Order.Ascending}}),

OrphanFeeRecords = List.Transform(
Table.ToRecords(OrphanFeeDatesOnly),
(r as record) =>
    let
        orphanDate = r[disbursement_date],
        orphanAmt = r[TargetFee],

        CandidateRows = Table.SelectRows(
            FeeLinesForTemplate,
            each [Posting Date] <> null and [Posting Date] <= orphanDate
        ),

        TemplateRow =
            if Table.RowCount(CandidateRows) > 0 then
                Table.Last(CandidateRows)
            else
                Table.First(FeeLinesForTemplate),

        NewRow =
            Record.TransformFields(
                TemplateRow,
                {
                    {"Posting Date", each orphanDate},
                    {"Document Date", each orphanDate},
                    {"Amount", each orphanAmt}
                }
            )
    in
        NewRow),

OrphanFeeLinesFinal =
if List.Count(OrphanFeeRecords) = 0 then
    #table(Table.ColumnNames(FeeLinesFinal), {})
else
    Table.FromRecords(OrphanFeeRecords, Table.ColumnNames(FeeLinesFinal))
```

### What it does
For each orphan fee date:
- finds the latest fee line on or before that date
- uses that fee row as a template
- updates posting date, document date, and amount with the orphan fee values

### Why it matters
This ensures fee-only disbursement days still generate a journal line even with no same-day income row available to allocate against.

### Risks / gotchas
- This is a **template-copy workaround**, not a pure source-driven derivation.
- If `FeeLinesFinal` is empty, `Table.First(FeeLinesForTemplate)` would fail.
- Copied dimensions (fund, department, project, etc.) come from a previous fee row, which may or may not match the true business expectation for the orphan day.

---

## Block 27 – Combine income, fee, and orphan fee rows

### Code

```m
Output = Table.Combine({IncomeLines, FeeLinesFinal, OrphanFeeLinesFinal}),

#"Removed Columns6" = Table.RemoveColumns(Output, {"TargetFee"})
```

### What it does
Combines the three main row sets into a single dataset and removes the helper field `TargetFee`.

### Why it matters
This forms the main Addition journal dataset before BC export shaping.

### Risks / gotchas
- Structural compatibility between the three tables is assumed.
- If upstream schema changes differently in one branch, combine can break or introduce nulls.

---

## Block 28 – Add temporary tag and retype dates

### Code

```m
#"Added Custom" = Table.AddColumn(#"Removed Columns6", "Custom", each "Braintree Addition"),

#"Typed Dates Again" = Table.TransformColumns(
#"Added Custom",
{
    {"Document Date", each FxToDate(_), type date},
    {"Posting Date", each FxToDate(_), type date}
}
)
```

### What it does
- Adds a temporary marker column `Custom = Braintree Addition`
- Reasserts date typing on `Document Date` and `Posting Date`

### Why it matters
Date retyping is a safe consistency check before final document-number generation.

### Risks / gotchas
- The `Custom` column is later removed and is not part of final output.

---

## Block 29 – Create Document No. and External Document No.

### Code

```m
#"Inserted Merged Column1" = Table.AddColumn(
#"Typed Dates Again",
"Document No. ",
each Text.Combine({"Addition - ", Date.ToText([Posting Date], "dd/MM/yyyy", "en-GB")}),
type text
),

#"Duplicated Column" = Table.DuplicateColumn(#"Inserted Merged Column1", "Document No. ", "Document No.  - Copy"),
#"Renamed Columns5" = Table.RenameColumns(#"Duplicated Column", {{"Document No.  - Copy", "External Document No."}, {"Fund ID", "Fund Code"}})
```

### What it does
- Creates `Document No. ` (note the trailing space in the column name) as:

```text
Addition - dd/MM/yyyy
```

- Duplicates it into `External Document No.`
- Renames `Fund ID` to `Fund Code`

### Why it matters
This is a critical step because it defines the outward-facing document identifier used in the final output.

### Risks / gotchas
- There is already an earlier `Fund Code` from the `FUND PQ` merge. Renaming `Fund ID` to `Fund Code` here may effectively replace the earlier meaning.
- The column name `Document No. ` includes a trailing space, which is easy to miss and can cause developer confusion later.

---

## Block 30 – Overwrite Description based on Account No.

### Code

```m
#"Added Custom1" = Table.AddColumn(
#"Renamed Columns5",
"Custom.1",
each if Text.StartsWith(Text.From([#"Account No."]), "10") then "ER Braintree Addition - Income"
else if Text.StartsWith(Text.From([#"Account No."]), "25") then "ER Braintree Addition - Fees"
else "Braintree Addition"
),

#"Renamed Columns6" = Table.RenameColumns(#"Added Custom1", {{"Custom.1", "Description"}})
```

### What it does
Creates a **new `Description`** based on `Account No.` prefix:
- accounts starting with `10` → `ER Braintree Addition - Income`
- accounts starting with `25` → `ER Braintree Addition - Fees`
- otherwise → `Braintree Addition`

### Why it matters
This overwrites earlier narrative descriptions and creates a cleaner final journal description standard.

### Risks / gotchas
- Earlier `Description` logic built from source descriptions is effectively discarded here.
- Prefix-based logic assumes numerically meaningful account naming.

---

## Block 31 – Add BC-required static / blank fields

### Code

```m
#"Removed Columns7" = Table.RemoveColumns(#"Renamed Columns6", {"Custom"}),
#"Added Custom2" = Table.AddColumn(#"Removed Columns7", "Account Type", each "G/L Account"),
#"Added Custom3" = Table.AddColumn(#"Added Custom2", "Document Type", each null),
#"Added Custom4" = Table.AddColumn(#"Added Custom3", "Currency", each null),
#"Added Custom5" = Table.AddColumn(#"Added Custom4", "Major Appeal Code", each "NA"),
#"Added Custom6" = Table.AddColumn(#"Added Custom5", "Bal. Account No.", each null),
#"Added Custom7" = Table.AddColumn(#"Added Custom6", "Bal. Account Type", each "BANK ACCOUNT")
```

### What it does
Adds the standard final BC-style columns:
- `Account Type = G/L Account`
- `Document Type = null`
- `Currency = null`
- `Major Appeal Code = NA`
- `Bal. Account No. = null`
- `Bal. Account Type = BANK ACCOUNT`

### Why it matters
This shapes the Addition rows into the same general schema expected by the final journal output.

### Risks / gotchas
- `Document Type` and `Currency` are intentionally blank.
- `Bal. Account Type` is not blank here; it is set to `BANK ACCOUNT`, unlike in your other journal where these fields were forced null.
- So the final addition output behaves differently to the Daily Banking journal in this respect.

---

## Block 32 – Arrange BC output columns and append Bank Journal

### Code

```m
#"Reordered Columns1" = Table.ReorderColumns(...),

#"Appended Query" = Table.Combine({#"Reordered Columns1", #"Bank Journal"})
```

### What it does
- Reorders the Addition output columns into BC layout
- Appends rows from `Bank Journal`

### Why it matters
This is the point where the query transitions from “Addition-only” rows into the **full final journal**, including bank-side rows supplied by another query.

### Risks / gotchas
- Structural alignment with `Bank Journal` is essential.
- If `Bank Journal` columns drift from this schema, the append may produce nulls or misaligned data.

---

## Block 33 – Reconcile two Document No. sources after append

### Code

```m
#"Reordered Columns2" = Table.ReorderColumns(
#"Appended Query",
{"Posting Date", "Amount", "Document Date", "Description", "Document No. ", "Document No.", "External Document No.", "Account No.", "Appeal Code", "Project Code", "Fund Code", "Department Code", "Account Type", "Document Type", "Currency", "Year Code", "Major Appeal Code", "Bal. Account No.", "Bal. Account Type"}
),

#"Duplicated Column1" = Table.AddColumn(#"Reordered Columns2", "Document No.  - Copy", each [#"Document No. "], type text),
#"Removed Columns8" = Table.RemoveColumns(#"Duplicated Column1", {"Document No.  - Copy"}),
#"Renamed Columns7" = Table.RenameColumns(#"Removed Columns8", {{"Document No. ", "Document No.1"}, {"Document No.", "Document No.2"}}),

#"Added Custom8" = Table.AddColumn(
#"Renamed Columns7",
"Document No.",
each if [Document No.1] <> null then [Document No.1] else [Document No.2]
),

#"Removed Columns9" = Table.RemoveColumns(#"Reordered Columns3", {"Document No.1", "Document No.2"})
```

### What it does
After appending `Bank Journal`, the query has two possible document-number columns:
- `Document No. ` (from Addition rows)
- `Document No.` (likely from Bank Journal)

It reconciles them into one final `Document No.` by taking:
- `Document No.1` if present
- otherwise `Document No.2`

### Why it matters
This is a smart way to reconcile schemas between appended sources that do not use the same document-number column naming.

### Risks / gotchas
- The use of `Document No. ` with trailing space vs `Document No.` is easy to miss.
- This logic depends on only one of those fields being populated per row.
- If both are populated but conflicting, the `Document No.1` value wins.

---

## Block 34 – Final shape, add Line No., enforce dates, remove zeros

### Code

```m
#"Added Index" = Table.AddIndexColumn(#"Reordered Columns4", "Index", 1, 1, Int64.Type),
#"Renamed Columns8" = Table.RenameColumns(#"Added Index", {{"Index", "Line No."}}),

#"Changed Type2" = Table.TransformColumns(
#"Reordered Columns5",
{
    {"Document Date", each FxToDate(_), type date},
    {"Posting Date", each FxToDate(_), type date}
}
),
#"Filtered Rows" = Table.SelectRows(#"Changed Type2", each ([Amount] <> 0))
```

### What it does
- Creates a sequential `Line No.` starting at 1
- Performs one final date normalisation pass
- Removes zero-amount rows

### Why it matters
This is the final quality-control and output step before returning the completed journal.

### Risks / gotchas
- If bank or addition logic legitimately produces a zero row that should be retained, it will be removed.
- Final date normalisation is a good safety net and helps catch type drift.

---

## Block 35 – Final output

### Code

```m
in
#"Filtered Rows"
```

### What it does
Returns the fully transformed Addition journal table.

### Why it matters
This is the table users consume or downstream queries import.

### Risks / gotchas
- Whatever schema and values exist at this point are the final deliverable.
- Any temporary assumptions earlier in the query only matter if they survive to here.

---

# 4. Functional summary of the most important logic

## Posting Date logic
- Source `Gift Date` becomes `Document Date`
- `Braintree PQ.Disbursement Date` becomes `Posting Date`

## Income account mapping
- `Gift subtype` is mapped through `VAT PQ` to derive `Account No.`

## Department mapping
- `Department Code Description` is mapped through `DEP PQ`

## Fund logic
- There is an early fund mapping via `FUND PQ`
- But later `Fund ID` is renamed to `Fund Code`, which may override or bypass the earlier mapped fund meaning

## Fee logic
- Fees come from `PP Fees PQ`
- Daily fees are proportionally allocated across grouped income rows
- Rounding differences are corrected on the first fee row of each day
- Fee-only days are handled by copying the nearest previous fee row as a template

## Final append
- Addition rows are appended to `Bank Journal`
- Two document-number fields are reconciled into one final `Document No.`

---

# 5. Biggest things to watch in this query

## 1. Fund Code ambiguity
There are two competing ideas of fund here:
- `FUND PQ.Fund` renamed to `Fund Code`
- later `Fund ID` renamed to `Fund Code`

That is one of the most important review points if output looks wrong.

## 2. Posting Date dependency on Braintree match
If `Gift ID` does not match `Transaction ID`, posting date becomes null and many downstream groupings / descriptions / fees can misbehave.

## 3. Fee allocation rounding adjustment
The first fee row of the day absorbs the rounding correction. This is normal, but users may ask why one row differs slightly.

## 4. Orphan fee template-copy logic
This is pragmatic, but it is not fully source-pure.
Dimensions on orphan fee rows are copied from the nearest previous fee row.

## 5. Trailing-space column name
`Document No. ` has a trailing space. This is easy to miss and can cause maintenance confusion.

## 6. Final Description overwrite
Earlier descriptive text is discarded and replaced by account-prefix-based descriptions.
That may be intentional, but it is worth knowing.

---

# 6. Recommended future improvements

## Improvement 1 – Resolve Fund Code ambiguity
Decide whether final `Fund Code` should come from:
- `FUND PQ.Fund`, or
- `Fund ID`

Then keep only one approach.

---

## Improvement 2 – Standardise all join keys like Gift subtype
`Gift subtype` is well normalised; other lookup keys should ideally follow the same pattern.

Suggested pre-merge pattern:

```m
Text.Upper(Text.Trim(Text.Clean(Text.From(_))))
```

---

## Improvement 3 – Make orphan fee logic more defensive
Add a guard in case `FeeLinesFinal` is empty before calling `Table.First(FeeLinesForTemplate)`.

---

## Improvement 4 – Comment the account-based Description override
Because earlier descriptions are overwritten later, add a comment to make that explicit.

---

## Improvement 5 – Align Addition vs Bank Journal schema earlier
If possible, standardise `Document No.` column names before append to avoid the later reconciliation dance.

---

# 7. Short developer summary

> The Addition journal query takes source gift data, attaches Braintree disbursement dates, groups income lines, allocates Braintree fees proportionally by day, creates orphan fee rows where needed, appends bank journal rows, reconciles document numbering, and outputs a final Business Central-ready journal.

---

# 8. Suggested handover note for future maintainers

If this query starts producing wrong numbers or missing rows, check these first:

1. **`Gift ID` ↔ `Braintree PQ[Transaction ID]` match quality**
2. **`VAT PQ`, `DEP PQ`, and `FUND PQ` lookup integrity**
3. **Whether posting dates are null after Braintree expansion**
4. **Daily fee totals vs grouped income totals**
5. **Whether orphan fee rows are being templated from sensible prior rows**
6. **Whether `Fund Code` is coming from the intended source**
7. **Document number reconciliation after appending `Bank Journal`**
