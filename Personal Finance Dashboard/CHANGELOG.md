# Changelog

All notable changes to the Personal Finance Planner are documented here.

## [2.3.0] - 2026-08-03

### Added — new Trends tab
- **Monthly Summary**: Income/Expenses/Net Savings/Cumulative for every month from Jan 2024
  through Dec 2027 (48 months — fixed range with growth room), plus a full-history Income vs
  Expenses line chart.
- **Category × Month matrix**: every category (padded to 60 rows for future growth), one column
  per month, with a color-scale heatmap.
- **Category trend picker**: dropdown + line chart for any single category's monthly spend.
- **Vendor Summary**: every vendor (padded to 220 rows), with Total (all-time), YTD, Avg Monthly
  (trailing 12 months), transaction count, and last transaction date — sort by clicking a column
  header.
- **Vendor trend picker**: dropdown + line chart for any single vendor's monthly spend.
- All of the above combine Fact_Expenses/Fact_Income with BankImportRaw (netting Returns against
  Expenses), matching the existing combined-source convention used by the Dashboard/Budget/
  Vendors totals.
- Applied directly to the user's in-progress workbook (which already had ~730 real transactions
  entered) rather than regenerated from the template, so no data was lost.

## [2.2.0] - 2026-07-30

### Added
- Bank Import's TransactionType (column H) now has a third option, **Return**, alongside
  Expense/Income (still a dropdown you can override the sign-based guess with). A Return
  reduces that category's and vendor's spend everywhere it's counted — Budget Actual, Vendor
  CurMonthSpend/YTDSpend, Dashboard's Monthly Expenses KPI, the 13-month trend chart, Need vs
  Want %, and the Cut-Back Analyzer — instead of counting as income or double-counting as a
  second expense. Every formula that previously added `BankImportRaw` rows where
  `TransactionType="Expense"` now also subtracts rows where `TransactionType="Return"`.

## [2.1.0] - 2026-07-30

### Changed — Bank Import now feeds every report directly, no copy step
Previously, Bank Import was a scratch/staging table: you'd paste a CSV, categorize it, then
manually copy the finished rows into Expenses/Income and clear the staging rows. That meant
nothing you pasted into Bank Import ever showed up on the Dashboard until you did that copy —
confusing, and not what "import" implies. Bank Import is now a **permanent ledger** for
bank-derived transactions, on equal footing with Expenses/Income for manually entered ones:

- Every total that previously read only `Fact_Expenses`/`Fact_Income` now also sums matching
  `BankImportRaw` rows (filtered by `TransactionType`): Dashboard's Monthly Income/Expenses KPIs,
  the 13-month `ChartData` trend (feeds the Income vs Expenses and Spending Trend charts),
  `Budget`'s Actual/Status/Remaining per category (feeds the Spending Breakdown pie, Budget
  Performance chart, and Budget Status widget), `Dim_Vendor`'s CurMonthSpend/YTDSpend (feeds the
  Largest Vendors widget), Need vs Want %, and the Cut-Back Analyzer's four category rows.
- Rows are no longer copied elsewhere or deleted after import — deleting a Bank Import row now
  removes that transaction from every report, the same as deleting a row from Expenses would.
- The Vendor auto-fill ("remembers" a description from a past import) no longer depends on
  Description history copied into Fact_Expenses/Fact_Income (that copy step no longer exists).
  It now looks at earlier rows within Bank Import itself, using a range that only ever includes
  prior rows (never the current row) to avoid a circular reference.
- **Known scope limit, documented in-sheet and in the README**: the Dashboard's "Top Expenses"
  and "Recent Transactions" widgets still only read `Fact_Expenses` — ranking individual
  transactions across two separate tables without dynamic-array formulas is fragile to get right
  without a live Excel to verify against, so this was intentionally left out rather than shipped
  half-working.

## [2.0.1] - 2026-07-30

### Fixed
- **Every dropdown in the workbook was broken.** `add_list_validation` (in `build/common.py`) was
  writing formula1 values with a leading `=` (e.g. `=VendorNameList`) straight into the raw XML.
  Excel's `<formula1>` element for data validation must not include that `=` (same rule as a
  cell's `<f>` formula element) — with it there, Excel silently fails to parse the validation and
  no dropdown arrow ever renders. Fixed by stripping a leading `=` in the helper, which fixes
  every dropdown in the workbook (Vendor, Category, Account, Payment Method, Need/Want, Yes/No,
  Source, Goal name) in one place.

### Added
- **Bank Import now remembers vendor/category choices.** Added a `Description` column to
  `Fact_Expenses` and `Fact_Income` so the raw bank/CSV description text persists permanently once
  you copy a row over from Bank Import. The Bank Import tab's VendorName column now auto-fills
  from an exact match against that history (`Fact_Expenses[Description]` / `Fact_Income[SourceName]`
  keyed on `Fact_Income[Description]`) before falling back to a manual dropdown pick — so a
  recurring bill or subscription with identical description text only needs to be categorized
  once. Category continues to auto-fill from the resolved vendor's default category as before.
  Added an "Other Income Source" catch-all row to `Dim_Vendor` so miscellaneous deposits have
  somewhere to land in the same dropdown.

## [2.0.0] - 2026-07-29

### Changed — simplified to a lean personal planner
- Reduced from 23 sheets to 13 (11 visible + 2 hidden helpers), per user request to simplify.
- Removed: Debt, Investments, Bills, Subscriptions, Net Worth, Monthly Reports, Annual Reports,
  Cash Flow Forecast, Financial Health Score, Data Model tab, README tab (key points folded
  into Data Entry).
- Income simplified to Payroll (Dim_Source now just "Employer Payroll Inc" / "Other").
- Expenses simplified to a flat category list (no subcategories): Mortgage, Utilities, Car,
  Groceries, Insurance, Healthcare, Dining & Entertainment, Personal Care, Shopping,
  Miscellaneous.
- Savings promoted to its own top-level tab (goals + transfers + progress + KPIs), no longer
  bundled under a "Savings Goals" tab alongside Debt/Investments.
- Dashboard trimmed to 8 KPIs and 5 charts (Income vs Expenses, Spending Breakdown, Spending
  Trend, Savings Growth, Budget Performance) matching the remaining data; Cut-Back Analyzer
  simplified to categories that still exist (Dining & Entertainment, Shopping, Personal Care,
  Non-Essential Wants).

### Added
- **Bank Import tab**: paste Date/Description/Amount from a bank/credit-card CSV export
  directly into a staging table; pick a Vendor per row and Category auto-suggests from that
  vendor's default category; a Sign Convention toggle splits Expense vs Income by the amount's
  sign; ready-to-copy AbsAmount column. Works immediately, no Power Query setup required.
- "Where your data lives" guidance in the Data Entry tab and README (data is stored entirely
  inside the workbook's Excel Tables; flags the git-repo privacy consideration for real data).

### Fixed
- A layout bug where the Dashboard's "Cut-Back Analyzer" section header overwrote the last row
  of the "Budget Status" widget when the expense category count differed from the "Recent
  Transactions" widget's fixed row count; the Cut-Back Analyzer's start row is now computed from
  the taller of the two widgets above it.

## [1.0.0] - 2026-07-29

### Added
- Initial release: 23-sheet workbook built on a star schema data model (fact tables for
  Income/Expenses/Bills/Debt Payments/Investments/Savings Transfers; dimension tables for
  Category/Subcategory/Vendor/Account/Payment Method/Debt/Investment/Goal/Subscription).
- Dashboard with 12 KPIs, 8 charts, and widgets; Budget; Bills; Debt (Snowball/Avalanche); 
  Savings Goals; Investments; Net Worth; Monthly/Annual Reports; Cash Flow Forecast; Financial
  Health Score; Data Model documentation tab.
- ~13 months of sample/demo transactions; starter Power Query M-code for bank-export cleanup.

### Known limitations
- Automated LibreOffice-based formula recalculation could not be run in the build sandbox due
  to a pre-existing environment defect (LibreOffice's type-detection subsystem fails on all
  documents, including its own bundled samples). Formulas were manually reviewed instead; open
  in Excel and press Ctrl+Alt+F9 on first use. This remains true for 2.0.0.
- No native PivotTables or Power Pivot Data Model are included, since those are authored inside
  the Excel application itself and cannot be generated by an external script — equivalent
  formula-driven reports are provided instead.
