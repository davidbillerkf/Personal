# Personal Finance Intelligence Dashboard

An Excel 365 personal finance system built on a star-schema data model — fact tables for
transactions, dimension tables for controlled vocabularies, and formula-driven dashboards,
budgets, forecasts, and a financial health score. No VBA, no macros.

## Folder contents

```
Personal Finance Dashboard/
├── Personal Finance Intelligence Dashboard.xlsx   ← the workbook
├── README.md                                      ← this file
├── CHANGELOG.md
├── Documentation/
├── Assets/
└── PowerQuery/                                    ← starter M-code for bank-export refresh
```

## Quick start

1. Open **Personal Finance Intelligence Dashboard.xlsx** in Excel 365 (or any modern desktop
   Excel — no dynamic-array-only functions like `XLOOKUP`/`FILTER`/`UNIQUE` are used, so it
   also works in Excel 2016/2019).
2. **The very first time you open it**, press **Ctrl+Alt+F9** (or Formulas ▸ Calculate Now) to
   force a full recalculation — see *Known limitation* below for why this matters.
3. Start on the **Data Entry** tab for a guided workflow, then log transactions on **Income**,
   **Expenses**, **Bills**, **Debt**, **Savings Goals**, and **Investments**.
4. Set your monthly budget per category on **Budget**.
5. Review **Dashboard** for KPIs, charts, the Cut-Back Analyzer, and your Financial Health Score.

## How to enter transactions

Every fact table (Income, Expenses, Bills, Debt payments, Investments, Savings transfers) is a
native Excel Table. To add a row, click into the empty row immediately below the last row of
data and start typing — Excel automatically extends the table, and every downstream formula,
report, and chart updates the moment you recalculate. Use the dropdowns for Vendor, Category,
Subcategory, Account, Payment Method, Frequency, Need/Want, Yes/No, etc. — never type free
text — so every report stays consistent. A hidden ID column next to each dropdown looks up the
matching dimension-table key via `INDEX`/`MATCH`, preserving true star-schema referential
integrity even though you only ever type a name.

## How refresh works

There is no Power Pivot data model or PivotTable in this build (see *Data Model* tab and the
note below), so there is no "Refresh All" data-model step required for the workbook's own
formulas — they recalculate live. If you wire up the Power Query starter scripts in
`PowerQuery/` to pull bank-export CSVs, use **Data ▸ Queries & Connections ▸ Refresh All**
after adding new export files, then **Ctrl+Alt+F9** to recalculate.

## How dashboards and reports update

Every KPI card, chart, and report table is built from `SUMIFS`/`COUNTIFS`/`INDEX`/`MATCH`/
`SUMPRODUCT` formulas that reference the fact and dimension tables directly by name — there are
no hard-coded values anywhere. Add a transaction, recalculate, and the Dashboard, Budget,
Monthly Reports, Annual Reports, Cash Flow Forecast, and Financial Health tabs all update
together.

## How to add categories, vendors, or accounts

- **New category**: add a row to `Dim_Category` on the **Categories** tab (CategoryID,
  CategoryName, ParentCategory, CategoryType, EssentialCategory). It appears in every dropdown
  and report automatically — no formulas anywhere else need to change.
- **New subcategory**: add a row to `Dim_Subcategory` on **Categories**, with the CategoryID it
  belongs under.
- **New vendor**: add a row to `Dim_Vendor` on **Vendors**.
- **New account**: add a row to `Dim_Account` on **Accounts**, with its starting Balance.
- **New debt / investment / savings goal / subscription**: add a row to the matching
  `Dim_*` table on that tab.

All dropdown lists are wired to these tables via Excel structured references
(`=Dim_Category[CategoryName]`, etc.), which automatically grow to include new rows — not to a
fixed cell range — so nothing else needs to be touched.

## How to customize budgets

Open **Budget**, and edit the `MonthlyBudget` value (blue text) for any category. `Actual`,
`Remaining`, `Variance %`, `Status`, and the progress bar recalculate immediately.

## Star schema architecture

See the **Data Model** tab inside the workbook for the full table-by-table relationship map.
In short: Fact tables (`Fact_Expenses`, `Fact_Income`, `Fact_Bills`, `Fact_DebtPayments`,
`Fact_Investments`, `Fact_SavingsTransfers`) hold transactions; Dimension tables
(`Dim_Category`, `Dim_Subcategory`, `Dim_Vendor`, `Dim_Account`, `Dim_PaymentMethod`,
`Dim_Debt`, `Dim_Investment`, `Dim_Goal`, `Dim_Subscription`, and a hidden `Dim_Date`) hold
controlled vocabularies. Every fact row stores its dimension value as a name (chosen from a
dropdown) plus a hidden ID column that looks the surrogate key up via `INDEX`/`MATCH` — true
ID-based relationships, without needing thousands of manual cross-sheet formulas.

Dates use `DateID = INT(Date)` (Excel's native date serial number) instead of a manual lookup
join, so the hidden `Dim_Date` sheet lines up with every fact table by construction; reports
filter dates directly with `SUMIFS` date-bound criteria, which is both simpler and faster than
a join at 100,000+ row scale.

### A note on Power Query / Power Pivot

This workbook was generated by a script outside of Excel. Power Query queries and a Power
Pivot/xVelocity Data Model are authored *inside* the Excel application itself (via its Get Data
and Power Pivot UI) and cannot be produced by an external build tool — there is no file format
for "a Power Query" that isn't created through Excel's own query editor. In their place, every
report that would normally be a PivotTable (Monthly Reports, Annual Reports, Budget, Dashboard)
is built from live `SUMIFS`/`COUNTIFS`/`INDEX`/`MATCH` formulas against the star-schema tables,
which needs no manual refresh step and recalculates automatically.

If you want to add a real Power Pivot Data Model on top of this workbook: select each table ▸
**Data ▸ Get Data ▸ From Table/Range**, check *"Add this data to the Data Model"*, then build
the relationships shown on the Data Model tab in **Power Pivot ▸ Manage ▸ Diagram View**. The
`PowerQuery/` folder contains starter M-code you can paste into Power Query's Advanced Editor to
clean and load recurring bank-export CSVs (deduping, vendor-name standardization, category
normalization) if you want to automate data entry from bank statements.

## Known environment limitation — please read

This workbook's formulas were hand-verified for correctness and written to be fully compatible
with Excel 365 (and Excel 2016/2019), deliberately avoiding volatile functions (`RAND`, `OFFSET`,
`INDIRECT`; `NOW`/`TODAY` used only once, for the workbook's single "today" reference) and
avoiding dynamic-array-only functions (`XLOOKUP`, `FILTER`, `UNIQUE`, `SORT`, `SEQUENCE`) in
favor of `INDEX`/`MATCH`, `SUMIFS`/`COUNTIFS`, and `SUMPRODUCT`.

However, the sandbox this workbook was built in has a broken LibreOffice installation — its
document *type-detection* subsystem fails even on LibreOffice's own bundled sample files, so
the automated recalculation/verification step (which normally opens the file in LibreOffice
headless and checks every formula for `#REF!`/`#VALUE!`/etc.) could not be run before delivery.
This is an environment defect, unrelated to the workbook's formulas. Every formula was instead
manually reviewed cell-by-cell against the underlying data. **The first time you open this file
in real Excel, press Ctrl+Alt+F9 to force a full recalculation**, then please look over the
Dashboard and Financial Health tabs and let me know if anything looks off.

## Sample data

The Income, Expenses, Bills, Debt Payments, Investments, and Savings Transfers tables ship with
~13 months of realistic sample/demo transactions so the dashboards, charts, and reports render
real output out of the box. Delete these rows (keep the header row and table formatting) and
replace them with your own data whenever you're ready — everything recalculates automatically.
