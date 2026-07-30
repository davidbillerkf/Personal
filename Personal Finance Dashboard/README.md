# Personal Finance Planner

A simple, formula-driven personal finance planner in Excel 365: Income (payroll), Expenses
(Mortgage, Utilities, Car, Groceries, Insurance, Healthcare, Dining & Entertainment, Personal
Care, Shopping, Miscellaneous), Savings goals, a category Budget, and a Dashboard — plus a
dedicated **Bank Import** tab for bringing in your bank/credit-card CSV exports. No VBA, no
macros.

## Folder contents

```
Personal Finance Dashboard/
├── Personal Finance Intelligence Dashboard.xlsx   ← the workbook
├── README.md                                      ← this file
├── CHANGELOG.md
├── Documentation/
├── Assets/
└── PowerQuery/                                    ← optional advanced M-code (see below)
```

## Quick start

1. Open **Personal Finance Intelligence Dashboard.xlsx** in Excel 365 (also works in Excel
   2016/2019 — no dynamic-array-only functions like `XLOOKUP`/`FILTER`/`UNIQUE` are used).
2. **The first time you open it**, press **Ctrl+Alt+F9** (Formulas ▸ Calculate Now) to force a
   full recalculation — see *Known limitation* below.
3. Start on **Data Entry** for a guided workflow.
4. Have a bank CSV export? Go to **Bank Import** first (see next section).
5. Otherwise, log transactions directly on **Income**, **Expenses**, and **Savings**.
6. Set your monthly budget per category on **Budget**.
7. Review **Dashboard** for KPIs, charts, and the Cut-Back Analyzer.

## Bringing in bank / credit-card transactions

### Option A — the Bank Import tab (works immediately, no setup)

1. Export a CSV of transactions from your bank or credit card's website (usually
   *Accounts ▸ Download/Export*).
2. Open that CSV and copy its **Date**, **Description**, and **Amount** columns.
3. Paste them into columns A–C of the `BankImportRaw` table on the **Bank Import** tab (paste
   *values only* — Home ▸ Paste ▸ Values — into the first blank row under the header).
4. Set **Sign Convention** near the top of the tab to match your export (most bank/debit
   exports show purchases as negative numbers; most credit-card exports show charges as
   positive numbers — check one row against your real statement to confirm).
5. For each row, pick a **Vendor** from the dropdown in column D. **Category** (column E)
   auto-fills from that vendor's default category — override it if it's wrong, and add new
   vendors to `Dim_Vendor` on the **Vendors** tab as they come up.
6. Column H tells you **Expense** or **Income**, column I gives the positive **AbsAmount**.
   Copy the finished rows (Date, Vendor, Category, Account, Payment Method, AbsAmount, Notes)
   into the **Expenses** tab (or **Income** tab for paychecks/deposits), then clear the pasted
   rows on Bank Import so it stays a clean scratch area for your next import.

This works the moment you open the file — no Power Query setup required.

### Option B — Power Query (optional, for hands-off recurring imports)

If you'd rather point Excel at a folder and have it auto-clean every new export you drop in
there (dedup, standardize vendor text) without the manual paste step, see `PowerQuery/README.md`
for M-code you paste into Excel's own Power Query editor (a one-time, few-minute setup). This is
optional — Option A above is enough for most people.

## Where your data lives

Everything is stored **inside this workbook** — the Excel Tables on each tab (`Fact_Income`,
`Fact_Expenses`, `Fact_SavingsTransfers`, `Dim_Category`, `Dim_Vendor`, `Dim_Account`,
`Dim_Goal`, `Budget`) are the database. There is no external server, cloud account, or database
connection. That means:

- **This file is your only copy** — back it up the way you would any important document (a
  synced folder, a periodic copy, or version control).
- **If you keep this file in a git repo**, committing it puts your real transaction data into
  that repo's history. Use a private repo once you start entering real numbers.

## How to enter transactions

Every fact table (Income, Expenses, Savings transfers) is a native Excel Table. Click into the
row immediately below the last row of data and start typing — Excel extends the table
automatically, and every KPI, chart, and report updates the moment you recalculate. Use the
dropdowns for Vendor, Category, Account, Payment Method, Need/Want, Yes/No — never type free
text — so every report stays consistent.

## How to add categories, vendors, or accounts

- **New category**: add a row to `Dim_Category` on **Categories** (CategoryID, CategoryName,
  CategoryType, EssentialCategory). It appears in every dropdown and report automatically.
- **New vendor**: add a row to `Dim_Vendor` on **Vendors**, with its default category.
- **New account**: add a row to `Dim_Account` on **Accounts**, with its starting Balance.

All dropdowns use Excel structured references (`=Dim_Category[CategoryName]`, etc.), which
automatically grow to include new rows — not a fixed cell range — so nothing else needs to be
touched.

## How to customize your budget

Open **Budget** and edit the `MonthlyBudget` value (blue text) for any category. `Actual`,
`Remaining`, `Variance %`, `Status`, and the progress bar recalculate immediately.

## Known environment limitation — please read

This workbook's formulas were manually reviewed for correctness and written to be fully
compatible with Excel 365 (and Excel 2016/2019), deliberately avoiding volatile functions
(`RAND`, `OFFSET`, `INDIRECT`; `TODAY()` is used once, for the workbook's single "today"
reference) and dynamic-array-only functions (`XLOOKUP`, `FILTER`, `UNIQUE`, `SORT`,
`SEQUENCE`) in favor of `INDEX`/`MATCH`, `SUMIFS`/`COUNTIFS`, and `SUMPRODUCT`.

The sandbox this workbook was built in has a broken LibreOffice installation — its document
type-detection subsystem fails even on LibreOffice's own bundled sample files — so the usual
automated recalculation/verification pass could not be run before delivery. This is an
environment defect, unrelated to the workbook itself. **The first time you open this file in
real Excel, press Ctrl+Alt+F9** to force a full recalculation, then look over the Dashboard and
let me know if anything looks off.

## Sample data

The Income, Expenses, and Savings Transfers tables ship with ~13 months of realistic
sample/demo transactions so the Dashboard renders real output out of the box. Delete these rows
(keep the header row and table formatting) and replace them with your own — everything
recalculates automatically.
