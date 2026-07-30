# Power Query starter scripts (optional — advanced)

**Most people don't need this folder.** The workbook's **Bank Import** tab already lets you
paste a bank/credit-card CSV export and get it cleaned and categorized with no setup at all —
see the main README for that workflow.

The `.pq` files here are **M code** for a more hands-off alternative: point Excel at a folder,
and every CSV you drop in gets deduplicated and cleaned automatically on **Data ▸ Refresh All**,
without you pasting anything by hand. They are starting points, not plug-and-play connections —
a Power Query connection is stored inside the workbook's own file format and can only be created
through Excel's Get Data UI, so paste this M code in rather than expecting it to run standalone.

## How to use

1. In Excel, go to **Data ▸ Get Data ▸ Launch Power Query Editor** (or **Get Data ▸ From
   File ▸ From Folder** if you export one CSV per statement into a single folder).
2. Create a new **Blank Query**, then open **Home ▸ Advanced Editor**.
3. Paste in the contents of `clean_expenses_import.pq` (or `clean_vendor_names.pq`), replacing
   the placeholder file path with your export folder/file.
4. Click **Done**, then **Close & Load To...** ▸ *Only Create Connection* (or load to a new
   sheet if you want to review the cleaned rows before copying them into `Fact_Expenses`).
5. Whenever you have new bank exports, drop them in the watched folder and click
   **Data ▸ Refresh All**.

## Files

- `clean_expenses_import.pq` — reads a folder of CSV exports, removes duplicate transactions,
  standardizes date/amount types, and trims/uppercases vendor text for consistent matching
  against `Dim_Vendor`.
- `clean_vendor_names.pq` — a reusable standardization step: strips trailing store numbers
  ("AMAZON.COM*1A2B3" → "AMAZON.COM"), common suffixes, and extra whitespace so the same
  merchant always maps to the same `Dim_Vendor` row.
