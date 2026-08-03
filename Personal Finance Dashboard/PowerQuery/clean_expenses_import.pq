// Power Query M — paste into Home > Advanced Editor on a new Blank Query.
// Reads every CSV in a folder (one file per bank/credit-card export), standardizes
// columns, removes duplicate transactions, and prepares rows ready to paste into
// the Expenses tab's Fact_Expenses table.
//
// Replace the folder path below with wherever you save your bank exports.
let
    SourceFolder = "C:\Finance\BankExports",
    Source = Folder.Files(SourceFolder),
    FilterCsv = Table.SelectRows(Source, each Text.EndsWith([Name], ".csv")),
    CombinedBinaries = Table.SelectColumns(FilterCsv, {"Content", "Name"}),
    ParsedFiles = Table.AddColumn(CombinedBinaries, "Data",
        each Csv.Document([Content], [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv])),
    ExpandedData = Table.ExpandTableColumn(ParsedFiles, "Data",
        {"Column1", "Column2", "Column3", "Column4"},
        {"RawDate", "RawDescription", "RawAmount", "RawCategory"}),
    RemoveFileNameCol = Table.RemoveColumns(ExpandedData, {"Content", "Name"}),

    // Standardize types
    ChangedTypes = Table.TransformColumnTypes(RemoveFileNameCol, {
        {"RawDate", type date},
        {"RawAmount", type number}
    }),

    // Trim / clean the vendor description text
    CleanedVendor = Table.TransformColumns(ChangedTypes, {
        {"RawDescription", each Text.Trim(Text.Upper(_)), type text}
    }),

    // Absolute value for Amount (bank exports often show expenses as negative)
    AbsAmount = Table.TransformColumns(CleanedVendor, {
        {"RawAmount", each Number.Abs(_), type number}
    }),

    Renamed = Table.RenameColumns(AbsAmount, {
        {"RawDate", "Date"},
        {"RawDescription", "VendorRaw"},
        {"RawAmount", "Amount"},
        {"RawCategory", "BankCategory"}
    }),

    // Remove exact-duplicate transactions (same date + vendor text + amount),
    // which happens when the same statement is exported twice.
    Deduplicated = Table.Distinct(Renamed, {"Date", "VendorRaw", "Amount"}),

    // Drop rows Power Query couldn't parse a date/amount for (blank export rows, footers)
    Validated = Table.SelectRows(Deduplicated, each [Date] <> null and [Amount] <> null),

    Sorted = Table.Sort(Validated, {{"Date", Order.Ascending}})
in
    Sorted
