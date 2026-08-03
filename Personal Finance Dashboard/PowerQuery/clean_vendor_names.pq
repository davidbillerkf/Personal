// Power Query M — reusable vendor-name standardization function.
// Paste into a new Blank Query's Advanced Editor, name the query "fnCleanVendorName",
// then call it as an Invoke Custom Function step against your VendorRaw column:
//   Table.AddColumn(PreviousStep, "VendorName", each fnCleanVendorName([VendorRaw]))
//
// Strips trailing POS/store reference numbers, common suffixes, and repeated
// whitespace so the same merchant always resolves to the same Dim_Vendor row
// (e.g. "AMAZON.COM*1A2B3C4D5" and "AMAZON.COM*9Z8Y7X6" both -> "AMAZON.COM").
let
    fnCleanVendorName = (rawName as text) as text =>
    let
        Upper = Text.Upper(Text.Trim(rawName)),

        // Drop anything after a "*", "#", or long digit run — these are usually
        // store numbers, terminal IDs, or reference codes appended by the processor.
        SplitOnStar = Text.BeforeDelimiter(Upper, "*"),
        SplitOnHash = Text.BeforeDelimiter(SplitOnStar, "#"),

        // Remove common noise suffixes some processors append
        RemoveSuffixes = List.Accumulate(
            {" POS", " DEBIT", " PURCHASE", " ONLINE", " WEB ID:"},
            SplitOnHash,
            (state, suffix) => Text.Replace(state, suffix, "")
        ),

        // Collapse repeated whitespace to a single space
        CollapseSpaces = Text.Combine(
            List.Select(Text.Split(RemoveSuffixes, " "), each Text.Trim(_) <> ""),
            " "
        ),

        Final = Text.Trim(CollapseSpaces)
    in
        Final
in
    fnCleanVendorName
