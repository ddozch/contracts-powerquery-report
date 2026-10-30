// Loaded to tblContracts on the Contracts sheet. Types are set explicitly so the table
// keeps its columns and formats when the source has no records.
let
    Source = Table.SelectColumns(ParsedRows, {"ContractNo", "Client", "ContractType", "Manager", "Drafted", "Signed", "ValueEUR"}),
    Renamed = Table.RenameColumns(Source, {{"ContractNo", "Contract No"}, {"ContractType", "Contract Type"}, {"ValueEUR", "Value EUR"}}),
    // Month columns let the report match whole months with a plain "=" comparison
    DraftedMonth = Table.AddColumn(Renamed, "Drafted Month", each if [Drafted] = null then null else Date.StartOfMonth([Drafted])),
    SignedMonth = Table.AddColumn(DraftedMonth, "Signed Month", each if [Signed] = null then null else Date.StartOfMonth([Signed])),
    Status = Table.AddColumn(SignedMonth, "Status", each if [Signed] = null then "Pending" else "Signed"),
    Typed = Table.TransformColumnTypes(
        Status,
        {
            {"Contract No", type text}, {"Client", type text}, {"Contract Type", type text}, {"Manager", type text},
            {"Drafted", type date}, {"Signed", type date}, {"Value EUR", type number},
            {"Drafted Month", type date}, {"Signed Month", type date}, {"Status", type text}
        },
        "en-US")
in
    Typed
