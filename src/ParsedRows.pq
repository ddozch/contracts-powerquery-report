// Raw cells -> typed values plus the list of problems found in each row. A problem cell
// becomes null, but the row is always kept, so totals never lose a contract silently.
let
    Source = SourceRows,
    Clean = (cell as record) as nullable text => fnCleanText(cell[Value]),
    CanonicalType = (cell as record) as text =>
        let
            Cleaned = Clean(cell),
            Position = if Cleaned = null then -1 else List.PositionOf(ContractTypes, Cleaned, Occurrence.First, Comparer.OrdinalIgnoreCase)
        in
            if Cleaned = null then "(no type)"
            else if Position >= 0 then ContractTypes{Position}
            else Cleaned,
    // One issue row per problem cell: {Contract No, Field, Problem, Original Value}
    IssueOf = (contractNo as nullable text, field as text, parsed as record) as list =>
        if parsed[Problem] = null then {} else {{contractNo, field, parsed[Problem], parsed[Original]}},
    ParseRow = (row as record) as record =>
        let
            contractNo = Clean(row[ContractNo]),
            manager = Clean(row[Manager]),
            drafted = fnParseDate(row[Drafted]),
            signed = fnParseDate(row[Signed]),
            value = fnParseNumber(row[ValueEUR]),
            draftedIssues =
                if drafted[Problem] = null and drafted[Value] = null then {{contractNo, "Drafted", "Missing drafted date", null}}
                else IssueOf(contractNo, "Drafted", drafted)
        in
            [
                ContractNo = contractNo,
                Client = Clean(row[Client]),
                ContractType = CanonicalType(row[ContractType]),
                Manager = if manager = null then "(no manager)" else manager,
                Drafted = drafted[Value],
                Signed = signed[Value],
                ValueEUR = value[Value],
                Issues = draftedIssues & IssueOf(contractNo, "Signed", signed) & IssueOf(contractNo, "Value (EUR)", value)
            ],
    Parsed = Table.AddColumn(Source, "Parsed", ParseRow),
    Result = Table.ExpandRecordColumn(
        Table.SelectColumns(Parsed, {"Parsed"}),
        "Parsed",
        {"ContractNo", "Client", "ContractType", "Manager", "Drafted", "Signed", "ValueEUR", "Issues"})
in
    Result
