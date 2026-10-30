// Finds the contracts table in the source workbook and returns one row per record with the
// seven logical columns of HeaderMap. The table may sit on any sheet, below title rows and
// under renamed headers. Each cell is kept as [Value, Error], so an Excel error cell (#N/A)
// becomes data for the Data Issues sheet instead of failing the refresh.
let
    // A function call: the path is read inside this query (see SourcePath)
    Path = SourcePath(),
    Opened = try Excel.Workbook(File.Contents(Path), null, true),
    Book =
        if Opened[HasError] then
            error Error.Record("Contracts source", "Cannot open the source file " & Path & ". " & Opened[Error][Message])
        else
            Opened[Value],
    Sheets = Table.SelectRows(Book, each [Kind] = "Sheet"),

    // Every accepted header as {normalised name, logical field}
    Lookup = List.Buffer(List.Combine(List.Transform(
        Table.ToRecords(HeaderMap),
        (f) => List.Transform(f[Headers], (h) => {fnNormalizeHeader(h), f[Field]})))),
    LookupKeys = List.Buffer(List.Transform(Lookup, each _{0})),

    // One sheet row -> record of logical field = source column name (Column1, Column2, ...).
    // The leftmost column wins when two columns map to the same field.
    MatchRow = (row as record) as record =>
        List.Accumulate(Record.FieldNames(row), [], (found, column) =>
            let
                Key = try fnNormalizeHeader(Record.Field(row, column)) otherwise "",
                Position = if Key = "" then -1 else List.PositionOf(LookupKeys, Key),
                FieldName = if Position < 0 then null else Lookup{Position}{1}
            in
                if FieldName = null or Record.HasFields(found, FieldName) then found
                else Record.AddField(found, FieldName, column)),

    // Header row = the first of the first 15 rows that names at least 4 logical columns
    FindHeader = (sheet as record) as nullable record =>
        let
            Maps = List.Transform(Table.ToRecords(Table.FirstN(sheet[Data], 15)), MatchRow),
            Hits = List.Select(List.Positions(Maps), (i) => Record.FieldCount(Maps{i}) >= 4)
        in
            if Hits = {} then null
            else [Sheet = sheet[Name], Data = sheet[Data], HeaderIndex = Hits{0}, Columns = Maps{Hits{0}}],

    // The data sheet is the first one, in workbook order, with a header row
    Found = List.First(List.RemoveNulls(List.Transform(Table.ToRecords(Sheets), FindHeader)), null),
    RequiredLabels = Table.SelectRows(HeaderMap, each [Required])[Label],
    Located =
        if Found = null then
            error Error.Record("Contracts source",
                "No contracts table found in " & Path
                    & ". No sheet has a header row (within the first 15 rows) with at least 4 known column names."
                    & " Required columns: " & Text.Combine(RequiredLabels, ", ") & ".")
        else
            let
                Missing = Table.SelectRows(HeaderMap, (f) => f[Required] and not Record.HasFields(Found[Columns], f[Field]))[Label]
            in
                if Missing = {} then Found
                else error Error.Record("Contracts source",
                    "Required columns not found in " & Path & " (sheet """ & Found[Sheet] & """): "
                        & Text.Combine(Missing, ", ") & ". Check the header names against the HeaderMap query."),

    Fields = HeaderMap[Field],
    Body = Table.Skip(Located[Data], Located[HeaderIndex] + 1),

    // Power Query reports an Excel error cell as "Invalid cell value '#N/A'." in the UI
    // language and with local quote marks, so the code is cut out from the "#" onwards
    ErrorChars = {"A".."Z", "0".."9", "/", "!", "?", "_"},
    ErrorText = (e as record) as text =>
        let
            Message = Text.From(e[Message]),
            Start = Text.PositionOf(Message, "#"),
            Tail = if Start < 0 then "" else Text.Middle(Message, Start + 1),
            Stop = List.PositionOf(List.Transform(Text.ToList(Tail), each List.Contains(ErrorChars, _)), false),
            Code = if Stop < 0 then Tail else Text.Start(Tail, Stop)
        in
            if Code = "" then Message else "#" & Code,
    // try around the field access is what catches the error cell; optional columns may be absent
    ReadCell = (row as record, column as nullable text) as record =>
        if column = null then [Value = null, Error = null]
        else
            let
                Attempt = try Record.Field(row, column)
            in
                if Attempt[HasError] then [Value = null, Error = ErrorText(Attempt[Error])]
                else [Value = Attempt[Value], Error = null],
    WithCells = Table.AddColumn(Body, "Cells", (row) =>
        Record.FromList(
            List.Transform(Fields, (f) => ReadCell(row, Record.FieldOrDefault(Located[Columns], f, null))),
            Fields)),
    // Expanding with explicit names keeps the columns when the sheet has no data rows
    Cells = Table.ExpandRecordColumn(Table.SelectColumns(WithCells, {"Cells"}), "Cells", Fields),

    IsBlank = (cell as record) as logical => cell[Error] = null and fnCleanText(cell[Value]) = null,
    NonBlank = Table.SelectRows(Cells, (row) => not List.AllTrue(List.Transform(Fields, (f) => IsBlank(Record.Field(row, f)))))
in
    NonBlank
