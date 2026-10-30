// Path of the contracts register, read from the named cell SourcePath (Report!C4),
// so users can switch files without opening the Power Query editor.
// A relative path such as data\contracts_register.xlsx is read from the folder of this
// workbook (named cell WorkbookFolder, a CELL("filename") formula on the hidden Lists sheet),
// so a downloaded copy of the project refreshes without editing C4.
// A function, so the cells are read inside SourceRows, the query that opens the file: Excel's
// formula firewall stops a query that opens a file with a path taken from another query.
() as text =>
let
    Cell =
        try Excel.CurrentWorkbook(){[Name = "SourcePath"]}[Content]{0}[Column1]
        otherwise error Error.Record("Contracts source", "The named cell SourcePath is missing. It should be Report!C4."),
    // "Copy as path" in Explorer adds quotes, so they are trimmed too
    Path = if Cell = null then "" else Text.Trim(Text.From(Cell), {" ", """", "#(tab)", "#(00A0)"}),
    // A drive letter, a UNC share or a web address; a relative Windows path has no colon
    IsAbsolute = Text.StartsWith(Path, "\\") or Text.Contains(Path, ":"),
    Folder =
        let
            Value = try Excel.CurrentWorkbook(){[Name = "WorkbookFolder"]}[Content]{0}[Column1] otherwise null,
            Dir = if Value = null then "" else Text.From(Value)
        in
            if Dir = "" then
                error Error.Record("Contracts source", "Report!C4 holds a relative path, but the folder of this workbook is unknown. Save the workbook, or type the full path of the register in Report!C4.")
            else if Text.StartsWith(Dir, "http") then
                error Error.Record("Contracts source", "Report!C4 holds a relative path, but this workbook is opened from OneDrive or SharePoint, where its folder is a web address. Type the full path of the register in Report!C4.")
            else
                Dir,
    Result =
        if Path = "" then
            error Error.Record("Contracts source", "No source file. Type the path of the contracts register in Report!C4.")
        else if IsAbsolute then
            Path
        else
            Folder & Path
in
    Result
