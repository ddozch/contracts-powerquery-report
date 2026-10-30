// Loaded to tblIssues on the Data Issues sheet: one row per problem cell in the source.
// #table with an explicit type keeps the four typed columns when there are no issues.
let
    Rows = List.Combine(ParsedRows[Issues]),
    Result = #table(
        type table [#"Contract No" = nullable text, Field = text, Problem = text, #"Original Value" = nullable text],
        Rows)
in
    Result
