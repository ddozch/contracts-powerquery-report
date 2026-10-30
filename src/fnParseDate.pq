// Cell record [Value, Error] from SourceRows -> [Value = date or null, Problem, Original].
// Accepts real dates, datetimes (time dropped), Excel serial numbers and text in three
// digit groups split by "." "/" or "-": yyyy-m-d when the first group has 4 digits,
// d-m-yyyy (day first) when the last one has. Anything else is "Unreadable date", and so is
// a date before 1900.
(cell as record) as record =>
let
    Raw = cell[Value],
    CellText = fnCleanText(Raw),
    // Excel has no dates before 1900: such a value (a year typo like 05.03.1026, a negative
    // serial) would land in the table as a negative number and make YEAR() fail in the report
    MinDate = #date(1900, 1, 1),
    Ok = (d as date) as record => if d < MinDate then Bad else [Value = d, Problem = null, Original = CellText],
    Bad = [Value = null, Problem = "Unreadable date", Original = CellText],
    // The date is evaluated inside try here; a record field would be evaluated too late to catch
    Checked = (attempt as record) as record => if attempt[HasError] then Bad else Ok(attempt[Value]),
    IsDigits = (s as text) as logical => s <> "" and Text.Select(s, {"0".."9"}) = s,
    FromParts = (y as text, m as text, d as text) as record =>
        Checked(try #date(Number.FromText(y, "en-US"), Number.FromText(m, "en-US"), Number.FromText(d, "en-US"))),
    FromText = (t as text) as record =>
        let
            Parts = Text.SplitAny(t, "./-"),
            Lengths = List.Transform(Parts, Text.Length)
        in
            if List.Count(Parts) <> 3 or not List.AllTrue(List.Transform(Parts, IsDigits)) then Bad
            // Only the 4-digit group decides the order; #date rejects impossible parts such as day 31.02
            else if Lengths{0} = 4 then FromParts(Parts{0}, Parts{1}, Parts{2})
            else if Lengths{2} = 4 then FromParts(Parts{2}, Parts{1}, Parts{0})
            else Bad,
    Result =
        if cell[Error] <> null then [Value = null, Problem = "Error value", Original = cell[Error]]
        else if CellText = null then [Value = null, Problem = null, Original = null]
        else if Raw is date then Ok(Raw)
        else if Raw is datetime then Ok(DateTime.Date(Raw))
        else if Raw is datetimezone then Ok(DateTime.Date(DateTimeZone.RemoveZone(Raw)))
        // Excel serial date typed into a General cell; the fraction is a time of day
        else if Raw is number then Checked(try Date.From(Number.RoundDown(Raw)))
        else if Raw is text then FromText(CellText)
        else Bad
in
    Result
