// Cell record [Value, Error] from SourceRows -> [Value = number or null, Problem, Original].
// Text such as "12 500,00", "12,500.00" or "1.250,50 EUR" is read without the machine
// locale: spaces and "EUR" are removed, and when both "," and "." occur the last one is
// the decimal separator; a lone "," is a decimal comma.
(cell as record) as record =>
let
    Raw = cell[Value],
    CellText = fnCleanText(Raw),
    Bad = [Value = null, Problem = "Unreadable number", Original = CellText],
    FromText = (t as text) as record =>
        let
            // The spec removes the currency code "EUR" exactly as written
            Compact = Text.Remove(Text.Replace(t, "EUR", ""), {" ", "#(00A0)"}),
            LastComma = Text.PositionOf(Compact, ",", Occurrence.Last),
            LastDot = Text.PositionOf(Compact, ".", Occurrence.Last),
            Normalised =
                if LastComma >= 0 and LastDot >= 0 then
                    (if LastComma > LastDot then Text.Replace(Text.Remove(Compact, "."), ",", ".") else Text.Remove(Compact, ","))
                else
                    Text.Replace(Compact, ",", "."),
            // Plain number only: optional minus, digits, at most one dot with digits on both
            // sides, so "100," or ",50" (a digit probably missing) is reported, not guessed
            Digits = if Text.StartsWith(Normalised, "-") then Text.Middle(Normalised, 1) else Normalised,
            IsPlain =
                Text.Select(Digits, {"0".."9", "."}) = Digits
                and Text.Length(Text.Select(Digits, ".")) <= 1
                and Text.Select(Digits, {"0".."9"}) <> ""
                and not Text.StartsWith(Digits, ".")
                and not Text.EndsWith(Digits, "."),
            Attempt = if IsPlain then try Number.FromText(Normalised, "en-US") else null
        in
            if IsPlain and not Attempt[HasError] then [Value = Attempt[Value], Problem = null, Original = CellText]
            else Bad,
    Result =
        if cell[Error] <> null then [Value = null, Problem = "Error value", Original = cell[Error]]
        else if CellText = null then [Value = null, Problem = null, Original = null]
        else if Raw is number then [Value = Raw, Problem = null, Original = CellText]
        else if Raw is text then FromText(CellText)
        else Bad
in
    Result
