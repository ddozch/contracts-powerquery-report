// Cell value -> clean text: non-breaking spaces and line breaks become spaces, runs of
// whitespace collapse to one, ends are trimmed. Blank gives null, so "blank" means the
// same thing in every query.
(value as any) as nullable text =>
let
    AsText =
        if value = null then ""
        else if value is text then value
        // Fixed formats: the text must not depend on the machine locale
        else if value is number then Number.ToText(value, "G", "en-US")
        else if value is date then Date.ToText(value, "yyyy-MM-dd", "en-US")
        else if value is datetime then DateTime.ToText(value, "yyyy-MM-dd HH:mm:ss", "en-US")
        else if value is logical then (if value then "TRUE" else "FALSE")
        else Text.From(value),
    // Every Unicode white-space character, so thin and other typographic spaces collapse too
    Words = Text.SplitAny(
        AsText,
        " #(tab)#(cr)#(lf)#(000B)#(000C)#(0085)#(00A0)#(1680)"
            & "#(2000)#(2001)#(2002)#(2003)#(2004)#(2005)#(2006)#(2007)#(2008)#(2009)#(200A)"
            & "#(2028)#(2029)#(202F)#(205F)#(3000)"
    ),
    Cleaned = Text.Combine(List.Select(Words, each _ <> ""), " ")
in
    if Cleaned = "" then null else Cleaned
