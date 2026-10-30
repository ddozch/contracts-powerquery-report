// Header cell -> matching key: lower case, Polish and German letters folded to ASCII,
// everything except a-z and 0-9 dropped. "DATA" + line break + "PODPISANIA",
// "Data podpisania" and "data_podpisania" all give "datapodpisania".
(value as any) as text =>
let
    Lower = if value = null then "" else Text.Lower(Text.From(value), "en-US"),
    Folds = {
        {"ą", "a"}, {"ć", "c"}, {"ę", "e"}, {"ł", "l"}, {"ń", "n"}, {"ó", "o"},
        {"ś", "s"}, {"ź", "z"}, {"ż", "z"}, {"ä", "a"}, {"ö", "o"}, {"ü", "u"}, {"ß", "ss"}
    },
    Folded = List.Accumulate(Folds, Lower, (t, f) => Text.Replace(t, f{0}, f{1})),
    Key = Text.Select(Folded, {"a".."z", "0".."9"})
in
    Key
