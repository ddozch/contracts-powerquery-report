// Logical columns of the contracts register and the header names accepted for each.
// Headers are compared with fnNormalizeHeader, so case, Polish letters, spaces, line breaks
// and punctuation in the source file do not matter. Add a new spelling here, nowhere else.
#table(
    type table [Field = text, Label = text, Required = logical, Headers = list],
    {
        {"ContractNo",   "Contract No",   true,  {"Contract No", "Contract Number", "Nr umowy", "Numer umowy"}},
        {"Client",       "Client",        false, {"Client", "Customer", "Klient"}},
        {"ContractType", "Contract Type", true,  {"Contract Type", "Type", "Rodzaj umowy", "Typ umowy"}},
        {"Manager",      "Manager",       false, {"Manager", "Account Manager", "Opiekun", "Handlowiec"}},
        {"Drafted",      "Drafted",       true,  {"Drafted", "Drafted On", "Data sporządzenia", "Created"}},
        {"Signed",       "Signed",        false, {"Signed", "Signed On", "Data podpisania"}},
        {"ValueEUR",     "Value (EUR)",   false, {"Value (EUR)", "Value EUR", "Value", "Wartość (EUR)", "Wartość"}}
    }
)
