#!/usr/bin/env python3
"""Generate the demo contracts register in three layouts and the figures the report must show.

Writes:
  data/variant_a/contracts_register.xlsx      clean layout: one sheet, English headers in row 1
  data/variant_b/contracts_register.xlsx      messy layout: instruction sheet first, title rows,
                                              Polish headers with line breaks, other column order
  data/variant_empty/contracts_register.xlsx  header row only
  data/contracts_register.xlsx                copy of variant B, the default source of the report
  data/expected.json                          expected report figures per variant: report year 2026
                                              with both filters on "All", the exact spellings in
                                              tblContracts, and one block per filter scenario

Every record is generated as truth first: what each cell is meant to contain (a date, blank, an
error or unreadable text). Variants A and B then write the same truth in their own mix of
representations (real dates, text dates, serial numbers, text numbers ...), so both must give the
same figures. The expected figures come from the truth with the parsing rules of the spec,
implemented here independently of the Power Query code. Two self-checks guard the generator:
the rendered cells must parse back to the truth, and the written files, read back with openpyxl,
must give the same figures. Fixed seed: the output is the same on every run.
The scenarios (report year 2025, Manager and Contract type filters, names with * ? < > or over
255 characters) check the year limits, the pending definition and each section's filter rule.

Usage:  python tools/generate_demo_data.py [--seed 2026]
"""
from __future__ import annotations

import argparse
import hashlib
import json
import math
import random
import re
import shutil
from collections import Counter
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from decimal import Decimal
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "data"
REGISTER = "contracts_register.xlsx"

REPORT_YEAR = 2026
RECORDS = 1200
FIRST_DRAFTED, LAST_DRAFTED = date(2025, 10, 1), date(2026, 10, 28)
REGISTER_DATE = LAST_DRAFTED   # the register is a snapshot: later signatures have not happened yet
MIN_EDGE = 4                   # each representation at least this often per variant and field (spec: 3)
BLANK_ROWS = 8                 # fully blank rows between records, per variant
NBSP = " "

# canonical contract type: (share, min EUR, max EUR)
TYPES = {
    "Installation contract": (0.40, 5_000, 60_000),
    "Service agreement": (0.25, 500, 8_000),
    "Maintenance contract": (0.20, 1_000, 15_000),
    "Annex": (0.15, 500, 20_000),
}
MANAGERS = ["Leo Brooks", "Mia Lane", "Nina Hart", "Owen Price", "Sam Carter"]
CLIENT_WORDS = ["Harborview", "Oakridge", "Northgate", "Riverside", "Summit", "Clearview", "Westbrook",
                "Lakeside", "Pinecrest", "Redwood", "Fairway", "Stonehill", "Brightway", "Ironbridge"]
CLIENT_TYPES = ["Properties", "Offices", "Residences", "Business Park", "Housing", "Hotels", "Campus"]
NOTES = ["Klient prosi o zmianę terminu", "Czeka na podpis zarządu", "Wysłano skan do klienta",
         "Do weryfikacji wartości", "Oryginał w segregatorze"]

# Deliberate problem cells, shared by variants A and B (record counts)
SPECIAL = dict(drafted_blank=5, drafted_error=6, drafted_bad=12, signed_error=6, signed_bad=12,
               value_blank=8, value_error=6, value_bad=12, type_blank=3, manager_blank=18)
# Unreadable cells added to each variant's own unreadable texts, in both variants:
# dates Excel cannot show (year typo, day before 1900, negative serial in a General cell) and
# numbers with a separator but a digit missing on one side
PRE_1900 = ("05.03.1026", "31.12.1899", -5)
HALF_NUMBERS = ("100,", ",50", "5.")
# Free text outside the type list that COUNTIFS would read as a pattern ("Annex*" also matches
# "Annex", "<other>" means less than "other>") or cannot take as a criterion (over 255 chars)
LONG_TYPE = ("Installation contract for the second building, see the e-mail from the property manager "
             "of 14 March: the framework price list applies, the start date is still to be agreed with "
             "the crews, and the client wants the Wi-Fi survey added later as an annex to this contract")
assert len(LONG_TYPE) > 255
OTHER_TYPES = {"Consulting": 5, "Network audit": 3, "Annex*": 3, "<other>": 3, LONG_TYPE: 3}
# Placeholders typed instead of a manager name; "?" as a pattern also matches "-"
OTHER_MANAGERS = {"?": 3, "-": 3, "<to be assigned>": 3}
PATTERN_CHARS = set("*?~<>=")
BOUNDARY_DAYS = (date(2025, 12, 31), date(2026, 1, 1))   # year limits, tested with report years 2025 and 2026
# Filter settings the test also checks, besides report year 2026 with both filters on "All"
SCENARIOS = [
    dict(name="2025, all", report_year=2025, manager="All", type="All"),
    dict(name="2025, Leo Brooks + Installation contract", report_year=2025, manager="Leo Brooks",
         type="Installation contract"),
    dict(name="2026, Mia Lane + Service agreement", report_year=2026, manager="Mia Lane", type="Service agreement"),
    dict(name="2026, (no manager)", report_year=2026, manager="(no manager)", type="All"),
    dict(name="2026, manager ?", report_year=2026, manager="?", type="All"),
    dict(name="2026, type Annex*", report_year=2026, manager="All", type="Annex*"),
    dict(name="2026, type <other>", report_year=2026, manager="All", type="<other>"),
    dict(name="2026, type over 255 chars", report_year=2026, manager="All", type=LONG_TYPE),
]

# Logical columns and the header names accepted for them (compared after normalize_header)
COLUMNS = ("ContractNo", "Client", "ContractType", "Manager", "Drafted", "Signed", "ValueEUR")
REQUIRED = ("ContractNo", "ContractType", "Drafted")
ACCEPTED = {
    "ContractNo": ["Contract No", "Contract Number", "Nr umowy", "Numer umowy"],
    "Client": ["Client", "Customer", "Klient"],
    "ContractType": ["Contract Type", "Type", "Rodzaj umowy", "Typ umowy"],
    "Manager": ["Manager", "Account Manager", "Opiekun", "Handlowiec"],
    "Drafted": ["Drafted", "Drafted On", "Data sporządzenia", "Created"],
    "Signed": ["Signed", "Signed On", "Data podpisania"],
    "ValueEUR": ["Value (EUR)", "Value EUR", "Value", "Wartość (EUR)", "Wartość"],
}
PROBLEMS = ("Error value", "Unreadable date", "Unreadable number", "Missing drafted date")
FIELD_LABELS = {"Drafted": "Drafted", "Signed": "Signed", "ValueEUR": "Value (EUR)"}

DATE_REPS = ("date", "datetime", "text dd.mm.yyyy", "text d.m.yyyy", "text yyyy-mm-dd", "text dd/mm/yyyy",
             "serial number")
VALUE_REPS = ("number", "text 12 500,00", "text 12,500.00", "text 12 500,00 EUR (nbsp)")
TYPE_REPS = ("canonical", "lower case + trailing space", "upper case + double spaces", "leading space",
             "non-breaking space")
MANAGER_REPS = ("clean", "leading/trailing spaces", "non-breaking space")

VARIANTS = {
    # Clean register: mostly typed cells, a few of every messy representation
    "a": dict(
        sheet="Register", intro=None, titles=[], date_format="yyyy-mm-dd",
        columns=[("ContractNo", "Contract No"), ("Client", "Client"), ("ContractType", "Contract Type"),
                 ("Manager", "Manager"), ("Drafted", "Drafted"), ("Signed", "Signed"),
                 ("ValueEUR", "Value (EUR)")],
        dates=dict(zip(DATE_REPS, (82, 3, 4, 2, 3, 3, 3))),
        values=dict(zip(VALUE_REPS, (90, 4, 4, 2))),
        types=dict(zip(TYPE_REPS, (94, 2, 2, 1, 1))),
        managers=dict(zip(MANAGER_REPS, (95, 3, 2))),
        bad_dates=("brak", "TBD", "31.02.2026", *PRE_1900), bad_values=("n/a", "TBC", "see annex", *HALF_NUMBERS),
        messy_text=0.0, padded_other_types=False),
    # Office register: instruction sheet, title rows, Polish headers, many text dates and numbers
    "b": dict(
        sheet="Umowy 2026", date_format="dd.mm.yyyy",
        intro=("Instrukcja", ["Rejestr umów: dane w arkuszu Umowy 2026.",
                              "Nowe umowy dopisuj na końcu tabeli, nie zmieniaj kolejności kolumn."]),
        titles=["Rejestr umów 2026", "Stan na 28.10.2026, prowadzi biuro sprzedaży"],
        columns=[("Lp", "Lp."), ("ContractNo", "Nr\numowy"), ("Drafted", "Data\nsporządzenia"),
                 ("Client", "Klient"), ("ContractType", "Rodzaj\numowy"), ("Manager", "Opiekun"),
                 ("ValueEUR", "Wartość (EUR)"), ("Signed", "DATA\nPODPISANIA"), ("Note", "Uwagi")],
        dates=dict(zip(DATE_REPS, (45, 8, 20, 6, 8, 6, 7))),
        values=dict(zip(VALUE_REPS, (60, 20, 10, 10))),
        types=dict(zip(TYPE_REPS, (80, 5, 5, 5, 5))),
        managers=dict(zip(MANAGER_REPS, (85, 10, 5))),
        bad_dates=("brak", "TBD", "31.02.2026", *PRE_1900),
        bad_values=("n/a", "do ustalenia", "TBC", *HALF_NUMBERS),
        messy_text=0.03, padded_other_types=True),
}
WIDTHS = dict(Lp=6, ContractNo=14, Client=24, ContractType=24, Manager=16, Drafted=14, Signed=14,
              ValueEUR=16, Note=32)

HEADER_FONT = Font(bold=True)
HEADER_FILL = PatternFill("solid", fgColor="DDEBF7")
TITLE_FONT = Font(bold=True, size=14)
WRAP = Alignment(wrap_text=True, vertical="top")


@dataclass(frozen=True)
class XlError:
    """An Excel error cell. openpyxl writes the string "#N/A" as a real error cell."""
    code: str = "#N/A"


# ------------------------------------------------- parsing rules (independent of the M code)

FOLD = str.maketrans({"ą": "a", "ć": "c", "ę": "e", "ł": "l", "ń": "n", "ó": "o", "ś": "s", "ź": "z",
                      "ż": "z", "ä": "a", "ö": "o", "ü": "u", "ß": "ss"})
# ASCII digits only, like the M code; \d would also accept other scripts' digits
DATE_TEXT = re.compile(r"^([0-9]+)[./-]([0-9]+)[./-]([0-9]+)$")
PLAIN_NUMBER = re.compile(r"^-?[0-9]+(\.[0-9]+)?$")   # a dot needs digits on both sides
EXCEL_EPOCH = date(1899, 12, 30)   # Excel serial 1 = 1900-01-01 for all serials after Feb 1900
MIN_DATE = date(1900, 1, 1)        # Excel cannot show earlier dates: they are unreadable


def normalize_header(v) -> str:
    return re.sub(r"[^a-z0-9]", "", str(v).lower().translate(FOLD))


HEADER_KEYS = {normalize_header(name): col for col, names in ACCEPTED.items() for name in names}
# "Value (EUR)" and "Value EUR" share a key, which is fine; a key shared by two columns is not
assert len(HEADER_KEYS) == len({(normalize_header(n), c) for c, ns in ACCEPTED.items() for n in ns}), \
    "accepted header names of different columns collide"


def clean_text(v) -> str | None:
    if v is None:
        return None
    s = re.sub(r"\s+", " ", str(v).replace(NBSP, " ")).strip()
    return s or None


def is_blank(v) -> bool:
    return v is None or (isinstance(v, str) and clean_text(v) is None)


def parse_date(v) -> tuple[date | None, str | None]:
    """Return (date or None, problem or None)."""
    if isinstance(v, XlError):
        return None, "Error value"
    if is_blank(v):
        return None, None
    d = None
    if isinstance(v, datetime):          # before date: datetime is a subclass of date
        d = v.date()
    elif isinstance(v, date):
        d = v
    elif isinstance(v, (int, float)) and not isinstance(v, bool):
        try:
            d = EXCEL_EPOCH + timedelta(days=math.floor(v))
        except OverflowError:            # beyond year 9999
            pass
    elif m := DATE_TEXT.match(clean_text(v)):
        a, b, c = m.groups()
        try:
            if len(a) == 4:
                d = date(int(a), int(b), int(c))
            elif len(c) == 4:
                d = date(int(c), int(b), int(a))
        except ValueError:               # impossible date such as 31.02.2026, or year 0
            pass
    if d is None or d < MIN_DATE:
        return None, "Unreadable date"
    return d, None


def parse_value(v) -> tuple[Decimal | None, str | None]:
    """Return (amount or None, problem or None). Decimal keeps cent sums exact."""
    if isinstance(v, XlError):
        return None, "Error value"
    if is_blank(v):
        return None, None
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return Decimal(str(v)), None
    s = clean_text(v).replace(" ", "").replace("EUR", "")
    if "," in s and "." in s:            # the separator that comes last is the decimal one
        s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") else s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    if not PLAIN_NUMBER.match(s):
        return None, "Unreadable number"
    return Decimal(s), None


def parse_type(v) -> str:
    s = clean_text(v)
    if s is None:
        return "(no type)"
    return next((t for t in TYPES if t.lower() == s.lower()), s)


def parse_row(raw: dict) -> tuple[dict | None, list[tuple]]:
    """One register row -> (parsed row or None for a fully blank row, issues)."""
    if all(is_blank(raw.get(c)) for c in COLUMNS):
        return None, []
    no = clean_text(raw.get("ContractNo"))
    parsed = dict(no=no, client=clean_text(raw.get("Client")), type=parse_type(raw.get("ContractType")),
                  manager=clean_text(raw.get("Manager")) or "(no manager)")
    issues = []
    for key, parser in (("Drafted", parse_date), ("Signed", parse_date), ("ValueEUR", parse_value)):
        v = raw.get(key)
        parsed[key], problem = parser(v)
        if problem is None and key == "Drafted" and parsed[key] is None:
            problem = "Missing drafted date"
        if problem:
            issues.append((no, FIELD_LABELS[key], problem, v.code if isinstance(v, XlError) else str(v or "")))
    return parsed, issues


def report_figures(rows: list[dict], year: int, manager: str = "All", ctype: str = "All") -> dict:
    """The figures that depend on the Report inputs: report year, Manager and Contract type filters.
    KPIs, monthly table and List use both filters, by type only the manager filter, by manager only
    the type filter. A filter matches the whole name exactly; like Excel's "=" it ignores case."""
    def in_year(d):
        return d is not None and d.year == year

    def money(rs):
        return float(round(sum((r["ValueEUR"] or 0) for r in rs), 2))

    def is_(selected, value):
        return selected == "All" or value.lower() == selected.lower()

    drafted = [r for r in rows if in_year(r["Drafted"])]
    signed = [r for r in rows if in_year(r["Signed"])]

    def pending(rs):                     # drafted in the year and no valid signed date
        return sum(1 for r in rs if in_year(r["Drafted"]) and r["Signed"] is None)

    both = lambda r: is_(manager, r["manager"]) and is_(ctype, r["type"])
    d_both, s_both = [r for r in drafted if both(r)], [r for r in signed if both(r)]

    def block(key, value, other_key, other_value, with_pending=False):
        d = [r for r in drafted if r[key] == value and is_(other_value, r[other_key])]
        s = [r for r in signed if r[key] == value and is_(other_value, r[other_key])]
        out = dict(drafted=len(d), signed=len(s), signed_value=money(s))
        if with_pending:
            out["pending"] = pending(d)
        return out

    return dict(
        kpi=dict(drafted=len(d_both), signed=len(s_both), signed_value=money(s_both), pending=pending(d_both)),
        monthly=[dict(month=f"{year}-{m:02d}",
                      drafted=sum(1 for r in d_both if r["Drafted"].month == m),
                      signed=sum(1 for r in s_both if r["Signed"].month == m),
                      signed_value=money(r for r in s_both if r["Signed"].month == m)) for m in range(1, 13)],
        # Rows: every type and manager in the data, whatever the year
        by_type={t: block("type", t, "manager", manager) for t in sorted({r["type"] for r in rows}, key=str.lower)},
        by_manager={m: block("manager", m, "type", ctype, True)
                    for m in sorted({r["manager"] for r in rows}, key=str.lower)},
        list_rows=sum(1 for r in rows if (in_year(r["Drafted"]) or in_year(r["Signed"])) and both(r)),
    )


def figures(rows: list[dict], issues: list[tuple], scenarios: list[dict] = ()) -> dict:
    """The numbers the report shows, per the metric definitions of the spec: report year 2026 with
    both filters on "All", plus one block per extra scenario (other year, filters set)."""
    # Excel UNIQUE and "=" ignore case, so names that differ only in case would merge there
    for key in ("type", "manager", "client"):
        names = {r[key] for r in rows if r[key] is not None}
        assert len(names) == len({n.lower() for n in names}), f"{key} names differ only in case"

    base = report_figures(rows, REPORT_YEAR)
    problems = Counter(p for _, _, p, _ in issues)
    out = dict(
        contracts_rows=len(rows),
        kpi=dict(**base["kpi"], issues=len(issues)),
        monthly=base["monthly"], by_type=base["by_type"], by_manager=base["by_manager"],
        issues_by_problem={p: problems[p] for p in PROBLEMS if problems[p]},
        list_rows=base["list_rows"],
        # Exact spellings in tblContracts. Excel's UNIQUE and "=" ignore case, so only these
        # show whether the type canonicalisation and the text cleaning work
        exact_counts=dict(
            contract_type=dict(sorted(Counter(r["type"] for r in rows).items())),
            manager=dict(sorted(Counter(r["manager"] for r in rows).items())),
            status=dict(sorted(Counter("Signed" if r["Signed"] else "Pending" for r in rows).items())),
            client=dict(sorted(Counter(r["client"] or "" for r in rows).items())),
        ),
        # SHA-256 of the Contract No column in table order, joined with "\n" (UTF-8)
        contract_no_sha256=hashlib.sha256("\n".join(r["no"] or "" for r in rows).encode("utf-8")).hexdigest(),
    )
    out["scenarios"] = [dict(**s, **report_figures(rows, s["report_year"], s["manager"], s["type"]))
                        for s in scenarios]
    return out


# ------------------------------------------------------------------- truth records

def make_records(rng: random.Random) -> list[dict]:
    workdays = [FIRST_DRAFTED + timedelta(n) for n in range((LAST_DRAFTED - FIRST_DRAFTED).days + 1)]
    workdays = [d for d in workdays if d.weekday() < 5]
    clients = sorted({f"{rng.choice(CLIENT_WORDS)} {rng.choice(CLIENT_TYPES)}" for _ in range(90)})
    names, shares = list(TYPES), [TYPES[t][0] for t in TYPES]
    recs = []
    for _ in range(RECORDS):
        drafted = rng.choice(workdays)
        ctype = rng.choices(names, shares)[0]
        lo, hi = TYPES[ctype][1:]
        cents = round(rng.triangular(lo, hi, lo + (hi - lo) * 0.15) * 100)   # big contracts are rarer
        signed = None
        if rng.random() < 0.9:
            signed = drafted + timedelta(days=rng.randint(2, 60))
            while signed.weekday() >= 5:
                signed += timedelta(days=1)
            if signed > REGISTER_DATE:
                signed = None
        recs.append(dict(client=rng.choice(clients), type=ctype, manager=rng.choice(MANAGERS),
                         drafted=drafted, signed=signed, cents=cents,
                         drafted_state="ok", signed_state="ok" if signed else "blank", value_state="ok",
                         note=rng.choice(NOTES) if rng.random() < 0.08 else None))
    recs.sort(key=lambda r: r["drafted"])
    per_year = Counter()
    for r in recs:
        yy = r["drafted"].year % 100
        per_year[yy] += 1
        r["no"] = f"CTR-{yy}-{per_year[yy]:04d}"

    # Problem cells: each kind in distinct records, so every planned count really appears
    for field, kinds in (("drafted_state", ("blank", "error", "bad")), ("signed_state", ("error", "bad")),
                         ("value_state", ("blank", "error", "bad"))):
        prefix = field.split("_")[0]
        picks = rng.sample(range(len(recs)), sum(SPECIAL[f"{prefix}_{k}"] for k in kinds))
        for kind in kinds:
            for _ in range(SPECIAL[f"{prefix}_{kind}"]):
                recs[picks.pop()][field] = kind
    picks = rng.sample(range(len(recs)), sum(OTHER_TYPES.values()) + SPECIAL["type_blank"])
    for name, count in [*OTHER_TYPES.items(), (None, SPECIAL["type_blank"])]:
        for _ in range(count):
            recs[picks.pop()]["type"] = name
    picks = rng.sample(range(len(recs)), sum(OTHER_MANAGERS.values()) + SPECIAL["manager_blank"])
    for name, count in [*OTHER_MANAGERS.items(), (None, SPECIAL["manager_blank"])]:
        for _ in range(count):
            recs[picks.pop()]["manager"] = name

    # The year limits must be exercised: valid dates on 31 Dec and 1 Jan in both columns
    for day in BOUNDARY_DAYS:
        for key in ("drafted", "signed"):
            n = sum(1 for r in recs if r[key] == day and r[key + "_state"] == "ok")
            assert n >= 2, f"only {n} valid {key} dates on {day}: choose another seed"
    return recs


def truth(recs: list[dict]) -> tuple[list[dict], list[tuple]]:
    """What the report must load from every record: parsed values and issues, straight from the truth."""
    rows, issues = [], []
    for r in recs:
        rows.append(dict(no=r["no"], client=r["client"], type=r["type"] or "(no type)",
                         manager=r["manager"] or "(no manager)",
                         Drafted=r["drafted"] if r["drafted_state"] == "ok" else None,
                         Signed=r["signed"] if r["signed_state"] == "ok" else None,
                         ValueEUR=Decimal(r["cents"]) / 100 if r["value_state"] == "ok" else None))
        for key, state in (("Drafted", r["drafted_state"]), ("Signed", r["signed_state"]),
                           ("ValueEUR", r["value_state"])):
            problem = {"error": "Error value", "bad": "Unreadable number" if key == "ValueEUR" else "Unreadable date",
                       "blank": "Missing drafted date" if key == "Drafted" else None}.get(state)
            if problem:
                issues.append((r["no"], FIELD_LABELS[key], problem, ""))
    return rows, issues


# ------------------------------------------------------------- per-variant rendering

def assign(rng: random.Random, cells: list[int], weights: dict, eligible=lambda cell, rep: True) -> dict:
    """Pick a representation per cell by weight, then top up rare ones to MIN_EDGE."""
    reps = {}
    for cell in cells:
        options = [r for r in weights if eligible(cell, r)]
        reps[cell] = rng.choices(options, [weights[r] for r in options])[0]
    counts = Counter(reps.values())
    common = max(weights, key=weights.get)
    for rep in weights:
        while counts[rep] < MIN_EDGE:
            cell = rng.choice([c for c in cells if reps[c] == common and eligible(c, rep)])
            reps[cell] = rep
            counts[rep] += 1
            counts[common] -= 1
    return reps


def cycle_texts(rng: random.Random, cells: list[int], texts: tuple) -> dict:
    """Spread the unreadable texts evenly so each one appears."""
    cells = cells[:]
    rng.shuffle(cells)
    return {c: texts[n % len(texts)] for n, c in enumerate(cells)}


def render_date(d: date, rep: str, rng: random.Random):
    if rep == "date":
        return d
    if rep == "datetime":                # never midnight, so the time part is real
        return datetime(d.year, d.month, d.day, rng.randint(8, 17), rng.randint(1, 59))
    if rep == "text dd.mm.yyyy":
        return d.strftime("%d.%m.%Y")
    if rep == "text d.m.yyyy":
        return f"{d.day}.{d.month}.{d.year}"
    if rep == "text yyyy-mm-dd":
        return d.isoformat()
    if rep == "text dd/mm/yyyy":
        return d.strftime("%d/%m/%Y")
    return (d - EXCEL_EPOCH).days        # serial number in a General cell


def render_value(cents: int, rep: str):
    whole, frac = divmod(cents, 100)
    grouped = f"{whole:,}"
    if rep == "number":
        return cents / 100
    if rep == "text 12 500,00":
        return f"{grouped.replace(',', ' ')},{frac:02d}"
    if rep == "text 12,500.00":
        return f"{grouped}.{frac:02d}"
    return f"{grouped.replace(',', NBSP)},{frac:02d} EUR"


def render_type(t: str, rep: str) -> str:
    if rep == "lower case + trailing space":
        return t.lower() + " "
    if rep == "upper case + double spaces":
        return "  ".join(t.upper().split())
    if rep == "leading space":
        return " " + t
    if rep == "non-breaking space":
        return t.replace(" ", NBSP)
    return t


def render_manager(m: str, rep: str) -> str:
    if rep == "leading/trailing spaces":
        return f"  {m} "
    if rep == "non-breaking space":
        return m.replace(" ", NBSP)
    return m


def build_variant(name: str, recs: list[dict], cfg: dict, seed: int) -> tuple[list[dict], set, dict]:
    """Raw cell values per record for one variant, rows that get a blank row before them, edge counts."""
    rng = random.Random(f"{seed}:{name}")   # own stream: the truth does not depend on variant settings
    raws = [dict(Lp=i + 1, Note=r["note"]) for i, r in enumerate(recs)]
    edge = {label: Counter() for label in ("Drafted", "Signed", "Value (EUR)", "Contract Type", "Manager", "Text")}

    for key, field in (("drafted", "Drafted"), ("signed", "Signed")):
        state = key + "_state"
        ok = [i for i, r in enumerate(recs) if r[state] == "ok"]
        reps = assign(rng, ok, cfg["dates"], lambda i, rep: rep != "text d.m.yyyy"
                      or recs[i][key].day < 10 or recs[i][key].month < 10)
        bad = cycle_texts(rng, [i for i, r in enumerate(recs) if r[state] == "bad"], cfg["bad_dates"])
        for i, r in enumerate(recs):
            if r[state] == "ok":
                raws[i][field] = render_date(r[key], reps[i], rng)
                edge[field][reps[i]] += 1
            elif r[state] == "blank":
                raws[i][field] = None
                edge[field]["blank"] += 1
            elif r[state] == "error":
                raws[i][field] = XlError()
                edge[field]["#N/A"] += 1
            else:
                raws[i][field] = bad[i]
                edge[field]["before 1900 (text or negative serial)" if bad[i] in PRE_1900 else "unreadable text"] += 1

    ok = [i for i, r in enumerate(recs) if r["value_state"] == "ok"]
    reps = assign(rng, ok, cfg["values"], lambda i, rep: rep == "number" or recs[i]["cents"] >= 100_000)
    bad = cycle_texts(rng, [i for i, r in enumerate(recs) if r["value_state"] == "bad"], cfg["bad_values"])
    for i, r in enumerate(recs):
        st = r["value_state"]
        raws[i]["ValueEUR"] = (render_value(r["cents"], reps[i]) if st == "ok" else None if st == "blank"
                               else XlError() if st == "error" else bad[i])
        label = (reps[i] if st == "ok" else {"blank": "blank", "error": "#N/A"}.get(st)
                 or ("separator with a digit missing" if bad[i] in HALF_NUMBERS else "unreadable text"))
        edge["Value (EUR)"][label] += 1

    canonical = [i for i, r in enumerate(recs) if r["type"] in TYPES]
    reps = assign(rng, canonical, cfg["types"], lambda i, rep: rep != "non-breaking space" or " " in recs[i]["type"])
    for i, r in enumerate(recs):
        if r["type"] in TYPES:
            raws[i]["ContractType"] = render_type(r["type"], reps[i])
            edge["Contract Type"]["canonical" if reps[i] == "canonical" else "other case or extra spaces"] += 1
        elif r["type"] is None:
            raws[i]["ContractType"] = None
            edge["Contract Type"]["blank"] += 1
        else:
            padded = cfg["padded_other_types"] and rng.random() < 0.5
            raws[i]["ContractType"] = f" {r['type']} " if padded else r["type"]
            edge["Contract Type"]["outside the list"] += 1
            if PATTERN_CHARS & set(r["type"]):
                edge["Contract Type"]["outside the list, with * ? < >"] += 1
            if len(r["type"]) > 255:
                edge["Contract Type"]["outside the list, over 255 characters"] += 1

    named = [i for i, r in enumerate(recs) if r["manager"] in MANAGERS]
    reps = assign(rng, named, cfg["managers"])
    for i, r in enumerate(recs):
        if r["manager"] in MANAGERS:
            raws[i]["Manager"] = render_manager(r["manager"], reps[i])
            edge["Manager"][reps[i]] += 1
        elif r["manager"] is None:
            raws[i]["Manager"] = None
            edge["Manager"]["blank"] += 1
        else:                            # a placeholder such as "?" stays as it is typed
            raws[i]["Manager"] = r["manager"]
            edge["Manager"]["placeholder with ? - < >"] += 1

    for i, r in enumerate(recs):
        raws[i]["ContractNo"] = r["no"]
        raws[i]["Client"] = r["client"]
        if rng.random() < cfg["messy_text"]:
            raws[i]["ContractNo"] = r["no"] + " "
            edge["Text"]["Contract No with trailing space"] += 1
        if rng.random() < cfg["messy_text"]:
            raws[i]["Client"] = r["client"].replace(" ", rng.choice(["\n", "  "]), 1)
            edge["Text"]["Client with line break or double space"] += 1

    blank_before = set(rng.sample(range(1, len(recs)), BLANK_ROWS))
    edge["Blank rows"] = Counter({"fully blank rows": BLANK_ROWS})
    order = [*DATE_REPS, *VALUE_REPS, "canonical", "other case or extra spaces", "outside the list",
             *MANAGER_REPS, "blank", "#N/A", "unreadable text"]
    rank = lambda rep: order.index(rep) if rep in order else len(order)
    return raws, blank_before, {k: {rep: v[rep] for rep in sorted(v, key=rank)} for k, v in edge.items() if v}


# ------------------------------------------------------------------------- writing

def write_register(path: Path, cfg: dict, raws: list[dict], blank_before: set) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    if cfg["intro"]:
        ws.title, lines = cfg["intro"]
        for n, line in enumerate(lines, start=1):
            ws.cell(n, 1, line)
        ws.column_dimensions["A"].width = 80
        ws = wb.create_sheet(cfg["sheet"])
    else:
        ws.title = cfg["sheet"]

    row = 0
    for n, title in enumerate(cfg["titles"]):
        row += 1
        cell = ws.cell(row, 1, title)
        if n == 0:
            cell.font = TITLE_FONT
    if cfg["titles"]:
        row += 1                         # one blank row between the titles and the header
    row += 1
    for col, (key, label) in enumerate(cfg["columns"], start=1):
        cell = ws.cell(row, col, label)
        cell.font, cell.fill, cell.alignment = HEADER_FONT, HEADER_FILL, WRAP
        ws.column_dimensions[get_column_letter(col)].width = WIDTHS[key]
    ws.freeze_panes = ws.cell(row + 1, 1)

    for i, raw in enumerate(raws):
        row += 2 if i in blank_before else 1
        for col, (key, _) in enumerate(cfg["columns"], start=1):
            v = raw.get(key)
            if v is None:
                continue
            cell = ws.cell(row, col, v.code if isinstance(v, XlError) else v)
            if isinstance(v, XlError):
                assert cell.data_type == "e", "openpyxl did not write an error cell"
            elif isinstance(v, datetime):
                cell.number_format = cfg["date_format"] + " hh:mm"
            elif isinstance(v, date):
                cell.number_format = cfg["date_format"]
            elif isinstance(v, float):
                cell.number_format = "#,##0.00"
    wb.save(path)


def read_back(path: Path) -> tuple[list[dict], list[tuple]]:
    """Self-check only: find the header like the report does and parse every cell of the written file."""
    wb = load_workbook(path)
    for ws in wb.worksheets:
        grid = [[XlError(c.value) if c.data_type == "e" else c.value for c in row] for row in ws.iter_rows()]
        for h, header in enumerate(grid[:15]):
            cols = {}
            for j, v in enumerate(header):
                key = HEADER_KEYS.get(normalize_header(v)) if isinstance(v, str) else None
                if key and key not in cols:
                    cols[key] = j
            if len(cols) < 4:
                continue
            missing = [c for c in REQUIRED if c not in cols]
            if missing:
                raise ValueError(f"{path.name}: missing columns {missing}")
            rows, issues = [], []
            for line in grid[h + 1:]:
                parsed, row_issues = parse_row({k: line[j] if j < len(line) else None for k, j in cols.items()})
                if parsed:
                    rows.append(parsed)
                    issues += row_issues
            return rows, issues
    raise ValueError(f"{path.name}: no header row in the first 15 rows of any sheet")


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate the demo contracts register and expected figures.")
    parser.add_argument("--seed", type=int, default=2026, help="random seed (default 2026)")
    args = parser.parse_args()

    recs = make_records(random.Random(args.seed))
    truth_rows, truth_issues = truth(recs)
    expected = figures(truth_rows, truth_issues, SCENARIOS)
    result = dict(report_year=REPORT_YEAR, seed=args.seed, records=len(recs), variants={}, edge_cases={})

    for name in ("a", "b"):
        cfg = VARIANTS[name]
        raws, blank_before, edge = build_variant(name, recs, cfg, args.seed)
        # Self-check 1: every rendered cell parses back to the truth
        lines = []
        for i, raw in enumerate(raws):
            if i in blank_before:
                lines.append({})
            lines.append(raw)
        parsed = [parse_row(raw) for raw in lines]
        rows = [p for p, _ in parsed if p]
        issues = [x for _, row_issues in parsed for x in row_issues]
        assert rows == truth_rows, f"variant {name}: rendered cells do not parse to the truth"
        assert [x[:3] for x in issues] == [x[:3] for x in truth_issues], f"variant {name}: issues differ"
        path = DATA / f"variant_{name}" / REGISTER
        write_register(path, cfg, raws, blank_before)
        # Self-check 2: the written file gives the same figures
        assert figures(*read_back(path), SCENARIOS) == expected, f"variant {name}: file read back gives other figures"
        result["variants"][name] = expected
        result["edge_cases"][name] = edge

    empty_path = DATA / "variant_empty" / REGISTER
    write_register(empty_path, VARIANTS["a"], [], set())
    result["variants"]["empty"] = figures([], [], SCENARIOS)
    assert figures(*read_back(empty_path), SCENARIOS) == result["variants"]["empty"]

    shutil.copyfile(DATA / "variant_b" / REGISTER, DATA / REGISTER)
    (DATA / "expected.json").write_text(json.dumps(result, indent=2, ensure_ascii=False) + "\n", encoding="utf-8",
                                        newline="\n")

    years = Counter(r["no"][:6] for r in recs)
    print(f"{len(recs)} records ({', '.join(f'{k} {v}' for k, v in sorted(years.items()))}), "
          f"{BLANK_ROWS} blank rows per variant, report year {REPORT_YEAR}")
    print("Edge cases (cells per variant):")
    ea, eb = result["edge_cases"]["a"], result["edge_cases"]["b"]
    for field in dict.fromkeys([*ea, *eb]):
        for rep in dict.fromkeys([*ea.get(field, {}), *eb.get(field, {})]):
            print(f"  {field:<14} {rep:<40} a {ea.get(field, {}).get(rep, 0):>5}   b {eb.get(field, {}).get(rep, 0):>5}")
    k = expected["kpi"]
    print(f"Expected {REPORT_YEAR} (a = b): drafted {k['drafted']}, signed {k['signed']}, "
          f"signed value {k['signed_value']:,.2f} EUR, pending {k['pending']}, data issues {k['issues']}, "
          f"contracts rows {expected['contracts_rows']}, list rows {expected['list_rows']}")
    print("Issues by problem: " + ", ".join(f"{p} {n}" for p, n in expected["issues_by_problem"].items()))
    print("By type: " + ", ".join(f"{t[:30]} {v['drafted']}/{v['signed']}" for t, v in expected["by_type"].items())
          + "  (drafted/signed)")
    for day in BOUNDARY_DAYS:
        print(f"Valid dates on {day}: drafted {sum(1 for r in truth_rows if r['Drafted'] == day)}, "
              f"signed {sum(1 for r in truth_rows if r['Signed'] == day)}")
    print("Scenarios (drafted/signed/pending, list rows):")
    for s in expected["scenarios"]:
        k = s["kpi"]
        print(f"  {s['name']:<42} {k['drafted']}/{k['signed']}/{k['pending']}, list {s['list_rows']}")
    print(f"Wrote {', '.join(str(p.relative_to(ROOT)) for p in sorted(DATA.rglob('*')) if p.is_file())}")


if __name__ == "__main__":
    main()
