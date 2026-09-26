"""Lint an authored workbook for the mistakes a library most often makes,
against the machine-checkable part of the delivery checklist
(references/delivery-checklist.md).

Usage:
    python scripts/check_workbook.py model.xlsx [--json]

Needs only Python 3.9+ and openpyxl, which ChatGPT and Claude code sandboxes
ship with. It reads the file and changes nothing.

Each finding is FAIL (breaks the checklist), REVIEW (needs a person to judge)
or INFO (for the handover). The exit code is 1 when anything FAILs.

This is a heuristic lint, not a proof. It reads formulas as text with patterns
rather than parsing them as Excel does, so a clean run means none of the
common mistakes below were found, not that the workbook is correct. Known
limits: formulas using LET or LAMBDA skip the undefined-name check, because
their local names are not defined names; a function on none of its lists is
reported for review rather than judged; and a table is recognised by a header
row of two or more text labels in the first five rows.

What it cannot check, and so leaves to the checklist: invented figures, sources
and owners on Assumptions, hardcoded figures inside calculation ranges, visible
error values (a library writes no cached results; open the file in Excel), and
the handover itself.
"""

from __future__ import annotations

import json
import re
import sys
from collections import defaultdict

try:
    import openpyxl
    from openpyxl.utils import get_column_letter
    from openpyxl.utils.formulas import FORMULAE
except ImportError:  # pragma: no cover
    sys.exit("openpyxl is required: pip install openpyxl")

# Functions that must be written with a prefix, and the prefix each needs:
# the table in references/formula-prefixes.md, plus the other functions Excel
# added after 2007.
XLWS = {"FILTER", "SORT"}
XLFN = {
    "XLOOKUP", "XMATCH", "SORTBY", "UNIQUE", "SEQUENCE", "RANDARRAY", "LET",
    "LAMBDA", "ISOMITTED", "TEXTSPLIT", "TEXTBEFORE", "TEXTAFTER", "VSTACK",
    "HSTACK", "TOCOL", "TOROW", "CHOOSECOLS", "CHOOSEROWS", "BYROW", "BYCOL",
    "MAP", "REDUCE", "SCAN", "MAKEARRAY", "IFS", "SWITCH", "MAXIFS", "MINIFS",
    "CONCAT", "TEXTJOIN", "ANCHORARRAY", "SINGLE",
} | set("""
ACOT ACOTH AGGREGATE ARABIC ARRAYTOTEXT BASE BETA.DIST BETA.INV BINOM.DIST
BINOM.DIST.RANGE BINOM.INV BITAND BITLSHIFT BITOR BITRSHIFT BITXOR CEILING.MATH
CEILING.PRECISE CHISQ.DIST CHISQ.DIST.RT CHISQ.INV CHISQ.INV.RT CHISQ.TEST
COMBINA CONFIDENCE.NORM CONFIDENCE.T COT COTH COVARIANCE.P COVARIANCE.S CSC CSCH
DAYS DECIMAL DROP ERF.PRECISE ERFC.PRECISE EXPAND EXPON.DIST F.DIST F.DIST.RT
F.INV F.INV.RT F.TEST FILTERXML FLOOR.MATH FLOOR.PRECISE FORECAST.ETS
FORECAST.ETS.CONFINT FORECAST.ETS.SEASONALITY FORECAST.ETS.STAT FORECAST.LINEAR
FORMULATEXT GAMMA GAMMA.DIST GAMMA.INV GAMMALN.PRECISE GAUSS GROUPBY
HYPGEOM.DIST IFNA IMAGE IMCOSH IMCOT IMCSC IMCSCH IMSEC IMSECH IMSINH IMTAN
ISFORMULA ISOWEEKNUM LOGNORM.DIST LOGNORM.INV MODE.MULT MODE.SNGL MUNIT
NEGBINOM.DIST NETWORKDAYS.INTL NORM.DIST NORM.INV NORM.S.DIST NORM.S.INV
NUMBERVALUE PDURATION PERCENTILE.EXC PERCENTILE.INC PERCENTOF PERCENTRANK.EXC
PERCENTRANK.INC PERMUTATIONA PHI PIVOTBY POISSON.DIST QUARTILE.EXC QUARTILE.INC
RANK.AVG RANK.EQ REGEXEXTRACT REGEXREPLACE REGEXTEST RRI SEC SECH SHEET SHEETS
SKEW.P STDEV.P STDEV.S STOCKHISTORY T.DIST T.DIST.2T T.DIST.RT T.INV T.INV.2T
T.TEST TAKE TRIMRANGE UNICHAR UNICODE VALUETOTEXT VAR.P VAR.S WEBSERVICE
WEIBULL.DIST WORKDAY.INTL WRAPCOLS WRAPROWS XOR Z.TEST
""".split())
# Functions whose prefix this lint does not judge either way.
UNJUDGED = {"ISO.CEILING", "ECMA.CEILING"}
# Functions Excel 2007 already had, from openpyxl's catalogue of the built-in
# functions: a prefix on one of these is also #NAME?.
LEGACY = set(FORMULAE) - XLFN - XLWS - UNJUDGED
KNOWN = LEGACY | XLFN | XLWS | UNJUDGED

LIBRARY_NAMES = ("openpyxl", "sheetjs", "xlsx", "xlsxwriter", "python",
                 "pandas", "exceljs", "apache poi", "epplus")
BLUE, BLACK, GREEN = "0000FF", "000000", "008000"
DATE_TEXT = re.compile(
    r"^\s*(\d{1,2}[/.-]\d{1,2}[/.-]\d{2,4}|\d{4}-\d{2}-\d{2})\s*$")
CELL_REF = re.compile(r"^\$?[A-Z]{1,3}\$?\d+$")
HEADER_SCAN_ROWS = 5

findings: list[dict[str, str]] = []


def add(level: str, rule: str, where: str, detail: str) -> None:
    findings.append(
        {"level": level, "rule": rule, "where": where, "detail": detail})


def rgb(font) -> str | None:
    color = font.color if font else None
    if color is None or color.type != "rgb" or not isinstance(color.rgb, str):
        return None
    return color.rgb[-6:].upper()


def strip_strings(formula: str) -> str:
    return re.sub(r'"[^"]*"', '""', formula)


def header_row(ws) -> int | None:
    """The first of the top rows holding two or more text labels with rows
    under it, which is where a table's header sits under an optional title."""
    for row in ws.iter_rows(min_row=1, max_row=min(HEADER_SCAN_ROWS,
                                                    ws.max_row)):
        labels = [c.value for c in row if isinstance(c.value, str)
                  and c.value.strip() and not c.value.startswith("=")]
        if len(labels) >= 2 and ws.max_row > row[0].row:
            return row[0].row
    return None


def check_formula(where: str, formula: str, lookups: dict[str, list[str]],
                  names_used: dict[str, set[str]],
                  unknown_calls: dict[str, set[str]]) -> None:
    body = strip_strings(formula)
    sheet = where.split("!")[0]
    # Excel reads function names without regard to case, and so does this.
    for match in re.finditer(r"((?:_xlfn\.)?(?:_xlws\.)?)([A-Z][A-Z0-9._]*)\(",
                             body, re.IGNORECASE):
        prefix, name = match.group(1).lower(), match.group(2).upper()
        if name in XLWS and prefix != "_xlfn._xlws.":
            add("FAIL", "prefix", where,
                f"{name} must be written _xlfn._xlws.{name}")
        elif name in XLFN and prefix != "_xlfn.":
            add("FAIL", "prefix", where, f"{name} must be written _xlfn.{name}")
        elif name in LEGACY and prefix:
            add("FAIL", "prefix", where,
                f"{name} existed in Excel 2007 and takes no prefix")
        elif name not in KNOWN:
            unknown_calls.setdefault(name, set()).add(where)
        if name in {"XLOOKUP", "XMATCH"}:
            lookups["XLOOKUP"].append(where)
        elif name in {"VLOOKUP", "HLOOKUP", "INDEX", "MATCH"}:
            lookups["INDEX/MATCH"].append(where)
    if re.search(r"#(REF|NAME|VALUE|DIV/0|N/A)[!?]?", body):
        add("FAIL", "error-in-formula", where, "the formula text holds an error")
    if re.search(r"(?:^|[^A-Za-z_])'?Assumptions'?!\$?[A-Z]{1,3}\$?\d+", body,
                 re.IGNORECASE):
        add("FAIL", "named-range", where,
            "points at an Assumptions cell by address; use its named range")

    # LET and LAMBDA declare local names that are not defined names.
    if re.search(r"(?:_xlfn\.)?(?:LET|LAMBDA)\(", body, re.IGNORECASE):
        return
    # Drop what is not a name before looking for names: error literals and
    # the insides of structured references such as Table1[Amount].
    names_body = re.sub(r"#[A-Z0-9/]+[!?]?", " ", body, flags=re.IGNORECASE)
    names_body = re.sub(r"\[[^\]]*\]", "[]", names_body)
    # A quoted sheet name such as 'My Sheet'! holds words that are not names.
    bare_body = re.sub(r"'(?:[^']|'')+'!", "Q!", names_body)
    for token in re.findall(r"(?<![A-Za-z0-9_.!$':@#\[])([A-Za-z_][A-Za-z0-9_.]*)"
                            r"(?![A-Za-z0-9_.(!:\[])", bare_body):
        upper = token.upper()
        if (CELL_REF.match(upper) or upper in {"TRUE", "FALSE"}
                or token.lower().startswith("_xl")):
            continue
        names_used.setdefault(token, set()).add(sheet)
    # A sheet-qualified name, Sheet1!Tax_Rate or 'My Sheet'!Tax_Rate, is looked
    # up on the sheet it names. Recorded as "Sheet1!Tax_Rate" for main().
    for match in re.finditer(r"(?:'((?:[^']|'')+)'|([A-Za-z_][A-Za-z0-9_.]*))!"
                             r"([A-Za-z_][A-Za-z0-9_.]*)(?![A-Za-z0-9_.(!:])",
                             names_body):
        target = (match.group(1) or "").replace("''", "'") or match.group(2)
        name = match.group(3)
        if not CELL_REF.match(name.upper()):
            names_used.setdefault(f"{target}!{name}", set()).add(sheet)


def main(path: str, as_json: bool) -> int:
    wb = openpyxl.load_workbook(path)
    props = wb.properties
    for field in ("creator", "lastModifiedBy"):
        value = (getattr(props, field) or "").strip()
        if not value:
            add("FAIL", "properties", "workbook", f"{field} is not set")
        elif any(lib in value.lower() for lib in LIBRARY_NAMES):
            add("FAIL", "properties", "workbook",
                f"{field} reads '{value}', a library name")
    if not (props.title or "").strip():
        add("FAIL", "properties", "workbook", "title is not set")

    # A workbook-scoped name resolves on every sheet; a worksheet-scoped name
    # only on its own sheet.
    workbook_names = {name.upper() for name in wb.defined_names}
    sheet_names = {ws.title: {name.upper() for name in ws.defined_names}
                   for ws in wb.worksheets}
    names_used: dict[str, set[str]] = {}
    unknown_calls: dict[str, set[str]] = {}
    lookups: dict[str, list[str]] = defaultdict(list)

    for ws in wb.worksheets:
        sheet = ws.title
        # A cover, summary or notes sheet is not held to the table rules.
        header = header_row(ws)
        is_data = header is not None and sheet.lower() != "assumptions"

        for rng in ws.merged_cells.ranges:
            if rng.max_row > 1:
                add("FAIL", "merged-cells", f"{sheet}!{rng.coord}",
                    "merged beyond the title row")
        if is_data and not ws.freeze_panes:
            add("FAIL", "freeze", sheet, "header row is not frozen")
        if is_data and not ws.auto_filter.ref:
            add("FAIL", "autofilter", sheet, "autofilter is not on")

        # Rows down to the header hold titles and labels: formulas there are
        # checked, the data rules are not.
        last_label_row = header or 1
        formats: dict[int, set[str]] = defaultdict(set)
        for row in ws.iter_rows(min_row=1):
            for cell in row:
                where = f"{sheet}!{cell.coordinate}"
                value = cell.value
                in_labels = cell.row <= last_label_row
                is_formula = isinstance(value, str) and value.startswith("=")
                if in_labels and not is_formula:
                    continue
                if cell.comment is not None:
                    add("REVIEW", "comments", where,
                        "cell comment; keep only if the user asked for it")
                if value is None:
                    continue
                colour = rgb(cell.font)
                if is_formula:
                    check_formula(where, value, lookups, names_used,
                                  unknown_calls)
                    want = GREEN if "!" in strip_strings(value) else BLACK
                    if colour not in (None, want) or (
                            colour is None and want == GREEN):
                        add("FAIL", "font-colour", where,
                            f"formula font should be {want}, is {colour}")
                elif isinstance(value, (int, float)) and not isinstance(
                        value, bool):
                    if colour != BLUE:
                        add("REVIEW", "font-colour", where,
                            f"hardcoded number in {colour or 'default'}; an "
                            f"input is blue {BLUE}, a figure in a calculation "
                            "range should be a formula")
                elif isinstance(value, str):
                    if DATE_TEXT.match(value):
                        add("FAIL", "date-as-text", where,
                            f"'{value}' is text that looks like a date")
                    if value.strip().lower() == "needs input":
                        add("INFO", "needs-input", where,
                            "list this in the handover")
                # Text labels carry no number format worth comparing.
                if not in_labels and (not isinstance(value, str)
                                      or is_formula):
                    formats[cell.column].add(cell.number_format)
        for column, found in formats.items():
            if len(found) > 1:
                add("FAIL", "number-format",
                    f"{sheet}!{get_column_letter(column)}",
                    "mixes formats: " + ", ".join(sorted(found)))

    if lookups["XLOOKUP"] and lookups["INDEX/MATCH"]:
        add("FAIL", "lookup-family", "workbook",
            f"XLOOKUP ({lookups['XLOOKUP'][0]}) and INDEX/MATCH or VLOOKUP "
            f"({lookups['INDEX/MATCH'][0]}) are both used")
    for name in sorted(unknown_calls):
        # A named LAMBDA is called like a function and is a defined name.
        if name in workbook_names:
            continue
        add("REVIEW", "unknown-function", sorted(unknown_calls[name])[0],
            f"{name} is not a function this lint knows; check its spelling "
            "and whether it needs the _xlfn. prefix")
    for name in sorted(names_used):
        for sheet in sorted(names_used[name]):
            if "!" in name:
                target, bare = name.rsplit("!", 1)
                key = bare.upper()
                if target in sheet_names and (
                        key in sheet_names[target] or key in workbook_names):
                    continue
                add("REVIEW", "defined-names", sheet,
                    f"'{name}' names a range that is not defined on "
                    f"'{target}'; an unregistered name shows #NAME?")
                continue
            key = name.upper()
            if key in workbook_names or key in sheet_names.get(sheet, set()):
                continue
            add("REVIEW", "defined-names", sheet,
                f"'{name}' reads like a named range but is not defined for "
                "this sheet; an unregistered name shows #NAME?")

    failed = sum(f["level"] == "FAIL" for f in findings)
    if as_json:
        print(json.dumps({"file": path, "failed": failed,
                          "findings": findings}, indent=2))
    else:
        for level in ("FAIL", "REVIEW", "INFO"):
            for f in (f for f in findings if f["level"] == level):
                print(f"{level:6} {f['rule']:16} {f['where']:22} {f['detail']}")
        print(f"\n{failed} failing, "
              f"{sum(f['level'] == 'REVIEW' for f in findings)} to review. "
              "This is a heuristic lint: the rest of "
              "references/delivery-checklist.md is still yours to apply.")
    return 1 if failed else 0


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if a != "--json"]
    if len(args) != 1:
        sys.exit(__doc__)
    sys.exit(main(args[0], "--json" in sys.argv[1:]))
