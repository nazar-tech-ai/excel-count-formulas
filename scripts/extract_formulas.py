"""Build the formula catalog for the workbook.

Reads workbook/Multiples_Examples_of_Count.xlsx and writes:
  - docs/formulas.md   one section per sheet (topic, formulas, results)
  - docs/formulas.csv  every formula, one row per cell (for filtering in Excel)
  - README.md          refreshes the sheet index between the SHEETS markers

Usage (from the repo root):
    python scripts/extract_formulas.py
    python scripts/extract_formulas.py --workbook path/to/other.xlsx
"""

from __future__ import annotations

import argparse
import csv
import re
from collections import Counter
from pathlib import Path

import openpyxl
from openpyxl.worksheet.formula import ArrayFormula

ROOT = Path(__file__).resolve().parent.parent
DEFAULT_WORKBOOK = ROOT / "workbook" / "Multiples_Examples_of_Count.xlsx"
DOCS_DIR = ROOT / "docs"
README = ROOT / "README.md"

START_MARK = "<!-- SHEETS:START -->"
END_MARK = "<!-- SHEETS:END -->"

# Helper formulas used only to display other formulas as text; not examples.
HELPER_FUNCTIONS = ("FORMULATEXT",)

FUNC_RE = re.compile(r"([A-Z][A-Z0-9.]*)\(")
FUNC_NAME_RE = re.compile(r"[A-Z][A-Z0-9.]+")  # a bare function name, 2+ chars
# Row numbers inside cell references, e.g. B7 -> B#, $C$12 -> $C$#
CELL_ROW_RE = re.compile(r"\b(\$?[A-Z]{1,3})(\$?)\d+\b")


def clean(formula: str) -> str:
    """Drop the internal prefixes Excel stores for newer functions."""
    return formula.replace("_xlfn._xlws.", "").replace("_xlfn.", "")


def pattern_key(formula: str) -> str:
    """Normalise row numbers so B7..B17 style fill-downs group together."""
    return CELL_ROW_RE.sub(r"\1\2#", formula)


def sheet_topic(ws) -> str:
    """The sheet's heading: first plain-text cell in columns A-B of the top rows.

    Columns A-B only, because a 'Functions' header often sits further right.
    """
    for row in ws.iter_rows(min_row=1, max_row=3, min_col=1, max_col=2):
        for cell in row:
            v = cell.value
            if isinstance(v, str) and not v.startswith("=") and v != "Related Functions":
                return v.strip()
    return ws.title


def related_functions(ws) -> list[str]:
    """Names listed under the 'Related Functions' label, if the sheet has one."""
    for row in ws.iter_rows(min_row=1, max_row=8):
        for cell in row:
            if cell.value == "Related Functions":
                names = []
                r = cell.row + 1
                # The list can wrap onto a second row. The first row is always
                # read; further rows only while every entry looks like a function
                # name (e.g. SUMPRODUCT), so sample data is never picked up.
                while r <= ws.max_row:
                    found = [c.value.strip() for c in ws[r]
                             if isinstance(c.value, str) and not c.value.startswith("=")]
                    if not found:
                        break
                    if r > cell.row + 1 and not all(FUNC_NAME_RE.fullmatch(f) for f in found):
                        break
                    names += [f.upper() for f in found]
                    r += 1
                return names
    return []


def read_formulas(ws, ws_values):
    """Yield (cell_ref, formula_text, is_array, cached_result) for each formula."""
    for row in ws.iter_rows():
        for cell in row:
            v = cell.value
            if isinstance(v, ArrayFormula):
                text, ref, is_array = v.text, v.ref, True
            elif isinstance(v, str) and v.startswith("="):
                text, ref, is_array = v, cell.coordinate, False
            else:
                continue
            if any(h in text.upper() for h in HELPER_FUNCTIONS):
                continue
            result = ws_values[cell.coordinate].value
            yield ref, clean(text), is_array, result


def fmt_result(value) -> str:
    if value is None:
        return ""
    if hasattr(value, "isoformat"):
        return value.strftime("%Y-%m-%d")
    return str(value)


def md_code(text: str) -> str:
    return "`" + text.replace("|", "\\|") + "`"


def build(workbook_path: Path) -> None:
    wb = openpyxl.load_workbook(workbook_path)
    wb_values = openpyxl.load_workbook(workbook_path, data_only=True)

    sheets = []
    all_rows = []
    function_counter: Counter[str] = Counter()

    for ws in wb.worksheets:
        ws_values = wb_values[ws.title]
        items = list(read_formulas(ws, ws_values))
        topic = sheet_topic(ws)
        related = related_functions(ws)
        used = Counter()
        for _, text, _, _ in items:
            used.update(FUNC_RE.findall(text))
        function_counter.update(used.keys())  # count sheets, not occurrences

        # Group fill-down copies: keep the first, remember how many there are.
        groups: dict[str, dict] = {}
        for ref, text, is_array, result in items:
            key = pattern_key(text)
            if key not in groups:
                groups[key] = {"ref": ref, "text": text, "array": is_array,
                               "result": result, "count": 0}
            groups[key]["count"] += 1
            all_rows.append([ws.title, topic, ref, text,
                             "array" if is_array else "normal", fmt_result(result)])

        tables = sorted(ws.tables.keys())
        sheets.append({
            "name": ws.title, "topic": topic, "related": related,
            "tables": tables, "groups": list(groups.values()),
            "functions": sorted(used), "total": len(items),
        })

    DOCS_DIR.mkdir(exist_ok=True)
    write_markdown(sheets, function_counter, workbook_path)
    write_csv(all_rows)
    update_readme(sheets)
    print(f"{len(sheets)} sheets, {len(all_rows)} formulas -> docs/formulas.md, docs/formulas.csv, README.md")


def write_markdown(sheets, function_counter, workbook_path: Path) -> None:
    out = ["# Formula catalog", "",
           f"Generated from `{workbook_path.name}` by `scripts/extract_formulas.py`. "
           "Do not edit by hand; re-run the script instead.", "",
           "Formulas that are filled down a column (same pattern, different row) "
           "are shown once with a cell count.", "",
           "## Functions used across the workbook", "",
           "| Function | Sheets using it |", "| --- | --- |"]
    for fn, n in sorted(function_counter.items(), key=lambda kv: (-kv[1], kv[0])):
        out.append(f"| `{fn}` | {n} |")
    out.append("")

    for s in sheets:
        out.append(f"## {s['name']}: {s['topic']}")
        out.append("")
        if s["related"]:
            out.append("**Related functions:** " + ", ".join(f"`{r}`" for r in s["related"]) + "  ")
        if s["functions"]:
            out.append("**Functions used:** " + ", ".join(f"`{f}`" for f in s["functions"]) + "  ")
        if s["tables"]:
            out.append("**Excel tables:** " + ", ".join(f"`{t}`" for t in s["tables"]) + "  ")
        out.append("")
        if not s["groups"]:
            out.append("_No formulas on this sheet (explanation or layout only)._")
            out.append("")
            continue
        out += ["| Cell | Type | Formula | Result | Cells |", "| --- | --- | --- | --- | --- |"]
        for g in s["groups"]:
            kind = "array" if g["array"] else "normal"
            out.append(f"| {g['ref']} | {kind} | {md_code(g['text'])} | "
                       f"{md_code(fmt_result(g['result'])) if g['result'] not in (None, '') else ''} | "
                       f"{g['count']} |")
        out.append("")
    (DOCS_DIR / "formulas.md").write_text("\n".join(out), encoding="utf-8")


def write_csv(rows) -> None:
    with open(DOCS_DIR / "formulas.csv", "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh)
        w.writerow(["sheet", "topic", "cell", "formula", "type", "cached_result"])
        w.writerows(rows)


def update_readme(sheets) -> None:
    if not README.exists():
        return
    lines = ["| # | Sheet | Topic | Main functions | Formulas |",
             "| --- | --- | --- | --- | --- |"]
    for i, s in enumerate(sheets):
        funcs = ", ".join(s["related"] or s["functions"][:4])
        anchor = re.sub(r"[^a-z0-9 -]", "", f"{s['name']}: {s['topic']}".lower()).replace(" ", "-")
        lines.append(f"| {i} | [{s['name']}](docs/formulas.md#{anchor}) | "
                     f"{s['topic']} | {funcs} | {s['total']} |")
    block = f"{START_MARK}\n" + "\n".join(lines) + f"\n{END_MARK}"
    text = README.read_text(encoding="utf-8")
    pattern = re.compile(re.escape(START_MARK) + r".*?" + re.escape(END_MARK), re.S)
    if pattern.search(text):
        README.write_text(pattern.sub(lambda _: block, text), encoding="utf-8")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--workbook", type=Path, default=DEFAULT_WORKBOOK)
    args = parser.parse_args()
    build(args.workbook)
