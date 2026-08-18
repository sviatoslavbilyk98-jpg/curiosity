from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime, timezone
from io import BytesIO
from typing import Any, Iterable
from xml.sax.saxutils import escape
import re
import zipfile

INVALID_SHEET_CHARS = re.compile(r"[\\/*?:\[\]]")


def sanitize_sheet_name(name: str, used: set[str]) -> str:
    base = INVALID_SHEET_CHARS.sub("_", (name or "Стройка").strip())[:31] or "Стройка"
    if base.casefold() in {"накладні".casefold(), "результат".casefold()}:
        base = f"Проект {base}"[:31]
    candidate = base
    index = 2
    while candidate.casefold() in used:
        suffix = f" ({index})"
        candidate = (base[: 31 - len(suffix)] + suffix)
        index += 1
    used.add(candidate.casefold())
    return candidate


def col_letter(col: int) -> str:
    result = ""
    while col:
        col, rem = divmod(col - 1, 26)
        result = chr(65 + rem) + result
    return result


def cell_ref(row: int, col: int) -> str:
    return f"{col_letter(col)}{row}"


@dataclass
class Cell:
    value: Any = None
    style: int = 0
    formula: str | None = None


@dataclass
class Sheet:
    name: str
    cells: dict[tuple[int, int], Cell] = field(default_factory=dict)
    merges: list[str] = field(default_factory=list)
    widths: dict[int, float] = field(default_factory=dict)
    row_heights: dict[int, float] = field(default_factory=dict)
    freeze_row: int | None = None
    auto_filter: str | None = None

    def set(self, row: int, col: int, value: Any = None, style: int = 0, formula: str | None = None) -> None:
        self.cells[(row, col)] = Cell(value=value, style=style, formula=formula)

    def merge(self, start_row: int, start_col: int, end_row: int, end_col: int) -> None:
        self.merges.append(f"{cell_ref(start_row, start_col)}:{cell_ref(end_row, end_col)}")

    def set_row(self, row: int, values: Iterable[Any], start_col: int = 1, style: int = 0) -> None:
        for offset, value in enumerate(values):
            self.set(row, start_col + offset, value, style)


def _xml_text(value: str) -> str:
    preserve = " xml:space=\"preserve\"" if value[:1].isspace() or value[-1:].isspace() else ""
    return f"<is><t{preserve}>{escape(value)}</t></is>"


def _cell_xml(row: int, col: int, cell: Cell) -> str:
    ref = cell_ref(row, col)
    style_attr = f' s="{cell.style}"' if cell.style else ""
    if cell.formula is not None:
        formula = cell.formula[1:] if cell.formula.startswith("=") else cell.formula
        cached_value = cell.value if isinstance(cell.value, (int, float)) and not isinstance(cell.value, bool) else 0
        return f'<c r="{ref}"{style_attr}><f>{escape(formula)}</f><v>{cached_value}</v></c>'
    if cell.value is None:
        return f'<c r="{ref}"{style_attr}/>'
    if isinstance(cell.value, bool):
        return f'<c r="{ref}" t="b"{style_attr}><v>{1 if cell.value else 0}</v></c>'
    if isinstance(cell.value, (int, float)):
        return f'<c r="{ref}"{style_attr}><v>{cell.value}</v></c>'
    return f'<c r="{ref}" t="inlineStr"{style_attr}>{_xml_text(str(cell.value))}</c>'


def _worksheet_xml(sheet: Sheet) -> str:
    if sheet.cells:
        max_row = max(r for r, _ in sheet.cells)
        max_col = max(c for _, c in sheet.cells)
    else:
        max_row, max_col = 1, 1
    rows: dict[int, list[tuple[int, Cell]]] = {}
    for (row, col), cell in sheet.cells.items():
        rows.setdefault(row, []).append((col, cell))
    row_xml = []
    for row in sorted(rows):
        height = sheet.row_heights.get(row)
        ht_attr = f' ht="{height}" customHeight="1"' if height else ""
        cells = "".join(_cell_xml(row, col, cell) for col, cell in sorted(rows[row]))
        row_xml.append(f'<row r="{row}"{ht_attr}>{cells}</row>')

    cols_xml = ""
    if sheet.widths:
        pieces = []
        for col, width in sorted(sheet.widths.items()):
            pieces.append(f'<col min="{col}" max="{col}" width="{width}" customWidth="1"/>')
        cols_xml = f"<cols>{''.join(pieces)}</cols>"

    pane_xml = ""
    if sheet.freeze_row and sheet.freeze_row > 1:
        split = sheet.freeze_row - 1
        top = f"A{sheet.freeze_row}"
        pane_xml = f'<pane ySplit="{split}" topLeftCell="{top}" activePane="bottomLeft" state="frozen"/>'

    merges_xml = ""
    if sheet.merges:
        merges_xml = f'<mergeCells count="{len(sheet.merges)}">' + "".join(
            f'<mergeCell ref="{escape(m)}"/>' for m in sheet.merges
        ) + "</mergeCells>"

    filter_xml = f'<autoFilter ref="{escape(sheet.auto_filter)}"/>' if sheet.auto_filter else ""
    dimension = f"A1:{cell_ref(max_row, max_col)}"
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f'<dimension ref="{dimension}"/>'
        '<sheetViews><sheetView tabSelected="0" workbookViewId="0">'
        f'{pane_xml}'
        '</sheetView></sheetViews>'
        '<sheetFormatPr defaultRowHeight="15"/>'
        f'{cols_xml}'
        f'<sheetData>{"".join(row_xml)}</sheetData>'
        f'{filter_xml}{merges_xml}'
        '<pageMargins left="0.3" right="0.3" top="0.5" bottom="0.5" header="0.2" footer="0.2"/>'
        '<pageSetup orientation="landscape" fitToWidth="1" fitToHeight="0"/>'
        '</worksheet>'
    )


def _styles_xml() -> str:
    # Style ids documented in README and used below.
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="3">
    <numFmt numFmtId="164" formatCode="# ##0.00 [$₴-uk-UA]"/>
    <numFmt numFmtId="165" formatCode="0.00%"/>
    <numFmt numFmtId="166" formatCode="yyyy-mm-dd"/>
  </numFmts>
  <fonts count="6">
    <font><sz val="11"/><name val="Calibri"/><family val="2"/></font>
    <font><b/><sz val="14"/><name val="Calibri"/><family val="2"/><color rgb="FFFFFFFF"/></font>
    <font><b/><sz val="11"/><name val="Calibri"/><family val="2"/><color rgb="FFFFFFFF"/></font>
    <font><b/><sz val="11"/><name val="Calibri"/><family val="2"/><color rgb="FF111827"/></font>
    <font><b/><sz val="11"/><name val="Calibri"/><family val="2"/><color rgb="FF991B1B"/></font>
    <font><i/><sz val="10"/><name val="Calibri"/><family val="2"/><color rgb="FF475569"/></font>
  </fonts>
  <fills count="10">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFDB2777"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF0284C7"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFEDD5"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFBBF24"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF1F2937"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFD1FAE5"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFDE68A"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFF1F5F9"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="3">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border><left style="thin"><color rgb="FF334155"/></left><right style="thin"><color rgb="FF334155"/></right><top style="thin"><color rgb="FF334155"/></top><bottom style="thin"><color rgb="FF334155"/></bottom><diagonal/></border>
    <border><left style="medium"><color rgb="FF111827"/></left><right style="medium"><color rgb="FF111827"/></right><top style="medium"><color rgb="FF111827"/></top><bottom style="medium"><color rgb="FF111827"/></bottom><diagonal/></border>
  </borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="16">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="2" xfId="0" applyFill="1" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="2" fillId="3" borderId="1" xfId="0" applyFill="1" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="3" fillId="4" borderId="1" xfId="0" applyFill="1" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="3" fillId="5" borderId="1" xfId="0" applyFill="1" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="2" fillId="6" borderId="2" xfId="0" applyFill="1" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="166" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="165" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="0" fontId="3" fillId="9" borderId="1" xfId="0" applyFill="1" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="164" fontId="3" fillId="7" borderId="2" xfId="0" applyNumberFormat="1" applyFill="1" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="0" fontId="3" fillId="7" borderId="2" xfId="0" applyFill="1" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="164" fontId="3" fillId="8" borderId="2" xfId="0" applyNumberFormat="1" applyFill="1" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="164" fontId="4" fillId="8" borderId="2" xfId="0" applyNumberFormat="1" applyFill="1" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="165" fontId="3" fillId="7" borderId="2" xfId="0" applyNumberFormat="1" applyFill="1" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
  <dxfs count="0"/>
  <tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>'''


def _content_types(sheet_count: int) -> str:
    overrides = "".join(
        f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        for i in range(1, sheet_count + 1)
    )
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  {overrides}
</Types>'''


def _workbook_xml(sheets: list[Sheet]) -> str:
    sheet_nodes = "".join(
        f'<sheet name="{escape(sheet.name)}" sheetId="{idx}" r:id="rId{idx}"/>'
        for idx, sheet in enumerate(sheets, 1)
    )
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <bookViews><workbookView xWindow="0" yWindow="0" windowWidth="28800" windowHeight="16000"/></bookViews>
  <sheets>{sheet_nodes}</sheets>
  <calcPr calcId="191029" calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/>
</workbook>'''


def _workbook_rels(sheet_count: int) -> str:
    rels = "".join(
        f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>'
        for i in range(1, sheet_count + 1)
    )
    rels += f'<Relationship Id="rId{sheet_count + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{rels}</Relationships>'''


def _root_rels() -> str:
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>'''


def _core_props() -> str:
    now = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:creator>СтройФинанс</dc:creator><cp:lastModifiedBy>СтройФинанс</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified>
</cp:coreProperties>'''


def _app_props(sheet_names: list[str]) -> str:
    titles = "".join(f"<vt:lpstr>{escape(n)}</vt:lpstr>" for n in sheet_names)
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>СтройФинанс</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop>
  <HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Листы</vt:lpstr></vt:variant><vt:variant><vt:i4>{len(sheet_names)}</vt:i4></vt:variant></vt:vector></HeadingPairs>
  <TitlesOfParts><vt:vector size="{len(sheet_names)}" baseType="lpstr">{titles}</vt:vector></TitlesOfParts>
  <Company></Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>1.0</AppVersion>
</Properties>'''


def _excel_serial(date_text: str | None) -> float | None:
    if not date_text:
        return None
    try:
        dt = datetime.strptime(date_text[:10], "%Y-%m-%d")
    except (TypeError, ValueError):
        return None
    return (dt - datetime(1899, 12, 30)).days


def _sum_formula(cells: list[str]) -> str:
    if not cells:
        return "0"
    return "+".join(cells)


def _build_project_sheet(project: dict[str, Any], sheet_name: str, metrics: dict[str, Any], rate: float, result_row: int | None = None) -> Sheet:
    sheet = Sheet(sheet_name)
    sheet.widths = {1: 16, 2: 34, 3: 15, 4: 15, 5: 18, 6: 16, 7: 3, 8: 35, 9: 18}
    sheet.row_heights[1] = 28
    sheet.set(1, 1, project.get("name", sheet_name), 1)
    sheet.merge(1, 1, 1, 9)
    sheet.set(2, 8, "Ключевые показатели", 5)
    sheet.merge(2, 8, 2, 9)

    summary_labels = [
        "Общий оборот",
        "Общие безналичные расходы",
        "Остаток безналичных средств",
        f"Расход на конвертацию ({rate:.0%})",
        "Заработная плата",
        "Прямые наличные расходы",
        "Прямые наличные расходы всего",
        "Прибыль объекта до общих накладных",
        "Доля общих накладных",
        "Чистая прибыль",
    ]
    for idx, label in enumerate(summary_labels, start=3):
        sheet.set(idx, 8, label, 9)
        sheet.set(idx, 9, None, 6)

    entries = project.get("entries", [])
    groups = {
        "revenue": [e for e in entries if e.get("section") == "revenue"],
        "cashless": [e for e in entries if e.get("section") == "cashless"],
        "asset": [e for e in entries if e.get("section") == "asset"],
        "cash": [e for e in entries if e.get("section") == "cash"],
        "payroll": [e for e in entries if e.get("section") == "payroll"],
    }
    row = 3
    total_cells: dict[str, str] = {}

    def add_section(title: str, section_key: str, section_style: int, show_category: bool = False) -> None:
        nonlocal row
        sheet.set(row, 1, title, section_style)
        sheet.merge(row, 1, row, 6)
        row += 1
        headers = ["№", "Описание", "Категория", "Дата", "Сумма", "Примечание"]
        sheet.set_row(row, headers, 1, 5)
        row += 1
        start_data = row
        data = groups[section_key]
        if not data:
            sheet.set(row, 1, 1, 13)
            sheet.set(row, 2, "", 13)
            sheet.set(row, 3, "", 13)
            sheet.set(row, 4, "", 13)
            sheet.set(row, 5, 0, 6)
            sheet.set(row, 6, "", 13)
            row += 1
        else:
            for i, entry in enumerate(data, start=1):
                sheet.set(row, 1, i, 13)
                sheet.set(row, 2, entry.get("description", ""), 13)
                sheet.set(row, 3, entry.get("category", "") if show_category else "", 13)
                serial = _excel_serial(entry.get("date"))
                if serial is None:
                    sheet.set(row, 4, entry.get("date", ""), 13)
                else:
                    sheet.set(row, 4, serial, 7)
                sheet.set(row, 5, float(entry.get("amount", 0) or 0), 6)
                sheet.set(row, 6, entry.get("note", ""), 13)
                row += 1
        end_data = row - 1
        sheet.set(row, 1, "Итого", 11)
        sheet.merge(row, 1, row, 4)
        section_total = sum(float(item.get("amount", 0) or 0) for item in data)
        sheet.set(row, 5, section_total, 10, formula=f"SUM(E{start_data}:E{end_data})")
        sheet.set(row, 6, "", 11)
        total_cells[section_key] = f"E{row}"
        row += 2

    add_section("1. Приходная часть", "revenue", 2)
    add_section("2. Общие безналичные расходы", "cashless", 2)
    add_section("2.1 Безналичные расходы на основные средства", "asset", 3)
    add_section("3. Прямые наличные расходы объекта", "cash", 4, True)
    add_section("4. Дневник заработной платы", "payroll", 3, True)

    revenue_value = float(metrics.get("turnover", sum(float(e.get("amount", 0) or 0) for e in groups["revenue"])))
    cashless_value = float(metrics.get("cashless", sum(float(e.get("amount", 0) or 0) for e in groups["cashless"] + groups["asset"])))
    balance_value = float(metrics.get("balance", revenue_value - cashless_value))
    conversion_value = float(metrics.get("conversion", balance_value * rate))
    payroll_value = sum(float(e.get("amount", 0) or 0) for e in groups["payroll"])
    cash_value = sum(float(e.get("amount", 0) or 0) for e in groups["cash"])
    direct_cash_value = float(metrics.get("cash_expenses", payroll_value + cash_value))
    profit_before_value = float(metrics.get("profit_before_overhead", revenue_value - cashless_value - conversion_value - direct_cash_value))
    allocated_value = float(metrics.get("allocated_overhead", 0))
    net_value = float(metrics.get("net_profit", profit_before_value - allocated_value))
    sheet.set(3, 9, revenue_value, 10, formula=f"{total_cells['revenue']}")
    sheet.set(4, 9, cashless_value, 10, formula=f"{total_cells['cashless']}+{total_cells['asset']}")
    sheet.set(5, 9, balance_value, 10, formula="I3-I4")
    sheet.set(6, 9, conversion_value, 10, formula=f"I5*{rate}")
    sheet.set(7, 9, payroll_value, 10, formula=f"{total_cells['payroll']}")
    sheet.set(8, 9, cash_value, 10, formula=f"{total_cells['cash']}")
    sheet.set(9, 9, direct_cash_value, 10, formula="I7+I8")
    sheet.set(10, 9, profit_before_value, 10, formula="I3-I4-I6-I9")
    if result_row is not None:
        sheet.set(11, 9, allocated_value, 12, formula=f"'Результат'!E{result_row}")
        sheet.set(12, 9, net_value, 14, formula=f"'Результат'!I{result_row}")
    else:
        sheet.set(11, 9, allocated_value, 12)
        sheet.set(12, 9, net_value, 14)
    sheet.freeze_row = 3
    return sheet


def _build_overhead_sheet(overhead_entries: list[dict[str, Any]], rate: float) -> Sheet:
    overhead_sheet = Sheet("Накладні")
    overhead_sheet.widths = {1: 18, 2: 32, 3: 22, 4: 16, 5: 18, 6: 18, 7: 4, 8: 26, 9: 18}
    overhead_sheet.row_heights[1] = 28
    overhead_sheet.set(1, 1, "ОБЩИЕ НАКЛАДНЫЕ РАСХОДЫ", 1)
    overhead_sheet.merge(1, 1, 1, 9)
    overhead_sheet.set(2, 8, "Сводка", 5)
    overhead_sheet.merge(2, 8, 2, 9)
    overhead_sheet.set(3, 8, "Безналичные накладные", 9)
    overhead_sheet.set(4, 8, f"Корректировка конвертации {rate:.0%}", 9)
    overhead_sheet.set(5, 8, "Наличные накладные", 9)
    overhead_sheet.set(6, 8, "Накладные к распределению", 11)
    for r in range(3, 7):
        overhead_sheet.set(r, 9, None, 10 if r < 6 else 12)

    overhead_sheet.set_row(3, ["№", "Описание", "Группа", "Категория", "Дата", "Сумма"], 1, 5)
    row = 4
    cashless_cells: list[str] = []
    cash_cells: list[str] = []
    for idx, entry in enumerate(overhead_entries, start=1):
        overhead_sheet.set(row, 1, idx, 13)
        overhead_sheet.set(row, 2, entry.get("description", ""), 13)
        group = entry.get("group_type", "cashless")
        overhead_sheet.set(row, 3, "Безналичные" if group == "cashless" else "Наличные", 13)
        overhead_sheet.set(row, 4, entry.get("category", ""), 13)
        serial = _excel_serial(entry.get("date"))
        if serial is None:
            overhead_sheet.set(row, 5, entry.get("date", ""), 13)
        else:
            overhead_sheet.set(row, 5, serial, 7)
        overhead_sheet.set(row, 6, float(entry.get("amount", 0) or 0), 6)
        (cashless_cells if group == "cashless" else cash_cells).append(f"F{row}")
        row += 1
    if not overhead_entries:
        overhead_sheet.set_row(row, [1, "", "Безналичные", "", "", 0], 1, 13)
        overhead_sheet.set(row, 6, 0, 6)
        cashless_cells.append(f"F{row}")
        row += 1
    overhead_sheet.set(row, 1, "Итого", 11)
    overhead_sheet.merge(row, 1, row, 5)
    overhead_sheet.set(row, 6, sum(float(e.get("amount", 0) or 0) for e in overhead_entries), 10, formula=f"SUM(F4:F{row-1})")
    overhead_cashless_value = sum(float(e.get("amount", 0) or 0) for e in overhead_entries if e.get("group_type") == "cashless")
    overhead_cash_value = sum(float(e.get("amount", 0) or 0) for e in overhead_entries if e.get("group_type") == "cash")
    overhead_adjustment_value = overhead_cashless_value * rate
    distributable_value = overhead_cashless_value + overhead_cash_value - overhead_adjustment_value
    overhead_sheet.set(3, 9, overhead_cashless_value, 10, formula=_sum_formula(cashless_cells))
    overhead_sheet.set(4, 9, overhead_adjustment_value, 10, formula=f"I3*{rate}")
    overhead_sheet.set(5, 9, overhead_cash_value, 10, formula=_sum_formula(cash_cells))
    overhead_sheet.set(6, 9, distributable_value, 12, formula="I3+I5-I4")
    overhead_sheet.freeze_row = 4
    overhead_sheet.auto_filter = f"A3:F{max(4, row-1)}"
    return overhead_sheet


def _build_result_sheet(projects: list[dict[str, Any]], project_sheet_names: list[str], result_rows: list[dict[str, Any]], result_totals: dict[str, Any], rate: float, linked: bool = True) -> Sheet:
    result_sheet = Sheet("Результат")
    result_sheet.widths = {1: 8, 2: 32, 3: 20, 4: 20, 5: 20, 6: 20, 7: 20, 8: 20, 9: 20, 10: 16}
    result_sheet.row_heights[1] = 28
    result_sheet.set(1, 1, "СВОДНЫЙ РЕЗУЛЬТАТ ПО ВСЕМ СТРОЙКАМ", 1)
    result_sheet.merge(1, 1, 1, 10)
    headers = [
        "№ п/п", "Наименование", "Лист", "Общий оборот", "Накладные (доля)",
        "Общие безналичные расходы", "Расход на конвертацию", "Наличные расходы",
        "Чистая прибыль", "Процент прибыли",
    ]
    result_sheet.set_row(2, headers, 1, 5)
    first_project_row = 3
    project_metrics = {row.get("project_id"): row for row in result_rows}
    for idx, (project, sheet_name) in enumerate(zip(projects, project_sheet_names), start=1):
        r = 2 + idx
        quoted = sheet_name.replace("'", "''")
        result_sheet.set(r, 1, idx, 13)
        result_sheet.set(r, 2, project.get("name", sheet_name), 13)
        result_sheet.set(r, 3, sheet_name, 13)
        metrics = project_metrics.get(project.get("id"), {})
        if linked:
            result_sheet.set(r, 4, float(metrics.get("turnover", 0)), 6, formula=f"'{quoted}'!I3")
            result_sheet.set(r, 5, float(metrics.get("allocated_overhead", 0)), 6, formula=f"IFERROR($E${3+len(projects)}*D{r}/$D${3+len(projects)},0)")
            result_sheet.set(r, 6, float(metrics.get("cashless", 0)), 6, formula=f"'{quoted}'!I4")
            result_sheet.set(r, 7, float(metrics.get("conversion", 0)), 6, formula=f"'{quoted}'!I6")
            result_sheet.set(r, 8, float(metrics.get("cash_expenses", 0)), 6, formula=f"'{quoted}'!I9")
            result_sheet.set(r, 9, float(metrics.get("net_profit", 0)), 14, formula=f"D{r}-E{r}-F{r}-G{r}-H{r}")
            result_sheet.set(r, 10, float(metrics.get("profit_percent", 0)), 8, formula=f"IFERROR(I{r}/D{r},0)")
        else:
            result_sheet.set(r, 4, float(metrics.get("turnover", 0)), 6)
            result_sheet.set(r, 5, float(metrics.get("allocated_overhead", 0)), 6)
            result_sheet.set(r, 6, float(metrics.get("cashless", 0)), 6)
            result_sheet.set(r, 7, float(metrics.get("conversion", 0)), 6)
            result_sheet.set(r, 8, float(metrics.get("cash_expenses", 0)), 6)
            result_sheet.set(r, 9, float(metrics.get("net_profit", 0)), 14)
            result_sheet.set(r, 10, float(metrics.get("profit_percent", 0)), 8)

    total_row = 3 + len(projects)
    if not projects:
        first_project_row = total_row
    result_sheet.set(total_row, 1, "", 11)
    result_sheet.set(total_row, 2, "ВСЕГО", 11)
    result_sheet.set(total_row, 3, "", 11)
    total_values = {
        4: float(result_totals.get("turnover", 0)),
        5: float(result_totals.get("allocated_overhead", 0)),
        6: float(result_totals.get("cashless", 0)),
        7: float(result_totals.get("conversion", 0)),
        8: float(result_totals.get("cash_expenses", 0)),
        9: float(result_totals.get("net_profit", 0)),
    }
    for col in range(4, 10):
        formula = f"SUM({cell_ref(first_project_row, col)}:{cell_ref(total_row-1, col)})" if projects else "0"
        if col == 5:
            formula = "'Накладні'!I6" if linked else None
        style = 10 if col != 9 else 14
        result_sheet.set(total_row, col, total_values[col], style, formula=formula)
    result_sheet.set(total_row, 10, float(result_totals.get("profit_percent", 0)), 15, formula=f"IFERROR(I{total_row}/D{total_row},0)")
    result_sheet.set(total_row + 2, 2, "Коэффициент общих накладных", 9)
    result_sheet.set(total_row + 2, 3, float(result_totals.get("overhead_coefficient", 0)), 15, formula=f"IFERROR(E{total_row}/D{total_row},0)")
    result_sheet.set(total_row + 2, 5, "Ставка конвертации", 9)
    result_sheet.set(total_row + 2, 6, rate, 15)
    result_sheet.freeze_row = 3
    result_sheet.auto_filter = f"A2:J{max(2, total_row-1)}"
    return result_sheet


def _zip_workbook(sheets: list[Sheet]) -> bytes:
    buffer = BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", _content_types(len(sheets)))
        zf.writestr("_rels/.rels", _root_rels())
        zf.writestr("xl/workbook.xml", _workbook_xml(sheets))
        zf.writestr("xl/_rels/workbook.xml.rels", _workbook_rels(len(sheets)))
        zf.writestr("xl/styles.xml", _styles_xml())
        zf.writestr("docProps/core.xml", _core_props())
        zf.writestr("docProps/app.xml", _app_props([s.name for s in sheets]))
        for idx, sheet in enumerate(sheets, start=1):
            zf.writestr(f"xl/worksheets/sheet{idx}.xml", _worksheet_xml(sheet))
    return buffer.getvalue()


def _build_contracts_sheet(contracts: list[dict[str, Any]], project_name: str) -> Sheet:
    sheet = Sheet("Договори")
    sheet.widths = {1: 8, 2: 40, 3: 26, 4: 18, 5: 18, 6: 22}
    sheet.row_heights[1] = 28
    sheet.set(1, 1, "ДОГОВОРЫ И АКТЫ ВЫПОЛНЕННЫХ РАБОТ", 1)
    sheet.merge(1, 1, 1, 6)
    sheet.set(2, 1, f"Стройка: {project_name}", 9)
    sheet.merge(2, 1, 2, 6)

    sheet.set(4, 1, "ДОГОВОРЫ", 5)
    sheet.merge(4, 1, 4, 4)
    sheet.set_row(5, ["№", "Название", "Дата подписания", "Сумма"], 1, 5)
    row = 6
    contract_cells: list[str] = []
    for idx, contract in enumerate(contracts, start=1):
        sheet.set(row, 1, idx, 13)
        sheet.set(row, 2, contract.get("name", ""), 13)
        serial = _excel_serial(contract.get("date"))
        if serial is None:
            sheet.set(row, 3, contract.get("date", ""), 13)
        else:
            sheet.set(row, 3, serial, 7)
        sheet.set(row, 4, float(contract.get("amount", 0) or 0), 6)
        contract_cells.append(f"D{row}")
        row += 1
    if not contracts:
        sheet.set_row(row, [1, "", "", 0], 1, 13)
        sheet.set(row, 4, 0, 6)
        contract_cells.append(f"D{row}")
        row += 1
    contracts_total_row = row
    sheet.set(row, 1, "Итого договоры", 11)
    sheet.merge(row, 1, row, 3)
    contracts_total = sum(float(c.get("amount", 0) or 0) for c in contracts)
    sheet.set(row, 4, contracts_total, 10, formula=_sum_formula(contract_cells))
    row += 2

    sheet.set(row, 1, "АКТЫ ВЫПОЛНЕННЫХ РАБОТ", 5)
    sheet.merge(row, 1, row, 6)
    row += 1
    sheet.set_row(row, ["№", "Договор", "Номер акта", "Дата подписания", "Сумма"], 1, 5)
    acts_row_start = row + 1
    row += 1
    act_cells: list[str] = []
    acts: list[tuple[dict[str, Any], dict[str, Any]]] = []
    for contract in contracts:
        for act in contract.get("acts", []):
            acts.append((contract, act))
    for idx, (contract, act) in enumerate(acts, start=1):
        sheet.set(row, 1, idx, 13)
        sheet.set(row, 2, contract.get("name", ""), 13)
        sheet.set(row, 3, act.get("number", ""), 13)
        serial = _excel_serial(act.get("date"))
        if serial is None:
            sheet.set(row, 4, act.get("date", ""), 13)
        else:
            sheet.set(row, 4, serial, 7)
        sheet.set(row, 5, float(act.get("amount", 0) or 0), 6)
        act_cells.append(f"E{row}")
        row += 1
    if not acts:
        sheet.set_row(row, [1, "", "", "", 0], 1, 13)
        sheet.set(row, 5, 0, 6)
        act_cells.append(f"E{row}")
        row += 1
    acts_total_row = row
    sheet.set(row, 1, "Итого акты", 11)
    sheet.merge(row, 1, row, 4)
    acts_total = sum(float(act.get("amount", 0) or 0) for _, act in acts)
    sheet.set(row, 5, acts_total, 10, formula=_sum_formula(act_cells))
    row += 1
    sheet.set(row, 1, "Остаток (договоры − акты)", 11)
    sheet.merge(row, 1, row, 4)
    balance = contracts_total - acts_total
    sheet.set(row, 5, balance, 14, formula=f"D{contracts_total_row}-E{acts_total_row}")
    sheet.freeze_row = 4
    return sheet


def build_workbook(state: dict[str, Any]) -> bytes:
    projects = state.get("projects", [])
    overhead_entries = state.get("overheads", [])
    rate = float(state.get("settings", {}).get("conversion_rate", 0.14) or 0.14)
    result_data = state.get("results", {})
    result_rows = result_data.get("rows", [])
    result_totals = result_data.get("totals", {})
    project_metrics = {row.get("project_id"): row for row in result_rows}

    used: set[str] = set()
    project_sheet_names = [sanitize_sheet_name(p.get("name", "Стройка"), used) for p in projects]

    sheets: list[Sheet] = []
    for idx, (project, sheet_name) in enumerate(zip(projects, project_sheet_names)):
        sheets.append(_build_project_sheet(project, sheet_name, project_metrics.get(project.get("id"), {}), rate, result_row=3 + idx))
    sheets.append(_build_overhead_sheet(overhead_entries, rate))
    sheets.append(_build_result_sheet(projects, project_sheet_names, result_rows, result_totals, rate, linked=True))
    contracts = state.get("contracts", [])
    if contracts:
        sheets.append(_build_contracts_sheet(contracts, projects[0].get("name", "Стройка") if projects else "Стройка"))
    return _zip_workbook(sheets)


def build_page_workbook(state: dict[str, Any], page: str, project_id: int | None = None) -> bytes:
    projects = state.get("projects", [])
    overhead_entries = state.get("overheads", [])
    rate = float(state.get("settings", {}).get("conversion_rate", 0.14) or 0.14)
    result_data = state.get("results", {})
    result_rows = result_data.get("rows", [])
    result_totals = result_data.get("totals", {})
    project_metrics = {row.get("project_id"): row for row in result_rows}

    used: set[str] = set()
    if page == "overheads":
        sheets = [_build_overhead_sheet(overhead_entries, rate)]
    elif page in ("results", "dashboard"):
        project_sheet_names = [sanitize_sheet_name(p.get("name", "Стройка"), used) for p in projects]
        sheets = [_build_result_sheet(projects, project_sheet_names, result_rows, result_totals, rate, linked=False)]
    elif page == "project":
        project = next((p for p in projects if p.get("id") == project_id), None)
        if project is None:
            raise ValueError("Стройка не найдена")
        sheet_name = sanitize_sheet_name(project.get("name", "Стройка"), used)
        sheets = [_build_project_sheet(project, sheet_name, project_metrics.get(project_id, {}), rate)]
    elif page == "contracts":
        project = next((p for p in projects if p.get("id") == project_id), None)
        project_name = project.get("name", "Стройка") if project else "Стройка"
        project_contracts = [c for c in state.get("contracts", []) if c.get("project_id") == project_id]
        sheets = [_build_contracts_sheet(project_contracts, project_name)]
    else:
        raise ValueError("Неизвестный тип экспорта")
    return _zip_workbook(sheets)
