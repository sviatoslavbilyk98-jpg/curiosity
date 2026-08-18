from __future__ import annotations

import re
import zipfile
from datetime import datetime, timedelta
from io import BytesIO
from typing import Any
from xml.etree import ElementTree as ET

MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"
RID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
EPOCH = datetime(1899, 12, 30)

SECTION_KEYWORDS = [
    ("asset", ("основн",)),
    ("cashless", ("безналичн",)),
    ("cash", ("наличные",)),
    ("payroll", ("заработн", "дневник")),
    ("revenue", ("приходн",)),
]
DATE_FORMATS = ("%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y", "%d.%m.%y", "%d/%m/%y", "%Y/%m/%d")
RATE_RE = re.compile(r"(\d+(?:[.,]\d+)?)\s*%")


def _col_number(col: str) -> int:
    num = 0
    for ch in col:
        num = num * 26 + (ord(ch) - 64)
    return num


def _col_letter(num: int) -> str:
    letters = ""
    while num:
        num, rem = divmod(num - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def _read_shared(zf: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    return ["".join(t.text or "" for t in si.iter(MAIN + "t")) for si in root.findall(MAIN + "si")]


def _parse_grid(zf: zipfile.ZipFile, path: str, shared: list[str]) -> dict[tuple[int, str], str]:
    root = ET.fromstring(zf.read(path))
    grid: dict[tuple[int, str], str] = {}
    for cell in root.iter(MAIN + "c"):
        ref = cell.get("r")
        if not ref:
            continue
        match = re.match(r"([A-Z]+)(\d+)", ref)
        if not match:
            continue
        col, row = match.group(1), int(match.group(2))
        value = ""
        vnode = cell.find(MAIN + "v")
        cell_type = cell.get("t")
        if cell_type == "s" and vnode is not None and vnode.text is not None:
            idx = int(vnode.text)
            if 0 <= idx < len(shared):
                value = shared[idx]
        elif cell_type == "inlineStr":
            value = "".join(t.text or "" for t in cell.iter(MAIN + "t"))
        elif vnode is not None and vnode.text is not None:
            value = vnode.text
        if value:
            grid[(row, col)] = value
    return grid


def _read_sheets(raw: bytes) -> list[dict[str, Any]]:
    try:
        zf = zipfile.ZipFile(BytesIO(raw))
    except (zipfile.BadZipFile, OSError):
        raise ValueError("Файл не является книгой Excel (.xlsx)")
    with zf:
        if "xl/workbook.xml" not in zf.namelist():
            raise ValueError("Файл не является книгой Excel (.xlsx)")
        shared = _read_shared(zf)
        wb = ET.fromstring(zf.read("xl/workbook.xml"))
        rels = {}
        rel_path = "xl/_rels/workbook.xml.rels"
        if rel_path in zf.namelist():
            rel_root = ET.fromstring(zf.read(rel_path))
            for rel in rel_root.findall(REL + "Relationship"):
                rels[rel.get("Id")] = rel.get("Target", "")
        sheets = []
        for node in wb.iter(MAIN + "sheet"):
            name = node.get("name", "")
            rid = node.get(RID)
            target = rels.get(rid or "", "")
            if not target:
                continue
            if target.startswith("/"):
                path = target.lstrip("/")
            else:
                path = "xl/" + target.lstrip("/")
            sheets.append({"name": name, "grid": _parse_grid(zf, path, shared)})
    return sheets


def _get(grid: dict[tuple[int, str], str], row: int, col: str) -> str:
    return (grid.get((row, col)) or "").strip()


def _parse_date(value: str) -> str:
    text = value.strip()
    if not text:
        return ""
    try:
        serial = float(text)
    except ValueError:
        serial = None
    if serial is not None:
        if 0 < serial < 3000000:
            try:
                return (EPOCH + timedelta(days=int(serial))).strftime("%Y-%m-%d")
            except OverflowError:
                return text
        serial = None
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(text[:10], fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return text


def _parse_amount(value: str) -> float:
    text = value.strip().replace(" ", "").replace(",", ".")
    try:
        return round(float(text), 2)
    except ValueError:
        return 0.0


def _detect_section(title: str) -> str | None:
    low = title.casefold()
    for section, keywords in SECTION_KEYWORDS:
        for keyword in keywords:
            if keyword in low:
                return section
    return None


def _parse_project(grid: dict[tuple[int, str], str]) -> dict[str, Any] | None:
    name = _get(grid, 1, "A")
    if not name:
        return None
    entries: list[dict[str, Any]] = []
    section: str | None = None
    waiting_header = False
    rows = sorted({row for row, _ in grid})
    for row in rows:
        a = _get(grid, row, "A")
        b = _get(grid, row, "B")
        detected = _detect_section(a)
        if detected is not None:
            section = detected
            waiting_header = True
            continue
        if section is None:
            continue
        if waiting_header:
            if b.casefold() == "описание":
                waiting_header = False
            continue
        if a.casefold() in {"итого", "всего"}:
            section = None
            continue
        description = b
        amount = _parse_amount(_get(grid, row, "E"))
        if not description and not amount:
            continue
        category = _get(grid, row, "C") if section in ("cash", "payroll") else ""
        entries.append({
            "section": section,
            "category": category,
            "description": description,
            "amount": amount,
            "date": _parse_date(_get(grid, row, "D")),
            "note": _get(grid, row, "F"),
        })
    return {"name": name, "entries": entries}


def _parse_overheads(grid: dict[tuple[int, str], str]) -> list[dict[str, Any]]:
    entries: list[dict[str, Any]] = []
    header_row: int | None = None
    for row, col in sorted(grid):
        if col == "B" and _get(grid, row, "B").casefold() == "описание" and _get(grid, row, "C").casefold() == "группа":
            header_row = row
            break
    if header_row is None:
        for row, col in sorted(grid):
            if col == "A" and _get(grid, row, "A").casefold() == "№":
                header_row = row
                break
    if header_row is None:
        return entries
    for row in sorted({r for r, _ in grid}):
        if row <= header_row:
            continue
        a = _get(grid, row, "A")
        if a.casefold() in {"итого", "всего"}:
            break
        description = _get(grid, row, "B")
        amount = _parse_amount(_get(grid, row, "F"))
        if not description and not amount:
            continue
        group = _get(grid, row, "C")
        group_type = "cash" if group.casefold() == "наличные" else "cashless"
        entries.append({
            "group_type": group_type,
            "category": _get(grid, row, "D"),
            "description": description,
            "amount": amount,
            "date": _parse_date(_get(grid, row, "E")),
        })
    return entries


def _extract_rate(sheets: list[dict[str, Any]]) -> float:
    for sheet in sheets:
        grid = sheet["grid"]
        for (row, col), value in grid.items():
            if "ставка" in value.casefold() and "конверт" in value.casefold():
                start = _col_number(col) + 1
                for col_num in range(start, start + 10):
                    candidate = _get(grid, row, _col_letter(col_num))
                    if candidate:
                        try:
                            rate = float(candidate.replace(",", "."))
                        except ValueError:
                            continue
                        if 0 <= rate <= 1:
                            return rate
    for sheet in sheets:
        for value in sheet["grid"].values():
            if "конверт" in value.casefold():
                match = RATE_RE.search(value)
                if match:
                    rate = float(match.group(1).replace(",", ".")) / 100
                    if 0 <= rate <= 1:
                        return rate
    return 0.14


def parse_workbook(raw: bytes) -> dict[str, Any]:
    sheets = _read_sheets(raw)
    if not sheets:
        raise ValueError("В книге нет листов")
    rate = _extract_rate(sheets)
    projects: list[dict[str, Any]] = []
    overheads: list[dict[str, Any]] = []
    for sheet in sheets:
        name_low = sheet["name"].casefold()
        if "накладн" in name_low:
            overheads.extend(_parse_overheads(sheet["grid"]))
            continue
        if "результат" in name_low:
            continue
        project = _parse_project(sheet["grid"])
        if project is not None:
            projects.append(project)
    if not projects and not overheads:
        raise ValueError("Не удалось распознать данные в выбранной книге Excel")
    return {
        "settings": {"conversion_rate": rate},
        "projects": projects,
        "overheads": overheads,
    }
