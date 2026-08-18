from __future__ import annotations

from datetime import datetime, timezone
from io import BytesIO
from typing import Any
from xml.sax.saxutils import escape
import zipfile


def _para(text: str, bold: bool = False, size: int = 22, color: str | None = None, italic: bool = False) -> str:
    color_xml = f'<w:color w:val="{color}"/>' if color else ""
    rpr = f'<w:rPr>{"<w:b/>" if bold else ""}{"<w:i/>" if italic else ""}<w:sz w:val="{size}"/>{color_xml}</w:rPr>'
    return f'<w:p><w:pPr><w:spacing w:after="80" w:line="240" w:lineRule="auto"/></w:pPr><w:r>{rpr}<w:t xml:space="preserve">{escape(text)}</w:t></w:r></w:p>'


def _heading(text: str, level: int = 1) -> str:
    if level == 1:
        return f'<w:p><w:pPr><w:spacing w:before="240" w:after="120"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="32"/></w:rPr><w:t xml:space="preserve">{escape(text)}</w:t></w:r></w:p>'
    return f'<w:p><w:pPr><w:spacing w:before="200" w:after="80"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="24"/></w:rPr><w:t xml:space="preserve">{escape(text)}</w:t></w:r></w:p>'


def _table(headers: list[str], rows: list[list[str]]) -> str:
    def cell(text: str, header: bool = False) -> str:
        rpr = '<w:rPr><w:b/><w:sz w:val="20"/></w:rPr>' if header else '<w:sz w:val="20"/>'
        shading = '<w:shd w:val="clear" w:color="auto" w:fill="D9E2F3"/>' if header else ""
        return f'<w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/>{shading}</w:tcPr><w:p><w:pPr><w:spacing w:after="40" w:line="240" w:lineRule="auto"/></w:pPr><w:r>{rpr}<w:t xml:space="preserve">{escape(text)}</w:t></w:r></w:p></w:tc>'
    border = ('<w:tblBorders>'
              '<w:top w:val="single" w:sz="4" w:space="0" w:color="A9B7CC"/>'
              '<w:left w:val="single" w:sz="4" w:space="0" w:color="A9B7CC"/>'
              '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="A9B7CC"/>'
              '<w:right w:val="single" w:sz="4" w:space="0" w:color="A9B7CC"/>'
              '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="A9B7CC"/>'
              '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="A9B7CC"/>'
              '</w:tblBorders>')
    head = "".join(cell(h, True) for h in headers)
    body = "".join("".join(cell(t) for t in r) for r in rows)
    return (f'<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>{border}'
            f'<w:tblLayout w:type="autofit"/><w:tblLook w:val="04A0"/></w:tblPr>'
            f'<w:tblGrid>{"".join("<w:gridCol w:w=\"2000\"/>" for _ in headers)}</w:tblGrid>'
            f'<w:tr>{head}</w:tr>{"".join(f"<w:tr>{r}</w:tr>" for r in rows)}</w:tbl>'
            f'<w:p><w:pPr><w:spacing w:after="120"/></w:pPr></w:p>')


def build_exec_docx(state: dict[str, Any], project_id: int | None = None) -> bytes:
    projects = state.get("projects", [])
    contracts = state.get("contracts", [])
    if project_id is not None:
        project = next((p for p in projects if p.get("id") == project_id), None)
        if project is not None:
            projects = [project]
            contracts = [c for c in contracts if c.get("project_id") == project_id]

    body = [_heading("ВИКОНАВЧА ДОКУМЕНТАЦІЯ", 1)]
    generated = datetime.now(timezone.utc).strftime("%d.%m.%Y %H:%M")
    body.append(_para(f"Сформовано: {generated}", italic=True, size=18))

    for project in projects:
        body.append(_heading(str(project.get("name", "")), 1))
        p_contracts = [c for c in contracts if c.get("project_id") == project.get("id")]
        contracts_sum = sum(float(c.get("amount", 0) or 0) for c in p_contracts)
        acts = [(c, a) for c in p_contracts for a in c.get("acts", [])]
        acts_sum = sum(float(a.get("amount", 0) or 0) for _, a in acts)

        body.append(_heading("Договори", 2))
        if p_contracts:
            rows = [[c.get("name", ""), c.get("date", ""), f"{float(c.get('amount', 0) or 0):,.2f}"] for c in p_contracts]
            body.append(_table(["Назва договору", "Дата підписання", "Сума, грн"], rows))
        else:
            body.append(_para("Договорів немає."))
        body.append(_para(f"Усього за договорами: {contracts_sum:,.2f} грн", bold=True))

        body.append(_heading("Акти виконаних робіт", 2))
        if acts:
            rows = [[c.get("name", ""), a.get("number", ""), a.get("date", ""), f"{float(a.get('amount', 0) or 0):,.2f}"] for c, a in acts]
            body.append(_table(["Договір", "№ акта", "Дата", "Сума, грн"], rows))
        else:
            body.append(_para("Актів немає."))
        body.append(_para(f"Усього за актами: {acts_sum:,.2f} грн", bold=True))

        balance = contracts_sum - acts_sum
        color = "C00000" if balance < 0 else "1F6E43"
        body.append(_para(f"Залишок (договори − акти): {balance:,.2f} грн", bold=True, color=color))

    document_xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                    '<w:body>' + "".join(body) +
                    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>'
                    '</w:body></w:document>')

    content_types = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                     '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                     '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                     '<Default Extension="xml" ContentType="application/xml"/>'
                     '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
                     '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
                     '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
                     '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
                     '</Types>')

    root_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                 '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                 '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
                 '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
                 '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
                 '</Relationships>')

    doc_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
                '</Relationships>')

    styles_xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                  '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                  '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:eastAsia="Calibri" w:hAnsi="Calibri"/>'
                  '<w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="uk-UA"/></w:rPr></w:rPrDefault></w:docDefaults>'
                  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>'
                  '<w:qFormat/><w:pPr><w:spacing w:line="240" w:lineRule="auto"/></w:pPr></w:style></w:styles>')

    core_xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
                'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
                'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
                '<dc:title>Виконавча документація</dc:title>'
                f'<dcterms:created xsi:type="dcterms:W3CDTF">{datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")}</dcterms:created>'
                '</cp:coreProperties>')

    app_xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
               '<Application>СтройФинанс</Application></Properties>')

    buf = BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", root_rels)
        zf.writestr("word/document.xml", document_xml)
        zf.writestr("word/_rels/document.xml.rels", doc_rels)
        zf.writestr("word/styles.xml", styles_xml)
        zf.writestr("docProps/core.xml", core_xml)
        zf.writestr("docProps/app.xml", app_xml)
    return buf.getvalue()
