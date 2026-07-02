from __future__ import annotations

import io
from pathlib import Path

from docx import Document

from app.models.translation_unit import TranslationUnit, UnitType


def _get_style_name(document: Document, preferred: str | None, fallback: str) -> str:
    for style_name in [preferred, fallback]:
        if style_name and style_name in document.styles:
            return style_name
    return "Normal"


def _append_unit_to_document(document: Document, unit: TranslationUnit) -> None:
    if unit.type == UnitType.TABLE_CELL:
        return

    if unit.type == UnitType.HEADING:
        level = unit.formatting.get("level") or unit.metadata.get("level") or 1
        style_name = _get_style_name(document, f"Heading {level}", "Heading 1")
        document.add_paragraph(unit.text, style=style_name)
        return

    if unit.type == UnitType.LIST_ITEM:
        preferred_style = unit.formatting.get("style_name") or unit.metadata.get("style_name")
        style_name = _get_style_name(document, preferred_style, "List Bullet")
        document.add_paragraph(unit.text, style=style_name)
        return

    if unit.type == UnitType.QUOTE:
        style_name = _get_style_name(document, "Intense Quote", "Quote")
        document.add_paragraph(unit.text, style=style_name)
        return

    if unit.type == UnitType.CODE_BLOCK:
        style_name = _get_style_name(document, "Code", "Normal")
        document.add_paragraph(unit.text, style=style_name)
        return

    if unit.type == UnitType.CAPTION:
        style_name = _get_style_name(document, "Caption", "Normal")
        document.add_paragraph(unit.text, style=style_name)
        return

    document.add_paragraph(unit.text)


def build_docx_from_units(units: list[TranslationUnit], output_path: str | Path | None = None) -> Path | bytes:
    """Rebuild a DOCX document from translation units and optionally save it to disk."""
    document = Document()

    paragraph_units = [unit for unit in units if unit.type != UnitType.TABLE_CELL]
    table_units = [unit for unit in units if unit.type == UnitType.TABLE_CELL]

    for unit in paragraph_units:
        _append_unit_to_document(document, unit)

    if table_units:
        max_row = max(unit.location.get("row", 0) for unit in table_units)
        max_col = max(unit.location.get("column", 0) for unit in table_units)

        table = document.add_table(rows=max_row + 1, cols=max_col + 1)

        for unit in table_units:
            row_index = unit.location.get("row", 0)
            column_index = unit.location.get("column", 0)
            table.cell(row_index, column_index).text = unit.text

    if output_path is None:
        document_bytes = io.BytesIO()
        document.save(document_bytes)
        return document_bytes.getvalue()

    output = Path(output_path)
    document.save(output)
    return output
