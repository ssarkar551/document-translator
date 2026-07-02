from __future__ import annotations

from pathlib import Path

from docx import Document

from app.models.translation_unit import TranslationUnit, UnitType


def build_docx_from_units(units: list[TranslationUnit], output_path: str | Path) -> Path:
    """Rebuild a DOCX document from translation units."""
    document = Document()

    paragraph_units = [unit for unit in units if unit.type == UnitType.PARAGRAPH]
    table_units = [unit for unit in units if unit.type == UnitType.TABLE_CELL]

    for unit in paragraph_units:
        document.add_paragraph(unit.text)

    if table_units:
        max_row = max(unit.location.get("row", 0) for unit in table_units)
        max_col = max(unit.location.get("column", 0) for unit in table_units)

        table = document.add_table(rows=max_row + 1, cols=max_col + 1)

        for unit in table_units:
            row_index = unit.location.get("row", 0)
            column_index = unit.location.get("column", 0)
            table.cell(row_index, column_index).text = unit.text

    output = Path(output_path)
    document.save(output)
    return output
