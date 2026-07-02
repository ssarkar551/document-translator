from __future__ import annotations

from docx import Document

from app.models.translation_unit import TranslationUnit, UnitType


def _classify_paragraph_type(paragraph: object, paragraph_index: int) -> tuple[UnitType, str]:
    style_name = (getattr(paragraph, "style", None).name or "").lower()
    if style_name.startswith("heading"):
        return UnitType.HEADING, f"h{paragraph_index}"
    if "list" in style_name:
        return UnitType.LIST_ITEM, f"li{paragraph_index}"
    if "quote" in style_name:
        return UnitType.QUOTE, f"q{paragraph_index}"
    if "code" in style_name:
        return UnitType.CODE_BLOCK, f"code{paragraph_index}"
    if "caption" in style_name:
        return UnitType.CAPTION, f"cap{paragraph_index}"
    return UnitType.PARAGRAPH, f"p{paragraph_index}"


def extract_translation_units_from_docx(file_path: str, language: str = "en") -> list[TranslationUnit]:
    """Extract translatable units from a DOCX file as a list of typed units."""
    document = Document(file_path)
    units: list[TranslationUnit] = []

    for paragraph_index, paragraph in enumerate(document.paragraphs):
        text = paragraph.text.strip()
        if not text:
            continue

        unit_type, unit_id = _classify_paragraph_type(paragraph, paragraph_index)
        units.append(
            TranslationUnit(
                id=unit_id,
                text=text,
                type=unit_type,
                location={"paragraph_index": paragraph_index},
                document_type="docx",
                formatting={
                    "preserve_whitespace": True,
                    "style_name": getattr(paragraph.style, "name", None),
                },
                metadata={"language": language, "char_count": len(text)},
            )
        )

    for table_index, table in enumerate(document.tables):
        for row_index, row in enumerate(table.rows):
            for column_index, cell in enumerate(row.cells):
                text = " ".join(
                    paragraph.text.strip()
                    for paragraph in cell.paragraphs
                    if paragraph.text.strip()
                ).strip()
                if not text:
                    continue

                units.append(
                    TranslationUnit(
                        id=f"t{table_index}_r{row_index}_c{column_index}",
                        text=text,
                        type=UnitType.TABLE_CELL,
                        location={
                            "table_index": table_index,
                            "row": row_index,
                            "column": column_index,
                        },
                        document_type="docx",
                        formatting={"preserve_whitespace": True},
                        metadata={"language": language, "char_count": len(text)},
                    )
                )

    return units


def extract_text_from_docx(file_path: str, language: str = "en") -> list[str]:
    """Backward-compatible helper returning paragraph text from a DOCX file."""
    return [
        unit.text
        for unit in extract_translation_units_from_docx(file_path, language=language)
        if unit.type == UnitType.PARAGRAPH
    ]
