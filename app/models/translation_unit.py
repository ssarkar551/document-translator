from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

from enum import StrEnum

class UnitType(StrEnum):
    PARAGRAPH = "paragraph"
    HEADING = "heading"
    TABLE_CELL = "table_cell"
    LIST_ITEM = "list_item"
    QUOTE = "quote"
    CODE_BLOCK = "code_block"
    CAPTION = "caption"

@dataclass(slots=True)
class TranslationUnit:
    """Represents a single translatable unit extracted from a document."""

    id: str
    text: str
    type: UnitType | str = UnitType.PARAGRAPH
    location: dict[str, Any] = field(default_factory=dict)
    document_type: str | None = None
    formatting: dict[str, Any] = field(default_factory=dict)
    metadata: dict[str, Any] = field(default_factory=dict)
    translated_text: str | None = None

    def __post_init__(self) -> None:
        if not self.id:
            raise ValueError("TranslationUnit requires a non-empty id")

    @property
    def is_translated(self) -> bool:
        return bool(self.translated_text and self.translated_text.strip())

    def to_dict(self) -> dict[str, Any]:
        data: dict[str, Any] = {
            "id": self.id,
            "text": self.text,
            "type": self.type,
            "location": self.location,
        }

        if self.document_type is not None:
            data["document_type"] = self.document_type
        if self.formatting:
            data["formatting"] = self.formatting
        if self.metadata:
            data["metadata"] = self.metadata
        if self.translated_text is not None:
            data["translated_text"] = self.translated_text

        return data

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "TranslationUnit":
        return cls(
            id=data.get("id", ""),
            text=data.get("text", ""),
            type=data.get("type", "paragraph"),
            location=dict(data.get("location", {})),
            document_type=data.get("document_type"),
            formatting=dict(data.get("formatting", {})),
            metadata=dict(data.get("metadata", {})),
            translated_text=data.get("translated_text"),
        )
