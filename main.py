from __future__ import annotations

import tempfile
from pathlib import Path

from fastapi import FastAPI, File, UploadFile
from fastapi.responses import FileResponse, JSONResponse

from app.builder.docx import build_docx_from_units
from app.extractors.docx import extract_translation_units_from_docx
from app.models.translation_unit import TranslationUnit

app = FastAPI(title="Document Translator")


@app.post("/extract-docx")
async def extract_docx(file: UploadFile = File(...), language: str = "en") -> JSONResponse:
    """Extract translation units from an uploaded DOCX file."""
    if not file.filename or not file.filename.lower().endswith(".docx"):
        return JSONResponse(status_code=400, content={"error": "Please upload a .docx file"})

    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as temp_file:
        content = await file.read()
        temp_file.write(content)
        temp_path = Path(temp_file.name)

    try:
        units = extract_translation_units_from_docx(str(temp_path), language=language)
        return JSONResponse(content=[unit.to_dict() for unit in units])
    finally:
        temp_path.unlink(missing_ok=True)


@app.post("/build-docx")
async def build_docx(units: list[dict], language: str = "en") -> FileResponse:
    """Build a DOCX file from translation units and return it as a download."""
    translation_units = [TranslationUnit.from_dict(unit) for unit in units]

    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as temp_file:
        output_path = Path(temp_file.name)

    built_path = build_docx_from_units(translation_units, output_path)
    return FileResponse(path=built_path, filename="built.docx", media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


def main() -> None:
    import uvicorn

    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)


if __name__ == "__main__":
    main()
