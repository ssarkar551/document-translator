# Document Translator

A small FastAPI service for extracting translation units from DOCX files and rebuilding DOCX documents from those units.

## Features

- Upload a DOCX file and extract translation units
- Rebuild a DOCX file from a list of translation units
- Supports paragraph and table-cell units
- Simple JSON-based payloads for integration

## Project structure

- app/extractors/docx.py: extracts translation units from DOCX files
- app/builder/docx.py: rebuilds DOCX files from translation units
- app/models/translation_unit.py: translation unit model
- main.py: FastAPI application

## Run locally

Install dependencies:

```bash
pip install -r requirements.txt
```

Start the server:

```bash
uvicorn main:app --reload
```

The API will be available at:

- http://127.0.0.1:8000/docs

## API endpoints

### POST /extract-docx

Upload a DOCX file and receive a JSON array of extracted translation units.

Example:

```bash
curl -X POST "http://127.0.0.1:8000/extract-docx?language=fr" \
  -F "file=@sample.docx"
```

### POST /build-docx

Send a JSON array of translation units and receive a rebuilt DOCX file.

Example:

```bash
curl -X POST http://127.0.0.1:8000/build-docx \
  -H "Content-Type: application/json" \
  -d '[{"id":"p0","text":"Invoice","type":"paragraph","location":{"paragraph_index":0}}]'
```

## Docker

Build the image:

```bash
docker build -t document-translator .
```

Run the container:

```bash
docker run -p 8000:8000 document-translator
```
