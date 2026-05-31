# docx-editor

A single-file Flask web app that opens documents (`.docx .doc .odt .rtf .txt .html`), edits them in a browser-based Quill.js rich-text editor, and saves back to `.docx` (or `.txt`). Conversion runs through Quill Delta JSON as the intermediate format.

## Stack
- Python 3.11 + Flask
- python-docx, odfpy, striprtf, beautifulsoup4 for format parsing; antiword (system pkg) for legacy `.doc`
- Quill.js front-end, served inline via `render_template_string`

## Commands
- `./run.sh` — start the app (`python3 main.py`); serves on `http://localhost:5123`
- `pip install -r requirements.txt` — install Python deps
- `docker build -t docx-editor . && docker run -p 8000:8000 docx-editor` — containerized run (listens on `PORT`, default 8000 in Docker)

## Key files
- `main.py` — entire app: format→Delta readers, Delta→docx writer, Flask routes (`/open`, `/save`, `/saveas`, `/reset`, `/state`), inline HTML/Quill UI
- `requirements.txt` — Python dependencies
- `Dockerfile` — Python 3.11-slim image; installs antiword for `.doc` support
- `run.sh` — local launcher

## Conventions
- Single-module design — all logic lives in `main.py`; no package structure.
- Quill Delta JSON is the canonical in-memory document model; all readers/writers convert to/from it.
- Runtime config via env vars: `PORT` (server port), `DOCKER=1` (suppresses auto-opening a browser).
- Editable document state is held in the in-memory `state` dict (single active document, not multi-user).
