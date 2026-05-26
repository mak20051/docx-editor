"""
DOCX Rich Text Editor — Flask + Quill.js
http://localhost:5123

Data flow:
  open .docx ──▶ docx_to_delta() ──▶ Quill Delta JSON ──▶ editor
  editor     ──▶ Quill Delta JSON ──▶ delta_to_docx()  ──▶ .docx

Formatting preserved: bold, italic, underline, strikethrough,
  H1/H2/H3, bullet list, ordered list, blockquote, code block.
"""

import json
import threading
import webbrowser
from pathlib import Path

from flask import Flask, jsonify, render_template_string, request
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

app = Flask(__name__)
PORT = 5123

state: dict = {"filepath": None, "delta": {"ops": [{"insert": "\n"}]}}


# ── DOCX → Quill Delta ────────────────────────────────────────────────────────

def docx_to_delta(doc: Document) -> dict:
    """Convert python-docx Document to Quill Delta ops."""
    ops = []
    for para in doc.paragraphs:
        style = para.style.name if para.style else "Normal"

        if not para.runs and not para.text:
            ops.append({"insert": "\n"})
            continue

        for run in para.runs:
            if not run.text:
                continue
            attrs: dict = {}
            if run.bold is True:
                attrs["bold"] = True
            if run.italic is True:
                attrs["italic"] = True
            if run.underline is True:
                attrs["underline"] = True
            if run.font.strike is True:
                attrs["strike"] = True
            op: dict = {"insert": run.text}
            if attrs:
                op["attributes"] = attrs
            ops.append(op)

        # paragraph terminator — carries line-level attributes
        line: dict = {}
        lname = style.lower()
        if "heading 1" in lname:
            line["header"] = 1
        elif "heading 2" in lname:
            line["header"] = 2
        elif "heading 3" in lname:
            line["header"] = 3
        elif lname in ("list bullet", "list bullet 2"):
            line["list"] = "bullet"
        elif lname in ("list number", "list number 2"):
            line["list"] = "ordered"
        elif lname == "block text":
            line["blockquote"] = True

        nl: dict = {"insert": "\n"}
        if line:
            nl["attributes"] = line
        ops.append(nl)

    return {"ops": ops}


# ── Quill Delta → DOCX ────────────────────────────────────────────────────────

def delta_to_docx(ops: list, path: str):
    """Convert Quill Delta ops to a .docx file."""
    doc = Document()
    doc.styles["Normal"].font.size = Pt(11)

    # Buffer of (text, inline_attrs) for the current paragraph
    buf: list[tuple[str, dict]] = []

    def flush(line_attrs: dict):
        header = line_attrs.get("header")
        lst    = line_attrs.get("list")
        quote  = line_attrs.get("blockquote")
        code   = line_attrs.get("code-block")

        if header == 1:
            para = doc.add_heading("", level=1)
        elif header == 2:
            para = doc.add_heading("", level=2)
        elif header == 3:
            para = doc.add_heading("", level=3)
        elif lst == "bullet":
            para = doc.add_paragraph(style="List Bullet")
        elif lst == "ordered":
            para = doc.add_paragraph(style="List Number")
        elif quote:
            para = doc.add_paragraph(style="Quote")
        else:
            para = doc.add_paragraph()

        for text, attrs in buf:
            run = para.add_run(text)
            if attrs.get("bold"):
                run.bold = True
            if attrs.get("italic"):
                run.italic = True
            if attrs.get("underline"):
                run.underline = True
            if attrs.get("strike"):
                run.font.strike = True
            if code or line_attrs.get("code-block"):
                run.font.name = "Courier New"
                run.font.size = Pt(10)

        buf.clear()

    INLINE = {"bold", "italic", "underline", "strike", "color", "background", "link", "code"}
    LINE   = {"header", "list", "blockquote", "code-block", "align", "indent", "direction"}

    for op in ops:
        insert = op.get("insert")
        if not isinstance(insert, str):
            continue
        attrs = op.get("attributes") or {}

        inline_attrs = {k: v for k, v in attrs.items() if k in INLINE}
        line_attrs   = {k: v for k, v in attrs.items() if k in LINE}

        parts = insert.split("\n")
        for i, part in enumerate(parts):
            if part:
                buf.append((part, inline_attrs))
            if i < len(parts) - 1:
                flush(line_attrs)

    if buf:
        flush({})

    doc.save(path)


# ── HTML ──────────────────────────────────────────────────────────────────────

HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>DOCX Editor</title>
<link href="https://cdn.quilljs.com/1.3.7/quill.snow.css" rel="stylesheet">
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  html, body { height: 100%; background: #1e1e2e; color: #cdd6f4; font-family: system-ui, sans-serif; }

  /* ── file toolbar ── */
  #file-bar {
    background: #181825;
    padding: 8px 14px;
    display: flex;
    align-items: center;
    gap: 6px;
    border-bottom: 1px solid #313244;
    flex-shrink: 0;
  }
  #file-bar button {
    background: #313244; color: #cdd6f4;
    border: none; border-radius: 6px;
    padding: 6px 14px; font-size: 13px; cursor: pointer;
    font-family: inherit; transition: background .15s;
    white-space: nowrap;
  }
  #file-bar button:hover { background: #89b4fa; color: #1e1e2e; }
  #filepath {
    margin-left: 10px; color: #6c7086; font-size: 12px;
    flex: 1; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
  }
  #modified-dot { color: #f38ba8; font-size: 16px; display: none; }

  /* ── Quill overrides ── */
  #quill-wrapper {
    display: flex;
    flex-direction: column;
    height: calc(100vh - 82px); /* file-bar + status */
  }

  .ql-toolbar.ql-snow {
    background: #24273a;
    border: none;
    border-bottom: 1px solid #313244;
    flex-shrink: 0;
  }
  .ql-toolbar .ql-stroke { stroke: #cdd6f4; }
  .ql-toolbar .ql-fill   { fill:   #cdd6f4; }
  .ql-toolbar .ql-picker-label { color: #cdd6f4; }
  .ql-toolbar button:hover .ql-stroke,
  .ql-toolbar .ql-active .ql-stroke { stroke: #89b4fa; }
  .ql-toolbar button:hover .ql-fill,
  .ql-toolbar .ql-active .ql-fill   { fill:   #89b4fa; }
  .ql-toolbar .ql-picker-label:hover,
  .ql-toolbar .ql-active { color: #89b4fa; }

  .ql-container.ql-snow {
    border: none;
    flex: 1;
    overflow: hidden;
    background: #1e1e2e;
    font-size: 15px;
  }
  .ql-editor {
    color: #cdd6f4;
    padding: 32px 80px;
    line-height: 1.75;
    height: 100%;
    overflow-y: auto;
  }
  .ql-editor.ql-blank::before { color: #45475a; font-style: normal; }
  .ql-editor h1, .ql-editor h2, .ql-editor h3 { color: #89b4fa; }
  .ql-editor blockquote {
    border-left: 4px solid #89b4fa;
    color: #a6adc8;
    padding-left: 16px;
    margin-left: 0;
  }
  .ql-editor pre.ql-syntax {
    background: #11111b;
    border-radius: 6px;
    color: #a6e3a1;
    font-size: 13px;
  }
  .ql-editor a { color: #89dceb; }

  /* ── status bar ── */
  #status {
    background: #11111b;
    padding: 4px 16px;
    font-size: 11px; color: #89b4fa;
    border-top: 1px solid #313244;
    flex-shrink: 0;
    height: 24px;
  }

  #file-input { display: none; }
</style>
</head>
<body>

<div id="file-bar">
  <button onclick="triggerOpen()">Open…</button>
  <button onclick="newDoc()">New</button>
  <button onclick="saveDoc()">Save</button>
  <button onclick="saveAs()">Save As…</button>
  <span id="filepath">Untitled</span>
  <span id="modified-dot">●</span>
</div>

<div id="quill-wrapper">
  <div id="toolbar">
    <span class="ql-formats">
      <select class="ql-header">
        <option selected></option>
        <option value="1">H1</option>
        <option value="2">H2</option>
        <option value="3">H3</option>
      </select>
    </span>
    <span class="ql-formats">
      <button class="ql-bold"></button>
      <button class="ql-italic"></button>
      <button class="ql-underline"></button>
      <button class="ql-strike"></button>
    </span>
    <span class="ql-formats">
      <button class="ql-list" value="ordered"></button>
      <button class="ql-list" value="bullet"></button>
    </span>
    <span class="ql-formats">
      <button class="ql-blockquote"></button>
      <button class="ql-code-block"></button>
    </span>
    <span class="ql-formats">
      <button class="ql-link"></button>
    </span>
    <span class="ql-formats">
      <button class="ql-clean"></button>
    </span>
  </div>
  <div id="editor"></div>
</div>

<div id="status">Ready</div>
<input type="file" id="file-input" accept=".docx" onchange="openFile(this)">

<script src="https://cdn.quilljs.com/1.3.7/quill.min.js"></script>
<script>
const quill = new Quill('#editor', {
  theme: 'snow',
  modules: { toolbar: '#toolbar' },
  placeholder: 'Start typing or open a .docx file…',
});

const statusEl  = document.getElementById('status');
const filepathEl = document.getElementById('filepath');
const dot        = document.getElementById('modified-dot');
let modified = false;

// ── helpers ──────────────────────────────────────────────────────────────────
function setStatus(msg)    { statusEl.textContent = msg; }
function setFilepath(p)    { filepathEl.textContent = p ? p.split('/').pop() : 'Untitled'; filepathEl.title = p || ''; }
function setModified(v)    { modified = v; dot.style.display = v ? 'inline' : 'none'; }

quill.on('text-change', () => setModified(true));

// ── load initial state ────────────────────────────────────────────────────────
(async () => {
  const d = await (await fetch('/state')).json();
  quill.setContents(d.delta, 'silent');
  setFilepath(d.filepath);
  setModified(false);
})();

// ── keyboard shortcuts ────────────────────────────────────────────────────────
document.addEventListener('keydown', e => {
  if (e.metaKey && !e.shiftKey && e.key === 's') { e.preventDefault(); saveDoc(); }
  if (e.metaKey &&  e.shiftKey && e.key === 's') { e.preventDefault(); saveAs(); }
  if (e.metaKey && e.key === 'o') { e.preventDefault(); triggerOpen(); }
  if (e.metaKey && e.key === 'n') { e.preventDefault(); newDoc(); }
});

// ── file commands ─────────────────────────────────────────────────────────────
function triggerOpen() { document.getElementById('file-input').click(); }

async function openFile(input) {
  if (!input.files.length) return;
  const fd = new FormData();
  fd.append('file', input.files[0]);
  input.value = '';
  const d = await (await fetch('/open', { method: 'POST', body: fd })).json();
  if (d.error) { setStatus('Error: ' + d.error); return; }
  quill.setContents(d.delta, 'silent');
  setFilepath(d.filepath);
  setModified(false);
  setStatus(`Opened: ${d.filepath}  (${d.paragraphs} paragraphs)`);
}

function newDoc() {
  if (modified && !confirm('Discard unsaved changes?')) return;
  quill.setContents([{ insert: '\n' }], 'silent');
  setFilepath(null);
  setModified(false);
  fetch('/reset', { method: 'POST' });
  setStatus('New document');
}

async function saveDoc() {
  const delta = quill.getContents();
  const d = await (await fetch('/save', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ delta: delta }),
  })).json();
  if (d.error)       { setStatus('Save failed: ' + d.error); return; }
  if (d.needs_saveas) { saveAs(); return; }
  setModified(false);
  setFilepath(d.filepath);
  setStatus('Saved: ' + d.filepath);
}

async function saveAs() {
  const cur = filepathEl.textContent;
  const name = prompt('Filename (saved to Desktop):', cur !== 'Untitled' ? cur : 'document.docx');
  if (!name) return;
  const delta = quill.getContents();
  const d = await (await fetch('/saveas', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ delta: delta, name }),
  })).json();
  if (d.error) { setStatus('Save failed: ' + d.error); return; }
  setModified(false);
  setFilepath(d.filepath);
  setStatus('Saved: ' + d.filepath);
}
</script>
</body>
</html>
"""


# ── Flask routes ──────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template_string(HTML)


@app.route("/state")
def get_state():
    return jsonify({"filepath": state["filepath"], "delta": state["delta"]})


@app.route("/open", methods=["POST"])
def open_docx():
    f = request.files.get("file")
    if not f:
        return jsonify({"error": "no file"})
    try:
        tmp = Path("/tmp") / f.filename
        f.save(tmp)
        doc = Document(tmp)
        delta = docx_to_delta(doc)
        state["filepath"] = str(tmp)
        state["delta"] = delta
        return jsonify({"delta": delta, "filepath": str(tmp), "paragraphs": len(doc.paragraphs)})
    except Exception as exc:
        return jsonify({"error": str(exc)})


@app.route("/reset", methods=["POST"])
def reset():
    state["filepath"] = None
    state["delta"] = {"ops": [{"insert": "\n"}]}
    return jsonify({"ok": True})


@app.route("/save", methods=["POST"])
def save():
    data = request.get_json()
    if not state["filepath"]:
        return jsonify({"needs_saveas": True})
    try:
        delta_to_docx(data["delta"]["ops"], state["filepath"])
        state["delta"] = data["delta"]
        return jsonify({"filepath": state["filepath"]})
    except Exception as exc:
        return jsonify({"error": str(exc)})


@app.route("/saveas", methods=["POST"])
def saveas():
    data = request.get_json()
    name = data.get("name", "document.docx")
    if not name.endswith(".docx"):
        name += ".docx"
    dest = str(Path.home() / "Desktop" / name)
    try:
        delta_to_docx(data["delta"]["ops"], dest)
        state["filepath"] = dest
        state["delta"] = data["delta"]
        return jsonify({"filepath": dest})
    except Exception as exc:
        return jsonify({"error": str(exc)})


# ── main ──────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    url = f"http://localhost:{PORT}"
    threading.Timer(0.9, lambda: webbrowser.open(url)).start()
    print(f"DOCX Rich Text Editor → {url}")
    app.run(port=PORT, debug=False)
