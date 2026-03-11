# Docling Local Review App

Tiny local-first QA/review app for comparing source documents (PDF/PPTX) with Docling outputs (HTML/Markdown/JSON), with optional Railway hosting.

## Why this app

- Runs fully local on a normal laptop.
- Pure Python + Streamlit (no Docker, no Node, no external services).
- Handles partial assets gracefully (for example JSON + PDF only).
- Includes robust JSON block parsing for schema variation.

## Features

- Two-pane review UI:
  - Left: source preview (PDF embed or PPTX fallback slide cards).
  - Right: Docling output preview (HTML / Markdown / JSON).
- Top toolbar:
  - Multi-file uploader.
  - Local folder path loader.
  - Document selector.
  - View selector.
  - Simple theme toggle and pane split slider.
- Metadata pane:
  - Extracted blocks table (id, label/type, page, snippet, path).
  - Filters (search, label/type, page).
  - Select a block and inspect raw metadata.
  - Export selected block JSON.
- Search:
  - Search across the currently visible right-pane content.
  - Shows highlighted snippets.
- Status panel:
  - Clearly shows loaded/missing assets per document.
- Persistent recent folders:
  - Stored in `~/.docling_review_config.json`.
- Sample data:
  - `sample_data/sample_docling.json` for quick testing.

## Project structure

```text
docling-vis/
  app.py
  loaders.py
  parsers.py
  ui_components.py
  utils.py
  requirements.txt
  pyproject.toml
  railway.toml
  Procfile
  sample_data/
    sample_docling.json
  README.md
```

## Setup and run

1. Create virtual environment:

```bash
python -m venv .venv
```

2. Activate:

Windows (PowerShell):

```powershell
.venv\Scripts\Activate.ps1
```

macOS/Linux:

```bash
source .venv/bin/activate
```

3. Install dependencies:

```bash
pip install -r requirements.txt
```

4. Run app:

```bash
streamlit run app.py
```

## How document matching works

- The app groups files by normalized filename stem.
- Example: `invoice.pdf`, `invoice.docling.json`, `invoice.md`, `invoice.html` map to one document key.
- Supported extensions:
  - Source: `.pdf`, `.pptx`
  - Docling outputs: `.html`, `.htm`, `.md`, `.markdown`, `.json`

## PPTX preview strategy

- Streamlit cannot reliably render native PPTX visuals cross-platform without external tooling.
- This app uses a Python-only fallback:
  - Parses PPTX text with `python-pptx`.
  - Generates slide-card images with Pillow for in-app preview.
- This is reliable and local, but not full-fidelity slide rendering.

## Deploy on Railway

This repo includes [railway.toml](railway.toml) with a Streamlit start command that binds to Railway's required host/port:

```toml
startCommand = "streamlit run app.py --server.headless true --server.address 0.0.0.0 --server.port $PORT"
```

Typical flow:

1. Push this project to GitHub.
2. Create a new Railway project from that repo.
3. Railway will install Python dependencies and start the app using `railway.toml`.
4. Open the generated Railway URL.

Important for hosted usage:

- The folder-path loader reads the server filesystem, not your laptop filesystem.
- For normal review usage on Railway, upload files in the UI.

## Known limitations

- PPTX preview is text-first fallback, not pixel-perfect visual rendering.
- HTML sanitization removes scripts and unsupported tags; complex interactive HTML will not behave the same as a browser.
- Block-to-HTML exact visual highlight is approximate; fallback is focused snippet + metadata panel.
- Folder scanning is non-recursive (current folder only).
- On Railway, "local folder path" points to container storage, not user-local paths.

## Future improvements

- Optional recursive folder scan.
- Better HTML block anchoring and scroll-to-target behavior.
- Optional page/block synchronization using bbox overlays.
- Inline diff mode between Markdown and HTML text.
- Persist full app state per document (selected view, filters, selected block).
- Optional plugin parser adapters for known Docling schema variants.

