<!-- .github/copilot-instructions.md -->
# Repo guide for AI coding agents

This project converts documents (PDF, DOCX, PPTX, images) to plain text using a mix of native extractors and OCR (Tesseract + EasyOCR). Below are the minimal, actionable facts an AI coding agent needs to be productive here.

1. Big picture
- **Two-layer design**: `core/doc_to_text.py` is the conversion engine (format-specific extractors + OCR merging). `app/app_streamlit.py` is a thin Streamlit UI that calls into the core functions.
- **Primary flow**: input file -> format-specific extractor (DOCX/PPTX/PDF-text) -> if text insufficient -> render pages to images -> OCR with both pytesseract & easyocr -> merge words by confidence -> write `.txt` alongside input.

2. Key files & symbols
- `core/doc_to_text.py`: main logic. Important symbols: `convert_any_to_text(filepath, do_spell_correct)`, `extract_important_details(text)`, `extract_text_from_pdf_with_mixed_strategy`, `ocr_pytesseract_with_conf`, `ocr_easyocr_with_conf`, `merge_word_lists_by_conf`, `lightly_spell_correct_text`.
- `app/app_streamlit.py`: Streamlit frontend. It uses `convert_any_to_text` and `extract_important_details` and exposes spell-correction UI.
- `README.md`: contains install/run quick-start (system deps note: Tesseract + Poppler).

3. Environment & developer workflows
- System prerequisites (manual):
  - Install Tesseract OCR and ensure `tesseract` is on `PATH` or set `pytesseract.pytesseract.tesseract_cmd` in code (example commented in `doc_to_text.py`).
  - Install Poppler for `pdf2image` (add to PATH or set `POPPLER_PATH` env on Windows).
- Python setup: `pip install -r requirements.txt`.
- Run the UI locally: `streamlit run app/app_streamlit.py` (PowerShell: `streamlit run app/app_streamlit.py`).
- CLI usage: `python core/doc_to_text.py <path-to-file>` (script entrypoint at bottom of `doc_to_text.py`).

4. Project-specific patterns & conventions
- Prefer native format extractors first: the code tries `extract_text_from_text_pdf` / `extract_text_from_docx` / `extract_text_from_pptx` and only falls back to OCR when native text is too short.
- OCR merging strategy: both Tesseract and EasyOCR are run on preprocessed images; `merge_word_lists_by_conf` picks per-position words by higher confidence and computes an average confidence (0..1). Tests/changes should respect this merging approach.
- Confidence semantics: functions return an estimated average confidence in 0..1. The Streamlit UI warns when confidence < 0.5 — keep that threshold in mind when changing scoring behavior.
- Spell correction is intentionally light and optional: controlled by `DO_SPELL_CORRECTION` (module-level default) and `do_spell_correct` param on `convert_any_to_text`.
- Heavy objects are singletons at module import: `reader_easyocr = easyocr.Reader(...)` and `GLOBAL_SPELLER` — avoid re-instantiating in hot paths.

5. Known limitations the agent should be aware of
- `.doc` and `.ppt` (binary Office) are not supported; code raises RuntimeError and expects conversion to `.docx`/`.pptx` externally (LibreOffice/pandoc). Don't attempt blind fixes without adding conversion tooling.
- Output file location: `convert_any_to_text` writes a `.txt` next to the input file (`Path(filepath).with_suffix('.txt')`). The repository has `outputs/converted_text/` and `outputs/extracted_details/` but core code does not use those folders — consider this when adding orchestration.
- Streamlit front-end contains a small bug in filename suffix extraction (see `app/app_streamlit.py`): the current lines that compute `suffix` are malformed. A simple, non-invasive fix is to replace with `suffix = uploaded.name.split('.')[-1].lower()`.

6. Editing & testing guidance (concrete checks)
- When changing OCR/image path logic:
  - Run `streamlit run app/app_streamlit.py` and upload representative files (text PDF, scanned PDF, DOCX, image).
  - Inspect the produced `.txt` file written next to the temp file (Streamlit saves uploads to a temp file by default). Confirm confidence values printed.
- When modifying `reader_easyocr` or `pytesseract` settings, note these are global at import time — restart runtime to test changes.
- Manual smoke test commands (Windows PowerShell):
  - `pip install -r requirements.txt`
  - Set Tesseract path if needed: in PowerShell `setx PATH "$env:PATH;C:\Program Files\Tesseract-OCR"` then restart shell.
  - `streamlit run app/app_streamlit.py`
  - Or CLI: `python core/doc_to_text.py examples/sample.txt` (example text input exists in `examples/`).

7. PR checklist for proposed changes
- Keep the core merging strategy intact unless intentionally replacing it with clearer metrics.
- Add or update README notes when changing system dependencies (Poppler/Tesseract).
- Ensure heavy initializations remain module-level singletons or document the change.
- If adding output orchestration to `outputs/`, make it opt-in and preserve current behavior by default.

8. When you can't discover something
- If a runtime behavior or CI step is not discoverable from files, run the app locally and report observed behavior and exact commands used.

If anything here is unclear or you'd like a different focus (more on testing, CI, or refactoring tasks), tell me which part to expand and I'll iterate.
