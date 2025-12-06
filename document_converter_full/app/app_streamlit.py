# app_streamlit.py
import sys
import os
# Ensure project root is on sys.path so "core" imports always work
ROOT_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if ROOT_DIR not in sys.path:
    sys.path.insert(0, ROOT_DIR)

import shutil
import streamlit as st
import json
from pathlib import Path
import tempfile

# detect presence of tesseract binary (pytesseract requires this)
TESSERACT_AVAILABLE = shutil.which("tesseract") is not None

# detect EasyOCR Python package (optional fallback)
try:
    import easyocr  # noqa: F401
    HAS_EASYOCR_LOCAL = True
except Exception:
    HAS_EASYOCR_LOCAL = False

# import core functions (now that ROOT_DIR is on sys.path)
from core.doc_to_text import convert_any_to_text, extract_important_details

# show an upfront banner if image OCR won't work on this host
if not TESSERACT_AVAILABLE and not HAS_EASYOCR_LOCAL:
    st.warning(
        "Image OCR (scanned PDFs / JPG/PNG) is currently unavailable on this host:\n\n"
        "- The **Tesseract** binary is not installed in the environment, and\n"
        "- **EasyOCR** Python package is not present.\n\n"
        "Result: uploading images or scanned PDFs will fail. To enable full image OCR either:\n"
        "1. Deploy the app using Docker (recommended) — Dockerfile installs Tesseract & Poppler, OR\n"
        "2. Run the app on a machine with Tesseract installed and on PATH.\n\n"
        "See README for deployment instructions."
    )

st.set_page_config(page_title="Doc → TXT Converter", layout="wide")

st.title("Convert Document → TXT (Tesseract + EasyOCR)")
st.markdown("Supports: PDF (text/scanned), DOCX, PPTX, images. Use clean scans for best OCR.")

uploaded = st.file_uploader(
    "Upload a document",
    type=["pdf", "docx", "pptx", "jpg", "jpeg", "png", "tiff", "bmp", "webp"]
)
spell_corr = st.checkbox("Enable light spell-correction", value=False)

if uploaded:
    suffix = uploaded.name.split('.')[-1].lower()
    # write uploaded file to a temp file with the correct suffix
    with tempfile.NamedTemporaryFile(delete=False, suffix=f".{suffix}") as tmp:
        tmp.write(uploaded.read())
        tmp_path = tmp.name

    st.info("File saved. Starting conversion...")
    try:
        txt_path, confidence, text = convert_any_to_text(tmp_path, do_spell_correct=spell_corr)
        st.success(f"Saved TXT: {txt_path} — Confidence: {confidence*100:.2f}%")
        st.download_button("Download TXT", data=text, file_name=Path(txt_path).name, mime="text/plain")

        # Important details
        details = extract_important_details(text)
        st.subheader("Important details (heuristic)")
        st.json(details)
        st.download_button("Download details JSON", data=json.dumps(details, indent=2), file_name="important_details.json", mime="application/json")

        st.subheader("Extracted text preview")
        st.text_area("Text preview", value=text[:20000], height=400)

        if confidence < 0.5:
            st.warning("Low OCR confidence. Consider re-uploading a cleaner or higher-DPI scan, or enabling spell-correction.")

    except RuntimeError as rte:
        # runtime errors from converter (friendly message)
        st.error(f"Conversion failed: {rte}")
    except Exception as e:
        # unexpected: show user-friendly message and log full details
        st.error("Conversion failed due to an unexpected error. Check logs for details.")
        import logging
        logging.exception("Unexpected error during conversion")
