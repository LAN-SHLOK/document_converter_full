# 📄 Document Converter – OCR + Text Extraction (Tesseract + EasyOCR)

A powerful multi-format **Document → Text converter** built using:
- **Tesseract OCR**
- **EasyOCR**
- **PyMuPDF**
- **pdf2image**
- **python-docx / python-pptx**

Supports:
- 🖼️ **Images:** JPG, PNG, TIFF, BMP, WEBP  
- 📄 **PDFs:** Text-based & Scanned PDFs  
- 📝 **DOCX** files (Office Word)  
- 📊 **PPTX** files (Office PowerPoint)

Also includes:
- 🔍 Automatic **Important Details Extractor** (Emails, Phones, Dates, Key:Value pairs)
- 📁 Automatic output folder generation  
- 🧹 Optional spell correction  
- 🖥️ Full **Streamlit Web App UI**

---

# 🚀 Features

### ✔ Convert ANY document to `.txt`  
- Multi-OCR merge: **Tesseract + EasyOCR**
- PDF text-mode detection → uses direct extraction when possible

### ✔ Important details extraction  
Automatically extracts:
- Emails  
- Phone numbers  
- Date formats  
- Key:Value structured text  

### ✔ Clean Web UI  
Built with **Streamlit**, featuring:
- File upload  
- OCR progress  
- Download TXT output  
- Download JSON details  
- Text preview panel  

### ✔ High accuracy  
Image preprocessing:
- Grayscale  
- Denoising  
- Adaptive thresholding  

### ✔ Works locally or in Docker

---

# 📦 Requirements

Install system packages:

### **Windows**
- Tesseract OCR → https://github.com/UB-Mannheim/tesseract/wiki  
- Poppler for Windows → add `bin/` to PATH

### **Linux**
```bash
sudo apt install tesseract-ocr poppler-utils
