# Keyword_Analysis_Ocean.py

This script scans PDF documents for specified keywords, including content found inside tables. It is designed to support structured document analysis, especially for research involving ocean-related reports or scientific papers.

## ✨ Features

- Extracts both **plain text** and **tables** from PDFs using `pdfplumber`
- Searches for user-defined **keywords** (case-insensitive)
- Outputs matches with page numbers into a formatted **Excel file**

## 📦 Installation

```bash
pip install pdfplumber xlsxwriter

project-folder/
├── Keyword_Analysis_Ocean.py
├── data/
│   ├── Norway/
│   │   ├── file1.pdf
│   │   ├── file2.pdf
│   ├── Japan/
│   │   ├── file1.pdf
│   │   ├── file2.pdf
│   └── Brazil/
│       ├── file1.pdf
│       ├── file2.pdf
