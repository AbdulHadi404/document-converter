# Quick Start: Markdown to PDF

## Convert SYSTEM_VERIFIERS_AND_ANALYTICS_DOCUMENTATION.md to PDF

### Prerequisites

1. **Install Pandoc** (if not already installed)
   ```bash
   # Windows (Scoop)
   scoop install pandoc
   
   # Windows (Chocolatey)
   choco install pandoc
   
   # Linux
   sudo apt-get install pandoc
   
   # macOS
   brew install pandoc
   ```

2. **Install LaTeX** (for beautiful PDF output)
   ```bash
   # Windows: Download and install MiKTeX from https://miktex.org/download
   # Linux
   sudo apt-get install texlive-xetex texlive-latex-extra
   
   # macOS
   brew install --cask mactex
   ```

   **Alternative (lighter):** wkhtmltopdf
   ```bash
   # Windows: Download from https://wkhtmltopdf.org/downloads.html
   # Linux
   sudo apt-get install wkhtmltopdf
   
   # macOS
   brew install wkhtmltopdf
   ```

### Conversion Steps

1. **Navigate to document-converter directory**
   ```bash
   cd document-converter
   ```

2. **Convert the markdown file**
   ```bash
   # Windows
   python doc_converter.py ..\SYSTEM_VERIFIERS_AND_ANALYTICS_DOCUMENTATION.md
   
   # Or use the batch script
   convert_md_to_pdf.bat ..\SYSTEM_VERIFIERS_AND_ANALYTICS_DOCUMENTATION.md
   
   # Linux/macOS
   python doc_converter.py ../SYSTEM_VERIFIERS_AND_ANALYTICS_DOCUMENTATION.md
   ```

3. **Find your PDF**
   - The PDF will be saved in the same directory as the markdown file (or in the output directory if specified)
   - Default output: `output/SYSTEM_VERIFIERS_AND_ANALYTICS_DOCUMENTATION.pdf`

### What You Get

The PDF will have:
- ✅ Professional typography with Georgia font
- ✅ Beautifully styled headers with borders
- ✅ Syntax-highlighted code blocks
- ✅ Properly formatted tables
- ✅ Embedded images
- ✅ Automatic table of contents
- ✅ Numbered sections
- ✅ Page numbers
- ✅ Clickable links

### Troubleshooting

**Error: "Pandoc not found"**
- Install Pandoc (see Prerequisites above)
- Verify: `pandoc --version`

**Error: "PDF engine not found"**
- Install LaTeX (MiKTeX/TeX Live) or wkhtmltopdf
- Verify: `xelatex --version` or `wkhtmltopdf --version`

**Images not showing in PDF**
- Ensure image paths in markdown are relative to the markdown file
- Images will be automatically embedded in the PDF

**Conversion takes too long**
- Large documents may take 1-2 minutes
- First-time LaTeX compilation may be slower (package installation)

---

**Need help?** Check the main README.md for detailed documentation.

