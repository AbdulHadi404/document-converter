@echo off
REM Markdown to PDF Converter
REM Usage: convert_md_to_pdf.bat [markdown_file] [output_directory]

echo ==========================================
echo Markdown to PDF Converter
echo ==========================================
echo.
echo Converting markdown to beautiful PDF...
echo.

if "%1"=="" (
    echo Usage: convert_md_to_pdf.bat [markdown_file] [output_directory]
    echo Example: convert_md_to_pdf.bat SYSTEM_VERIFIERS_AND_ANALYTICS_DOCUMENTATION.md
    echo.
    pause
    exit /b 1
)

python doc_converter.py %1 %2

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ✓ Conversion successful!
) else (
    echo.
    echo ✗ Conversion failed. Check error messages above.
)

echo.
pause

