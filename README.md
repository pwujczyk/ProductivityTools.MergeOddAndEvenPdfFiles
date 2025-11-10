# Merge Odd and Even PDFs

This tool helps you combine two PDF files—one containing odd pages and one containing even pages—into a single, properly interleaved PDF document. This is common when scanning two-sided documents using a scanner without a duplexer/ADF that handles dual-sided scanning automatically.

## Prerequisites

- **PowerShell** (Standard on Windows)
- **iTextSharp.dll**: The script attempts to download this automatically if not found.

## Usage

1. Open PowerShell.
2. Navigate to this directory.
3. Run the script:

```powershell
.\Merge-OddEvenPdf.ps1 -OddPdfPath "C:\Path\To\Odd.pdf" -EvenPdfPath "C:\Path\To\Even.pdf" -OutputPdfPath "C:\Path\To\Merged.pdf"
```

### Parameters

- `-OddPdfPath`: Path to the PDF containing pages 1, 3, 5, etc.
- `-EvenPdfPath`: Path to the PDF containing pages 2, 4, 6, etc.
- `-OutputPdfPath`: Path where the final PDF will be saved.
- `-ReverseEven`: (Optional) Use this switch if your even pages were scanned in reverse order (e.g., 6, 4, 2) which often happens when flipping a stack of pages.

## Example

```powershell
# Standard merge
.\Merge-OddEvenPdf.ps1 -OddPdfPath ".\odd.pdf" -EvenPdfPath ".\even.pdf" -OutputPdfPath ".\result.pdf"

# Merge with reversed even pages
.\Merge-OddEvenPdf.ps1 -OddPdfPath ".\odd.pdf" -EvenPdfPath ".\even.pdf" -OutputPdfPath ".\result.pdf" -ReverseEven
```
