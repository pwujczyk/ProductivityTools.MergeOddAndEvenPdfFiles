<#
.SYNOPSIS
    Merges an Odd-pages PDF and an Even-pages PDF into a single interleaved PDF.
    
.DESCRIPTION
    This script takes two PDF files: one containing odd pages (1, 3, 5...) and one containing even pages (2, 4, 6...).
    It combines them into a single output PDF with pages interleaved (1, 2, 3, 4...).
    It requires iTextSharp.dll. If not found, it attempts to download it from NuGet.

.PARAMETER OddPdfPath
    Path to the PDF file containing odd pages.

.PARAMETER EvenPdfPath
    Path to the PDF file containing even pages.

.PARAMETER OutputPdfPath
    Path where the merged PDF will be saved.

.PARAMETER ReverseEven
    If set, the even pages are treated as being in reverse order (e.g., 6, 4, 2).
    Use this if you scanned the even pages by flipping the stack without reordering.

.EXAMPLE
    .\Merge-OddEvenPdf.ps1 -OddPdfPath "odd.pdf" -EvenPdfPath "even.pdf" -OutputPdfPath "merged.pdf"
    
.EXAMPLE
    .\Merge-OddEvenPdf.ps1 -OddPdfPath "odd.pdf" -EvenPdfPath "even.pdf" -OutputPdfPath "merged.pdf" -ReverseEven
#>


[CmdletBinding()]
param(
    [string]$OddPdfPath = "C:\Users\pawel\Documents\Scan\2025.12.14- Zmywarka_odd.pdf",

    [string]$EvenPdfPath = "C:\Users\pawel\Documents\Scan\2025.12.14- Zmywarka_even.pdf",

    [string]$OutputPdfPath = "C:\Users\pawel\Documents\Scan\2025.12.14- Zmywarka_out.pdf",

    [switch]$ReverseEven
)

Set-StrictMode -Version Latest

# --- Dependency Management ---
$ScriptDir = $PSScriptRoot

$Dependencies = @(
    @{
        Name = "BouncyCastle.Crypto.dll"
        PackageName = "BouncyCastle"
        Version = "1.8.9"
        Url = "https://www.nuget.org/api/v2/package/BouncyCastle/1.8.9"
    },
    @{
        Name = "itextsharp.dll"
        PackageName = "iTextSharp"
        Version = "5.5.13.3"
        Url = "https://www.nuget.org/api/v2/package/iTextSharp/5.5.13.3"
    }
)

function Install-Dependency {
    param($Dep)
    $dllPath = Join-Path $ScriptDir $Dep.Name
    if (-not (Test-Path $dllPath)) {
        Write-Warning "$($Dep.Name) not found."
        $confirmation = Read-Host "Do you want to download $($Dep.PackageName) from NuGet? (Y/N)"
        if ($confirmation -eq 'Y') {
            Write-Host "Downloading $($Dep.PackageName)..."
            $zipPath = Join-Path $ScriptDir "temp_$($Dep.PackageName).zip"
            $extractPath = Join-Path $ScriptDir "temp_$($Dep.PackageName)"
            
            try {
                Invoke-WebRequest -Uri $Dep.Url -OutFile $zipPath
                Expand-Archive -Path $zipPath -DestinationPath $extractPath -Force
                
                # Find the DLL
                $foundDll = Get-ChildItem -Path $extractPath -Filter $Dep.Name -Recurse | Select-Object -First 1
                if ($foundDll) {
                    Copy-Item -Path $foundDll.FullName -Destination $dllPath
                    Write-Host "$($Dep.Name) successfully installed." -ForegroundColor Green
                } else {
                    Throw "Could not find $($Dep.Name) in the downloaded package."
                }
            }
            catch {
                Write-Error "Failed to download or install $($Dep.PackageName): $_"
                exit 1
            }
            finally {
                if (Test-Path $extractPath) { Remove-Item -Path $extractPath -Recurse -Force }
                if (Test-Path $zipPath) { Remove-Item -Path $zipPath -Force }
            }
        } else {
            Write-Error "$($Dep.Name) is required."
            exit 1
        }
    }
    return $dllPath
}

# --- Load Assemblies ---
foreach ($dep in $Dependencies) {
    $path = Install-Dependency -Dep $dep
    try {
        Add-Type -Path $path
    }
    catch {
        Write-Error "Failed to load $($dep.Name). Ensure the file is not blocked."
        exit 1
    }
}

# --- Logic ---
function Merge-Pdfs {
    param($Odd, $Even, $Out, $RevEven)
    # Sanitize input paths (remove quotes if pasted interactively)
    $Odd = "$Odd".Trim('"').Trim("'")
    $Even = "$Even".Trim('"').Trim("'")
    $Out = "$Out".Trim('"').Trim("'")

    if (-not (Test-Path $Odd)) { Throw "Odd PDF not found: $Odd" }
    if (-not (Test-Path $Even)) { Throw "Even PDF not found: $Even" }

    $readerOdd = $null
    $readerEven = $null
    $doc = $null
    $copy = $null
    $fs = $null

    try {
        $readerOdd = New-Object iTextSharp.text.pdf.PdfReader($Odd)
        $readerEven = New-Object iTextSharp.text.pdf.PdfReader($Even)

        $fs = [System.IO.File]::Create($Out)
        $doc = New-Object iTextSharp.text.Document
        $copy = New-Object iTextSharp.text.pdf.PdfCopy($doc, $fs)

        $doc.Open()

        $countOdd = $readerOdd.NumberOfPages
        $countEven = $readerEven.NumberOfPages
        $maxPages = [Math]::Max($countOdd, $countEven)

        Write-Host "Merging files..."
        Write-Host "Odd File: $countOdd pages"
        Write-Host "Even File: $countEven pages"

        for ($i = 1; $i -le $maxPages; $i++) {
            # Add Odd Page
            if ($i -le $countOdd) {
                # ImportPage returns a PdfImportedPage
                $page = $copy.GetImportedPage($readerOdd, $i)
                $copy.AddPage($page)
            }

            # Add Even Page
            if ($i -le $countEven) {
                if ($RevEven) {
                    $evenIndex = $countEven - $i + 1
                    # Double check we don't go out of bounds if logical mismatch, but loop bounds protect us mostly
                    # However if countEven < countOdd, we might try to access 0 or negative index if logic was wrong?
                    # No, i ranges 1 to max. 
                    # If i > countEven, we don't enter this block. 
                    # So $evenIndex logic is safe relative to i.
                } else {
                    $evenIndex = $i
                }

                $page = $copy.GetImportedPage($readerEven, $evenIndex)
                $copy.AddPage($page)
            }
        }

        Write-Host "Successfully created: $Out" -ForegroundColor Green
    }
    catch {
        Write-Error "An error occurred during merging: $_"
    }
    finally {
        if ($doc -ne $null) { $doc.Close() }
        if ($readerOdd -ne $null) { $readerOdd.Close() }
        if ($readerEven -ne $null) { $readerEven.Close() }
        if ($fs -ne $null) { $fs.Dispose() }
    }
}

Merge-Pdfs -Odd $OddPdfPath -Even $EvenPdfPath -Out $OutputPdfPath -RevEven $ReverseEven
