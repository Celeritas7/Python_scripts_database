Write-Host ""
Write-Host "Select an operation for CURRENT FOLDER:" -ForegroundColor Cyan
Write-Host "1) Images → PDF"
Write-Host "2) Shrink images"
Write-Host "3) DOCX → PDF"
Write-Host ""

$choice = Read-Host "Enter choice (1-3)"

switch ($choice) {
    "1" {
        python "C:\tools\py-utils\images_to_pdf_converter_R001.py"
    }
    "2" {
        python "C:\tools\py-utils\Image_shrink_R002.py"
    }
    "3" {
        python "C:\tools\py-utils\docx_to_pdf_converter_R001.py"
    }
    default {
        Write-Host "Invalid choice" -ForegroundColor Red
    }
}