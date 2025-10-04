# WTFPL ChatGPT-4-turbo

 param(
    [Parameter(Mandatory = $true)][string]$InputFile,
    [Parameter(Mandatory = $true)][string]$OutputFile
)

$ext = [IO.Path]::GetExtension($InputFile).ToLower()

if ($ext -eq ".doc" -or $ext -eq ".docx") {
    $format = 17  # wdFormatPDF

    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0

    try {
        $doc = $word.Documents.Open($InputFile, [ref]$false, [ref]$true)
        $doc.SaveAs([ref]$OutputFile, [ref]$format)
        $doc.Close()
        Write-Host "✅ Word документ конвертирован в PDF"
    } catch {
        Write-Host "❌ Ошибка при конвертации Word: $_"
        exit 1
    } finally {
        $word.Quit()
    }

} elseif ($ext -eq ".ppt" -or $ext -eq ".pptx") {
    $format = 32  # ppSaveAsPDF

    $ppt = New-Object -ComObject PowerPoint.Application
    # $ppt.Visible = [Microsoft.Office.Core.MsoTriState]::msoFalse  ← Удалено

    try {
        $presentation = $ppt.Presentations.Open($InputFile, [ref]$false, [ref]$true)
        $presentation.SaveAs([ref]$OutputFile, [ref]$format)
        $presentation.Close()
        Write-Host "✅ PowerPoint документ конвертирован в PDF"
    } catch {
        Write-Host "❌ Ошибка при конвертации PowerPoint: $_"
        exit 1
    } finally {
        $ppt.Quit()
    }
} elseif ($ext -eq ".heic") {
    try {
        Add-Type -AssemblyName System.Drawing
        $img = [System.Drawing.Image]::FromFile($InputFile)
        $img.Save($OutputFile, [System.Drawing.Imaging.ImageFormat]::Png)
        $img.Dispose()
        Write-Host "✅ HEIC изображение конвертировано в PNG"
    } catch {
        Write-Host "❌ Ошибка при конвертации HEIC: $_"
        exit 1
    }

} else {
    Write-Host "❌ Неподдерживаемый формат: $ext"
    exit 1
}

