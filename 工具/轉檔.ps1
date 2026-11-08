# 把整理好的 Markdown 轉成 Word（.docx）與 PDF
# 用法（在專案根目錄）：
#   .\工具\轉檔.ps1 整理結果\某份文件.md            # 產生同名 .docx 與 .pdf
#   .\工具\轉檔.ps1 整理結果\某份文件.md -只要Word
# 需要：pandoc（md → docx）、Microsoft Word（docx → pdf）
param(
    [Parameter(Mandatory = $true)][string]$Path,
    [switch]$只要Word
)
$ErrorActionPreference = 'Stop'

$md = (Resolve-Path $Path).Path
$docx = [IO.Path]::ChangeExtension($md, '.docx')
$pdf = [IO.Path]::ChangeExtension($md, '.pdf')
$reference = Join-Path $PSScriptRoot 'reference.docx'   # 中文字型樣式範本

try {
    $pandocArgs = @($md, '-o', $docx, '--from', 'gfm')
    if (Test-Path $reference) { $pandocArgs += "--reference-doc=$reference" }
    & pandoc @pandocArgs
    if ($LASTEXITCODE -ne 0) { throw "pandoc 轉換失敗（代碼 $LASTEXITCODE）" }
    Write-Host "Word：$docx"
}
catch {
    Write-Error "轉 Word 失敗：$($_.Exception.Message)"
    exit 1
}

if ($只要Word) { return }

$word = $null
$doc = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $doc = $word.Documents.Open($docx, $false, $true, $false)
    $doc.ExportAsFixedFormat($pdf, 17)   # 17 = PDF
    Write-Host "PDF：$pdf"
}
catch {
    Write-Error "轉 PDF 失敗（需要 Microsoft Word）：$($_.Exception.Message)"
    exit 1
}
finally {
    if ($doc) { $doc.Close(0) }
    if ($word) { $word.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word) }
}
