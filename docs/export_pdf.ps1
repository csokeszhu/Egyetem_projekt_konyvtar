param([string]$work = (Join-Path ([IO.Path]::GetTempPath()) "kef_export"))
# KEF_dokumentacio.docx -> KEF_dokumentacio.pdf (Microsoft Word kell hozza). A make_pdf.py hivja.
# Onalloan:  powershell -NoProfile -ExecutionPolicy Bypass -File docs\export_pdf.ps1
# (A fajl szandekosan ekezet nelkuli: a Windows PowerShell 5.1 a BOM nelkuli UTF-8 szkriptet
#  ANSI-kent olvassa.)
$log = Join-Path $work "export.log"
New-Item -ItemType Directory -Force $work | Out-Null
"start $(Get-Date -f HH:mm:ss)" | Out-File $log
$src = Join-Path $PSScriptRoot "KEF_dokumentacio.docx"
$dst = Join-Path $PSScriptRoot "KEF_dokumentacio.pdf"
$tmp = Join-Path $work ("d{0}.docx" -f (Get-Random))
$tmpPdf = $tmp -replace '\.docx$', '.pdf'
Copy-Item $src $tmp
$word = New-Object -ComObject Word.Application
"com ok $(Get-Date -f HH:mm:ss)" | Out-File $log -Append
$word.Visible = $false
$word.DisplayAlerts = 0
$doc = $word.Documents.Open($tmp, $false, $true)
"open ok $(Get-Date -f HH:mm:ss) pages=$($doc.ComputeStatistics(2))" | Out-File $log -Append
$doc.ExportAsFixedFormat($tmpPdf, 17)
"export ok $(Get-Date -f HH:mm:ss)" | Out-File $log -Append
$doc.Close($false)
$word.Quit()
"quit ok $(Get-Date -f HH:mm:ss)" | Out-File $log -Append
Copy-Item $tmpPdf $dst -Force
Remove-Item -LiteralPath $tmp, $tmpPdf -Force -ErrorAction SilentlyContinue
Write-Output "PDF: $dst"
