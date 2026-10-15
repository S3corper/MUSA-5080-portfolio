$ErrorActionPreference='Stop'
$taskDir=Split-Path $PSScriptRoot -Parent
$app=New-Object -ComObject Word.Application
$app.Visible=$false
$app.DisplayAlerts=0
try {
 $doc=$app.Documents.Open((Join-Path $taskDir 'Traffic Impact Written Answers.docx'),$false,$true)
 $doc.Repaginate()
 Write-Output ('Word pages: '+$doc.ComputeStatistics(2))
 $doc.ExportAsFixedFormat((Join-Path $PSScriptRoot 'written.pdf'),17)
 $doc.Close(0)
} finally {$app.Quit()}
