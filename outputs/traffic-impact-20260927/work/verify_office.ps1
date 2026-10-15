$ErrorActionPreference = 'Stop'
$taskDir = Split-Path $PSScriptRoot -Parent
$wordApp = $null
$wordDoc = $null
try {
  $wordApp = New-Object -ComObject Word.Application
  $wordApp.Visible = $false
  $wordApp.DisplayAlerts = 0
  $wordDoc = $wordApp.Documents.Open((Join-Path $taskDir 'Traffic Impact Written Answers.docx'), $false, $true)
  $wordDoc.Repaginate()
  Write-Output ('Word pages: ' + $wordDoc.ComputeStatistics(2))
  $wordDoc.ExportAsFixedFormat((Join-Path $PSScriptRoot 'written.pdf'),17)
} finally {
  if ($null -ne $wordDoc) { $wordDoc.Close(0) }
  if ($null -ne $wordApp) { $wordApp.Quit() }
}
$excelApp = $null
$excelBook = $null
try {
  $excelApp = New-Object -ComObject Excel.Application
  $excelApp.Visible = $false
  $excelApp.DisplayAlerts = $false
  $excelBook = $excelApp.Workbooks.Open((Join-Path $taskDir 'Traffic Impact Calculations.xlsx'),0,$true)
  $excelApp.CalculateFullRebuild()
  $lane = $excelBook.Worksheets.Item('SW added lane')
  $traffic = $excelBook.Worksheets.Item('Traffic analysis')
  $inputs = $excelBook.Worksheets.Item('Inputs')
  foreach ($pair in @(@('D6',3113.4666666666667),@('D7',32691400),@('D9',40864250))) {
    $actual = $lane.Range($pair[0]).Value2
    if ([Math]::Abs($actual-$pair[1]) -gt 0.00001) { throw ('Calculation mismatch '+$pair[0]+': '+$actual) }
    Write-Output ($pair[0]+': '+$actual)
  }
  $inputs.Range('G7').Value2 = 0
  $excelApp.CalculateFullRebuild()
  if ($traffic.Range('F8').Value2 -ne 10000) { throw 'Office input dependency failed' }
  $inputs.Range('G7').Value2 = 800000
  $inputs.Range('G14').Value2 = 0
  $excelApp.CalculateFullRebuild()
  if ($lane.Range('D6').Value2 -ne 0) { throw 'Lane input dependency failed' }
  $inputs.Range('G14').Value2 = 1
  $excelApp.CalculateFullRebuild()
  foreach ($sheet in $excelBook.Worksheets) {
    $cells = $sheet.UsedRange
    for ($r=1; $r -le $cells.Rows.Count; $r++) {
      for ($c=1; $c -le $cells.Columns.Count; $c++) {
        $txt = $cells.Cells.Item($r,$c).Text
        if ($txt -match '^#(REF!|DIV/0!|VALUE!|NAME\?|N/A|NUM!|NULL!|SPILL!|CALC!)$') { throw ('Excel error: '+$sheet.Name+' '+$txt) }
      }
    }
  }
  Write-Output 'Excel recalculation, input changes, and formula scan passed.'
} finally {
  if ($null -ne $excelBook) { $excelBook.Close($false) }
  if ($null -ne $excelApp) { $excelApp.Quit() }
}
