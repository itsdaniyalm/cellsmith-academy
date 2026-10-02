# Types every answer from the book into the practice workbook's yellow cells and
# confirms each self-check says Correct (and that a wrong answer says Not yet).
# Run after workbook.ps1. Does not save the workbook.
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$extract = Get-Content (Join-Path $root 'build\extract.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$path = Join-Path $root 'site\practice\cellsmith-academy-practice.xlsx'
$CHECK = [string][char]0x2713

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$failed = 0; $passed = 0
try {
  $wb = $xl.Workbooks.Open($path, 0, $true)
  foreach ($ch in $extract.chapters) {
    $list = @($extract.exercises | Where-Object { $_.chapter -eq $ch.id })
    if ($list.Count -eq 0) { continue }
    $ws = $wb.Worksheets.Item($ch.cell)
    $row = 1
    foreach ($e in $list) {
      # Find the "Exercise n" label, then the input cell is one row down in column B.
      while ([string]$ws.Cells.Item($row, 1).Value2 -ne "Exercise $($e.number)") { $row++; if ($row -gt 500) { throw "label not found: $($ch.cell) #$($e.number)" } }
      $inCell = $ws.Cells.Item($row + 1, 2)
      $status = $ws.Cells.Item($row + 2, 1)
      $inCell.Formula2 = $e.formula
      $xl.Calculate()
      $ok = ([string]$status.Text) -eq "$CHECK Correct"
      if ($ok) { $passed++ } else { $failed++; "FAIL $($ch.cell) #$($e.number): status '$($status.Text)' for $($e.formula)" }
      $inCell.ClearContents() | Out-Null
    }
  }
  # A wrong answer must not pass.
  $ws = $wb.Worksheets.Item('A1')
  $ws.Range('B5').Formula2 = '=47'
  $xl.Calculate()
  $wrong = [string]$ws.Range('A6').Text
  if ($wrong -ne 'Not yet') { $failed++; "FAIL wrong answer shows '$wrong'" } else { $passed++ }
  $wb.Close($false)
}
finally {
  $xl.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
  [GC]::Collect()
}
"$passed passed, $failed failed"
if ($failed) { exit 1 }
