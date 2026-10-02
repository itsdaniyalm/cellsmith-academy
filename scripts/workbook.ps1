# Builds the practice workbook and checks every example in real Excel.
#
#   1. node scripts/extract.mjs       (writes build/extract.json)
#   2. powershell -File scripts/workbook.ps1
#   3. node scripts/check.mjs          (compares results; --fill writes "→ ?" results)
#
# Needs Microsoft 365 Excel on Windows. Writes:
#   site/practice/cellsmith-academy-practice.xlsx
#   build/results.json   (what Excel calculated for each example and answer)
#   data/dataset.json    (the sample data, for the website's data page)
param([switch]$NoSave)

$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$extract = Get-Content (Join-Path $root 'build\extract.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$outXlsx = Join-Path $root 'site\practice\cellsmith-academy-practice.xlsx'
New-Item -ItemType Directory -Force (Split-Path $outXlsx), (Join-Path $root 'data') | Out-Null

$CHECK = [string][char]0x2713
$green = 65 * 65536 + 124 * 256 + 16          # #107C41 as BGR
$white = 16777215
$yellow = 207 * 65536 + 243 * 256 + 255       # #FFF3CF
$grey = 8421504

# ---------------------------------------------------------------------------
# The sample company: Brightside Supply Co. (deterministic, so the book's numbers never drift)
# ---------------------------------------------------------------------------
[long]$script:seed = 7
function Rnd { $script:seed = ($script:seed * 1103515245 + 12345) % 2147483648; return $script:seed / 2147483648 }
function Pick($items) { return $items[[int][math]::Floor((Rnd) * $items.Count)] }
function Iso([datetime]$d) { return $d.ToString('yyyy-MM-dd') }

$products = @(
  @('BS-101', 'Desk Lamp', 'Lighting', 34.00, 15.50),
  @('BS-102', 'Notebook', 'Paper', 6.50, 2.10),
  @('BS-103', 'Gel Pen Set', 'Writing', 9.00, 3.40),
  @('BS-104', 'Monitor Stand', 'Furniture', 49.00, 27.00),
  @('BS-105', 'Desk Chair', 'Furniture', 189.00, 112.00),
  @('BS-106', 'Cable Kit', 'Accessories', 14.00, 5.20),
  @('BS-107', 'Sticky Notes', 'Paper', 4.25, 1.30),
  @('BS-108', 'Whiteboard', 'Furniture', 79.00, 41.00)
)
$maxUnits = @{ 'BS-101' = 12; 'BS-102' = 60; 'BS-103' = 40; 'BS-104' = 10; 'BS-105' = 4; 'BS-106' = 25; 'BS-107' = 80; 'BS-108' = 5 }

$staff = @(
  @('Ava', 'Ava Thompson', 'North', [datetime]'2019-03-11'),
  @('Ben', 'Benjamin Okafor', 'North', [datetime]'2021-08-02'),
  @('Chloe', 'Chloe Martin', 'South', [datetime]'2018-01-15'),
  @('Dev', 'Devendra Rao', 'South', [datetime]'2023-05-22'),
  @('Elena', 'Elena Petrova', 'East', [datetime]'2020-10-05'),
  @('Farid', 'Farid Haddad', 'East', [datetime]'2022-02-14'),
  @('Grace', 'Grace Lee', 'West', [datetime]'2017-06-30'),
  @('Hiro', 'Hiro Tanaka', 'West', [datetime]'2024-09-09')
)
$repsByRegion = @{}
foreach ($s in $staff) { if (-not $repsByRegion[$s[2]]) { $repsByRegion[$s[2]] = @() }; $repsByRegion[$s[2]] += $s[0] }

$sales = @()
$start = [datetime]'2026-01-02'
for ($i = 0; $i -lt 48; $i++) {
  $region = Pick @('North', 'South', 'East', 'West')
  $rep = Pick $repsByRegion[$region]
  $p = Pick $products
  $units = 1 + [int][math]::Floor((Rnd) * $maxUnits[$p[0]])
  $disc = Pick @(0, 0, 0, 0, 0, 0, 0.05, 0.05, 0.1, 0.15)
  $r = Rnd
  $status = if ($r -lt 0.75) { 'Closed' } elseif ($r -lt 0.92) { 'Open' } else { 'Cancelled' }
  $date = $start.AddDays([math]::Floor($i * 3.7))
  $ship = if ($status -eq 'Closed') { $date.AddDays(1 + [int][math]::Floor((Rnd) * 7)) } else { $null }
  $sales += , @("SO-$(1001 + $i)", $date, $region, $rep, $p[0], $p[1], $units, $p[3], $disc, $status, $ship)
}

$leads = @(
  @('L-01', '  maria GARCIA ', 'Harbor & Pine', 'Austin, TX', '(512) 555-0142', 'Web'),
  @('L-02', "james o'neil", 'Redwood Analytics', 'Denver, CO', '720.555.0199', 'Referral'),
  @('L-03', 'PRIYA  SHARMA', 'Bluefin Logistics', 'Seattle, WA', '206 555 0175', 'Web'),
  @('L-04', 'li wei', 'Copperleaf Studio', 'Boston, MA', '617-555-0123', 'Event'),
  @('L-05', 'Sofia   Rossi', 'Northgate Clinic', 'Chicago, IL', '+1 312 555 0110', 'Web'),
  @('L-06', 'noah johnson ', 'Summit Foods', 'Miami, FL', '(305)555-0166', 'Referral')
)

$targets = @(
  @('North', 'Q1', 1500), @('South', 'Q1', 1200), @('East', 'Q1', 1800), @('West', 'Q1', 1500),
  @('North', 'Q2', 1200), @('South', 'Q2', 1500), @('East', 'Q2', 1200), @('West', 'Q2', 1000)
)
$tiers = @(@(0, 0.02), @(2000, 0.04), @(5000, 0.06), @(8000, 0.08))

# ---------------------------------------------------------------------------
function Put($ws, $row, $col, $rows) {
  $r = $rows.Count; $c = $rows[0].Count
  $arr = New-Object 'object[,]' $r, $c
  for ($i = 0; $i -lt $r; $i++) {
    for ($j = 0; $j -lt $c; $j++) {
      $v = $rows[$i][$j]
      if ($v -is [datetime]) { $v = $v.ToOADate() }
      $arr[$i, $j] = $v
    }
  }
  $range = $ws.Range($ws.Cells.Item($row, $col), $ws.Cells.Item($row + $r - 1, $col + $c - 1))
  $range.Value2 = $arr
  return , $range
}
function MakeTable($ws, $header, $rows, $name) {
  $range = Put $ws 1 1 (@(, $header) + $rows)
  $lo = $ws.ListObjects.Add(1, $range, [Type]::Missing, 1)
  $lo.Name = $name
  $lo.TableStyle = 'TableStyleMedium7'
  return $lo
}
function Fmt($lo, $col, $format) { $lo.ListColumns.Item($col).DataBodyRange.NumberFormat = $format }

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$results = @{ examples = @(); answers = @() }
try {
  $wb = $xl.Workbooks.Add()
  while ($wb.Worksheets.Count -gt 1) { $wb.Worksheets.Item($wb.Worksheets.Count).Delete() }
  $startWs = $wb.Worksheets.Item(1); $startWs.Name = 'Start here'
  $last = $startWs
  function NewSheet($name) { $ws = $wb.Worksheets.Add([Type]::Missing, $script:last); $ws.Name = $name; $script:last = $ws; return $ws }

  # --- data sheets -----------------------------------------------------------
  $ws = NewSheet 'Sales'
  $lo = MakeTable $ws @('Order ID', 'Date', 'Region', 'Rep', 'SKU', 'Product', 'Units', 'Price', 'Discount', 'Revenue', 'Status', 'Ship Date') `
    ($sales | ForEach-Object { , @($_[0], $_[1], $_[2], $_[3], $_[4], $_[5], $_[6], $_[7], $_[8], $null, $_[9], $_[10]) }) 'Sales'
  $lo.ListColumns.Item('Revenue').DataBodyRange.Formula2 = '=ROUND([@Units]*[@Price]*(1-[@Discount]),2)'
  Fmt $lo 'Date' 'yyyy-mm-dd'; Fmt $lo 'Ship Date' 'yyyy-mm-dd'; Fmt $lo 'Price' '#,##0.00'; Fmt $lo 'Revenue' '#,##0.00'; Fmt $lo 'Discount' '0%'
  $ws.Columns.AutoFit() | Out-Null

  $ws = NewSheet 'Products'
  $lo = MakeTable $ws @('SKU', 'Product', 'Category', 'Price', 'Cost') $products 'Products'
  Fmt $lo 'Price' '#,##0.00'; Fmt $lo 'Cost' '#,##0.00'
  $ws.Columns.AutoFit() | Out-Null

  $ws = NewSheet 'Staff'
  $lo = MakeTable $ws @('Rep', 'Name', 'Email', 'Region', 'Hire Date') ($staff | ForEach-Object {
      $parts = $_[1].ToLower().Split(' ')
      , @($_[0], $_[1], "$($_[0].ToLower()).$($parts[-1])@brightside.example", $_[2], $_[3])
    }) 'Staff'
  Fmt $lo 'Hire Date' 'yyyy-mm-dd'
  $ws.Columns.AutoFit() | Out-Null

  $ws = NewSheet 'Leads'
  $lo = MakeTable $ws @('Lead', 'Name', 'Company', 'Location', 'Phone', 'Source') $leads 'Leads'
  $ws.Columns.AutoFit() | Out-Null

  $ws = NewSheet 'Targets'
  $lo = MakeTable $ws @('Region', 'Quarter', 'Target') $targets 'Targets'
  Fmt $lo 'Target' '#,##0'
  $ws.Columns.AutoFit() | Out-Null

  $ws = NewSheet 'Tiers'
  $lo = MakeTable $ws @('Min Sales', 'Rate') $tiers 'Tiers'
  Fmt $lo 'Min Sales' '#,##0'; Fmt $lo 'Rate' '0%'
  $ws.Columns.AutoFit() | Out-Null

  $ws = NewSheet 'Settings'
  Put $ws 1 1 @(@('Setting', 'Value'), @('Tax rate', 0.0825), @('Report start', ([datetime]'2026-01-01')), @('Report end', ([datetime]'2026-06-30')), @('Minimum order', 100)) | Out-Null
  $ws.Range('A1:B1').Font.Bold = $true
  $ws.Range('B2').NumberFormat = '0.00%'; $ws.Range('B3:B4').NumberFormat = 'yyyy-mm-dd'
  $ws.Columns.AutoFit() | Out-Null
  $wb.Names.Add('TaxRate', '=Settings!$B$2') | Out-Null
  $wb.Names.Add('ReportStart', '=Settings!$B$3') | Out-Null
  $wb.Names.Add('ReportEnd', '=Settings!$B$4') | Out-Null
  $wb.Names.Add('MinOrder', '=Settings!$B$5') | Out-Null
  # Named LAMBDA functions used in chapter C5.
  $n = $wb.Names.Add('Margin', '=LAMBDA(price, cost, (price - cost) / price)'); $n.Comment = 'Profit as a fraction of price.'
  $n = $wb.Names.Add('WithTax', '=LAMBDA(amount, [rate], amount * (1 + IF(ISOMITTED(rate), TaxRate, rate)))'); $n.Comment = 'Adds tax at TaxRate, or at the rate you give.'
  $n = $wb.Names.Add('QuarterOf', '=LAMBDA(date, "Q" & ROUNDUP(MONTH(date) / 3, 0))'); $n.Comment = 'The calendar quarter of a date, like Q2.'
  $xl.Calculate()

  # --- check every example ------------------------------------------------
  function Evaluate($ws, $cell, $formula) {
    $res = @{ text = $null; values = $null; error = $null }
    try {
      $cell.Formula2 = $formula
      $xl.Calculate()
      # Show dates and times the same way for every reader, whatever their regional settings.
      $nf = [string]$cell.NumberFormat
      $target = if ($cell.HasSpill) { $cell.SpillingToRange } else { $cell }
      if ($nf -match '[dy]') { $target.NumberFormat = if ($nf -match 'h') { 'yyyy-mm-dd hh:mm' } else { 'yyyy-mm-dd' } }
      elseif ($nf -match 'h') { $target.NumberFormat = 'hh:mm' }
      if ($cell.HasSpill) {
        $spill = $cell.SpillingToRange
        $rows = @()
        for ($r = 1; $r -le $spill.Rows.Count; $r++) {
          $row = @()
          for ($c = 1; $c -le $spill.Columns.Count; $c++) { $row += [string]$spill.Cells.Item($r, $c).Text }
          $rows += , $row
        }
        $res.values = $rows
      }
      $res.text = [string]$cell.Text
    }
    catch { $res.error = $_.Exception.Message }
    return $res
  }

  $verify = NewSheet '__verify'
  $verify.Columns.Item(1).ColumnWidth = 120
  foreach ($ex in $extract.examples) {
    $cell = $verify.Range('A1')
    $res = Evaluate $verify $cell $ex.formula
    $res.id = $ex.id
    $results.examples += $res
    $verify.Cells.Clear() | Out-Null
  }

  # --- chapter practice sheets ---------------------------------------------
  $answers = NewSheet 'Answers'
  $answers.Columns.Item(2).ColumnWidth = 40
  $answerRow = 1
  foreach ($ch in $extract.chapters) {
    $list = @($extract.exercises | Where-Object { $_.chapter -eq $ch.id })
    if ($list.Count -eq 0) { continue }
    $ws = NewSheet $ch.cell
    $ws.Range('A1').Value2 = "$($ch.cell)  $($ch.title)"
    $ws.Range('A1').Font.Size = 16; $ws.Range('A1').Font.Bold = $true
    $ws.Range('A2').Value2 = "Type your formula in each yellow cell. The check below it turns green when your result matches."
    $ws.Range('A2').Font.Color = $grey
    $ws.Columns.Item(1).ColumnWidth = 14
    $ws.Columns.Item(2).ColumnWidth = 22
    $row = 4
    foreach ($e in $list) {
      $anchor = $answers.Cells.Item($answerRow, 2)
      $answers.Cells.Item($answerRow, 1).Value2 = "$($ch.cell) #$($e.number)"
      $res = Evaluate $answers $anchor $e.formula
      $res.chapter = $e.chapter; $res.number = $e.number
      $results.answers += $res
      $h = 1; $w = 1
      if ($anchor.HasSpill) { $h = $anchor.SpillingToRange.Rows.Count; $w = $anchor.SpillingToRange.Columns.Count }

      $ws.Cells.Item($row, 1).Value2 = "Exercise $($e.number)"
      $ws.Cells.Item($row, 1).Font.Bold = $true
      $ws.Cells.Item($row, 2).Value2 = $e.prompt
      $inCell = $ws.Cells.Item($row + 1, 2)
      $ws.Cells.Item($row + 1, 1).Value2 = 'Your formula'
      $ws.Cells.Item($row + 1, 1).Font.Color = $grey
      $inCell.Interior.Color = $yellow
      $in = $inCell.Address($false, $false)
      $ans = "Answers!" + $anchor.Address($true, $true)
      $status = $ws.Cells.Item($row + 2, 1)
      if ($h -gt 1 -or $w -gt 1) {
        $status.Formula2 = "=IF(ISBLANK($in),`"`",IFERROR(IF(AND(ROWS($in#)=ROWS($ans#),COLUMNS($in#)=COLUMNS($ans#),AND($in#=$ans#)),`"$CHECK Correct`",`"Not yet`"),`"Not yet`"))"
      }
      else {
        $status.Formula2 = "=IF(ISBLANK($in),`"`",IF(IFERROR(AND($in=$ans),FALSE),`"$CHECK Correct`",`"Not yet`"))"
      }
      $fc = $status.FormatConditions.Add(1, 3, "=`"$CHECK Correct`"")
      $fc.Interior.Color = 206 * 65536 + 239 * 256 + 198
      $fc.Font.Color = 0 * 65536 + 97 * 256 + 0
      $status.Font.Bold = $true
      $row += [math]::Max(4, $h + 3)
      $answerRow += $h + 1
    }
    $ws.Range('B4').EntireColumn.WrapText = $false
  }
  $answers.Visible = 2   # very hidden: answers are in the book

  # --- start sheet ---------------------------------------------------------
  $s = $startWs
  $s.Range('A1').Value2 = 'Cellsmith Academy practice workbook'
  $s.Range('A1').Font.Size = 20; $s.Range('A1').Font.Bold = $true
  $lines = @(
    'This workbook goes with the free course at itsdaniyalm.github.io/cellsmith-academy',
    '',
    'Data sheets (used by every example in the book):',
    '   Sales      48 orders from Brightside Supply Co., January to June 2026 (table: Sales)',
    '   Products   8 products with price and cost (table: Products)',
    '   Staff      8 sales reps (table: Staff)',
    '   Leads      6 messy contact records for the text chapter (table: Leads)',
    '   Targets    quarterly revenue targets by region (table: Targets)',
    '   Tiers      commission rates by sales band (table: Tiers)',
    '   Settings   named cells TaxRate, ReportStart, ReportEnd and MinOrder',
    '',
    'Exercise sheets are named after chapters (A1, A2, B1 ...). Type a formula in a yellow cell;',
    'the check below it shows Correct when the result matches. Answers are in the book.',
    '',
    'Made by Muhammad Daniyal. Free to use and share under CC BY 4.0.'
  )
  for ($i = 0; $i -lt $lines.Count; $i++) { $s.Cells.Item($i + 3, 1).Value2 = $lines[$i] }
  $s.Columns.Item(1).ColumnWidth = 100

  $verify.Delete()
  $startWs.Activate()

  if (-not $NoSave) {
    # Document properties need late binding from PowerShell.
    $props = $wb.BuiltinDocumentProperties
    $binding = [System.Reflection.BindingFlags]
    foreach ($prop in @(@('Title', 'Cellsmith Academy practice workbook'), @('Author', 'Cellsmith Academy'), @('Last Author', 'Cellsmith Academy'), @('Subject', 'Excel formula exercises'), @('Company', ''))) {
      $p = [System.__ComObject].InvokeMember('Item', $binding::GetProperty, $null, $props, @($prop[0]))
      [System.__ComObject].InvokeMember('Value', $binding::SetProperty, $null, $p, @($prop[1])) | Out-Null
    }
    if (Test-Path $outXlsx) { Remove-Item $outXlsx }
    $wb.SaveAs($outXlsx, 51)
  }

  # --- export the data for the website --------------------------------------
  $dataset = [ordered]@{}
  foreach ($name in 'Sales', 'Products', 'Staff', 'Leads', 'Targets', 'Tiers') {
    $lo = $null
    foreach ($w in $wb.Worksheets) { foreach ($l in $w.ListObjects) { if ($l.Name -eq $name) { $lo = $l } } }
    $cols = @(); foreach ($c in $lo.ListColumns) { $cols += $c.Name }
    $rows = @()
    for ($r = 1; $r -le $lo.ListRows.Count; $r++) {
      $row = @(); foreach ($c in 1..$cols.Count) { $row += [string]$lo.DataBodyRange.Cells.Item($r, $c).Text }
      $rows += , $row
    }
    $dataset[$name] = [ordered]@{ columns = $cols; rows = $rows }
  }
  $dataset | ConvertTo-Json -Depth 6 | Set-Content (Join-Path $root 'data\dataset.json') -Encoding UTF8
  $wb.Close($false)
}
finally {
  $xl.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
  [GC]::Collect()
}
$results | ConvertTo-Json -Depth 8 | Set-Content (Join-Path $root 'build\results.json') -Encoding UTF8
"checked $($results.examples.Count) examples and $($results.answers.Count) answers"
