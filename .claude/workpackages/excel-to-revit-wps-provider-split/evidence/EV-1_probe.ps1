# EV-1 — WPS ProgID + late-bound member probe
# Package: excel-to-revit-wps-provider-split
# Runbook source: 04_EVIDENCE_QUEUE.md (EV-1), results/T4.2_result.md
# Deviations from the operator runbook, deliberate and recorded:
#   - WPS version is read from et.exe FileVersionInfo instead of the About dialog (more precise, no GUI).
#   - The test workbook is a scratch COPY of the user's file so the original is never opened/locked.
#   - Read-Host is replaced by a hardcoded scratch path so the probe runs non-interactively.
#   - Enrichment: Worksheets.Count, Names.Count, sheet-name enumeration, Names enumeration.
# Everything else (member list, BindingFlags, ExportAsFixedFormat 8-arg shape) is unchanged.

$ErrorActionPreference = 'Continue'

$sourceWorkbook = 'C:\Users\ADMIN\Desktop\PA4\BULONG.xlsx'

function Write-Section($title) {
  Write-Output ''
  Write-Output ('=' * 78)
  Write-Output ("== $title")
  Write-Output ('=' * 78)
}

Write-Section 'STEP 0 - probe metadata'
Write-Output ("timestamp_local      : " + (Get-Date).ToString('yyyy-MM-dd HH:mm:ss zzz'))
Write-Output ("machine              : " + $env:COMPUTERNAME)
Write-Output ("powershell_version   : " + $PSVersionTable.PSVersion.ToString())
Write-Output ("process_is_64bit     : " + [Environment]::Is64BitProcess)
Write-Output ("os_is_64bit          : " + [Environment]::Is64BitOperatingSystem)
Write-Output ("source_workbook      : " + $sourceWorkbook)
Write-Output ("source_workbook_exists: " + (Test-Path -LiteralPath $sourceWorkbook))

Write-Section 'STEP 1/2 - tasklist BEFORE probe'
$before = & tasklist 2>&1 | Select-String -Pattern 'et\.exe|wps\.exe|excel\.exe' -CaseSensitive:$false
if ($before) { $before | ForEach-Object { Write-Output $_.Line } } else { Write-Output '(no matching process: et.exe / wps.exe / excel.exe)' }

Write-Section 'STEP 3 - WPS version / build (from et.exe FileVersionInfo)'
$etCandidates = @()
foreach ($root in @("$env:LOCALAPPDATA\Kingsoft\WPS Office", "$env:ProgramFiles\Kingsoft\WPS Office", "${env:ProgramFiles(x86)}\Kingsoft\WPS Office")) {
  if (Test-Path -LiteralPath $root) {
    $etCandidates += Get-ChildItem -LiteralPath $root -Recurse -Filter 'et.exe' -ErrorAction SilentlyContinue
    $etCandidates += Get-ChildItem -LiteralPath $root -Recurse -Filter 'wps.exe' -ErrorAction SilentlyContinue
  }
}
if ($etCandidates.Count -eq 0) {
  Write-Output '(et.exe / wps.exe not found under any known WPS install root)'
} else {
  $etCandidates | Select-Object -Property FullName,
    @{n='FileVersion';e={$_.VersionInfo.FileVersion}},
    @{n='ProductVersion';e={$_.VersionInfo.ProductVersion}},
    @{n='ProductName';e={$_.VersionInfo.ProductName}} |
    Format-List | Out-String -Width 200 | Write-Output
}

Write-Section 'STEP 4 - ProgID resolution probe'
$progIds = @('KET.Application','ET.Application','Kingsoft.ET.Application','KWPS.Application','Excel.Application')
$progRows = foreach ($id in $progIds) {
  try {
    $t = [Type]::GetTypeFromProgID($id, $false)
    if ($null -eq $t) {
      [pscustomobject]@{ ProgID = $id; Resolved = $false; TypeName = $null; CLSID = $null; Error = $null }
    }
    else {
      [pscustomobject]@{ ProgID = $id; Resolved = $true; TypeName = $t.FullName; CLSID = $t.GUID.Guid; Error = $null }
    }
  }
  catch {
    [pscustomobject]@{ ProgID = $id; Resolved = 'ERROR'; TypeName = $null; CLSID = $null; Error = $_.Exception.Message }
  }
}
$progRows | Format-Table -AutoSize | Out-String -Width 200 | Write-Output

Write-Section 'STEP 5 - prepare scratch copy of the test workbook'
$scratch = Join-Path $env:TEMP ('ArcTool_EV1_scratch_' + [guid]::NewGuid().ToString('N') + '.xlsx')
$scratchReady = $false
try {
  Copy-Item -LiteralPath $sourceWorkbook -Destination $scratch -ErrorAction Stop
  $scratchReady = Test-Path -LiteralPath $scratch
  Write-Output ("scratch_copy         : " + $scratch)
  Write-Output ("scratch_copy_exists  : " + $scratchReady)
  Write-Output ("scratch_size_bytes   : " + (Get-Item -LiteralPath $scratch).Length)
}
catch {
  Write-Output ("SCRATCH COPY FAILED  : " + $_.Exception.Message)
}

Write-Section 'STEP 6 - late-bound member probe on KET.Application'

$progId = 'KET.Application'
$type = [Type]::GetTypeFromProgID($progId, $false)
if ($null -eq $type) {
  Write-Output "FATAL: ProgID not registered: $progId - member probe cannot run."
}
elseif (-not $scratchReady) {
  Write-Output "FATAL: scratch workbook copy unavailable - member probe cannot run."
}
else {

$app = $null
$books = $null
$book = $null
$sheets = $null
$sheet = $null
$names = $null
$nameObj = $null
$range = $null
$pageSetup = $null
$flagsGet = [Reflection.BindingFlags]'Public,Instance,GetProperty'
$flagsSet = [Reflection.BindingFlags]'Public,Instance,SetProperty'
$flagsCall = [Reflection.BindingFlags]'Public,Instance,InvokeMethod'
$results = New-Object System.Collections.Generic.List[object]
function Add-Result($name, $present, $detail) {
  $results.Add([pscustomobject]@{ Member = $name; Present = $present; Detail = $detail })
}

try {
  try {
    $app = [Activator]::CreateInstance($type)
    Add-Result 'CreateInstance' 'present' ($app.GetType().FullName)
  } catch { Add-Result 'CreateInstance' 'absent/error' $_.Exception.Message }

  if ($app) {
    try { $type.InvokeMember('Visible', $flagsSet, $null, $app, @($false)); Add-Result 'Visible(set)' 'present' $null } catch { Add-Result 'Visible(set)' 'absent/error' $_.Exception.Message }
    try { $type.InvokeMember('DisplayAlerts', $flagsSet, $null, $app, @($false)); Add-Result 'DisplayAlerts(set)' 'present' $null } catch { Add-Result 'DisplayAlerts(set)' 'absent/error' $_.Exception.Message }
    try { $v = $type.InvokeMember('Version', $flagsGet, $null, $app, $null); Add-Result 'Application.Version(get)' 'present' ([string]$v) } catch { Add-Result 'Application.Version(get)' 'absent/error' $_.Exception.Message }
    try { $books = $type.InvokeMember('Workbooks', $flagsGet, $null, $app, $null); Add-Result 'Workbooks' 'present' $null } catch { Add-Result 'Workbooks' 'absent/error' $_.Exception.Message }
  }

  if ($books) {
    $bookType = $books.GetType()
    try { $book = $bookType.InvokeMember('Open', $flagsCall, $null, $books, @($scratch)); Add-Result 'Workbooks.Open' 'present' $scratch } catch { Add-Result 'Workbooks.Open' 'absent/error' $_.Exception.Message }
  }

  if ($book) {
    $bookType = $book.GetType()
    try { $sheets = $bookType.InvokeMember('Worksheets', $flagsGet, $null, $book, $null); Add-Result 'Worksheets' 'present' $null } catch { Add-Result 'Worksheets' 'absent/error' $_.Exception.Message }
    try { $names = $bookType.InvokeMember('Names', $flagsGet, $null, $book, $null); Add-Result 'Names' 'present' $null } catch { Add-Result 'Names' 'absent/error' $_.Exception.Message }
  }

  $sheetCount = $null
  if ($sheets) {
    $sheetsType = $sheets.GetType()
    try { $sheetCount = $sheetsType.InvokeMember('Count', $flagsGet, $null, $sheets, $null); Add-Result 'Worksheets.Count' 'present' ([string]$sheetCount) } catch { Add-Result 'Worksheets.Count' 'absent/error' $_.Exception.Message }
    try { $sheet = $sheetsType.InvokeMember('Item', $flagsGet, $null, $sheets, @(1)); Add-Result 'Worksheets.Item(1)' 'present' $null } catch { Add-Result 'Worksheets.Item(1)' 'absent/error' $_.Exception.Message }
    if ($sheetCount) {
      $sheetNames = @()
      for ($i = 1; $i -le [int]$sheetCount; $i++) {
        try {
          $s = $sheetsType.InvokeMember('Item', $flagsGet, $null, $sheets, @($i))
          $sheetNames += ('[' + $i + '] ' + [string]$s.GetType().InvokeMember('Name', $flagsGet, $null, $s, $null))
        } catch { $sheetNames += ('[' + $i + '] ERROR: ' + $_.Exception.Message) }
      }
      Add-Result 'Worksheets enumeration' 'present' ($sheetNames -join ' | ')
    }
  }

  if ($sheet) {
    $sheetType = $sheet.GetType()
    try { $n = $sheetType.InvokeMember('Name', $flagsGet, $null, $sheet, $null); Add-Result 'Worksheet.Name' 'present' ([string]$n) } catch { Add-Result 'Worksheet.Name' 'absent/error' $_.Exception.Message }
    try { $range = $sheetType.InvokeMember('UsedRange', $flagsGet, $null, $sheet, $null); Add-Result 'UsedRange' 'present' $null } catch { Add-Result 'UsedRange' 'absent/error' $_.Exception.Message }
    try { $pageSetup = $sheetType.InvokeMember('PageSetup', $flagsGet, $null, $sheet, $null); Add-Result 'PageSetup' 'present' $null } catch { Add-Result 'PageSetup' 'absent/error' $_.Exception.Message }
    try { $sheetType.InvokeMember('Range', $flagsGet, $null, $sheet, @('A1')) | Out-Null; Add-Result 'Range("A1")' 'present' $null } catch { Add-Result 'Range("A1")' 'absent/error' $_.Exception.Message }
  }

  if ($names) {
    $namesType = $names.GetType()
    $nameCount = $null
    try { $nameCount = $namesType.InvokeMember('Count', $flagsGet, $null, $names, $null); Add-Result 'Names.Count' 'present' ([string]$nameCount) } catch { Add-Result 'Names.Count' 'absent/error' $_.Exception.Message }
    try { $nameObj = $namesType.InvokeMember('Item', $flagsGet, $null, $names, @(1)); Add-Result 'Names.Item(1)' 'present' $null } catch { Add-Result 'Names.Item(1)' 'absent/error' $_.Exception.Message }
    if ($nameCount -and [int]$nameCount -gt 0) {
      $nameList = @()
      for ($i = 1; $i -le [int]$nameCount; $i++) {
        try {
          $nm = $namesType.InvokeMember('Item', $flagsGet, $null, $names, @($i))
          $nmType = $nm.GetType()
          $nmName = [string]$nmType.InvokeMember('Name', $flagsGet, $null, $nm, $null)
          $nmRefers = 'n/a'
          try { $nmRefers = [string]$nmType.InvokeMember('RefersTo', $flagsGet, $null, $nm, $null) } catch { $nmRefers = 'RefersTo ERROR: ' + $_.Exception.Message }
          $nameList += ('[' + $i + '] ' + $nmName + ' -> ' + $nmRefers)
        } catch { $nameList += ('[' + $i + '] ERROR: ' + $_.Exception.Message) }
      }
      Add-Result 'Names enumeration' 'present' ($nameList -join ' | ')
    }
  }

  if ($nameObj) {
    $nameType = $nameObj.GetType()
    try { $nameType.InvokeMember('Name', $flagsGet, $null, $nameObj, $null) | Out-Null; Add-Result 'Name.Name' 'present' $null } catch { Add-Result 'Name.Name' 'absent/error' $_.Exception.Message }
    try { $range = $nameType.InvokeMember('RefersToRange', $flagsGet, $null, $nameObj, $null); Add-Result 'RefersToRange' 'present' $null } catch { Add-Result 'RefersToRange' 'absent/error' $_.Exception.Message }
  }

  if ($range) {
    $rangeType = $range.GetType()
    try { $rangeType.InvokeMember('Worksheet', $flagsGet, $null, $range, $null) | Out-Null; Add-Result 'Range.Worksheet' 'present' $null } catch { Add-Result 'Range.Worksheet' 'absent/error' $_.Exception.Message }
    try { $addr = $rangeType.InvokeMember('Address', $flagsGet, $null, $range, @($false,$false)); Add-Result 'Range.Address(false,false)' 'present' ([string]$addr) } catch { Add-Result 'Range.Address(false,false)' 'absent/error' $_.Exception.Message }
    try { $addr2 = $rangeType.InvokeMember('Address', ($flagsGet -bor [Reflection.BindingFlags]'InvokeMethod'), $null, $range, @($false,$false)); Add-Result 'Range.Address(GetProperty|InvokeMethod)' 'present' ([string]$addr2) } catch { Add-Result 'Range.Address(GetProperty|InvokeMethod)' 'absent/error' $_.Exception.Message }
  }

  if ($pageSetup) {
    $psType = $pageSetup.GetType()
    foreach ($member in 'PrintArea','Zoom','FitToPagesWide','FitToPagesTall','TopMargin','BottomMargin','LeftMargin','RightMargin','PaperSize') {
      try { $val = $psType.InvokeMember($member, $flagsGet, $null, $pageSetup, $null); Add-Result "PageSetup.$member(get)" 'present' ([string]$val) } catch { Add-Result "PageSetup.$member(get)" 'absent/error' $_.Exception.Message }
    }
    # Setter probe: the MS provider WRITES these during normalization, so read-only presence is not enough.
    try { $psType.InvokeMember('Zoom', $flagsSet, $null, $pageSetup, @($false)); Add-Result 'PageSetup.Zoom(set false)' 'present' $null } catch { Add-Result 'PageSetup.Zoom(set false)' 'absent/error' $_.Exception.Message }
    try { $psType.InvokeMember('FitToPagesWide', $flagsSet, $null, $pageSetup, @(1)); Add-Result 'PageSetup.FitToPagesWide(set 1)' 'present' $null } catch { Add-Result 'PageSetup.FitToPagesWide(set 1)' 'absent/error' $_.Exception.Message }
    try { $psType.InvokeMember('FitToPagesTall', $flagsSet, $null, $pageSetup, @(1)); Add-Result 'PageSetup.FitToPagesTall(set 1)' 'present' $null } catch { Add-Result 'PageSetup.FitToPagesTall(set 1)' 'absent/error' $_.Exception.Message }
    foreach ($m in 'TopMargin','BottomMargin','LeftMargin','RightMargin') {
      try { $psType.InvokeMember($m, $flagsSet, $null, $pageSetup, @(0)); Add-Result "PageSetup.$m(set 0)" 'present' $null } catch { Add-Result "PageSetup.$m(set 0)" 'absent/error' $_.Exception.Message }
    }
    try { $psType.InvokeMember('PaperSize', $flagsSet, $null, $pageSetup, @(8)); Add-Result 'PageSetup.PaperSize(set 8 = xlPaperA3)' 'present' $null } catch { Add-Result 'PageSetup.PaperSize(set 8 = xlPaperA3)' 'absent/error' $_.Exception.Message }
    try { $psType.InvokeMember('PaperSize', $flagsSet, $null, $pageSetup, @(26)); Add-Result 'PageSetup.PaperSize(set 26 = xlPaperEsheet)' 'present' $null } catch { Add-Result 'PageSetup.PaperSize(set 26 = xlPaperEsheet)' 'absent/error' $_.Exception.Message }
    try { $psType.InvokeMember('PrintArea', $flagsSet, $null, $pageSetup, @('')); Add-Result 'PageSetup.PrintArea(set empty)' 'present' $null } catch { Add-Result 'PageSetup.PrintArea(set empty)' 'absent/error' $_.Exception.Message }
  }

  $tempPdf = $null
  if ($sheet) {
    $sheetType = $sheet.GetType()
    $tempPdf = Join-Path $env:TEMP ('ArcTool_EV1_' + [guid]::NewGuid().ToString('N') + '.pdf')
    try {
      $sheetType.InvokeMember('ExportAsFixedFormat', $flagsCall, $null, $sheet, @(0, $tempPdf, 0, $false, $false, 1, 1, $false))
      Add-Result 'ExportAsFixedFormat(8-arg)' 'present' $tempPdf
    } catch {
      Add-Result 'ExportAsFixedFormat(8-arg)' 'absent/error' $_.Exception.Message
    }
    Start-Sleep -Seconds 2
    if (Test-Path -LiteralPath $tempPdf) {
      Add-Result 'PDF File.Exists post-export' 'present' ('size_bytes=' + (Get-Item -LiteralPath $tempPdf).Length + ' path=' + $tempPdf)
    } else {
      Add-Result 'PDF File.Exists post-export' 'absent/error' ('file NOT created: ' + $tempPdf)
    }
  }
}
finally {
  if ($book) { try { $book.GetType().InvokeMember('Close', $flagsCall, $null, $book, @($false)); Add-Result 'Workbook.Close(false)' 'present' $null } catch { Add-Result 'Workbook.Close(false)' 'absent/error' $_.Exception.Message } }
  if ($app)  { try { $type.InvokeMember('Quit', $flagsCall, $null, $app, $null); Add-Result 'Application.Quit' 'present' $null } catch { Add-Result 'Application.Quit' 'absent/error' $_.Exception.Message } }
  foreach ($o in @($range, $pageSetup, $nameObj, $names, $sheet, $sheets, $book, $books, $app)) {
    if ($o) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch {} }
  }
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

Write-Output ''
Write-Output 'MEMBER PROBE RESULTS'
$results | Format-Table -Wrap -AutoSize | Out-String -Width 220 | Write-Output

Write-Output ''
Write-Output 'ABSENT / ERROR SUMMARY'
$bad = $results | Where-Object { $_.Present -ne 'present' }
if ($bad) { $bad | Format-Table -Wrap -AutoSize | Out-String -Width 220 | Write-Output } else { Write-Output '(none - every probed member responded)' }

}

Write-Section 'STEP 7 - tasklist AFTER probe'
Start-Sleep -Seconds 3
$after = & tasklist 2>&1 | Select-String -Pattern 'et\.exe|wps\.exe|excel\.exe' -CaseSensitive:$false
if ($after) { $after | ForEach-Object { Write-Output $_.Line } } else { Write-Output '(no matching process: et.exe / wps.exe / excel.exe)' }

Write-Section 'STEP 8 - temp artifacts left behind'
Get-ChildItem -LiteralPath $env:TEMP -Filter 'ArcTool_EV1_*' -ErrorAction SilentlyContinue |
  Select-Object FullName, Length, LastWriteTime | Format-Table -AutoSize | Out-String -Width 200 | Write-Output

Write-Output ''
Write-Output 'EV-1 PROBE COMPLETE'
