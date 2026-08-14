# EV-1b — WPS late-bound probe, re-run after the T5.1b patch.
#
# Purpose: prove (or disprove) that the ordered Workbooks.Open call-shape ladder introduced by
# T5.1b in ArcTool.Core/Services/Excel/WpsWorkbookPdfExporter.cs actually opens a workbook on this
# WPS build. EV-1 (pre-patch) failed at the single-argument shape with DISP_E_TYPEMISMATCH
# (0x80020005) and therefore never reached any downstream member.
#
# This script mirrors the patched C# ladder exactly:
#   1. Open(filePath)
#   2. Open(filePath, Type.Missing)
#   3. Open(filePath, Type.Missing, Type.Missing)
#   4. Open(filePath, Type.Missing, Type.Missing, Type.Missing, Type.Missing)
#
# If all four fail, an extra DIAGNOSTIC section probes additional shapes that are NOT in the
# patch, purely to give a follow-up patch real data instead of guesses.
#
# Deliberate deviations from the operator runbook (same as EV-1):
#  - version read from et.exe FileVersionInfo instead of the About dialog
#  - operates on a scratch copy so the sample workbook is never modified
#  - non-interactive: test workbook path is hardcoded below
#  - never kills pre-existing wps.exe/et.exe processes; only Quit()s the instance it creates

$ErrorActionPreference = 'Continue'

$sourceWorkbook = 'C:\Users\ADMIN\Desktop\PA4\BULONG.xlsx'

$flagsGet  = [Reflection.BindingFlags]'Public,Instance,GetProperty'
$flagsSet  = [Reflection.BindingFlags]'Public,Instance,SetProperty'
$flagsCall = [Reflection.BindingFlags]'Public,Instance,InvokeMethod'

function Write-Section([string]$title) {
    ''
    '=============================================================================='
    "== $title"
    '=============================================================================='
}

$probe = New-Object System.Collections.ArrayList
function Add-Probe([string]$member, [string]$state, [string]$detail) {
    [void]$probe.Add([pscustomobject]@{ Member = $member; Result = $state; Detail = $detail })
}

# ── STEP 0 ────────────────────────────────────────────────────────────────────────────────
Write-Section 'STEP 0 - probe metadata'
"timestamp_local       : {0}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss zzz')
"machine               : $env:COMPUTERNAME"
"powershell_version    : $($PSVersionTable.PSVersion)"
"process_is_64bit      : $([Environment]::Is64BitProcess)"
"probe_variant         : EV-1b (post T5.1b patch)"
"source_workbook       : $sourceWorkbook"
"source_workbook_exists: $(Test-Path -LiteralPath $sourceWorkbook)"

# ── STEP 1 ────────────────────────────────────────────────────────────────────────────────
Write-Section 'STEP 1 - tasklist BEFORE probe'
tasklist /FI "IMAGENAME eq wps.exe" /FO TABLE /NH
tasklist /FI "IMAGENAME eq et.exe"  /FO TABLE /NH

# ── STEP 2 ────────────────────────────────────────────────────────────────────────────────
Write-Section 'STEP 2 - WPS build (from et.exe FileVersionInfo)'
Get-ChildItem -Path "$env:LOCALAPPDATA\Kingsoft\WPS Office" -Filter 'et.exe' -Recurse -ErrorAction SilentlyContinue |
    ForEach-Object { $_.VersionInfo } |
    Select-Object FileName, FileVersion, ProductVersion |
    Format-List

# ── STEP 3 ────────────────────────────────────────────────────────────────────────────────
Write-Section 'STEP 3 - ProgID resolution (patched candidate order)'
$progIds = @('KET.Application', 'ET.Application', 'Kingsoft.ET.Application')
$resolved = $null
foreach ($id in $progIds) {
    $t = [Type]::GetTypeFromProgID($id, $false)
    "{0,-26} resolved={1}" -f $id, ($null -ne $t)
    if ($null -ne $t -and $null -eq $resolved) { $resolved = @{ ProgId = $id; Type = $t } }
}

if ($null -eq $resolved) {
    'FATAL: no WPS spreadsheet ProgID resolved for this user. EV-1b cannot continue.'
    return
}
"selected_progid       : $($resolved.ProgId)"

# ── STEP 4 ────────────────────────────────────────────────────────────────────────────────
Write-Section 'STEP 4 - scratch copy of the test workbook'
$scratch = Join-Path $env:TEMP ("ArcTool_EV1b_scratch_{0}.xlsx" -f ([guid]::NewGuid().ToString('N')))
Copy-Item -LiteralPath $sourceWorkbook -Destination $scratch -Force
"scratch_copy          : $scratch"
"scratch_copy_exists   : $(Test-Path -LiteralPath $scratch)"

$tempPdf = Join-Path $env:TEMP ("ArcTool_EV1b_export_{0}.pdf" -f ([guid]::NewGuid().ToString('N')))
"planned_export_pdf    : $tempPdf"

# ── STEP 5 ────────────────────────────────────────────────────────────────────────────────
Write-Section 'STEP 5 - activate the WPS application'
$app = $null
try {
    $app = [Activator]::CreateInstance($resolved.Type)
    Add-Probe 'CreateInstance' 'present' $app.GetType().FullName
} catch {
    Add-Probe 'CreateInstance' 'absent/error' $_.Exception.Message
}

$appType = if ($null -ne $app) { $app.GetType() } else { $null }

if ($null -ne $app) {
    try { $appType.InvokeMember('Visible',       $flagsSet, $null, $app, @($false)); Add-Probe 'Visible(set)'       'present' '' }
    catch { Add-Probe 'Visible(set)'       'absent/error' $_.Exception.Message }

    try { $appType.InvokeMember('DisplayAlerts', $flagsSet, $null, $app, @($false)); Add-Probe 'DisplayAlerts(set)' 'present' '' }
    catch { Add-Probe 'DisplayAlerts(set)' 'absent/error' $_.Exception.Message }

    try {
        $v = $appType.InvokeMember('Version', $flagsGet, $null, $app, @())
        Add-Probe 'Application.Version(get)' 'present' "$v"
    } catch { Add-Probe 'Application.Version(get)' 'absent/error' $_.Exception.Message }
}

# ── STEP 6 ────────────────────────────────────────────────────────────────────────────────
Write-Section 'STEP 6 - Workbooks.Open call-shape ladder (mirrors the T5.1b patch)'
$books    = $null
$bookType = $null
$book     = $null
$winningShape = $null

if ($null -ne $app) {
    try {
        $books = $appType.InvokeMember('Workbooks', $flagsGet, $null, $app, @())
        $bookType = $books.GetType()
        Add-Probe 'Workbooks' 'present' ''
    } catch {
        Add-Probe 'Workbooks' 'absent/error' $_.Exception.Message
    }
}

function Invoke-OpenShape([string]$label, [object[]]$callArgs) {
    if ($null -eq $script:books -or $null -ne $script:book) { return }
    try {
        $result = $script:bookType.InvokeMember('Open', $flagsCall, $null, $script:books, $callArgs)
        if ($null -ne $result) {
            $script:book = $result
            $script:winningShape = $label
            "  $label -> SUCCESS"
            Add-Probe "Workbooks.Open [$label]" 'present' 'opened workbook'
        } else {
            "  $label -> returned null"
            Add-Probe "Workbooks.Open [$label]" 'absent/error' 'returned null'
        }
    } catch {
        "  $label -> FAILED: $($_.Exception.Message)"
        Add-Probe "Workbooks.Open [$label]" 'absent/error' $_.Exception.Message
    }
}

# Build the same argument arrays the patched CreateWorkbookOpenArgs() produces.
function New-OpenArgs([int]$missingCount) {
    $a = New-Object object[] ($missingCount + 1)
    $a[0] = $scratch
    for ($i = 1; $i -lt $a.Length; $i++) { $a[$i] = [Type]::Missing }
    return $a
}

Invoke-OpenShape 'attempt#1 1-arg'                 @($scratch)
Invoke-OpenShape 'attempt#2 path + 1x Missing'     (New-OpenArgs 1)
Invoke-OpenShape 'attempt#3 path + 2x Missing'     (New-OpenArgs 2)
Invoke-OpenShape 'attempt#4 path + 4x Missing'     (New-OpenArgs 4)

if ($null -ne $winningShape) {
    "winning_shape         : $winningShape"
} else {
    'winning_shape         : NONE - every patched shape failed'
}

# ── STEP 6b — DIAGNOSTIC only, shapes NOT in the current patch ────────────────────────────
if ($null -eq $book -and $null -ne $books) {
    Write-Section 'STEP 6b - DIAGNOSTIC: extra shapes not present in the patch'
    'These exist only to inform a follow-up patch. A success here means the patched ladder is'
    'incomplete and needs the winning shape added.'
    ''

    function New-NullArgs([int]$nullCount) {
        $a = New-Object object[] ($nullCount + 1)
        $a[0] = $scratch
        for ($i = 1; $i -lt $a.Length; $i++) { $a[$i] = $null }
        return $a
    }
    function New-MissingValueArgs([int]$count) {
        $a = New-Object object[] ($count + 1)
        $a[0] = $scratch
        for ($i = 1; $i -lt $a.Length; $i++) { $a[$i] = [System.Reflection.Missing]::Value }
        return $a
    }

    Invoke-OpenShape 'diag full 15-arg Missing'      (New-OpenArgs 14)
    Invoke-OpenShape 'diag path + 1x $null'          (New-NullArgs 1)
    Invoke-OpenShape 'diag path + 2x $null'          (New-NullArgs 2)
    Invoke-OpenShape 'diag Missing.Value x1'         (New-MissingValueArgs 1)
    Invoke-OpenShape 'diag path + UpdateLinks=0'     @($scratch, 0)
    Invoke-OpenShape 'diag path + 0 + ReadOnly=$false' @($scratch, 0, $false)

    if ($null -ne $winningShape) {
        "diagnostic_winning_shape : $winningShape (NOT yet in the patch)"
    }
}

# ── STEP 7 ────────────────────────────────────────────────────────────────────────────────
Write-Section 'STEP 7 - downstream members (only reachable once Open succeeds)'
$firstSheetName = $null

if ($null -eq $book) {
    'SKIPPED - no workbook handle. Everything below stays UNVERIFIED, exactly as in EV-1.'
} else {
    $wbType = $book.GetType()

    # Worksheets
    try {
        $sheets = $wbType.InvokeMember('Worksheets', $flagsGet, $null, $book, @())
        Add-Probe 'Worksheets' 'present' ''
        $sType = $sheets.GetType()
        $count = $sType.InvokeMember('Count', $flagsGet, $null, $sheets, @())
        Add-Probe 'Worksheets.Count' 'present' "$count"
        for ($i = 1; $i -le [int]$count; $i++) {
            $ws = $sType.InvokeMember('Item', $flagsGet, $null, $sheets, @($i))
            $n  = $ws.GetType().InvokeMember('Name', $flagsGet, $null, $ws, @())
            if ($null -eq $firstSheetName) { $firstSheetName = "$n" }
            "  sheet #$i = $n"
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($ws)
        }
        Add-Probe 'Worksheets.Item/Name' 'present' "first='$firstSheetName'"
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($sheets)
    } catch {
        Add-Probe 'Worksheets walk' 'absent/error' $_.Exception.Message
    }

    # Names
    try {
        $names = $wbType.InvokeMember('Names', $flagsGet, $null, $book, @())
        $nCount = $names.GetType().InvokeMember('Count', $flagsGet, $null, $names, @())
        Add-Probe 'Names.Count' 'present' "$nCount"
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($names)
    } catch {
        Add-Probe 'Names.Count' 'absent/error' $_.Exception.Message
    }

    # PageSetup / Range / Address / export
    if ($null -ne $firstSheetName) {
        try {
            $sheets = $wbType.InvokeMember('Worksheets', $flagsGet, $null, $book, @())
            $sheet  = $sheets.GetType().InvokeMember('Item', $flagsGet, $null, $sheets, @($firstSheetName))
            $shType = $sheet.GetType()

            $ps = $null
            try {
                $ps = $shType.InvokeMember('PageSetup', $flagsGet, $null, $sheet, @())
                Add-Probe 'PageSetup' 'present' ''
            } catch { Add-Probe 'PageSetup' 'absent/error' $_.Exception.Message }

            $range = $null
            try {
                $range = $shType.InvokeMember('UsedRange', $flagsGet, $null, $sheet, @())
                Add-Probe 'UsedRange' 'present' ''
            } catch { Add-Probe 'UsedRange' 'absent/error' $_.Exception.Message }

            if ($null -ne $range) {
                $rType = $range.GetType()
                try {
                    $addr = $rType.InvokeMember('Address', $flagsGet, $null, $range, @($false, $false))
                    Add-Probe 'Range.Address(false,false)' 'present' "$addr"
                } catch {
                    Add-Probe 'Range.Address(false,false)' 'absent/error' $_.Exception.Message
                    try {
                        $addr = $rType.InvokeMember('Address', $flagsGet, $null, $range, @())
                        Add-Probe 'Range.Address(no-arg)' 'present' "$addr"
                    } catch { Add-Probe 'Range.Address(no-arg)' 'absent/error' $_.Exception.Message }
                }
            }

            if ($null -ne $ps) {
                $psType = $ps.GetType()
                foreach ($pair in @(
                    @{ n = 'Zoom';           v = $false },
                    @{ n = 'FitToPagesWide'; v = 1 },
                    @{ n = 'FitToPagesTall'; v = 1 },
                    @{ n = 'TopMargin';      v = 0.0 },
                    @{ n = 'BottomMargin';   v = 0.0 },
                    @{ n = 'LeftMargin';     v = 0.0 },
                    @{ n = 'RightMargin';    v = 0.0 }
                )) {
                    try {
                        $psType.InvokeMember($pair.n, $flagsSet, $null, $ps, @($pair.v))
                        Add-Probe "PageSetup.$($pair.n)(set)" 'present' ''
                    } catch { Add-Probe "PageSetup.$($pair.n)(set)" 'absent/error' $_.Exception.Message }
                }

                # PaperSize: 26 (Esheet) then 8 (A3) — the patched tolerable-by-design fallback.
                try {
                    $psType.InvokeMember('PaperSize', $flagsSet, $null, $ps, @(26))
                    Add-Probe 'PageSetup.PaperSize=26' 'present' ''
                } catch {
                    Add-Probe 'PageSetup.PaperSize=26' 'absent/error' $_.Exception.Message
                    try {
                        $psType.InvokeMember('PaperSize', $flagsSet, $null, $ps, @(8))
                        Add-Probe 'PageSetup.PaperSize=8' 'present' ''
                    } catch { Add-Probe 'PageSetup.PaperSize=8' 'absent/error' $_.Exception.Message }
                }
            }

            # ExportAsFixedFormat — the one member that is fatal if absent.
            try {
                $shType.InvokeMember('ExportAsFixedFormat', $flagsCall, $null, $sheet,
                    @(0, $tempPdf, 0, $false, $false, 1, 1, $false))
                Add-Probe 'ExportAsFixedFormat' 'present' 'call returned without throwing'
            } catch {
                Add-Probe 'ExportAsFixedFormat' 'absent/error' $_.Exception.Message
            }

            Start-Sleep -Seconds 2
            $pdfExists = Test-Path -LiteralPath $tempPdf
            $pdfSize   = if ($pdfExists) { (Get-Item -LiteralPath $tempPdf).Length } else { 0 }
            Add-Probe 'exported PDF on disk' $(if ($pdfExists) { 'present' } else { 'absent/error' }) "exists=$pdfExists size=$pdfSize"

            if ($null -ne $range) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($range) }
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($sheet)
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($sheets)
        } catch {
            Add-Probe 'export stage' 'absent/error' $_.Exception.Message
        }
    }
}

# ── STEP 8 ────────────────────────────────────────────────────────────────────────────────
Write-Section 'STEP 8 - cleanup (workbook -> app, child before parent)'
if ($null -ne $book) {
    try { $book.GetType().InvokeMember('Close', $flagsCall, $null, $book, @($false)); 'workbook Close(false) ok' }
    catch { "workbook Close failed: $($_.Exception.Message)" }
    try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($book) } catch { }
}
if ($null -ne $books) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($books) } catch { } }
if ($null -ne $app) {
    try { $appType.InvokeMember('Quit', $flagsCall, $null, $app, @()); 'app Quit() ok'; Add-Probe 'Application.Quit' 'present' '' }
    catch { "app Quit failed: $($_.Exception.Message)"; Add-Probe 'Application.Quit' 'absent/error' $_.Exception.Message }
    try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app) } catch { }
}
[GC]::Collect(); [GC]::WaitForPendingFinalizers()

# ── STEP 9 ────────────────────────────────────────────────────────────────────────────────
Write-Section 'STEP 9 - MEMBER PROBE RESULTS'
$probe | Format-Table -AutoSize -Wrap

Write-Section 'STEP 9b - ABSENT / ERROR SUMMARY'
$bad = $probe | Where-Object { $_.Result -ne 'present' }
if ($bad) { $bad | Format-Table -AutoSize -Wrap } else { 'none - every probed member responded' }

# ── STEP 10 ───────────────────────────────────────────────────────────────────────────────
Write-Section 'STEP 10 - tasklist AFTER probe'
tasklist /FI "IMAGENAME eq wps.exe" /FO TABLE /NH
tasklist /FI "IMAGENAME eq et.exe"  /FO TABLE /NH

Write-Section 'STEP 11 - temp artifacts left behind'
Get-ChildItem -Path $env:TEMP -Filter 'ArcTool_EV1b_*' -ErrorAction SilentlyContinue |
    Select-Object FullName, Length, LastWriteTime | Format-Table -AutoSize

''
"EV-1b VERDICT: workbook_open_succeeded = $($null -ne $book); winning_shape = $winningShape"
'EV-1b PROBE COMPLETE'
