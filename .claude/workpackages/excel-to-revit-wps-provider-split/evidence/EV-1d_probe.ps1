# EV-1d — narrow the WPS defect to its true origin.
#
# EV-1c produced a result that reframes everything: Application.Workbooks itself
# returned NULL on this build, in both headless and visible rounds. EV-1 / EV-1b
# had reported "Workbooks present" only because they never null-checked the value
# they got back -- InvokeMember simply did not throw.
#
# If Workbooks is null, then DISP_E_TYPEMISMATCH on Workbooks.Open was never a
# signature problem at all: it is what you get when the late binder is handed a
# null/VT_EMPTY target. So the whole T5.1b "try more argument shapes" direction
# was aimed at the wrong layer.
#
# This probe answers three questions:
#   Q1. Which server does KET.Application actually start, and is a type library
#       registered for it? (registry: CLSID -> LocalServer32 -> TypeLib)
#   Q2. Is Workbooks null deterministically, or is it a readiness/race issue?
#       (retry loop, three different binders, timing recorded)
#   Q3. Does attaching to an ALREADY-RUNNING WPS instance behave differently
#       from CreateInstance? (Marshal.GetActiveObject)
#
# Never touches the user's original workbook. No Revit involved.

$ErrorActionPreference = 'Continue'

$evidenceDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$outFile     = Join-Path $evidenceDir 'EV-1d_output.txt'

$lines = New-Object System.Collections.Generic.List[string]
function W([string]$s = '') { [void]$lines.Add($s) }
function Hdr([string]$t) { W ''; W ('=' * 78); W ("== " + $t); W ('=' * 78) }

$PublicInstance = [System.Reflection.BindingFlags]::Public -bor [System.Reflection.BindingFlags]::Instance
$FlagGetProp    = [System.Reflection.BindingFlags]::GetProperty -bor $PublicInstance
$FlagSetProp    = [System.Reflection.BindingFlags]::SetProperty -bor $PublicInstance
$FlagInvoke     = [System.Reflection.BindingFlags]::InvokeMethod -bor $PublicInstance

function RegVal([string]$path) {
    try {
        $k = Get-Item -Path ("Registry::" + $path) -ErrorAction Stop
        return $k.GetValue('')
    } catch { return $null }
}

Hdr 'STEP 0 - metadata'
W ("timestamp_local   : " + (Get-Date).ToString('yyyy-MM-dd HH:mm:ss zzz'))
W ("probe_variant     : EV-1d (origin isolation)")
W ("process_is_64bit  : " + [Environment]::Is64BitProcess)
W ("thread_culture    : " + [System.Threading.Thread]::CurrentThread.CurrentCulture.Name)


W ("powershell_version : " + $PSVersionTable.PSVersion.ToString())
W ("computer          : " + $env:COMPUTERNAME)
W ("user              : " + $env:USERNAME)

Hdr 'STEP 1 - registry / COM registration for KET.Application'
$progId = 'KET.Application'
$clsid = $null
try {
    $clsid = [Type]::GetTypeFromProgID($progId, $false).GUID.ToString('B')
    W ("ProgID            : $progId")
    W ("CLSID via Type    : $clsid")
} catch {
    W ('Type.GetTypeFromProgID failed -> ' + $_.Exception.Message)
}

$regClsid = RegVal ("HKEY_CLASSES_ROOT\\$progId\\CLSID")
W ("CLSID via registry: " + $regClsid)
if (-not $clsid -and $regClsid) { $clsid = $regClsid }

if ($clsid) {
    W ("HKCR\\CLSID\\$clsid\\LocalServer32 = " + (RegVal ("HKEY_CLASSES_ROOT\\CLSID\\$clsid\\LocalServer32")))
    W ("HKCR\\CLSID\\$clsid\\InprocServer32 = " + (RegVal ("HKEY_CLASSES_ROOT\\CLSID\\$clsid\\InprocServer32")))
    $typeLib = RegVal ("HKEY_CLASSES_ROOT\\CLSID\\$clsid\\TypeLib")
    W ("HKCR\\CLSID\\$clsid\\TypeLib = " + $typeLib)
    if ($typeLib) {
        W ("HKCR\\TypeLib\\$typeLib = " + (RegVal ("HKEY_CLASSES_ROOT\\TypeLib\\$typeLib")))
        foreach ($sub in @('1.0','1.1','1.2','1.3','1.4','1.5','1.6','1.7','1.8','1.9','2.0')) {
            $p = "HKEY_CLASSES_ROOT\\TypeLib\\$typeLib\\$sub\\0\\win64"
            $v = RegVal $p
            if ($v) { W ("$p = $v") }
            $p32 = "HKEY_CLASSES_ROOT\\TypeLib\\$typeLib\\$sub\\0\\win32"
            $v32 = RegVal $p32
            if ($v32) { W ("$p32 = $v32") }
        }
    }
}

Hdr 'STEP 2 - active processes before COM activation'
try {
    Get-Process wps,et -ErrorAction SilentlyContinue | Sort-Object ProcessName,Id | ForEach-Object {
        W (("{0,-6} pid={1,-6} mainWindow='{2}' path='{3}'" -f $_.ProcessName, $_.Id, $_.MainWindowTitle, $_.Path))
    }
} catch {
    W ('Get-Process pre-activation failed -> ' + $_.Exception.Message)
}

function TryGetWorkbooksViaReflection($app) {
    try {
        return $app.GetType().InvokeMember('Workbooks', $FlagGetProp, $null, $app, $null)
    } catch {
        throw
    }
}

function TryGetWorkbooksViaPsAdapter($app) {
    return $app.Workbooks
}

function TryGetWorkbooksViaCallByName($app) {
    return [Microsoft.VisualBasic.Interaction]::CallByName($app, 'Workbooks', [Microsoft.VisualBasic.CallType]::Get)
}

Add-Type -AssemblyName Microsoft.VisualBasic

function Probe-App([string]$label, $app, [bool]$ownsApp) {
    Hdr ("STEP 3 - probe app: " + $label)

    if ($null -eq $app) {
        W 'app is null'
        return
    }

    try { W ('app type            : ' + $app.GetType().FullName) } catch { }
    try { W ('Application.Version : ' + $app.GetType().InvokeMember('Version', $FlagGetProp, $null, $app, $null)) } catch { W ('Version get failed -> ' + $_.Exception.Message) }
    try { $app.GetType().InvokeMember('Visible', $FlagSetProp, $null, $app, @($false)) | Out-Null; W 'Visible=false set ok' } catch { W ('Visible=false set failed -> ' + $_.Exception.Message) }
    try { $app.GetType().InvokeMember('DisplayAlerts', $FlagSetProp, $null, $app, @($false)) | Out-Null; W 'DisplayAlerts=false set ok' } catch { W ('DisplayAlerts=false set failed -> ' + $_.Exception.Message) }

    $attempts = @(
        @{ Name = 'reflection once';      Fn = { TryGetWorkbooksViaReflection $app } },
        @{ Name = 'ps-adapter once';      Fn = { TryGetWorkbooksViaPsAdapter $app } },
        @{ Name = 'CallByName once';      Fn = { TryGetWorkbooksViaCallByName $app } }
    )

    foreach ($attempt in $attempts) {
        try {
            $obj = & $attempt.Fn
            if ($null -eq $obj) {
                W (("{0,-20} -> NULL" -f $attempt.Name))
            } else {
                W (("{0,-20} -> {1}" -f $attempt.Name, $obj.GetType().FullName))
                try {
                    $count = $obj.GetType().InvokeMember('Count', $FlagGetProp, $null, $obj, $null)
                    W (("{0,-20} Count = {1}" -f $attempt.Name, $count))
                } catch {
                    W (("{0,-20} Count failed -> {1}" -f $attempt.Name, $_.Exception.Message))
                }
            }
        } catch {
            W (("{0,-20} -> EXCEPTION: {1}" -f $attempt.Name, $_.Exception.Message))
        }
    }

    Hdr ("STEP 4 - readiness retry loop: " + $label)
    for ($i = 1; $i -le 10; $i++) {
        Start-Sleep -Milliseconds 500
        $stamp = Get-Date -Format 'HH:mm:ss.fff'
        try {
            $wb = TryGetWorkbooksViaReflection $app
            if ($null -eq $wb) {
                W ("[$stamp] retry #$i reflection -> NULL")
            } else {
                W ("[$stamp] retry #$i reflection -> " + $wb.GetType().FullName)
                try {
                    $count = $wb.GetType().InvokeMember('Count', $FlagGetProp, $null, $wb, $null)
                    W ("[$stamp] retry #$i Count = $count")
                } catch {
                    W ("[$stamp] retry #$i Count failed -> " + $_.Exception.Message)
                }
                break
            }
        } catch {
            W ("[$stamp] retry #$i reflection -> EXCEPTION: " + $_.Exception.Message)
        }
    }

    Hdr ("STEP 5 - active processes during probe: " + $label)
    try {
        Get-Process wps,et -ErrorAction SilentlyContinue | Sort-Object ProcessName,Id | ForEach-Object {
            W (("{0,-6} pid={1,-6} mainWindow='{2}' path='{3}'" -f $_.ProcessName, $_.Id, $_.MainWindowTitle, $_.Path))
        }
    } catch {
        W ('Get-Process during probe failed -> ' + $_.Exception.Message)
    }

    if ($ownsApp) {
        Hdr ("STEP 6 - cleanup owned app: " + $label)
        try { $app.GetType().InvokeMember('Quit', $FlagInvoke, $null, $app, $null) | Out-Null; W 'Quit() ok' } catch { W ('Quit() failed -> ' + $_.Exception.Message) }
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app); W 'ReleaseComObject(app) ok' } catch { W ('ReleaseComObject(app) failed -> ' + $_.Exception.Message) }
    }
}

Hdr 'STEP 3A - create new COM instance'
$app1 = $null
try {
    $t = [Type]::GetTypeFromProgID($progId, $false)
    if ($t) {
        $app1 = [Activator]::CreateInstance($t)
        W 'CreateInstance succeeded'
    } else {
        W 'Type.GetTypeFromProgID returned null'
    }
} catch {
    W ('CreateInstance failed -> ' + $_.Exception.Message)
}
Probe-App 'CreateInstance(KET.Application)' $app1 $true

Hdr 'STEP 3B - attach to running object table'
$app2 = $null
try {
    $app2 = [Runtime.InteropServices.Marshal]::GetActiveObject($progId)
    W 'Marshal.GetActiveObject succeeded'
} catch {
    W ('Marshal.GetActiveObject failed -> ' + $_.Exception.Message)
}
Probe-App 'GetActiveObject(KET.Application)' $app2 $false

Hdr 'STEP 7 - active processes after cleanup'
try {
    Get-Process wps,et -ErrorAction SilentlyContinue | Sort-Object ProcessName,Id | ForEach-Object {
        W (("{0,-6} pid={1,-6} mainWindow='{2}' path='{3}'" -f $_.ProcessName, $_.Id, $_.MainWindowTitle, $_.Path))
    }
} catch {
    W ('Get-Process post-cleanup failed -> ' + $_.Exception.Message)
}

W ''
W 'EV-1d PROBE COMPLETE'
$lines | Set-Content -LiteralPath $outFile -Encoding UTF8
Write-Host ('EV-1d output written to: ' + $outFile)

