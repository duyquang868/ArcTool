# EV-1c — WPS Workbooks.Open root-cause diagnostic
#
# EV-1 and EV-1b both failed at Workbooks.Open with DISP_E_TYPEMISMATCH (0x80020005),
# across every positional / Type.Missing argument shape. That rules out "wrong number of
# optional arguments" as the root cause.
#
# This probe stops guessing shapes and instead asks the server what it actually exposes:
#   1. Read the real ITypeInfo / FUNCDESC for Workbooks.Open (param count + VT types).
#   2. Compare several DIFFERENT late-binding binders on the same object:
#        - raw Type.InvokeMember (what the C# provider does today)
#        - PowerShell's own IDispatch adapter
#        - Microsoft.VisualBasic Interaction.CallByName (Office-friendly late binder)
#        - named-argument InvokeMember
#        - InvokeMethod | GetProperty flag combination
#        - forced en-US culture
#   3. Check whether Workbooks.Add works while Workbooks.Open fails (isolates string
#      marshaling from collection identity).
#   4. Retry with Visible = $true, in case WPS gates document automation on visibility.
#
# Read-only against the user's data: always operates on a scratch copy, never the original.
# No Revit involved.

$ErrorActionPreference = 'Continue'

$evidenceDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$outFile     = Join-Path $evidenceDir 'EV-1c_output.txt'
$sourceBook  = 'C:\Users\ADMIN\Desktop\PA4\BULONG.xlsx'

$lines = New-Object System.Collections.Generic.List[string]
function W([string]$s = '') { [void]$lines.Add($s) }
function Hdr([string]$t) { W ''; W ('=' * 78); W ("== " + $t); W ('=' * 78) }

$results = New-Object System.Collections.Generic.List[object]
function Rec([string]$name, [bool]$ok, [string]$detail) {
    $verdict = 'FAIL'
    if ($ok) { $verdict = 'OK' }
    [void]$results.Add([pscustomobject]@{ Test = $name; Result = $verdict; Detail = $detail })
}

# ---------------------------------------------------------------- late-bind helpers
$PublicInstance = [System.Reflection.BindingFlags]::Public -bor [System.Reflection.BindingFlags]::Instance
$FlagGetProp    = [System.Reflection.BindingFlags]::GetProperty -bor $PublicInstance
$FlagSetProp    = [System.Reflection.BindingFlags]::SetProperty -bor $PublicInstance
$FlagInvoke     = [System.Reflection.BindingFlags]::InvokeMethod -bor $PublicInstance
$FlagInvokeGet  = [System.Reflection.BindingFlags]::InvokeMethod -bor [System.Reflection.BindingFlags]::GetProperty -bor $PublicInstance

function GetProp($target, [string]$name) {
    $target.GetType().InvokeMember($name, $FlagGetProp, $null, $target, $null)
}
function SetProp($target, [string]$name, $value) {
    [void]$target.GetType().InvokeMember($name, $FlagSetProp, $null, $target, @($value))
}
function CallMethod($target, [string]$name, $argv) {
    $target.GetType().InvokeMember($name, $FlagInvoke, $null, $target, $argv)
}

# ---------------------------------------------------------------- ITypeInfo dumper
$dumperSource = @'
using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

public static class ComTypeInfoDumper
{
    [ComImport, Guid("00020400-0000-0000-C000-000000000046"),
     InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IDispatchLite
    {
        void GetTypeInfoCount(out uint pctinfo);
        void GetTypeInfo(uint iTInfo, int lcid, out ITypeInfo ppTInfo);
    }

    public static string Dump(object comObject, string memberFilter)
    {
        var sb = new System.Text.StringBuilder();

        IDispatchLite disp = comObject as IDispatchLite;
        if (disp == null)
            return "object could not be cast to IDispatch";

        uint count = 0;
        try { disp.GetTypeInfoCount(out count); }
        catch (Exception ex) { return "GetTypeInfoCount threw: " + ex.Message; }
        sb.AppendLine("GetTypeInfoCount = " + count);
        if (count == 0)
            return sb.ToString() + "no ITypeInfo exposed -> server is IDispatch-by-name only";

        ITypeInfo ti = null;
        try { disp.GetTypeInfo(0, 0, out ti); }
        catch (Exception ex) { return sb.ToString() + "GetTypeInfo threw: " + ex.Message; }
        if (ti == null)
            return sb.ToString() + "GetTypeInfo returned null";

        try
        {
            string n0, d0, h0; int c0;
            ti.GetDocumentation(-1, out n0, out d0, out c0, out h0);
            sb.AppendLine("interface name  = " + n0);
            sb.AppendLine("interface doc   = " + d0);
        }
        catch (Exception ex) { sb.AppendLine("GetDocumentation(-1) threw: " + ex.Message); }

        int cFuncs = 0;
        IntPtr pAttr = IntPtr.Zero;
        try
        {
            ti.GetTypeAttr(out pAttr);
            TYPEATTR attr = (TYPEATTR)Marshal.PtrToStructure(pAttr, typeof(TYPEATTR));
            sb.AppendLine("type guid       = " + attr.guid);
            sb.AppendLine("cFuncs          = " + attr.cFuncs);
            sb.AppendLine("cVars           = " + attr.cVars);
            cFuncs = attr.cFuncs;
        }
        catch (Exception ex) { sb.AppendLine("GetTypeAttr threw: " + ex.Message); }
        finally { if (pAttr != IntPtr.Zero) ti.ReleaseTypeAttr(pAttr); }

        int elemSize = Marshal.SizeOf(typeof(ELEMDESC));

        for (int i = 0; i < cFuncs; i++)
        {
            IntPtr pFunc = IntPtr.Zero;
            try
            {
                ti.GetFuncDesc(i, out pFunc);
                FUNCDESC fd = (FUNCDESC)Marshal.PtrToStructure(pFunc, typeof(FUNCDESC));

                string name = "?", doc = "", helpFile = ""; int ctx = 0;
                try { ti.GetDocumentation(fd.memid, out name, out doc, out ctx, out helpFile); }
                catch { }

                if (!string.IsNullOrEmpty(memberFilter) &&
                    string.Compare(name, memberFilter, StringComparison.OrdinalIgnoreCase) != 0)
                    continue;

                sb.AppendLine("---- " + name);
                sb.AppendLine("     memid       = 0x" + fd.memid.ToString("X8"));
                sb.AppendLine("     invkind     = " + fd.invkind);
                sb.AppendLine("     callconv    = " + fd.callconv);
                sb.AppendLine("     cParams     = " + fd.cParams);
                sb.AppendLine("     cParamsOpt  = " + fd.cParamsOpt);
                sb.AppendLine("     return vt   = " + fd.elemdescFunc.tdesc.vt);

                for (int p = 0; p < fd.cParams; p++)
                {
                    IntPtr pElem = new IntPtr(fd.lprgelemdescParam.ToInt64() + (long)(p * elemSize));
                    ELEMDESC ed = (ELEMDESC)Marshal.PtrToStructure(pElem, typeof(ELEMDESC));
                    sb.AppendLine("     param[" + p + "]   vt=" + ed.tdesc.vt +
                                  "  flags=" + ed.desc.paramdesc.wParamFlags);
                }
            }
            catch (Exception ex) { sb.AppendLine("func[" + i + "] inspection threw: " + ex.Message); }
            finally { if (pFunc != IntPtr.Zero) ti.ReleaseFuncDesc(pFunc); }
        }

        return sb.ToString();
    }
}
'@

Hdr 'STEP 0 - probe metadata'
W ("timestamp_local        : " + (Get-Date).ToString('yyyy-MM-dd HH:mm:ss zzz'))
W ("machine                : " + $env:COMPUTERNAME)
W ("powershell_version     : " + $PSVersionTable.PSVersion.ToString())
W ("process_is_64bit       : " + [Environment]::Is64BitProcess)
W ("probe_variant          : EV-1c (root-cause diagnostic)")
W ("thread_culture         : " + [System.Threading.Thread]::CurrentThread.CurrentCulture.Name)
W ("thread_ui_culture      : " + [System.Threading.Thread]::CurrentThread.CurrentUICulture.Name)
W ("os_ui_culture          : " + (Get-Culture).Name)
W ("source_workbook        : " + $sourceBook)
W ("source_workbook_exists : " + (Test-Path -LiteralPath $sourceBook))

$dumperLoaded = $false
try {
    Add-Type -TypeDefinition $dumperSource -ErrorAction Stop
    $dumperLoaded = $true
    W 'ComTypeInfoDumper       : compiled OK'
} catch {
    W ('ComTypeInfoDumper       : COMPILE FAILED -> ' + $_.Exception.Message)
}

$vbLoaded = $false
try {
    Add-Type -AssemblyName Microsoft.VisualBasic -ErrorAction Stop
    $vbLoaded = $true
    W 'Microsoft.VisualBasic   : loaded OK'
} catch {
    W ('Microsoft.VisualBasic   : LOAD FAILED -> ' + $_.Exception.Message)
}

# ---------------------------------------------------------------- scratch copy
Hdr 'STEP 1 - scratch copy of the test workbook'
$scratch = Join-Path $env:TEMP ('ArcTool_EV1c_scratch_' + ([guid]::NewGuid().ToString('N')) + '.xlsx')
$scratchFwd = $scratch.Replace('\', '/')
try {
    Copy-Item -LiteralPath $sourceBook -Destination $scratch -Force -ErrorAction Stop
    W ('scratch_copy       : ' + $scratch)
    W ('scratch_exists     : ' + (Test-Path -LiteralPath $scratch))
    W ('scratch_fwd_slash  : ' + $scratchFwd)
} catch {
    W ('scratch copy FAILED -> ' + $_.Exception.Message)
}

# ---------------------------------------------------------------- reusable round
function Invoke-OpenMatrix {
    param(
        [bool]$MakeVisible,
        [string]$RoundLabel
    )

    Hdr ("ROUND " + $RoundLabel + " - Visible = " + $MakeVisible)

    $t = $null
    try { $t = [Type]::GetTypeFromProgID('KET.Application', $false) } catch { }
    if ($null -eq $t) {
        W 'KET.Application did not resolve; round aborted.'
        Rec ("round " + $RoundLabel + " : ProgID KET.Application") $false 'not resolved'
        return
    }
    W 'KET.Application resolved.'

    $app = $null
    try {
        $app = [Activator]::CreateInstance($t)
        W ('CreateInstance ok  : ' + $app.GetType().FullName)
        Rec ("round " + $RoundLabel + " : CreateInstance") $true $app.GetType().FullName
    } catch {
        W ('CreateInstance FAILED -> ' + $_.Exception.Message)
        Rec ("round " + $RoundLabel + " : CreateInstance") $false $_.Exception.Message
        return
    }

    try { SetProp $app 'Visible' $MakeVisible; W ("Visible = " + $MakeVisible + " set ok") }
    catch { W ('Visible set FAILED -> ' + $_.Exception.Message) }

    try { SetProp $app 'DisplayAlerts' $false; W 'DisplayAlerts = false set ok' }
    catch { W ('DisplayAlerts set FAILED -> ' + $_.Exception.Message) }

    try { W ('Application.Version = ' + (GetProp $app 'Version')) }
    catch { W ('Version get FAILED -> ' + $_.Exception.Message) }

    # --- the Workbooks collection itself
    $workbooks = $null
    try {
        $workbooks = GetProp $app 'Workbooks'
        if ($null -eq $workbooks) {
            W 'Workbooks get returned null'
            Rec ("round " + $RoundLabel + " : GetProperty Workbooks") $false 'returned null'
            try { CallMethod $app 'Quit' $null } catch { }
            return
        }
        W ('Workbooks obtained : ' + $workbooks.GetType().FullName)
        Rec ("round " + $RoundLabel + " : GetProperty Workbooks") $true $workbooks.GetType().FullName
    } catch {
        W ('Workbooks get FAILED -> ' + $_.Exception.Message)
        Rec ("round " + $RoundLabel + " : GetProperty Workbooks") $false $_.Exception.Message
        try { CallMethod $app 'Quit' $null } catch { }
        return
    }

    # --- is it really a live collection?
    try {
        $cnt = GetProp $workbooks 'Count'
        W ('Workbooks.Count    = ' + $cnt)
        Rec ("round " + $RoundLabel + " : Workbooks.Count") $true ("" + $cnt)
    } catch {
        W ('Workbooks.Count FAILED -> ' + $_.Exception.Message)
        Rec ("round " + $RoundLabel + " : Workbooks.Count") $false $_.Exception.Message
    }

    # --- type library introspection (the definitive answer)
    if ($dumperLoaded) {
        Hdr ("ROUND " + $RoundLabel + " - ITypeInfo dump: Application")
        try { W ([ComTypeInfoDumper]::Dump($app, 'Workbooks')) }
        catch { W ('dump threw -> ' + $_.Exception.Message) }

        Hdr ("ROUND " + $RoundLabel + " - ITypeInfo dump: Workbooks collection (ALL members)")
        try { W ([ComTypeInfoDumper]::Dump($workbooks, '')) }
        catch { W ('dump threw -> ' + $_.Exception.Message) }
    }

    Hdr ("ROUND " + $RoundLabel + " - binder comparison for Workbooks.Open")

    $winner = $null
    $winnerName = ''

    function Try-Open([string]$label, [scriptblock]$action) {
        if ($null -ne $script:winner) { W ("  " + $label + " -> skipped (already open)"); return }
        try {
            $wb = & $action
            if ($null -ne $wb) {
                W ("  " + $label + " -> SUCCESS (" + $wb.GetType().FullName + ")")
                Rec ("open :: " + $label) $true 'workbook handle returned'
                $script:winner = $wb
                $script:winnerName = $label
            } else {
                W ("  " + $label + " -> returned null")
                Rec ("open :: " + $label) $false 'returned null'
            }
        } catch {
            $msg = $_.Exception.Message
            if ($null -ne $_.Exception.InnerException) { $msg = $msg + ' || inner: ' + $_.Exception.InnerException.Message }
            W ("  " + $label + " -> FAILED: " + $msg)
            Rec ("open :: " + $label) $false $msg
        }
    }

    $script:winner = $null
    $script:winnerName = ''

    # 1. exactly what the C# provider does today (baseline, must reproduce the defect)
    Try-Open ("A. Type.InvokeMember InvokeMethod, 1 positional arg") {
        $workbooks.GetType().InvokeMember('Open', $FlagInvoke, $null, $workbooks, @($scratch))
    }

    # 2. PowerShell's own IDispatch adapter -- a genuinely different binder
    Try-Open ("B. PowerShell COM adapter  \$workbooks.Open(path)") {
        $workbooks.Open($scratch)
    }

    # 3. same, reached through the app in one expression
    Try-Open ("C. PowerShell COM adapter  \$app.Workbooks.Open(path)") {
        $app.Workbooks.Open($scratch)
    }

    # 4. VB late binder (Office-tolerant, handles optional/ByRef quirks)
    if ($vbLoaded) {
        Try-Open ("D. VisualBasic Interaction.CallByName") {
            [Microsoft.VisualBasic.Interaction]::CallByName(
                $workbooks, 'Open', [Microsoft.VisualBasic.CallType]::Method, $scratch)
        }
    }

    # 5. named argument instead of positional
    Try-Open ("E. InvokeMember with named parameter 'Filename'") {
        $workbooks.GetType().InvokeMember('Open', $FlagInvoke, $null, $workbooks,
            @($scratch), $null, $null, @('Filename'))
    }

    # 6. DISPATCH_METHOD | DISPATCH_PROPERTYGET
    Try-Open ("F. InvokeMember InvokeMethod|GetProperty") {
        $workbooks.GetType().InvokeMember('Open', $FlagInvokeGet, $null, $workbooks, @($scratch))
    }

    # 7. forced en-US culture (LCID 1033) on the dispatch call
    Try-Open ("G. InvokeMember with CultureInfo en-US") {
        $workbooks.GetType().InvokeMember('Open', $FlagInvoke, $null, $workbooks,
            @($scratch), $null, (New-Object System.Globalization.CultureInfo('en-US')), $null)
    }

    # 8. invariant culture
    Try-Open ("H. InvokeMember with InvariantCulture") {
        $workbooks.GetType().InvokeMember('Open', $FlagInvoke, $null, $workbooks,
            @($scratch), $null, [System.Globalization.CultureInfo]::InvariantCulture, $null)
    }

    # 9. forward-slash path (some Kingsoft builds normalise differently)
    Try-Open ("I. PowerShell COM adapter, forward-slash path") {
        $workbooks.Open($scratchFwd)
    }

    # 10. explicit object[] with boxed string, InvokeMethod
    Try-Open ("J. InvokeMember, boxed object[] single element") {
        $a = New-Object 'object[]' 1
        $a[0] = [string]$scratch
        $workbooks.GetType().InvokeMember('Open', $FlagInvoke, $null, $workbooks, $a)
    }

    # --- isolate string marshaling from collection identity
    Hdr ("ROUND " + $RoundLabel + " - control tests (do OTHER Workbooks members work?)")

    try {
        $newBook = CallMethod $workbooks 'Add' $null
        if ($null -ne $newBook) {
            W '  Workbooks.Add() via InvokeMember -> SUCCESS'
            Rec ("control :: Workbooks.Add via InvokeMember") $true 'new workbook created'
            try { CallMethod $newBook 'Close' @($false) } catch { }
            try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($newBook) } catch { }
        } else {
            W '  Workbooks.Add() via InvokeMember -> returned null'
            Rec ("control :: Workbooks.Add via InvokeMember") $false 'returned null'
        }
    } catch {
        W ('  Workbooks.Add() via InvokeMember -> FAILED: ' + $_.Exception.Message)
        Rec ("control :: Workbooks.Add via InvokeMember") $false $_.Exception.Message
    }

    try {
        $newBook2 = $workbooks.Add()
        if ($null -ne $newBook2) {
            W '  Workbooks.Add() via PS adapter -> SUCCESS'
            Rec ("control :: Workbooks.Add via PS adapter") $true 'new workbook created'
            try { $newBook2.Close($false) } catch { }
            try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($newBook2) } catch { }
        } else {
            W '  Workbooks.Add() via PS adapter -> returned null'
            Rec ("control :: Workbooks.Add via PS adapter") $false 'returned null'
        }
    } catch {
        W ('  Workbooks.Add() via PS adapter -> FAILED: ' + $_.Exception.Message)
        Rec ("control :: Workbooks.Add via PS adapter") $false $_.Exception.Message
    }

    # a string-taking member that is NOT Open, to test BSTR marshaling in general
    try {
        $r = CallMethod $app 'Evaluate' @('1+1')
        W ('  Application.Evaluate("1+1") via InvokeMember -> ' + $r)
        Rec ("control :: Application.Evaluate string arg") $true ("" + $r)
    } catch {
        W ('  Application.Evaluate("1+1") via InvokeMember -> FAILED: ' + $_.Exception.Message)
        Rec ("control :: Application.Evaluate string arg") $false $_.Exception.Message
    }

    # --- if something opened, walk the downstream members EV-1 never reached
    if ($null -ne $script:winner) {
        Hdr ("ROUND " + $RoundLabel + " - downstream member walk (winner: " + $script:winnerName + ")")
        $wb = $script:winner

        try {
            $sheets = GetProp $wb 'Worksheets'
            $sc = GetProp $sheets 'Count'
            W ('  Worksheets.Count = ' + $sc)
            Rec 'downstream :: Worksheets.Count' $true ("" + $sc)

            $ws = $sheets.Item(1)
            $wsName = GetProp $ws 'Name'
            W ('  Worksheets(1).Name = ' + $wsName)
            Rec 'downstream :: Worksheets.Item(1).Name' $true $wsName

            try {
                $used = GetProp $ws 'UsedRange'
                $addr = $null
                try { $addr = $used.GetType().InvokeMember('Address', $FlagInvokeGet, $null, $used, @($false, $false)) }
                catch { $addr = $used.GetType().InvokeMember('Address', $FlagInvokeGet, $null, $used, $null) }
                W ('  UsedRange.Address = ' + $addr)
                Rec 'downstream :: UsedRange.Address' $true ("" + $addr)
            } catch {
                W ('  UsedRange FAILED -> ' + $_.Exception.Message)
                Rec 'downstream :: UsedRange.Address' $false $_.Exception.Message
            }

            try {
                $names = GetProp $wb 'Names'
                $nc = GetProp $names 'Count'
                W ('  Names.Count = ' + $nc)
                Rec 'downstream :: Names.Count' $true ("" + $nc)
            } catch {
                W ('  Names FAILED -> ' + $_.Exception.Message)
                Rec 'downstream :: Names.Count' $false $_.Exception.Message
            }

            try {
                $ps = GetProp $ws 'PageSetup'
                try { SetProp $ps 'Zoom' $false } catch { W ('    PageSetup.Zoom set failed: ' + $_.Exception.Message) }
                try { SetProp $ps 'FitToPagesWide' 1 } catch { W ('    FitToPagesWide set failed: ' + $_.Exception.Message) }
                try { SetProp $ps 'FitToPagesTall' 1 } catch { W ('    FitToPagesTall set failed: ' + $_.Exception.Message) }
                $paperOk = $false
                try { SetProp $ps 'PaperSize' 26; $paperOk = $true; W '    PaperSize = 26 (Esheet) accepted' }
                catch {
                    try { SetProp $ps 'PaperSize' 8; $paperOk = $true; W '    PaperSize = 8 (A3) accepted' }
                    catch { W ('    PaperSize rejected for both 26 and 8: ' + $_.Exception.Message) }
                }
                Rec 'downstream :: PageSetup normalisation' $true ("paperSizeAccepted=" + $paperOk)
            } catch {
                W ('  PageSetup FAILED -> ' + $_.Exception.Message)
                Rec 'downstream :: PageSetup normalisation' $false $_.Exception.Message
            }

            try {
                $pdf = Join-Path $env:TEMP ('ArcTool_EV1c_export_' + ([guid]::NewGuid().ToString('N')) + '.pdf')
                [void]$ws.GetType().InvokeMember('ExportAsFixedFormat', $FlagInvoke, $null, $ws,
                    @(0, $pdf, 0, $false, $false, 1, 1, $false))
                $exists = Test-Path -LiteralPath $pdf
                W ('  ExportAsFixedFormat returned; pdf exists = ' + $exists + ' -> ' + $pdf)
                Rec 'downstream :: ExportAsFixedFormat 8-arg' $exists ("pdf=" + $pdf)
            } catch {
                W ('  ExportAsFixedFormat FAILED -> ' + $_.Exception.Message)
                Rec 'downstream :: ExportAsFixedFormat 8-arg' $false $_.Exception.Message
            }
        } catch {
            W ('  downstream walk aborted -> ' + $_.Exception.Message)
        }

        try { CallMethod $wb 'Close' @($false) } catch { }
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wb) } catch { }
    }

    # --- cleanup, child before parent
    try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($workbooks) } catch { }
    try { CallMethod $app 'Quit' $null; W 'app Quit() ok' } catch { W ('app Quit FAILED -> ' + $_.Exception.Message) }
    try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app) } catch { }

    W ''
    if ($null -ne $script:winnerName -and $script:winnerName -ne '') {
        W ('ROUND ' + $RoundLabel + ' WINNER: ' + $script:winnerName)
    } else {
        W ('ROUND ' + $RoundLabel + ' WINNER: none')
    }
}

Invoke-OpenMatrix -MakeVisible $false -RoundLabel '1 (headless)'
Invoke-OpenMatrix -MakeVisible $true  -RoundLabel '2 (visible)'

# ---------------------------------------------------------------- summary
Hdr 'STEP 9 - full result matrix'
W (($results | Format-Table -AutoSize -Wrap | Out-String))

Hdr 'STEP 10 - successes only'
$ok = $results | Where-Object { $_.Result -eq 'OK' }
if ($ok) { W (($ok | Format-Table -AutoSize -Wrap | Out-String)) } else { W 'none' }

Hdr 'STEP 11 - cleanup'
try {
    if (Test-Path -LiteralPath $scratch) { Remove-Item -LiteralPath $scratch -Force -ErrorAction Stop; W 'scratch copy deleted' }
} catch { W ('scratch delete failed -> ' + $_.Exception.Message) }

W ''
W 'EV-1c PROBE COMPLETE'

$lines | Set-Content -LiteralPath $outFile -Encoding UTF8
Write-Host ('EV-1c output written to: ' + $outFile)
