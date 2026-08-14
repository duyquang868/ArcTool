# EV-1e — WPS workbook/document SURFACE DISCOVERY probe
#
# Purpose: EV-1d proved Application.Workbooks returns null on this machine across
# CreateInstance, GetActiveObject, 3 binders and a readiness loop. This probe does NOT
# retry Workbooks.Open argument permutations. It answers a different question:
#
#   Which member on this WPS automation object actually yields a usable
#   workbook/document surface?
#
# Method:
#   Stage A  - dump the real IDispatch ITypeInfo (member names + invkind) of KET.Application
#   Stage B  - GetIDsOfNames DISPID map for candidate member names (known-name vs unknown-name)
#   Stage C  - reflection GetProperty matrix over the same candidates, null-checked
#   Stage D  - application-level open surface: launch et.exe with a real file, attach via ROT,
#              then re-probe ActiveWorkbook / Workbooks / ActiveWindow / Documents / Sheets
#   Stage E  - if a workbook-ish object is found, dump ITypeInfo + probe downstream members
#   Stage F  - cleanup (close without saving, Quit, kill only PIDs this probe started)
#
# No Revit. No .rvt. No Revit MCP. Spreadsheet COM only.

$ErrorActionPreference = 'Continue'
[System.Threading.Thread]::CurrentThread.CurrentCulture   = 'en-US'
[System.Threading.Thread]::CurrentThread.CurrentUICulture = 'en-US'

function Section($t) {
    Write-Output ''
    Write-Output '=============================================================================='
    Write-Output ("== " + $t)
    Write-Output '=============================================================================='
}

$results = New-Object System.Collections.ArrayList
function AddResult($stage, $name, $result, $detail) {
    [void]$results.Add([pscustomobject]@{ Stage = $stage; Member = $name; Result = $result; Detail = $detail })
}

# ---------------------------------------------------------------- C# COM reflection helper
$cs = @'
using System;
using System.Text;
using System.Runtime.InteropServices;

namespace ArcToolProbe
{
    [ComImport, Guid("00020400-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IDispatchFull
    {
        void GetTypeInfoCount(out int pctinfo);
        void GetTypeInfo(int iTInfo, int lcid, out System.Runtime.InteropServices.ComTypes.ITypeInfo ppTInfo);
        [PreserveSig]
        int GetIDsOfNames(ref Guid riid,
            [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] rgszNames,
            int cNames, int lcid,
            [MarshalAs(UnmanagedType.LPArray)] int[] rgDispId);
    }

    public static class ComProbe
    {
        public static bool HasDispatch(object com)
        {
            return (com as IDispatchFull) != null;
        }

        // returns dispid, or a negative HRESULT-ish marker string via out param
        public static string GetDispId(object com, string name)
        {
            IDispatchFull d = com as IDispatchFull;
            if (d == null) { return "no IDispatch"; }
            Guid iidNull = Guid.Empty;
            string[] names = new string[] { name };
            int[] ids = new int[] { -1 };
            int hr = d.GetIDsOfNames(ref iidNull, names, 1, 1033, ids);
            if (hr == 0) { return "dispid=" + ids[0]; }
            if ((uint)hr == 0x80020006) { return "UNKNOWN_NAME"; }
            return "hr=0x" + ((uint)hr).ToString("X8");
        }

        public static string Dump(object com, int maxFuncs)
        {
            StringBuilder sb = new StringBuilder();
            try
            {
                IDispatchFull d = com as IDispatchFull;
                if (d == null) { return "target does not expose IDispatch"; }

                int count = 0;
                d.GetTypeInfoCount(out count);
                sb.AppendLine("GetTypeInfoCount = " + count);
                if (count == 0) { sb.AppendLine("NO TYPE INFO -> name binding is the only route"); return sb.ToString(); }

                System.Runtime.InteropServices.ComTypes.ITypeInfo ti = null;
                d.GetTypeInfo(0, 1033, out ti);
                if (ti == null) { sb.AppendLine("GetTypeInfo returned null"); return sb.ToString(); }

                string tn, doc, hf; int hc;
                ti.GetDocumentation(-1, out tn, out doc, out hc, out hf);
                sb.AppendLine("type name = " + tn);
                sb.AppendLine("type doc  = " + doc);

                IntPtr pAttr = IntPtr.Zero;
                int cFuncs = 0, cVars = 0;
                ti.GetTypeAttr(out pAttr);
                if (pAttr != IntPtr.Zero)
                {
                    System.Runtime.InteropServices.ComTypes.TYPEATTR attr =
                        (System.Runtime.InteropServices.ComTypes.TYPEATTR)Marshal.PtrToStructure(
                            pAttr, typeof(System.Runtime.InteropServices.ComTypes.TYPEATTR));
                    sb.AppendLine("guid      = " + attr.guid);
                    sb.AppendLine("typekind  = " + attr.typekind);
                    sb.AppendLine("cFuncs    = " + attr.cFuncs);
                    sb.AppendLine("cVars     = " + attr.cVars);
                    cFuncs = attr.cFuncs;
                    cVars = attr.cVars;
                    ti.ReleaseTypeAttr(pAttr);
                }

                int limit = cFuncs;
                if (maxFuncs > 0 && limit > maxFuncs) { limit = maxFuncs; }
                for (int i = 0; i < limit; i++)
                {
                    IntPtr pf = IntPtr.Zero;
                    try
                    {
                        ti.GetFuncDesc(i, out pf);
                        if (pf == IntPtr.Zero) { continue; }
                        System.Runtime.InteropServices.ComTypes.FUNCDESC fd =
                            (System.Runtime.InteropServices.ComTypes.FUNCDESC)Marshal.PtrToStructure(
                                pf, typeof(System.Runtime.InteropServices.ComTypes.FUNCDESC));
                        string n2, d2, hf2; int hc2;
                        ti.GetDocumentation(fd.memid, out n2, out d2, out hc2, out hf2);
                        sb.AppendLine("FUNC name=" + n2 + " memid=" + fd.memid + " invkind=" + fd.invkind + " cParams=" + fd.cParams);
                    }
                    catch (Exception ex) { sb.AppendLine("FUNC[" + i + "] dump error: " + ex.Message); }
                    finally { if (pf != IntPtr.Zero) { ti.ReleaseFuncDesc(pf); } }
                }
                if (cFuncs > limit) { sb.AppendLine("... " + (cFuncs - limit) + " more funcs not dumped"); }

                for (int i = 0; i < cVars; i++)
                {
                    IntPtr pv = IntPtr.Zero;
                    try
                    {
                        ti.GetVarDesc(i, out pv);
                        if (pv == IntPtr.Zero) { continue; }
                        System.Runtime.InteropServices.ComTypes.VARDESC vd =
                            (System.Runtime.InteropServices.ComTypes.VARDESC)Marshal.PtrToStructure(
                                pv, typeof(System.Runtime.InteropServices.ComTypes.VARDESC));
                        string n3, d3, hf3; int hc3;
                        ti.GetDocumentation(vd.memid, out n3, out d3, out hc3, out hf3);
                        sb.AppendLine("VAR  name=" + n3 + " memid=" + vd.memid);
                    }
                    catch (Exception ex) { sb.AppendLine("VAR[" + i + "] dump error: " + ex.Message); }
                    finally { if (pv != IntPtr.Zero) { ti.ReleaseVarDesc(pv); } }
                }
            }
            catch (Exception ex) { sb.AppendLine("dump failed: " + ex.Message); }
            return sb.ToString();
        }
    }
}
'@

Section 'STEP 0 - probe metadata'
Write-Output ("timestamp_local     : " + (Get-Date).ToString('yyyy-MM-dd HH:mm:ss zzz'))
Write-Output ("machine             : " + $env:COMPUTERNAME)
Write-Output ("powershell_version  : " + $PSVersionTable.PSVersion)
Write-Output ("process_is_64bit    : " + [Environment]::Is64BitProcess)
Write-Output  'probe_variant       : EV-1e (workbook/document surface discovery)'
Write-Output ("thread_culture      : " + [System.Threading.Thread]::CurrentThread.CurrentCulture.Name)

try {
    Add-Type -TypeDefinition $cs -Language CSharp -ErrorAction Stop
    Write-Output 'ComProbe helper     : COMPILED OK'
    $helperOk = $true
} catch {
    Write-Output ('ComProbe helper     : COMPILE FAILED -> ' + $_.Exception.Message)
    $helperOk = $false
}

$src = 'C:\Users\ADMIN\Desktop\PA4\BULONG.xlsx'
Write-Output ("source_workbook     : " + $src)
Write-Output ("source_exists       : " + (Test-Path $src))

$scratch = Join-Path $env:TEMP ('ArcTool_EV1e_' + ([guid]::NewGuid().ToString('N')) + '.xlsx')
if (Test-Path $src) {
    Copy-Item $src $scratch -Force
    Write-Output ("scratch_copy        : " + $scratch)
    Write-Output ("scratch_exists      : " + (Test-Path $scratch))
}

# ---------------------------------------------------------------- candidate member names
$candidates = @(
    'Workbooks','ActiveWorkbook','Documents','ActiveDocument','ActiveWindow','Windows',
    'Sheets','Worksheets','ActiveSheet','ActiveCell','Selection','RecentFiles',
    'Application','Parent','Books','WorkBook','Workbook','Document','ETApplication',
    'Version','Visible','DisplayAlerts','Name','Path','Caption','Range','Cells'
)
$openMethodCandidates = @('Open','OpenFile','OpenDocument','OpenWorkbook','Load','LoadFile','FileOpen','OpenText')

function Probe-Get($target, $name) {
    try {
        $v = $target.GetType().InvokeMember($name, [System.Reflection.BindingFlags]::GetProperty, $null, $target, $null)
        if ($null -eq $v) { return @{ ok = $false; detail = 'returned NULL' } }
        $tn = 'unknown'
        try { $tn = $v.GetType().FullName } catch { }
        return @{ ok = $true; detail = ('non-null <' + $tn + '>'); value = $v }
    } catch {
        $m = $_.Exception.Message
        if ($_.Exception.InnerException) { $m = $_.Exception.InnerException.Message }
        if ($m.Length -gt 90) { $m = $m.Substring(0, 90) }
        return @{ ok = $false; detail = ('ERR: ' + $m) }
    }
}

# =========================================================================== STAGE A / B / C
Section 'STAGE A - CreateInstance(KET.Application) + ITypeInfo dump'
$app = $null
try {
    $t = [Type]::GetTypeFromProgID('KET.Application', $false)
    if ($null -eq $t) { Write-Output 'KET.Application did NOT resolve - aborting' ; exit 1 }
    Write-Output 'KET.Application resolved'
    $app = [Activator]::CreateInstance($t)
    Write-Output ('CreateInstance ok   : ' + $app.GetType().FullName)
    AddResult 'A' 'CreateInstance' 'OK' $app.GetType().FullName
} catch {
    Write-Output ('CreateInstance FAILED: ' + $_.Exception.Message)
    AddResult 'A' 'CreateInstance' 'FAIL' $_.Exception.Message
}

if ($app -and $helperOk) {
    Write-Output ''
    Write-Output '--- ITypeInfo dump of the application object (max 250 funcs) ---'
    Write-Output ([ArcToolProbe.ComProbe]::Dump($app, 250))
}

if ($app -and $helperOk) {
    Section 'STAGE B - GetIDsOfNames DISPID map (known-name vs UNKNOWN_NAME)'
    foreach ($c in ($candidates + $openMethodCandidates)) {
        $r = [ArcToolProbe.ComProbe]::GetDispId($app, $c)
        Write-Output ("  {0,-18} -> {1}" -f $c, $r)
        $bRes = 'UNKNOWN'
        if ($r -like 'dispid=*') { $bRes = 'KNOWN' }
        AddResult 'B' ('dispid:' + $c) $bRes $r
    }
}

if ($app) {
    Section 'STAGE C - reflection GetProperty matrix (null-checked), no document open yet'
    foreach ($c in $candidates) {
        $r = Probe-Get $app $c
        $res = 'NULL/ERR'
        if ($r.ok) { $res = 'NON-NULL' }
        Write-Output ("  {0,-18} -> {1}" -f $c, $r.detail)
        AddResult 'C' $c $res $r.detail
    }
}

# =========================================================================== STAGE D
Section 'STAGE D - application-level open surface via et.exe + ROT attach'
Write-Output 'Rationale: ActiveWorkbook/ActiveWindow can only be non-null once a document is open.'
Write-Output 'KET LocalServer32 launches the spreadsheet as a MODE of the WPS shell, so this stage'
Write-Output 'tests whether the document surface appears after the app really owns a file.'

$etExe = 'C:\Users\ADMIN\AppData\Local\Kingsoft\WPS Office\12.1.0.28032\office6\et.exe'
Write-Output ("et_exe_exists       : " + (Test-Path $etExe))

$startedPid = $null
if ((Test-Path $etExe) -and (Test-Path $scratch)) {
    $before = @(Get-Process -Name 'et' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    try {
        $p = Start-Process -FilePath $etExe -ArgumentList ('"' + $scratch + '"') -PassThru
        $startedPid = $p.Id
        Write-Output ('launched et.exe pid : ' + $startedPid)
    } catch {
        Write-Output ('launch FAILED       : ' + $_.Exception.Message)
    }
    Start-Sleep -Seconds 8
    $after = @(Get-Process -Name 'et' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    Write-Output ('et pids before      : ' + ($before -join ','))
    Write-Output ('et pids after       : ' + ($after  -join ','))
}

$app2 = $null
Write-Output ''
Write-Output '--- ROT attach poll (up to 12 tries x 2s), checking the document surface each try ---'
$wbFound = $null
$wbFoundVia = ''
for ($i = 1; $i -le 12; $i++) {
    if ($null -eq $app2) {
        try { $app2 = [Runtime.InteropServices.Marshal]::GetActiveObject('KET.Application') } catch { $app2 = $null }
    }
    if ($app2) {
        foreach ($m in @('ActiveWorkbook','Workbooks','ActiveWindow','Documents','ActiveDocument','Sheets','Worksheets','ActiveSheet')) {
            $r = Probe-Get $app2 $m
            if ($r.ok) {
                Write-Output ("  try#{0} {1,-16} -> {2}" -f $i, $m, $r.detail)
                if ($null -eq $wbFound -and @('ActiveWorkbook','Workbooks','Documents','ActiveDocument') -contains $m) {
                    $wbFound = $r.value
                    $wbFoundVia = $m
                }
            } else {
                Write-Output ("  try#{0} {1,-16} -> {2}" -f $i, $m, $r.detail)
            }
        }
        if ($wbFound) { Write-Output ('  >>> document surface FOUND via ' + $wbFoundVia + ' on try#' + $i) ; break }
    } else {
        Write-Output ("  try#{0} GetActiveObject   -> not yet available" -f $i)
    }
    Start-Sleep -Seconds 2
}

if ($app2 -and $helperOk) {
    Write-Output ''
    Write-Output '--- ITypeInfo dump of the ROT-attached application object (max 250 funcs) ---'
    Write-Output ([ArcToolProbe.ComProbe]::Dump($app2, 250))
}

foreach ($m in @('ActiveWorkbook','Workbooks','ActiveWindow','Documents','ActiveDocument','Sheets','Worksheets','ActiveSheet')) {
    if ($app2) {
        $r = Probe-Get $app2 $m
        $res = 'NULL/ERR'
        if ($r.ok) { $res = 'NON-NULL' }
        AddResult 'D' ('rot:' + $m) $res $r.detail
    } else {
        AddResult 'D' ('rot:' + $m) 'NO-APP' 'GetActiveObject never returned an object'
    }
}

# =========================================================================== STAGE E
Section 'STAGE E - downstream surface on the discovered document object'
if ($wbFound) {
    Write-Output ('document surface via : ' + $wbFoundVia)
    try { Write-Output ('document type        : ' + $wbFound.GetType().FullName) } catch { }
    if ($helperOk) {
        Write-Output ''
        Write-Output '--- ITypeInfo dump of the document surface (max 250 funcs) ---'
        Write-Output ([ArcToolProbe.ComProbe]::Dump($wbFound, 250))
    }
    foreach ($m in @('Name','FullName','Path','Worksheets','Sheets','Names','ActiveSheet','Count','Item','Application','Parent')) {
        $r = Probe-Get $wbFound $m
        $res = 'NULL/ERR'
        if ($r.ok) { $res = 'NON-NULL' }
        Write-Output ("  {0,-14} -> {1}" -f $m, $r.detail)
        AddResult 'E' ($wbFoundVia + '.' + $m) $res $r.detail
    }
} else {
    Write-Output 'NO document surface found - Stage E SKIPPED.'
    Write-Output 'Consequence for the fix: the WPS provider must fail fast per contract, and must NOT'
    Write-Output 'assume Workbooks is a valid collection.'
    AddResult 'E' 'document-surface' 'ABSENT' 'no non-null workbook/document member found in stage C or D'
}

# =========================================================================== STAGE F
Section 'STAGE F - cleanup'
if ($wbFound) {
    try { $wbFound.GetType().InvokeMember('Close', [System.Reflection.BindingFlags]::InvokeMethod, $null, $wbFound, @($false)) | Out-Null ; Write-Output 'document Close(false) ok' } catch { Write-Output ('document Close failed: ' + $_.Exception.Message) }
    try { [Runtime.InteropServices.Marshal]::ReleaseComObject($wbFound) | Out-Null ; Write-Output 'ReleaseComObject(document) ok' } catch { }
}
foreach ($a in @($app2, $app)) {
    if ($a) {
        try { $a.GetType().InvokeMember('Quit', [System.Reflection.BindingFlags]::InvokeMethod, $null, $a, $null) | Out-Null ; Write-Output 'app Quit() ok' } catch { Write-Output ('app Quit failed: ' + $_.Exception.Message) }
        try { [Runtime.InteropServices.Marshal]::ReleaseComObject($a) | Out-Null ; Write-Output 'ReleaseComObject(app) ok' } catch { }
    }
}
if ($startedPid) {
    $still = Get-Process -Id $startedPid -ErrorAction SilentlyContinue
    if ($still) {
        try { Stop-Process -Id $startedPid -Force -ErrorAction Stop ; Write-Output ('killed the et.exe pid this probe started: ' + $startedPid) } catch { Write-Output ('could not kill pid ' + $startedPid + ': ' + $_.Exception.Message) }
    } else {
        Write-Output ('et.exe pid ' + $startedPid + ' already exited')
    }
}
if (Test-Path $scratch) { Remove-Item $scratch -Force -ErrorAction SilentlyContinue ; Write-Output 'scratch copy deleted' }

# =========================================================================== SUMMARY
Section 'SUMMARY - full matrix'
$results | Format-Table -AutoSize | Out-String -Width 200 | Write-Output

Section 'SUMMARY - NON-NULL / KNOWN only (this is the fix-relevant list)'
$results | Where-Object { $_.Result -eq 'NON-NULL' -or $_.Result -eq 'KNOWN' -or $_.Result -eq 'OK' } |
    Format-Table -AutoSize | Out-String -Width 200 | Write-Output

Section 'VERDICT'
if ($wbFound) {
    Write-Output ('EV-1e VERDICT: document surface FOUND via "' + $wbFoundVia + '"')
} else {
    Write-Output 'EV-1e VERDICT: NO usable workbook/document surface found on this WPS build'
}
Write-Output 'EV-1e PROBE COMPLETE'
