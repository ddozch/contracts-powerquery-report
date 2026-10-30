<#
.SYNOPSIS
    Golden-master test: refreshes ContractsReport.xlsx against each demo variant and compares
    every report figure with data\expected.json.

.DESCRIPTION
    expected.json is written by tools\generate_demo_data.py from the generator's own records,
    separately from the Power Query code. For each variant (a, b, empty) the test points the
    named cell SourcePath at data\variant_<v>\contracts_register.xlsx (variant b as a path
    relative to the workbook, the others as absolute paths), runs Data > Refresh All
    (Workbook.RefreshAll), recalculates and checks the row counts, KPIs, the monthly table,
    the by-type and by-manager tables, issue counts by problem and the rows on the List sheet.
    It also checks the exact (case-sensitive) spellings in tblContracts and the Contract No
    column, and then repeats the figure checks for every scenario in expected.json: another
    report year, the Manager and Contract type filters set.
    The workbook is closed without saving. Exit code 1 when any check fails.

.PARAMETER SourcePath
    Ad-hoc mode: refresh against this file instead and print the KPIs, the monthly table,
    issue counts by problem, or the refresh error message. Nothing is compared.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File tools\test_report.ps1

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File tools\test_report.ps1 -SourcePath C:\Temp\register.xlsx
#>
param(
    [string]$SourcePath
)

$ErrorActionPreference = 'Stop'
# Excel COM calls can fail with type-library errors under a non-English culture
[Threading.Thread]::CurrentThread.CurrentCulture = 'en-US'

$root = Split-Path $PSScriptRoot -Parent
$WorkbookPath = Join-Path $root 'ContractsReport.xlsx'
$Variants = @('a', 'b', 'empty')
$MoneyTolerance = 0.005
$Inv = [Globalization.CultureInfo]::InvariantCulture

# Excel rejects COM calls while Power Query is busy (RPC_E_CALL_REJECTED); this OLE message
# filter makes COM retry them. Same filter as in build_report.ps1.
if (-not ('ExcelBusyFilter' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

[ComImport, Guid("00000016-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IOleMessageFilter {
    [PreserveSig] int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo);
    [PreserveSig] int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType);
    [PreserveSig] int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType);
}

public class ExcelBusyFilter : IOleMessageFilter {
    [DllImport("ole32.dll")]
    private static extern int CoRegisterMessageFilter(IOleMessageFilter newFilter, out IOleMessageFilter oldFilter);
    public static void Register() { IOleMessageFilter old; CoRegisterMessageFilter(new ExcelBusyFilter(), out old); }
    public static void Revoke() { IOleMessageFilter old; CoRegisterMessageFilter(null, out old); }
    public int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo) { return 0; }
    public int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType) { return dwTickCount < 120000 ? 200 : -1; }
    public int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType) { return 2; }
}
'@
}

# --- Checks -----------------------------------------------------------------------------------
$script:passed = 0
$script:failed = 0

function Fail([string]$message) {
    $script:failed++
    Write-Host ('FAIL  ' + $message)
}

# Excel error values arrive through Value2 as Int32 codes; show them as such
function Show($value) {
    if ($null -eq $value) { return '(empty)' }
    if ($value -is [int] -and $value -lt -2146820000) { return "Excel error $value" }
    if ($value -is [double] -or $value -is [int]) { return $value.ToString('0.##', $Inv) }
    return "'$value'"
}

# A number check. "" or text in a figure cell fails, although PowerShell would cast "" to 0.
function Check([string]$name, $actual, $want, [double]$tolerance = 0) {
    $isNumber = $actual -is [double] -or ($actual -is [int] -and $actual -gt -2146820000)
    if ($isNumber -and [math]::Abs([double]$actual - [double]$want) -le $tolerance) {
        $script:passed++
    } else {
        Fail ('{0}: workbook {1}, expected {2}' -f $name, (Show $actual), ([double]$want).ToString('0.##', $Inv))
    }
}

function CheckText([string]$name, $actual, [string]$want) {
    if ([string]$actual -ceq $want) { $script:passed++ } else { Fail ("{0}: workbook '{1}', expected '{2}'" -f $name, $actual, $want) }
}

# --- Reading the workbook ---------------------------------------------------------------------
# Values of one table column as a flat list (Value2 is a 2-D array, or a scalar for one row)
function ColumnValues($lo, [string]$column) {
    $out = New-Object Collections.ArrayList
    if ($lo.ListRows.Count -eq 0) { return , $out }
    $values = $lo.ListColumns.Item($column).DataBodyRange.Value2
    if ($lo.ListRows.Count -eq 1) { [void]$out.Add($values) } else { foreach ($v in $values) { [void]$out.Add($v) } }
    return , $out
}

# Case-sensitive counts: a PowerShell @{} ignores case and would hide "ANNEX" next to "Annex"
function CountValues($lo, [string]$column) {
    $counts = New-Object 'Collections.Generic.Dictionary[string,int]' ([StringComparer]::Ordinal)
    foreach ($v in (ColumnValues $lo $column)) {
        $k = [string]$v
        $counts[$k] = (CountOf $counts $k) + 1
    }
    return , $counts
}

function CountOf($counts, [string]$key) {
    if ($counts.ContainsKey($key)) { return $counts[$key] }
    return 0
}

# Expected {name: count} object from expected.json as a case-sensitive dictionary
function ToCounts($obj) {
    $counts = New-Object 'Collections.Generic.Dictionary[string,int]' ([StringComparer]::Ordinal)
    foreach ($p in $obj.PSObject.Properties) { $counts[$p.Name] = [int]$p.Value }
    return , $counts
}

function Sha256Hex([string]$text) {
    $sha = [Security.Cryptography.SHA256]::Create()
    try { $bytes = $sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($text)) } finally { $sha.Dispose() }
    return (($bytes | ForEach-Object { $_.ToString('x2') }) -join '')
}

# Long names (a sentence typed as a contract type) are cut in FAIL lines
function Short([string]$name) {
    if ($name.Length -le 40) { return $name }
    return $name.Substring(0, 37) + '...'
}

# Rows of a dynamic array that starts in $cell; 0 for "" or the FILTER fallback text, -1 for an error
function SpillRows($cell, [string]$emptyText) {
    if ($cell.HasSpill) { return $cell.SpillingToRange.Rows.Count }
    if ([string]$cell.Text -in @('', $emptyText)) { return 0 }
    return -1
}

# Month cell (a date serial) as "yyyy-MM", the form expected.json uses
function MonthText($value) {
    if ($value -is [double]) { return [DateTime]::FromOADate($value).ToString('yyyy-MM', $Inv) }
    return (Show $value)
}

function HeaderNames($lo) {
    return (@($lo.HeaderRowRange.Value2) -join '|')
}

# RefreshAll does not report a failed refresh (DisplayAlerts is off) and keeps the old rows,
# which would let a broken variant pass with the previous variant's figures. A marker row is
# added to each table first: a refresh replaces it, a failed one leaves it in place.
$Marker = '__not refreshed__'
function RefreshAllChecked($wb, $tables) {
    foreach ($lo in $tables) {
        $row = $lo.ListRows.Add()
        $row.Range.Cells.Item(1, 1).Value2 = $Marker
    }
    $wb.RefreshAll()
    $wb.Application.CalculateUntilAsyncQueriesDone()
    foreach ($lo in $tables) {
        $n = $lo.ListRows.Count
        if ($n -gt 0 -and [string]$lo.DataBodyRange.Cells.Item($n, 1).Value2 -eq $Marker) {
            # Refreshing one table directly throws with Power Query's message
            try { [void]$lo.QueryTable.Refresh($false) } catch { return $_.Exception.Message }
            return "Table $($lo.Name) was not refreshed by Refresh All"
        }
    }
    return $null
}

function FindHashes($ws, [string]$range) {
    $bad = @()
    foreach ($cell in $ws.Range($range).Cells) {
        if ([string]$cell.Text -match '^#+$') { $bad += $cell.Address($false, $false) }
    }
    return $bad
}

# --- Report inputs ----------------------------------------------------------------------------
# PowerShell caches a COM setter per call site and value type, so the number has its own line
function SetInputs($rep, [double]$year, [string]$manager, [string]$type) {
    $rep.Range('C5').Value2 = $year
    $rep.Range('C6').Value2 = $manager
    $rep.Range('C7').Value2 = $type
}

# --- Figures that depend on the inputs --------------------------------------------------------
# KPIs, the monthly table, by type, by manager and the List rows against one expected block.
# $prefix names the scenario in FAIL lines.
function CheckFigures($ctx, $want, [string]$prefix) {
    $rep = $ctx.Report
    Check "${prefix}KPI drafted" $rep.Range('B10').Value2 $want.kpi.drafted
    Check "${prefix}KPI signed" $rep.Range('D10').Value2 $want.kpi.signed
    Check "${prefix}KPI signed value" $rep.Range('F10').Value2 $want.kpi.signed_value $MoneyTolerance
    Check "${prefix}KPI pending" $rep.Range('H10').Value2 $want.kpi.pending

    # Monthly table B14:E25
    $grid = $rep.Range('B14:E25').Value2
    for ($m = 0; $m -lt 12; $m++) {
        $row = $want.monthly[$m]
        $r = $m + 1
        CheckText "${prefix}month row $r" (MonthText $grid[$r, 1]) $row.month
        Check "${prefix}$($row.month) drafted" $grid[$r, 2] $row.drafted
        Check "${prefix}$($row.month) signed" $grid[$r, 3] $row.signed
        Check "${prefix}$($row.month) signed value" $grid[$r, 4] $row.signed_value $MoneyTolerance
    }

    CheckBreakdown "${prefix}by type" $rep.Range('G14') $want.by_type @('drafted', 'signed', 'signed_value')
    CheckBreakdown "${prefix}by manager" $rep.Range('L14') $want.by_manager @('drafted', 'signed', 'signed_value', 'pending')

    $listRows = SpillRows $ctx.List.Range('A3') 'No contracts'
    if ($listRows -lt 0) { Fail ("{0}List!A3 shows '{1}'" -f $prefix, $ctx.List.Range('A3').Text) }
    else { Check "${prefix}List rows" $listRows $want.list_rows }
}

# Case-sensitive {value: count} comparison, both ways: expected counts and no unexpected value
function CheckCounts([string]$label, $actual, $want) {
    foreach ($k in $want.Keys) { Check "$label '$(Short $k)'" (CountOf $actual $k) $want[$k] }
    foreach ($k in $actual.Keys) {
        if (-not $want.ContainsKey($k)) { Fail ("{0}: unexpected value '{1}' ({2} rows)" -f $label, (Short $k), $actual[$k]) }
    }
}

# --- One variant against expected.json --------------------------------------------------------
function TestVariant($ctx, [string]$variant, $want, [int]$year) {
    $rep = $ctx.Report
    $contracts = $ctx.Contracts
    $issues = $ctx.Issues
    SetInputs $rep $year 'All' 'All'

    $refreshError = RefreshAllChecked $ctx.Workbook @($contracts, $issues)
    if ($refreshError) { Fail "refresh failed: $refreshError"; return }
    $ctx.Excel.CalculateFull()

    # Tables keep their columns, also with no rows
    CheckText 'tblContracts columns' (HeaderNames $contracts) 'Contract No|Client|Contract Type|Manager|Drafted|Signed|Value EUR|Drafted Month|Signed Month|Status'
    CheckText 'tblIssues columns' (HeaderNames $issues) 'Contract No|Field|Problem|Original Value'
    Check 'tblContracts rows' $contracts.ListRows.Count $want.contracts_rows
    Check 'KPI data issues' $rep.Range('J10').Value2 $want.kpi.issues
    CheckFigures $ctx $want ''
    CheckCounts 'issues by problem' (CountValues $issues 'Problem') (ToCounts $want.issues_by_problem)

    # Exact spellings: Excel's UNIQUE and "=" ignore case, so the figures alone cannot show
    # whether "INSTALLATION  CONTRACT" became "Installation contract" or a client was cleaned
    foreach ($pair in @(@('Contract Type', 'contract_type'), @('Manager', 'manager'), @('Status', 'status'), @('Client', 'client'))) {
        CheckCounts "tblContracts[$($pair[0])]" (CountValues $contracts $pair[0]) (ToCounts $want.exact_counts.($pair[1]))
    }
    # Assigned first: piped straight from the function, the list would arrive as one object
    $numbers = ColumnValues $contracts 'Contract No'
    $joined = (@($numbers) | ForEach-Object { [string]$_ }) -join "`n"
    CheckText 'Contract No column (SHA-256)' (Sha256Hex $joined) $want.contract_no_sha256

    $hashes = @(FindHashes $rep 'B1:P40') + @(FindHashes $ctx.List 'A1:J60')
    if ($hashes.Count -eq 0) { $script:passed++ } else { Fail ('cells too narrow (###): ' + ($hashes -join ', ')) }

    # The same data with another report year or with the filters set; formulas only, no refresh
    foreach ($s in @($want.scenarios)) {
        SetInputs $rep $s.report_year $s.manager $s.type
        $ctx.Excel.CalculateFull()
        CheckFigures $ctx $s "[$($s.name)] "
    }
    SetInputs $rep $year 'All' 'All'
}

# A spilled breakdown (name column + figure columns) against an expected {name: {field: n}} map.
# With no data the formula shows "" in the first cell.
function CheckBreakdown([string]$label, $first, $want, [string[]]$fields) {
    $names = @($want.PSObject.Properties | ForEach-Object { $_.Name })
    if ($names.Count -eq 0) {
        CheckText "$label first cell" $first.Text ''
        return
    }
    $n = SpillRows $first ''
    if ($n -le 0) { Fail ("{0}: {1} shows '{2}'" -f $label, $first.Address($false, $false), $first.Text); return }
    Check "$label rows" $n $names.Count
    $grid = $first.Resize($n, $fields.Count + 1).Value2
    $seen = New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::Ordinal)
    for ($r = 1; $r -le $n; $r++) {
        $name = [string]$grid[$r, 1]
        $exp = $want.PSObject.Properties | Where-Object { $_.Name -ceq $name } | Select-Object -First 1
        if (-not $exp) { Fail "$label unexpected row '$(Short $name)'"; continue }
        [void]$seen.Add($name)
        for ($i = 0; $i -lt $fields.Count; $i++) {
            $tol = if ($fields[$i] -eq 'signed_value') { $MoneyTolerance } else { 0 }
            Check "$label '$(Short $name)' $($fields[$i])" $grid[$r, ($i + 2)] $exp.Value.($fields[$i]) $tol
        }
    }
    foreach ($name in $names) { if (-not $seen.Contains($name)) { Fail "$label missing row '$(Short $name)'" } }
}

# --- Ad-hoc run against any file --------------------------------------------------------------
function ShowFile($ctx) {
    $rep = $ctx.Report
    $refreshError = RefreshAllChecked $ctx.Workbook @($ctx.Contracts, $ctx.Issues)
    if ($refreshError) {
        Write-Host "Refresh failed: $refreshError"
        return $false
    }
    $ctx.Excel.CalculateFull()
    Write-Host ('Rows: tblContracts {0}, tblIssues {1}, List {2}' -f $ctx.Contracts.ListRows.Count, $ctx.Issues.ListRows.Count,
        (SpillRows $ctx.List.Range('A3') 'No contracts'))
    Write-Host $rep.Range('B1').Text
    foreach ($k in @(@('B10', 'Drafted'), @('D10', 'Signed'), @('F10', 'Signed value (EUR)'), @('H10', 'Pending signature'), @('J10', 'Data issues'))) {
        Write-Host ('  {0,-20} {1}' -f $k[1], (Show $rep.Range($k[0]).Value2))
    }
    Write-Host ('  {0,-9} {1,8} {2,8} {3,18}' -f 'Month', 'Drafted', 'Signed', 'Signed value')
    $grid = $rep.Range('B14:E25').Value2
    for ($r = 1; $r -le 12; $r++) {
        $month = MonthText $grid[$r, 1]
        $money = if ($grid[$r, 4] -is [double]) { $grid[$r, 4].ToString('#,##0.00', $Inv) } else { Show $grid[$r, 4] }
        Write-Host ('  {0,-9} {1,8} {2,8} {3,18}' -f $month, (Show $grid[$r, 2]), (Show $grid[$r, 3]), $money)
    }
    $counts = CountValues $ctx.Issues 'Problem'
    Write-Host 'Issues by problem:'
    if ($counts.Count -eq 0) { Write-Host '  (none)' }
    foreach ($k in ($counts.Keys | Sort-Object)) { Write-Host ('  {0,-22} {1}' -f $k, $counts[$k]) }
    return $true
}

# --- Run --------------------------------------------------------------------------------------
if (-not (Test-Path $WorkbookPath -PathType Leaf)) { throw "Workbook not found: $WorkbookPath. Run tools\build_report.ps1 first." }
if ($SourcePath) {
    $SourcePath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($SourcePath)
} else {
    $expected = Get-Content (Join-Path $root 'data\expected.json') -Raw -Encoding UTF8 | ConvertFrom-Json
}

# The run is a function: when it returns, the COM wrappers it created (workbook, sheets,
# tables, ranges) are out of scope and the GC below can free them. Excel does not exit while
# one of them is alive. On failure, Quit discards the open workbook (DisplayAlerts off).
function RunTests($excel) {
    $wb = $excel.Workbooks.Open($WorkbookPath)
    $ctx = @{
        Excel     = $excel
        Workbook  = $wb
        Report    = $wb.Worksheets.Item('Report')
        List      = $wb.Worksheets.Item('List')
        Contracts = $wb.Worksheets.Item('Contracts').ListObjects.Item('tblContracts')
        Issues    = $wb.Worksheets.Item('Data Issues').ListObjects.Item('tblIssues')
    }
    $rep = $ctx.Report

    if ($SourcePath) {
        # Filters on "All", the report year as saved in the workbook
        $rep.Range('C6').Value2 = 'All'
        $rep.Range('C7').Value2 = 'All'
        $rep.Range('C4').Value2 = $SourcePath
        Write-Host "Source: $SourcePath"
        if (-not (ShowFile $ctx)) { $script:exitCode = 1 }
    } else {
        # The shipped workbook points at the demo register relative to its own folder
        CheckText 'shipped Source file (Report!C4)' $rep.Range('C4').Value2 'data\contracts_register.xlsx'
        CheckText 'WorkbookFolder (Lists!D1)' $wb.Worksheets.Item('Lists').Range('D1').Value2 ($root + '\')
        foreach ($v in $Variants) {
            $file = Join-Path $root "data\variant_$v\contracts_register.xlsx"
            $want = $expected.variants.$v
            $before = @($script:passed, $script:failed)
            Write-Host "Variant ${v}: data\variant_$v\contracts_register.xlsx"
            if (-not $want) { Fail "variant $v is missing in expected.json" }
            elseif (-not (Test-Path $file -PathType Leaf)) { Fail "source file not found: $file" }
            else {
                # Variant b goes through a relative path, the others through an absolute one
                $path = if ($v -eq 'b') { "data\variant_$v\contracts_register.xlsx" } else { $file }
                $rep.Range('C4').Value2 = $path
                TestVariant $ctx $v $want ([int]$expected.report_year)
            }
            Write-Host ('Variant {0}: {1} checks passed, {2} failed' -f $v, ($script:passed - $before[0]), ($script:failed - $before[1]))
        }
        Write-Host ('Total: {0} checks passed, {1} failed' -f $script:passed, $script:failed)
        if ($script:failed -gt 0) { $script:exitCode = 1 }
    }
    $wb.Close($false)
}

[ExcelBusyFilter]::Register()
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$script:exitCode = 0
try {
    RunTests $excel
}
catch {
    $script:exitCode = 1
    Write-Host ('Test aborted: ' + $_.Exception.Message)
    Write-Host $_.InvocationInfo.PositionMessage
}
finally {
    # Quit throws when Excel has died; the release and the GC must run all the same
    try { $excel.Quit() } catch { }
    try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) } catch { }
    $excel = $null
    # The second pass frees the wrappers whose finalizers ran in the first
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    [ExcelBusyFilter]::Revoke()
}
exit $script:exitCode
