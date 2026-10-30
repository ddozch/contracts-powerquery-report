<#
.SYNOPSIS
    Builds ContractsReport.xlsx (no macros) from the Power Query files in src\ and refreshes it once.

.DESCRIPTION
    Every src\*.pq file becomes one query named after the file. Contracts and Issues are loaded
    to the tables tblContracts and tblIssues; all other queries are connection-only helpers.
    The Report, List and hidden Lists sheets get the formulas that summarise the two tables.
    Users refresh the finished workbook with Data > Refresh All.
    After Excel has closed, scrub_workbook.ps1 removes the local folder path that Excel saves
    inside the file, sets the Power Query locale to en-US and checks that no local path,
    computer name or Windows user name is left.
    If one is, the build stops with exit code 1; the workbook is saved all the same.

.PARAMETER SourcePath
    Contracts register the workbook reads, in Report!C4 (named cell SourcePath). A file inside
    the output folder is saved as a path relative to the workbook, any other file as an
    absolute path, which the check reports when it is in the user profile folder.
    Default: data\contracts_register.xlsx.

.PARAMETER Output
    Workbook to create. Default: ContractsReport.xlsx in the project folder.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File tools\build_report.ps1
#>
param(
    [string]$SourcePath = (Join-Path (Split-Path $PSScriptRoot -Parent) 'data\contracts_register.xlsx'),
    [string]$Output = (Join-Path (Split-Path $PSScriptRoot -Parent) 'ContractsReport.xlsx')
)

$ErrorActionPreference = 'Stop'
# Excel COM calls can fail with type-library errors under a non-English culture
[Threading.Thread]::CurrentThread.CurrentCulture = 'en-US'

$root = Split-Path $PSScriptRoot -Parent
$src = Join-Path $root 'src'
# Resolved against the PowerShell location, also for an output file that does not exist yet
$SourcePath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($SourcePath)
$Output = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Output)
if (-not (Test-Path $SourcePath -PathType Leaf)) { throw "Source file not found: $SourcePath" }

$ReportYear = 2026

# Query -> sheet and table it is loaded to; every other query stays connection-only
$Loaded = [ordered]@{
    'Contracts' = @('Contracts', 'tblContracts')
    'Issues'    = @('Data Issues', 'tblIssues')
}

# Excel colours are BGR integers
function Rgb([int]$r, [int]$g, [int]$b) { return $r + 256 * $g + 65536 * $b }
$ClrAccent = Rgb 22 82 76
$ClrAccentLight = Rgb 224 239 236
$ClrMuted = Rgb 90 96 94
$ClrInput = Rgb 255 250 230
$ClrWhite = Rgb 255 255 255

$FmtCount = '0'
$FmtMoney = '#,##0.00'
$FmtDate = 'yyyy-mm-dd'
$FmtMonth = '[$-409]mmm yyyy'

# --- Formula pieces (English names, comma separators, written with Formula2) ---------------
# LET names must not look like cell addresses ("y1" is cell Y1 and breaks references such as
# G14#), and "c" and "r" are reserved.
# Figures are sums of 0/1 conditions, not COUNTIFS/SUMIFS: a name is free text from the source,
# and as a COUNTIFS criterion "Annex*" would also count "Annex", "<other>" would mean "less than",
# and over 255 characters it fails. "=" compares the whole text, like the List sheet.

$LetCols = 'mgr,tblContracts[Manager],typ,tblContracts[Contract Type],drafted,tblContracts[Drafted],signed,tblContracts[Signed],amt,tblContracts[Value EUR]'
# 1 for the rows drafted / signed in the report year; empty dates compare as 0
$LetY = 'yStart,DATE(ReportYear,1,1),yEnd,DATE(ReportYear+1,1,1),dInYear,(drafted>=yStart)*(drafted<yEnd),sInYear,(signed>=yStart)*(signed<yEnd)'
$LetM = 'inMgr,((selManager="All")+(mgr=selManager))>0'
$LetT = 'inType,((selType="All")+(typ=selType))>0'

# Expands the placeholders above inside a formula template
function Fill([string]$template) {
    return $template.Replace('{C}', $LetCols).Replace('{Y}', $LetY).Replace('{M}', $LetM).Replace('{T}', $LetT)
}

# --- Helpers ----------------------------------------------------------------------------------
# PowerShell passes the user's regional LCID to Excel even after the culture change above, and
# Excel then reads format codes such as "General" or "#,##0.00" in that language. Passing
# en-US explicitly makes every format code English.
$EnUs = [Globalization.CultureInfo]'en-US'
function SetFormat($range, [string]$fmt) {
    [void][System.__ComObject].InvokeMember('NumberFormat', [Reflection.BindingFlags]::SetProperty, $null, $range, @($fmt), $EnUs)
}

# Excel rejects COM calls while Power Query or the calculation engine is busy
# (RPC_E_CALL_REJECTED). An OLE message filter makes COM retry them instead of failing.
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
    // Retry after 200 ms for up to two minutes, then give up
    public int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType) { return dwTickCount < 120000 ? 200 : -1; }
    public int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType) { return 2; }
}
'@
}

function AddSheet($wb, [string]$name) {
    $ws = $wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count))
    $ws.Name = $name
    return $ws
}

function AddName($wb, [string]$name, $range) {
    [void]$wb.Names.Add($name, "='" + $range.Worksheet.Name + "'!" + $range.Address())
}

function InputRow($wb, $ws, [int]$row, [string]$label, $value, [string]$fmt, [string]$name) {
    $ws.Cells.Item($row, 2).Value2 = $label
    $cell = $ws.Cells.Item($row, 3)
    SetFormat $cell $fmt
    # PowerShell caches the COM setter per call site and type, so text and numbers use separate lines
    if ($value -is [string]) { $cell.Value2 = $value } else { $cell.Value2 = [double]$value }
    $cell.Interior.Color = $ClrInput
    $cell.HorizontalAlignment = -4131    # xlLeft
    AddName $wb $name $cell
}

function SectionTitle($range, [string]$text) {
    $first = $range.Cells.Item(1, 1)
    $first.Value2 = $text
    $first.Font.Bold = $true
    $first.Font.Size = 12
    $first.Font.Color = $ClrAccent
    $edge = $range.Borders.Item(9)       # xlEdgeBottom
    $edge.LineStyle = 1
    $edge.Color = $ClrAccent
    $edge.Weight = 2
}

function HeaderStyle($range) {
    $range.Font.Bold = $true
    $range.Font.Color = $ClrWhite
    $range.Interior.Color = $ClrAccent
    $range.VerticalAlignment = -4108     # xlCenter
}

function HeaderRow($ws, [string]$firstCell, [string[]]$labels) {
    $start = $ws.Range($firstCell)
    for ($i = 0; $i -lt $labels.Count; $i++) { $start.Offset(0, $i).Value2 = $labels[$i] }
    HeaderStyle $ws.Range($start, $start.Offset(0, $labels.Count - 1))
}

function Kpi($ws, [string]$col, [string]$label, [string]$formula, [string]$fmt) {
    $l = $ws.Range($col + '9')
    $l.Value2 = $label
    $l.Font.Color = $ClrMuted
    $v = $ws.Range($col + '10')
    $v.Formula2 = $formula
    SetFormat $v $fmt
    $v.Font.Size = 16
    $v.Font.Bold = $true
    $v.HorizontalAlignment = -4131
}

function ListValidation($range, [string]$source) {
    $range.Validation.Delete()
    [void]$range.Validation.Add(3, 1, 1, $source)    # xlValidateList, xlValidAlertStop, xlBetween
}

# Loads one query to a table with the Mashup OLE DB provider (the pattern Excel itself records)
function LoadQuery($ws, [string]$query, [string]$table) {
    # $Workbook$ is a literal token of the provider, so the string is single-quoted
    $conn = 'OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=' + $query + ';Extended Properties=""'
    $lo = $ws.ListObjects.Add(0, $conn, [Type]::Missing, 1, $ws.Range('A1'))   # xlSrcExternal, xlYes
    $qt = $lo.QueryTable
    $qt.CommandType = 2                  # xlCmdSql
    $qt.CommandText = 'SELECT * FROM [' + $query + ']'
    $qt.BackgroundQuery = $false         # makes Workbook.RefreshAll synchronous
    $qt.RefreshStyle = 1                 # xlInsertDeleteCells
    $qt.AdjustColumnWidth = $true
    $qt.PreserveFormatting = $true       # keeps the number formats set below
    [void]$qt.Refresh($false)
    $lo.Name = $table
    $lo.TableStyle = 'TableStyleMedium1'
    try { $qt.WorkbookConnection.Name = 'Query - ' + $query } catch { }
    return $lo
}

function ColumnFormat($lo, [string]$column, [string]$fmt) {
    # Header and body; the table copies the format to rows added by later refreshes
    SetFormat $lo.ListColumns.Item($column).Range $fmt
}

# Window settings need the workbook window; failures only cost cosmetics
function WindowSetup($wb, $ws, [int]$freezeRows, [bool]$gridlines) {
    try {
        $ws.Activate()
        $win = $wb.Windows.Item(1)
        $win.FreezePanes = $false
        if ($freezeRows -gt 0) {
            $win.SplitColumn = 0
            $win.SplitRow = $freezeRows
            $win.FreezePanes = $true
        }
        $win.DisplayGridlines = $gridlines
    } catch {
        Write-Warning ("Window settings skipped on {0}: {1}" -f $ws.Name, $_.Exception.Message)
    }
}

# Column widths: the preset minimum, or wider when the refreshed values need it, but text
# columns at most $maximum: a whole sentence typed as a contract type must not widen the report
function FitColumns($ws, [string]$range, [hashtable]$minimum, [hashtable]$maximum = @{}) {
    [void]$ws.Range($range).Columns.AutoFit()
    foreach ($col in $minimum.Keys) {
        $c = $ws.Columns.Item($col)
        if ($c.ColumnWidth -lt $minimum[$col]) { $c.ColumnWidth = $minimum[$col] }
    }
    foreach ($col in $maximum.Keys) {
        $c = $ws.Columns.Item($col)
        if ($c.ColumnWidth -gt $maximum[$col]) { $c.ColumnWidth = $maximum[$col] }
    }
}

function FindHashes($ws, [string]$range) {
    $bad = @()
    foreach ($cell in $ws.Range($range).Cells) {
        if ([string]$cell.Text -match '^#+$') { $bad += $cell.Address($false, $false) }
    }
    return $bad
}

# --- Build ------------------------------------------------------------------------------------
# The build runs inside a function: when it returns, the COM wrappers it created (workbook,
# sheets, tables, ranges) are out of scope and the GC below can free them. Excel does not exit
# while one of them is alive. On failure, Quit discards the unsaved workbook (DisplayAlerts off).
function BuildWorkbook($excel) {
    $wb = $excel.Workbooks.Add()
    while ($wb.Worksheets.Count -gt 1) { $wb.Worksheets.Item($wb.Worksheets.Count).Delete() }
    $report = $wb.Worksheets.Item(1)
    $report.Name = 'Report'
    $list = AddSheet $wb 'List'
    $contractsWs = AddSheet $wb 'Contracts'
    $issuesWs = AddSheet $wb 'Data Issues'
    $lists = AddSheet $wb 'Lists'

    # Inputs first: the queries read SourcePath when the tables load
    InputRow $wb $report 4 'Source file' $SourcePath '@' 'SourcePath'
    InputRow $wb $report 5 'Report year' $ReportYear '0' 'ReportYear'
    InputRow $wb $report 6 'Manager' 'All' '@' 'selManager'
    InputRow $wb $report 7 'Contract type' 'All' '@' 'selType'

    # Privacy levels are not ignored (Excel's default): SourcePath is a function, so the path is
    # read inside SourceRows, the query that opens the file, and the formula firewall allows it.
    # Ignoring them is saved with a checksum that only the Windows account of the build can verify.
    $wb.Queries.FastCombine = $false
    $queryFiles = @(Get-ChildItem $src -Filter *.pq | Sort-Object Name)
    if ($queryFiles.Count -eq 0) { throw "No .pq files in $src" }
    foreach ($file in $queryFiles) {
        $m = [IO.File]::ReadAllText($file.FullName, [Text.Encoding]::UTF8)
        [void]$wb.Queries.Add($file.BaseName, $m)
    }
    foreach ($q in $Loaded.Keys) {
        if (-not (Test-Path (Join-Path $src "$q.pq"))) { throw "Query file src\$q.pq is missing" }
    }

    Write-Host "Loading queries from $SourcePath"
    $tables = @{}
    foreach ($q in $Loaded.Keys) {
        $sheetName, $tableName = $Loaded[$q]
        $tables[$tableName] = LoadQuery $wb.Worksheets.Item($sheetName) $q $tableName
    }
    $excel.CalculateUntilAsyncQueriesDone()
    $lo = $tables['tblContracts']
    foreach ($c in @('Drafted', 'Signed', 'Drafted Month', 'Signed Month')) { ColumnFormat $lo $c $FmtDate }
    ColumnFormat $lo 'Value EUR' $FmtMoney

    # Hidden drop-down sources for the filters
    $lists.Range('A1').Value2 = 'All'
    $lists.Range('A2').Formula2 = '=SORT(UNIQUE(FILTER(tblContracts[Manager],tblContracts[Manager]<>"","")))'
    $lists.Range('B1').Value2 = 'All'
    $lists.Range('B2').Formula2 = '=SORT(UNIQUE(FILTER(tblContracts[Contract Type],tblContracts[Contract Type]<>"","")))'
    # Folder of this workbook, read by the SourcePath function to resolve a relative path in C4.
    # CELL gives "C:\folder\[file.xlsx]Sheet", or "" before the first save
    $lists.Range('D1').Formula2 = '=IFERROR(LEFT(CELL("filename",$D$1),FIND("[",CELL("filename",$D$1))-1),"")'
    AddName $wb 'WorkbookFolder' $lists.Range('D1')

    # Report: title and filters
    $t = $report.Range('B1')
    $t.Formula2 = '="Contracts drafted and signed in "&ReportYear'
    $t.Font.Size = 20
    $t.Font.Bold = $true
    $t.Font.Color = $ClrAccent
    $report.Range('B2').Value2 = 'Refresh: Data > Refresh All'
    $report.Range('B2').Font.Color = $ClrMuted
    # The path is long; the light-yellow band shows where it can be typed
    $report.Range('C4:J4').Interior.Color = $ClrInput
    ListValidation $report.Range('C6') '=Lists!$A$1:$A$30'
    ListValidation $report.Range('C7') '=Lists!$B$1:$B$30'
    $report.Range('C5').Validation.Delete()
    [void]$report.Range('C5').Validation.Add(1, 1, 1, '2000', '2100')    # whole number, plain values
    foreach ($r in @(@('F6', 'Manager and type filters apply to the figures, the monthly table and the List sheet.'),
                     @('F7', 'By contract type uses only the manager filter; by manager only the type filter.'))) {
        $c = $report.Range($r[0])
        $c.Value2 = $r[1]
        $c.Font.Color = $ClrMuted
        $c.Font.Italic = $true
    }

    # KPIs for the report year
    $report.Range('B9:J10').Interior.Color = $ClrAccentLight
    Kpi $report 'B' 'Drafted' (Fill '=LET({C},{Y},{M},{T},SUM(dInYear*inMgr*inType))') $FmtCount
    Kpi $report 'D' 'Signed' (Fill '=LET({C},{Y},{M},{T},SUM(sInYear*inMgr*inType))') $FmtCount
    Kpi $report 'F' 'Signed value (EUR)' (Fill '=LET({C},{Y},{M},{T},SUM(sInYear*inMgr*inType*amt))') $FmtMoney
    # Column F is only the narrow gap after By month; a number cannot overflow into the empty
    # cell next to it, so the signed-value KPI spans F:G
    $kpiValue = $report.Range('F10:G10')
    $kpiValue.MergeCells = $true
    $kpiValue.HorizontalAlignment = -4131
    Kpi $report 'H' 'Pending signature' (Fill '=LET({C},{Y},{M},{T},SUM(dInYear*(signed="")*inMgr*inType))') $FmtCount
    Kpi $report 'J' 'Data issues' '=COUNTA(tblIssues[Problem])' $FmtCount

    # By month: both filters
    SectionTitle $report.Range('B12:E12') 'By month'
    HeaderRow $report 'B13' @('Month', 'Drafted', 'Signed', 'Signed value (EUR)')
    $report.Range('B14').Formula2 = '=DATE(ReportYear,SEQUENCE(12),1)'
    $report.Range('C14').Formula2 = Fill '=LET({C},{M},{T},dMonth,tblContracts[Drafted Month],MAP(B14#,LAMBDA(x,SUM((dMonth=x)*inMgr*inType))))'
    $report.Range('D14').Formula2 = Fill '=LET({C},{M},{T},sMonth,tblContracts[Signed Month],MAP(B14#,LAMBDA(x,SUM((sMonth=x)*inMgr*inType))))'
    $report.Range('E14').Formula2 = Fill '=LET({C},{M},{T},sMonth,tblContracts[Signed Month],MAP(B14#,LAMBDA(x,SUM((sMonth=x)*inMgr*inType*amt))))'
    SetFormat $report.Range('B14:B25') $FmtMonth
    $report.Range('B14:B25').HorizontalAlignment = -4131
    SetFormat $report.Range('C14:D25') $FmtCount
    SetFormat $report.Range('E14:E25') $FmtMoney

    # By contract type: every type in the data (any year), figures for the report year, manager filter only
    SectionTitle $report.Range('G12:J12') 'By contract type'
    HeaderRow $report 'G13' @('Contract type', 'Drafted', 'Signed', 'Signed value (EUR)')
    $report.Range('G14').Formula2 = '=SORT(UNIQUE(FILTER(tblContracts[Contract Type],tblContracts[Contract Type]<>"","")))'
    $report.Range('H14').Formula2 = Fill '=LET({C},{Y},{M},MAP(G14#,LAMBDA(x,IF(x="","",SUM((typ=x)*dInYear*inMgr)))))'
    $report.Range('I14').Formula2 = Fill '=LET({C},{Y},{M},MAP(G14#,LAMBDA(x,IF(x="","",SUM((typ=x)*sInYear*inMgr)))))'
    $report.Range('J14').Formula2 = Fill '=LET({C},{Y},{M},MAP(G14#,LAMBDA(x,IF(x="","",SUM((typ=x)*sInYear*inMgr*amt)))))'
    SetFormat $report.Range('H14:I60') $FmtCount
    SetFormat $report.Range('J14:J60') $FmtMoney

    # By manager: every manager in the data, "(no manager)" included, contract type filter only
    SectionTitle $report.Range('L12:P12') 'By manager'
    HeaderRow $report 'L13' @('Manager', 'Drafted', 'Signed', 'Signed value (EUR)', 'Pending')
    $report.Range('L14').Formula2 = '=SORT(UNIQUE(FILTER(tblContracts[Manager],tblContracts[Manager]<>"","")))'
    $report.Range('M14').Formula2 = Fill '=LET({C},{Y},{T},MAP(L14#,LAMBDA(x,IF(x="","",SUM((mgr=x)*dInYear*inType)))))'
    $report.Range('N14').Formula2 = Fill '=LET({C},{Y},{T},MAP(L14#,LAMBDA(x,IF(x="","",SUM((mgr=x)*sInYear*inType)))))'
    $report.Range('O14').Formula2 = Fill '=LET({C},{Y},{T},MAP(L14#,LAMBDA(x,IF(x="","",SUM((mgr=x)*sInYear*inType*amt)))))'
    $report.Range('P14').Formula2 = Fill '=LET({C},{Y},{T},MAP(L14#,LAMBDA(x,IF(x="","",SUM((mgr=x)*dInYear*(signed="")*inType)))))'
    # One range per call: a comma-separated address is read with the regional list separator
    foreach ($a in @('M14:N60', 'P14:P60')) { SetFormat $report.Range($a) $FmtCount }
    SetFormat $report.Range('O14:O60') $FmtMoney
    $report.Range('B13:P13').WrapText = $true

    # List: contracts drafted or signed in the report year, with the Report filters
    $t = $list.Range('A1')
    $t.Formula2 = '="Contracts drafted or signed in "&ReportYear&IF(selManager="All","",", manager "&selManager)&IF(selType="All","",", type "&selType)'
    $t.Font.Size = 14
    $t.Font.Bold = $true
    $t.Font.Color = $ClrAccent
    $list.Range('A2').Formula2 = '=tblContracts[#Headers]'
    HeaderStyle $list.Range('A2:J2')
    # Date ranges instead of YEAR(): one value YEAR() cannot read would turn the whole list into an error
    $list.Range('A3').Formula2 = Fill '=LET({C},{Y},{M},{T},FILTER(tblContracts,((dInYear+sInYear)>0)*inMgr*inType,"No contracts"))'
    # FILTER returns 0 for empty cells; the empty third section hides those zeros. No "@" text
    # section: Excel would treat the cells as Text and turn a re-entered formula into text
    foreach ($a in @('A:D', 'J:J')) { SetFormat $list.Range($a) 'General;-General;' }
    foreach ($a in @('E:F', 'H:I')) { SetFormat $list.Range($a) 'yyyy-mm-dd;;' }
    SetFormat $list.Range('G:G') '#,##0.00;-#,##0.00;'
    # Title and headers keep General: with the formats above the title would not overflow into B1
    SetFormat $list.Range('A1:J2') 'General'
    # Status follows the right-aligned Signed Month dates; the indent keeps the two apart
    $status = $list.Range('J:J')
    $status.HorizontalAlignment = -4131
    $status.IndentLevel = 1

    $excel.CalculateFull()

    # Widths: Report columns sized for full-year totals; F and K are equal gaps between the tables
    FitColumns $report 'B10:P60' @{ B = 16; C = 11; D = 11; E = 18; G = 24; H = 11; I = 11; J = 18; L = 18; M = 11; N = 11; O = 18; P = 11 } @{ G = 40; L = 28 }
    $report.Columns.Item('A').ColumnWidth = 2
    $report.Columns.Item('F').ColumnWidth = 4
    $report.Columns.Item('K').ColumnWidth = 4
    FitColumns $list 'A2:J5000' @{ A = 14; B = 24; C = 22; D = 16; E = 12; F = 12; G = 14; H = 13; I = 13; J = 10 } @{ B = 32; C = 40; D = 28 }
    $report.Rows.Item(13).AutoFit() | Out-Null

    WindowSetup $wb $issuesWs 1 $true
    WindowSetup $wb $contractsWs 1 $true
    WindowSetup $wb $list 2 $true
    WindowSetup $wb $report 0 $false
    $lists.Visible = 0                   # xlSheetHidden
    try { $report.Range('C4').Select() | Out-Null } catch { }

    $hashes = @(FindHashes $report 'B1:P40') + @(FindHashes $list 'A1:J60')
    if ($hashes.Count -gt 0) { Write-Warning ('Cells too narrow (###): ' + ($hashes -join ', ')) }

    Write-Host ('tblContracts: {0} rows, tblIssues: {1} rows' -f $tables['tblContracts'].ListRows.Count, $tables['tblIssues'].ListRows.Count)
    Write-Host ('{0}: drafted {1}, signed {2}, signed value {3}, pending {4}, data issues {5}' -f
        $report.Range('B1').Text, $report.Range('B10').Text, $report.Range('D10').Text, $report.Range('F10').Text,
        $report.Range('H10').Text, $report.Range('J10').Text)

    # The queries loaded with the absolute path. A register inside the output folder is saved as
    # a relative path, so a downloaded copy of the project refreshes as it is
    $outFolder = (Split-Path $Output -Parent).TrimEnd('\') + '\'
    if ($SourcePath.StartsWith($outFolder, [StringComparison]::OrdinalIgnoreCase)) {
        $report.Range('C4').Value2 = $SourcePath.Substring($outFolder.Length)
    }

    if (Test-Path $Output) { Remove-Item $Output }
    $wb.SaveAs($Output, 51)              # xlOpenXMLWorkbook: no macros
    $wb.Close($false)
    Write-Host "Built $Output"
}

[ExcelBusyFilter]::Register()       # needs an STA thread, the default of powershell.exe
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$failed = $false
try {
    BuildWorkbook $excel
}
catch {
    $failed = $true
    Write-Host ('Build failed: ' + $_.Exception.Message)
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
if ($failed) { exit 1 }

# Excel saves the folder of the workbook inside the file (xl/workbook.xml); remove it, then
# check the whole file for local data
try {
    & (Join-Path $PSScriptRoot 'scrub_workbook.ps1') -Path $Output
}
catch {
    Write-Host ('Scrub failed: ' + $_.Exception.Message)
    exit 1
}
