<#
.SYNOPSIS
    Removes the local folder path that Excel saves inside an .xlsx file and sets the Power Query
    locale to en-US, then checks that no local path, computer name or Windows user name is left.

.DESCRIPTION
    Excel writes the folder a workbook was saved in to xl/workbook.xml (x15ac:absPath), and a
    CELL("filename") formula keeps its last result, the full path of the file, in the sheet XML.
    Both would publish a local folder, often one with the Windows user name in it.
    This script deletes absPath and empties the saved result of every CELL("filename") formula;
    Excel calculates such a formula again when the workbook is opened.
    Power Query also writes the regional format of the build PC into the query package
    (Config/Package.xml, <Culture>). The queries do not use it, because every text-to-number
    and text-to-date conversion names en-US, so the script sets it to en-US: the file then
    shows no trace of the build PC's locale.
    It edits the file in place and needs no Excel; the file must not be open in Excel.
    build_report.ps1 runs it after every build.

    The check then reads every part of the file as UTF-8, Latin-1 and UTF-16 text. The Power
    Query part (DataMashup in customXml) is base64 with a zip package inside, so it is decoded
    and the query code, permissions and metadata are checked too. The script stops with an
    error when a part contains the user profile folder, the folder of the workbook or the
    computer name, or when the author fields hold the Windows user name; for example when
    Source file is a full path in the user folder. The workbook is saved all the same.

    The DataMashup also ends with a checksum protected with the key of the Windows account
    (DPAPI, the Permission Bindings). It holds no name or path and is kept as Excel wrote it.
    After the locale change it no longer matches the package on any account, the build account
    included, so Excel ignores it everywhere, as it would on any other Windows account. The
    workbook does not depend on it: it does not ignore privacy levels, so Refresh All works the
    same way.

.PARAMETER Path
    The .xlsx file to clean.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File tools\scrub_workbook.ps1 ContractsReport.xlsx
#>
param(
    [Parameter(Mandatory = $true)][string]$Path
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression

$Path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
if (-not (Test-Path $Path -PathType Leaf)) { throw "Workbook not found: $Path" }

# Excel writes its XML parts as UTF-8 without a byte order mark, customXml as UTF-16
$Utf8 = New-Object Text.UTF8Encoding($false)
$Latin1 = [Text.Encoding]::GetEncoding(28591)

# <mc:AlternateContent><mc:Choice Requires="x15"><x15ac:absPath url="C:\...\"/></mc:Choice></mc:AlternateContent>
$AbsPathPattern = '<mc:AlternateContent[^>]*>\s*<mc:Choice Requires="x15">\s*<x15ac:absPath [^>]*/>\s*</mc:Choice>\s*</mc:AlternateContent>'
# A formula cell with CELL("filename") followed by its saved result
$CellFilenamePattern = '(<c [^>]*>\s*<f[^>]*>[^<]*CELL\((?:"|&quot;)filename(?:"|&quot;)[^<]*</f>\s*)<v>[^<]*</v>'

# What must not be in the file
$folder = Split-Path $Path -Parent
$profileSlash = ([string]$env:USERPROFILE).Replace('\', '/')
$folderNeedle = if ($folder.Length -gt 3) { $folder } else { '' }    # not a bare drive such as C:\
$Needles = @()
foreach ($n in @(
        @('user profile folder', $env:USERPROFILE),
        @('user profile folder', $profileSlash),
        @('workbook folder', $folderNeedle),
        @('computer name', $env:COMPUTERNAME))) {
    if ($n[1]) { $Needles += New-Object PSObject -Property @{ Name = $n[0]; Value = $n[1] } }
}
$found = New-Object Collections.Generic.List[string]

function OpenZip([string]$mode) {
    $access = if ($mode -eq 'Read') { 'Read' } else { 'ReadWrite' }
    $stream = [IO.File]::Open($Path, 'Open', $access, 'None')
    return New-Object IO.Compression.ZipArchive($stream, [IO.Compression.ZipArchiveMode]::$mode)
}

function ReadBytes($entry) {
    $s = $entry.Open()
    try {
        $m = New-Object IO.MemoryStream
        $s.CopyTo($m)
        return , $m.ToArray()
    } finally { $s.Dispose() }
}

# Text of a part, with the encoding and byte order mark to write it back the same way
function ReadText([byte[]]$b) {
    if ($b.Length -ge 2 -and $b[0] -eq 0xFF -and $b[1] -eq 0xFE) { $enc = [Text.Encoding]::Unicode; $bom = 2 }
    elseif ($b.Length -ge 3 -and $b[0] -eq 0xEF -and $b[1] -eq 0xBB -and $b[2] -eq 0xBF) { $enc = $Utf8; $bom = 3 }
    else { $enc = $Utf8; $bom = 0 }
    return @{ Encoding = $enc; Bom = $bom; Text = $enc.GetString($b, $bom, $b.Length - $bom) }
}

# Records every needle found in the bytes; a zip inside (the Power Query package) is opened
# and each of its compressed entries checked as well
function CheckBytes([string]$label, [byte[]]$b, [int]$depth) {
    $texts = @($Utf8.GetString($b), $Latin1.GetString($b), [Text.Encoding]::Unicode.GetString($b))
    if ($b.Length -gt 1) { $texts += [Text.Encoding]::Unicode.GetString($b, 1, $b.Length - 1) }
    foreach ($n in $Needles) {
        foreach ($t in $texts) {
            if ($t.IndexOf($n.Value, [StringComparison]::OrdinalIgnoreCase) -ge 0) { $found.Add("${label}: $($n.Name)"); break }
        }
    }
    if ($depth -lt 3 -and $b.Length -ge 4 -and $b[0] -eq 0x50 -and $b[1] -eq 0x4B -and ($b[2] -eq 3 -or $b[2] -eq 5)) {
        try {
            $zip = New-Object IO.Compression.ZipArchive((New-Object IO.MemoryStream(, $b)), [IO.Compression.ZipArchiveMode]::Read)
            try { foreach ($e in $zip.Entries) { CheckBytes "$label > $($e.FullName)" (ReadBytes $e) ($depth + 1) } }
            finally { $zip.Dispose() }
        } catch { $found.Add("${label}: the zip inside cannot be read, so it is not checked") }
    }
}

# DataMashup (MS-QDEFF): a 4-byte version, then package (zip), permissions, metadata and
# permission bindings, each after its 4-byte length. Metadata is a version, the metadata XML
# and a zip, again each after its length.
function SplitMashup([byte[]]$b) {
    $parts = @()
    $o = 4
    foreach ($name in @('package', 'permissions', 'metadata', 'permission bindings')) {
        if ($o + 4 -gt $b.Length) { throw "it ends before the $name" }
        $len = [BitConverter]::ToInt32($b, $o)
        $o += 4
        if ($len -lt 0 -or $o + $len -gt $b.Length) { throw "the $name length is out of range" }
        $part = New-Object byte[] $len
        [Array]::Copy($b, $o, $part, 0, $len)
        $parts += , @($name, $part)
        $o += $len
        if ($name -eq 'metadata' -and $len -ge 8) {
            $xmlEnd = 8 + [BitConverter]::ToInt32($part, 4)
            if ($xmlEnd -ge 8 -and $xmlEnd + 4 -le $len) {
                $zipLen = [BitConverter]::ToInt32($part, $xmlEnd)
                if ($zipLen -gt 0 -and $xmlEnd + 4 + $zipLen -le $len) {
                    $content = New-Object byte[] $zipLen
                    [Array]::Copy($part, $xmlEnd + 4, $content, 0, $zipLen)
                    $parts += , @('metadata content', $content)
                }
            }
        }
    }
    return , $parts
}

# The locale written into the query package instead of the build PC's one
$MashupCulture = 'en-US'
$CulturePattern = '<Culture>[^<]*</Culture>'
$MashupPattern = '(<DataMashup\b[^>]*>)([^<]*)(</DataMashup>)'

# Bytes of a changed text part, with the byte order mark of the original
function TextBytes([byte[]]$original, $part, [string]$text) {
    $body = $part.Encoding.GetBytes($text)
    $out = New-Object byte[] ($part.Bom + $body.Length)
    [Array]::Copy($original, $out, $part.Bom)
    [Array]::Copy($body, 0, $out, $part.Bom, $body.Length)
    return , $out
}

# Culture written in Config/Package.xml of a query package (zip bytes)
function PackageCulture([byte[]]$pkg) {
    $zip = New-Object IO.Compression.ZipArchive((New-Object IO.MemoryStream(, $pkg)), [IO.Compression.ZipArchiveMode]::Read)
    try {
        $entry = $zip.GetEntry('Config/Package.xml')
        if (-not $entry) { throw 'the query package has no Config/Package.xml' }
        $m = [regex]::Match((ReadText (ReadBytes $entry)).Text, '<Culture>([^<]*)</Culture>')
        if ($m.Success) { return $m.Groups[1].Value } else { return '' }
    } finally { $zip.Dispose() }
}

# DataMashup bytes with the package culture set to $MashupCulture, or $null when it already is.
# The package is written again as a new zip with the same entries in the same order; the
# permissions, metadata and permission bindings after it keep their bytes.
function SetMashupCulture([byte[]]$b) {
    if ($b.Length -lt 8) { throw 'the DataMashup is too short' }
    $pkgLen = [BitConverter]::ToInt32($b, 4)
    if ($pkgLen -lt 0 -or 8 + $pkgLen -gt $b.Length) { throw 'the package length is out of range' }
    $pkg = New-Object byte[] $pkgLen
    [Array]::Copy($b, 8, $pkg, 0, $pkgLen)
    if ((PackageCulture $pkg) -ceq $MashupCulture) { return $null }

    $outStream = New-Object IO.MemoryStream
    $src = New-Object IO.Compression.ZipArchive((New-Object IO.MemoryStream(, $pkg)), [IO.Compression.ZipArchiveMode]::Read)
    $dst = New-Object IO.Compression.ZipArchive($outStream, [IO.Compression.ZipArchiveMode]::Create, $true)
    try {
        foreach ($e in $src.Entries) {
            $data = ReadBytes $e
            if ($e.FullName -eq 'Config/Package.xml') {
                $part = ReadText $data
                $text = [regex]::Replace($part.Text, $CulturePattern, "<Culture>$MashupCulture</Culture>")
                $data = TextBytes $data $part $text
            }
            $n = $dst.CreateEntry($e.FullName, [IO.Compression.CompressionLevel]::Optimal)
            $n.LastWriteTime = $e.LastWriteTime
            $s = $n.Open()
            try { $s.Write($data, 0, $data.Length) } finally { $s.Dispose() }
        }
    } finally { $dst.Dispose(); $src.Dispose() }

    $newPkg = $outStream.ToArray()
    $rest = $b.Length - 8 - $pkgLen
    $out = New-Object byte[] (8 + $newPkg.Length + $rest)
    [Array]::Copy($b, 0, $out, 0, 4)
    [Array]::Copy([BitConverter]::GetBytes([int]$newPkg.Length), 0, $out, 4, 4)
    [Array]::Copy($newPkg, 0, $out, 8, $newPkg.Length)
    [Array]::Copy($b, 8 + $pkgLen, $out, 8 + $newPkg.Length, $rest)
    return , $out
}

# Pass 1, read only: the new bytes of the parts that change
$changed = [ordered]@{}
$zip = OpenZip 'Read'
try {
    foreach ($entry in $zip.Entries) {
        $name = $entry.FullName
        $isBook = $name -eq 'xl/workbook.xml'
        $isSheet = $name -like 'xl/worksheets/*.xml'
        $isCustom = $name -like 'customXml/*.xml'
        if (-not ($isBook -or $isSheet -or $isCustom)) { continue }
        $bytes = ReadBytes $entry
        $part = ReadText $bytes
        $new = $part.Text
        if ($isBook) { $new = [regex]::Replace($new, $AbsPathPattern, '') }
        if ($isSheet) { $new = [regex]::Replace($new, $CellFilenamePattern, '$1<v/>') }
        if ($isCustom) {
            $new = [regex]::Replace($new, $MashupPattern, [Text.RegularExpressions.MatchEvaluator] {
                    param($m)
                    $mashup = SetMashupCulture ([Convert]::FromBase64String($m.Groups[2].Value))
                    if ($null -eq $mashup) { return $m.Value }
                    return $m.Groups[1].Value + [Convert]::ToBase64String($mashup) + $m.Groups[3].Value
                })
        }
        if ($new -cne $part.Text) { $changed[$name] = TextBytes $bytes $part $new }
    }
} finally { $zip.Dispose() }

# Pass 2: rewrite only those parts; the other parts keep their bytes, all keep their time stamp
if ($changed.Count -gt 0) {
    $zip = OpenZip 'Update'
    try {
        foreach ($name in $changed.Keys) {
            $entry = $zip.GetEntry($name)
            $stamp = $entry.LastWriteTime
            $stream = $entry.Open()
            try {
                $stream.SetLength(0)
                $bytes = $changed[$name]
                $stream.Write($bytes, 0, $bytes.Length)
            } finally { $stream.Dispose() }
            $entry.LastWriteTime = $stamp
        }
    } finally { $zip.Dispose() }
    Write-Host ('Removed the local folder path or set the query locale in: ' + (@($changed.Keys) -join ', '))
} else {
    Write-Host 'No local folder path to remove; the query locale is already en-US'
}

# Pass 3: check the saved file, every part
$count = 0
$zip = OpenZip 'Read'
try {
    foreach ($entry in $zip.Entries) {
        $name = $entry.FullName
        $bytes = ReadBytes $entry
        $count++
        CheckBytes $name $bytes 0
        if ($name -like 'customXml/*.xml') {
            foreach ($m in [regex]::Matches((ReadText $bytes).Text, '<DataMashup\b[^>]*>([^<]*)</DataMashup>')) {
                try { $mashup = [Convert]::FromBase64String($m.Groups[1].Value); $parts = SplitMashup $mashup }
                catch { $found.Add("${name}: the DataMashup cannot be decoded, so it is not checked ($($_.Exception.Message))"); continue }
                foreach ($p in $parts) { CheckBytes "$name > DataMashup $($p[0])" $p[1] 0 }
                $culture = PackageCulture $parts[0][1]
                if ($culture -cne $MashupCulture) { $found.Add("${name}: the query locale is '$culture', not $MashupCulture") }
            }
        }
        if ($name -eq 'docProps/core.xml' -and $env:USERNAME) {
            $author = '<(dc:creator|cp:lastModifiedBy)>\s*' + [regex]::Escape($env:USERNAME) + '\s*</'
            if ((ReadText $bytes).Text -match $author) { $found.Add("${name}: Windows user name as author") }
        }
    }
} finally { $zip.Dispose() }

if ($found.Count -gt 0) {
    throw ("Local data left in ${Path}:`n  " + ((@($found) | Select-Object -Unique) -join "`n  ") +
        "`nThe workbook is saved, but do not publish it like this.")
}
Write-Host "Checked $count parts: no local path, computer name or Windows user name"
