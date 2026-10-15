# Excel Sync - write texts edited in Illustrator back into the workbook (after Illustrator saves, or the panel's write-back button).
# Input: job file (UTF-8 JSON) { path, items: [{ ref, sheet, cell, kind, value, raw }] }
#   kind: num = a number (the cell keeps its number format) | typed = as if typed in Excel (dates, TRUE/FALSE)
#         | text = literal text (never turned into a number, date or formula)
#   raw:  the cell's value in the saved file. A cell is only written while it still holds that value, so edits
#         made in Excel and not saved yet are never overwritten. No raw (error values like #N/A) = not compared.
# Output (stdout, one JSON line):
#   {"ev":"done","live":true|false,"save":"saved|unsaved|readonly|nothing|error:..",
#    "results":[{"ref":"B3","status":"written|formula|changed|error","text":"..","message":".."}]}
#     changed: Excel holds something else (text = what it shows); formula: the cell has a formula, left alone
#   {"ev":"fatal","message":".."}
param([string]$jobFile)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8

function Say($obj) {
    [Console]::Out.WriteLine(($obj | ConvertTo-Json -Compress -Depth 6))
    [Console]::Out.Flush()
}

function Is-Empty($v) { return ($null -eq $v) -or ($v -is [DBNull]) -or (($v -is [string]) -and $v -eq '') }

# The workbook if it is already open in this Excel (same as excel-copy.ps1). Files synced by OneDrive /
# SharePoint report a URL as FullName, so match those by file name.
function Find-Open($xl, $file) {
    $name = ([string]$file -split '[\\/]')[-1]
    foreach ($w in $xl.Workbooks) {
        $full = [string]$w.FullName
        if ($full -ieq $file) { return $w }
        if ($full -match '^https?://' -and ([string]$w.Name) -ieq $name) { return $w }
    }
    return $null
}

# Does the cell (Value2) still hold the value of the saved file?
function Same-Value($v, $raw) {
    if ($null -eq $raw) { return (Is-Empty $v) }
    if ($raw -is [bool]) { return ($v -is [bool]) -and ($v -eq $raw) }
    if ($raw -is [string]) { return ($v -is [string]) -and ($v -ceq $raw) }
    try { return ($v -is [double]) -and ([double]$v -eq [double]$raw) } catch { return $false }
}

# Write the items into the open workbook. Returns @{ results = @(...); wrote = <cells written> }
function Write-Items($wb, $items) {
    $results = @()
    $wrote = 0
    foreach ($it in @($items)) {
        $r = @{ ref = [string]$it.ref }
        try {
            $ws = if ($it.sheet) { $wb.Worksheets.Item([string]$it.sheet) } else { $wb.Worksheets.Item(1) }
            $rg = $ws.Range([string]$it.cell)
            $hasRaw = @($it.PSObject.Properties.Name) -contains 'raw'
            if ($rg.HasFormula) {
                $r.status = 'formula'
            } elseif ($hasRaw -and -not (Same-Value $rg.Value2 $it.raw)) {
                $r.status = 'changed'
                $r.text = [string]$rg.Text
            } else {
                $v = [string]$it.value
                if ($it.kind -eq 'num') { $rg.Value2 = [double]$it.value }
                elseif ($it.kind -eq 'typed') { $rg.Value2 = $v }
                elseif ([string]$rg.NumberFormat -eq '@') { $rg.Value2 = $v }
                else { $rg.Value2 = "'" + $v }      # the apostrophe keeps it literal text, it isn't part of the value
                $r.status = 'written'
                $wrote++
            }
        } catch {
            $r.status = 'error'
            $r.message = $_.Exception.Message
        }
        $results += $r
    }
    return @{ results = $results; wrote = $wrote }
}

# Dot-sourced (tests): only the functions above
if ($MyInvocation.InvocationName -eq '.') { return }

$job = [IO.File]::ReadAllText($jobFile, [Text.Encoding]::UTF8) | ConvertFrom-Json
$xl = $null; $wb = $null; $own = $false; $live = $false; $alerts = $null; $updating = $null

try {
    # If the workbook is already open in Excel, write into that one (the user sees the change right away).
    try { $xl = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application') } catch { $xl = $null }
    if ($xl) { $wb = Find-Open $xl $job.path }
    if ($wb) {
        $live = $true
    } else {
        $xl = New-Object -ComObject Excel.Application
        $own = $true
        $xl.Visible = $false
        $xl.DisplayAlerts = $false
        # In use somewhere else: Excel opens it read-only and nothing is written.
        $wb = $xl.Workbooks.Open($job.path, 0, $false)
    }

    $results = @()
    $save = 'nothing'
    if ($wb.ReadOnly) {
        $save = 'readonly'
    } else {
        $wasSaved = [bool]$wb.Saved
        $alerts = $xl.DisplayAlerts
        $updating = $xl.ScreenUpdating
        $xl.DisplayAlerts = $false
        $xl.ScreenUpdating = $false
        $out = Write-Items $wb $job.items
        $results = $out.results
        if ($out.wrote) {
            # Open with other unsaved edits of the user: written, but saving is left to them (like the CMF sort).
            if ($live -and -not $wasSaved) { $save = 'unsaved' }
            else {
                try { $wb.Save(); $save = 'saved' }
                catch { $save = 'error:' + $_.Exception.Message }
            }
        }
    }
    Say @{ ev = 'done'; live = $live; save = $save; results = @($results) }
} catch {
    Say @{ ev = 'fatal'; message = $_.Exception.Message }
} finally {
    if ($xl -and $null -ne $updating) { try { $xl.ScreenUpdating = $updating } catch {} }
    if ($xl -and $null -ne $alerts) { try { $xl.DisplayAlerts = $alerts } catch {} }
    if ($own -and $wb) { try { $wb.Close($false) } catch {} }
    if ($own -and $xl) { try { $xl.Quit() } catch {} }
    if ($xl) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) } catch {} }
}
