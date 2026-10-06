# Excel Sync - copy ranges from Excel to the clipboard, one at a time.
# Protocol (stdout = JSON lines, stdin = one reply line per question):
#   {"ev":"ready"}
#   for each item:
#     item.cmf only: {"ev":"rows","i":0,"rows":[...]}  <- SKIP | PASS | STOP | {"head":..,"order":[..],"serial":[..],"edge":[..]}
#                                                        | {"write":{"head":..,"order":[..],"serial":[..]}}
#        after a "write" reply: {"ev":"rows","i":0,"after":true,"write":"saved|unsaved|readonly|error:..","rows":[...]}
#                                                     <- SKIP | PASS | STOP | {"head":..,...}
#     {"ev":"copied","i":0,...}                        <- NEXT | STOP
#   {"ev":"end"}
# item.cmf = { num, name }: 1-based columns (inside the range) of the serial number and the object name.
# A "write" reply sorts the rows of the workbook itself (job.write = open it writable) and saves it.
# Any other plan reorders only the copy: it is built in a temporary workbook.
param([string]$jobFile)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8

function Say($obj) {
    [Console]::Out.WriteLine(($obj | ConvertTo-Json -Compress -Depth 6))
    [Console]::Out.Flush()
}

function Is-Empty($v) { return ($null -eq $v) -or ($v -is [DBNull]) }

# Serial text, object name, hidden flag and "part of a vertical merge" for every row of the range.
function Get-Rows($rg, $numCol, $nameCol) {
    $list = New-Object System.Collections.ArrayList
    $rows = $rg.Rows.Count
    for ($r = 1; $r -le $rows; $r++) {
        $row = $rg.Rows.Item($r)
        $sc = $rg.Cells.Item($r, $numCol)
        $v = $sc.Value2
        # Use the value for numbers: a narrow column shows "##" as its text.
        if ($v -is [double]) { $s = [string]$v } else { $s = [string]$sc.Text }
        $vm = $false
        if ($row.MergeCells -ne $false) {
            foreach ($cell in $row.Cells) {
                if ($cell.MergeCells -and $cell.MergeArea.Rows.Count -gt 1) { $vm = $true; break }
            }
        }
        [void]$list.Add(@{ s = $s; n = [string]$rg.Cells.Item($r, $nameCol).Text; h = [bool]$row.EntireRow.Hidden; v = $vm })
    }
    return ,$list
}

# First problem per step, reported to the panel as "step@line: message".
$script:warned = @{}
function Describe($step, $err) {
    $msg = $err.Exception.Message
    if ($msg -match '^\w+@\d+: ') { return $msg }
    return "$step@$($err.InvocationInfo.ScriptLineNumber): $msg"
}
function Warn($step, $err) {
    if (-not $script:warned.ContainsKey($step)) { $script:warned[$step] = (Describe $step $err) }
}

# Number from a COM property. Excel returns DBNull when the value is mixed (e.g. merged cells),
# and assigning DBNull to another property fails with "Specified cast is not valid".
function Num($v, $default) {
    if (Is-Empty $v) { return $default }
    return [double]$v
}

function Copy-Edge($src, $dst, $idx) {
    try {
        $s = $src.Borders.Item($idx)
        $d = $dst.Borders.Item($idx)
        $ls = $s.LineStyle
        if ((Is-Empty $ls) -or $ls -eq -4142) { $d.LineStyle = -4142; return }
        $d.LineStyle = $ls
        if (-not (Is-Empty $s.Weight)) { $d.Weight = $s.Weight }
        if (-not (Is-Empty $s.Color)) { $d.Color = $s.Color }
    } catch { Warn 'borders' $_ }
}

# Rebuild the range in a temporary workbook with its rows in the planned order.
# Sizes are read from whole rows and columns of the source sheet: they never come back mixed.
function Build-Ordered($xl, $rg, $plan, $numCol) {
    $sws = $rg.Worksheet
    $row0 = [int]$rg.Row; $col0 = [int]$rg.Column
    $cols = [int]$rg.Columns.Count
    $order = @($plan.order); $serial = @($plan.serial); $edge = @($plan.edge)
    $n = $order.Count
    $head = [int]$plan.head

    try { $tmp = $xl.Workbooks.Add() } catch { throw (Describe 'workbook' $_) }
    try { $tmp.Windows.Item(1).Visible = $false } catch {}
    $ws = $tmp.Worksheets.Item(1)

    try {
        # Column widths in points. ColumnWidth depends on the workbook's default font, so scale until they match.
        for ($c = 1; $c -le $cols; $c++) {
            try {
                $src = $sws.Columns.Item($col0 + $c - 1)
                $dst = $ws.Columns.Item($c)
                if ($src.Hidden) { $dst.Hidden = $true; continue }
                $want = Num $src.Width 0
                for ($k = 0; $k -lt 4; $k++) {
                    $have = Num $dst.Width 0
                    if ($have -le 0 -or [Math]::Abs($have - $want) -le 0.3) { break }
                    $dst.ColumnWidth = [Math]::Min([double]255, (Num $dst.ColumnWidth 8.43) * $want / $have)
                }
            } catch { Warn 'widths' $_ }
        }

        # Header rows stay on top; copy them as one block so merged cells survive.
        try {
            if ($head -gt 0) {
                $blk = $sws.Range($sws.Cells.Item($row0, $col0), $sws.Cells.Item($row0 + $head - 1, $col0 + $cols - 1))
                [void]$blk.Copy($ws.Cells.Item(1, 1))
            }
            for ($i = $head; $i -lt $n; $i++) {
                [void]$rg.Rows.Item([int]$order[$i] + 1).Copy($ws.Cells.Item($i + 1, 1))
            }
        } catch { throw (Describe 'rows' $_) }

        for ($i = 0; $i -lt $n; $i++) {
            try {
                $src = $sws.Rows.Item($row0 + [int]$order[$i])
                $dst = $ws.Rows.Item($i + 1)
                if ($src.Hidden) { $dst.Hidden = $true } else { $dst.RowHeight = Num $src.RowHeight 15 }
            } catch { Warn 'heights' $_ }
        }

        # Formulas may point at rows that moved: keep the values Excel shows instead.
        $hasFormula = $true
        try { $hasFormula = $rg.HasFormula -ne $false } catch {}
        if ($hasFormula) {
            for ($i = 0; $i -lt $n; $i++) {
                $k = [int]$order[$i] + 1
                for ($c = 1; $c -le $cols; $c++) {
                    try {
                        $sc = $rg.Cells.Item($k, $c)
                        if ($sc.HasFormula -ne $true) { continue }
                        $dc = $ws.Cells.Item($i + 1, $c)
                        $v = $sc.Value2
                        if ($v -is [int]) { $dc.Value2 = [string]$sc.Text }   # error values such as #N/A
                        elseif (Is-Empty $v) { $dc.Value2 = '' }
                        else { $dc.Value2 = $v }
                    } catch { Warn 'formulas' $_ }
                }
            }
        }

        # Serial number column = callout numbers. The cell keeps its number format (e.g. 00 -> 01).
        try {
            for ($i = 0; $i -lt $n; $i++) {
                $sv = $serial[$i]
                if (Is-Empty $sv) { continue }
                $dc = $ws.Cells.Item($i + 1, $numCol)
                if ([string]$sv -eq '') { $dc.Value2 = '' } else { $dc.Value2 = [double]$sv }
            }
        } catch { throw (Describe 'serial' $_) }

        # Top/bottom borders stay where they were, so a thick bottom border stays at the bottom.
        for ($i = $head; $i -lt $n; $i++) {
            if (Is-Empty $edge[$i]) { continue }
            $p = [int]$edge[$i] + 1
            for ($c = 1; $c -le $cols; $c++) {
                $sc = $rg.Cells.Item($p, $c)
                $dc = $ws.Cells.Item($i + 1, $c)
                Copy-Edge $sc $dc 8   # xlEdgeTop
                Copy-Edge $sc $dc 9   # xlEdgeBottom
            }
        }

        try { $out = $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item($n, $cols)) } catch { throw (Describe 'range' $_) }
        return @{ wb = $tmp; range = $out }
    } catch {
        try { $tmp.Close($false) } catch {}
        throw
    }
}

function Get-Edge($cell, $idx) {
    try {
        $b = $cell.Borders.Item($idx)
        return ,@($b.LineStyle, $b.Weight, $b.Color)
    } catch { return $null }
}

function Set-Edge($cell, $idx, $e) {
    if ($null -eq $e) { return }
    try {
        $b = $cell.Borders.Item($idx)
        if ((Is-Empty $e[0]) -or $e[0] -eq -4142) { $b.LineStyle = -4142; return }
        $b.LineStyle = $e[0]
        if (-not (Is-Empty $e[1])) { $b.Weight = $e[1] }
        if (-not (Is-Empty $e[2])) { $b.Color = $e[2] }
    } catch { Warn 'borders' $_ }
}

# Sort the rows of the user's workbook into the planned order with Excel's own Sort (like sorting by hand),
# then write the serial numbers. Row heights and hidden rows travel with their rows; top/bottom borders
# stay at their positions, so a thick bottom border stays at the bottom.
function Write-Ordered($rg, $plan, $numCol) {
    $sws = $rg.Worksheet
    $row0 = [int]$rg.Row; $col0 = [int]$rg.Column
    $cols = [int]$rg.Columns.Count; $rows = [int]$rg.Rows.Count
    $order = @($plan.order); $serial = @($plan.serial)
    $head = [int]$plan.head
    if ($order.Count -ne $rows) { throw "sort@0: $($order.Count) rows planned, the range has $rows" }
    if ($rows -le $head) { return }
    $r1 = $row0 + $head; $r2 = $row0 + $rows - 1
    $sc = $col0 + $numCol - 1

    $target = @{}
    for ($j = $head; $j -lt $rows; $j++) { $target[[int]$order[$j]] = $j }

    $hidden = @{}; $height = @{}; $orig = @{}; $edges = @{}
    for ($k = $head; $k -lt $rows; $k++) {
        $row = $sws.Rows.Item($row0 + $k)
        $hidden[$k] = [bool]$row.Hidden
        if ($hidden[$k]) { $row.Hidden = $false }   # sorting skips hidden rows
        $height[$k] = Num $row.RowHeight 15
        $orig[$k] = $sws.Cells.Item($row0 + $k, $sc).Formula
        for ($c = 0; $c -lt $cols; $c++) {
            $cell = $sws.Cells.Item($row0 + $k, $col0 + $c)
            $edges["$k,$c"] = @((Get-Edge $cell 8), (Get-Edge $cell 9))
        }
    }

    try {
        # Sort key = target position, written into the serial column (replaced by the real serial below).
        for ($k = $head; $k -lt $rows; $k++) { $sws.Cells.Item($row0 + $k, $sc).Value2 = [double]($target[$k] + 1) }
        $sort = $sws.Sort
        $sort.SortFields.Clear()
        [void]$sort.SortFields.Add($sws.Range($sws.Cells.Item($r1, $sc), $sws.Cells.Item($r2, $sc)), 0, 1)
        $sort.SetRange($sws.Range($sws.Cells.Item($r1, $col0), $sws.Cells.Item($r2, $col0 + $cols - 1)))
        $sort.Header = 2           # xlNo
        $sort.MatchCase = $false
        $sort.Orientation = 1      # xlTopToBottom
        $sort.Apply()
        $sort.SortFields.Clear()
    } catch {
        $err = $_
        # Nothing moved: put the serial column and hidden rows back.
        for ($k = $head; $k -lt $rows; $k++) {
            try { $sws.Cells.Item($row0 + $k, $sc).Formula = $orig[$k] } catch {}
            if ($hidden[$k]) { try { $sws.Rows.Item($row0 + $k).Hidden = $true } catch {} }
        }
        throw (Describe 'sort' $err)
    }

    for ($j = $head; $j -lt $rows; $j++) {
        $k = [int]$order[$j]
        try {
            $cell = $sws.Cells.Item($row0 + $j, $sc)
            $sv = $serial[$j]
            if (Is-Empty $sv) { $cell.Formula = $orig[$k] }
            elseif ([string]$sv -eq '') { $cell.Value2 = '' }
            else { $cell.Value2 = [double]$sv }
        } catch { throw (Describe 'serial' $_) }
        $row = $sws.Rows.Item($row0 + $j)
        try { $row.RowHeight = $height[$k] } catch { Warn 'heights' $_ }
        if ($hidden[$k]) { try { $row.Hidden = $true } catch { Warn 'heights' $_ } }
        for ($c = 0; $c -lt $cols; $c++) {
            $cell = $sws.Cells.Item($row0 + $j, $col0 + $c)
            $e = $edges["$j,$c"]
            Set-Edge $cell 8 $e[0]
            Set-Edge $cell 9 $e[1]
        }
    }
}

# Sort the workbook and save it. Returns saved | unsaved | readonly | error:<message>.
# A workbook the user has open with other unsaved edits is sorted but left for the user to save.
function Save-Ordered($xl, $wb, $rg, $plan, $numCol, $live) {
    if ($wb.ReadOnly) { return 'readonly' }
    $wasSaved = [bool]$wb.Saved
    $xl.ScreenUpdating = $false
    try { $null = Write-Ordered $rg $plan $numCol }
    catch { return 'error:' + (Describe 'sort' $_) }
    finally { $xl.ScreenUpdating = $true }
    if ($live -and -not $wasSaved) { return 'unsaved' }
    $alerts = $xl.DisplayAlerts
    try { $xl.DisplayAlerts = $false; $wb.Save(); return 'saved' }
    catch { return 'error:' + (Describe 'save' $_) }
    finally { try { $xl.DisplayAlerts = $alerts } catch {} }
}

$job = [IO.File]::ReadAllText($jobFile, [Text.Encoding]::UTF8) | ConvertFrom-Json
$xl = $null; $wb = $null; $own = $false; $live = $false

try {
    # If the workbook is already open in Excel, use that instance (includes unsaved edits).
    try { $xl = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application') } catch { $xl = $null }
    if ($xl) {
        foreach ($w in $xl.Workbooks) {
            if ($w.FullName -ieq $job.path) { $wb = $w; break }
        }
    }
    if ($wb) {
        $live = $true
    } else {
        $xl = New-Object -ComObject Excel.Application
        $own = $true
        $xl.Visible = $false
        $xl.DisplayAlerts = $false
        $wb = $xl.Workbooks.Open($job.path, 0, (-not $job.write))
    }
    Say @{ ev = 'ready'; live = $live }

    for ($i = 0; $i -lt @($job.items).Count; $i++) {
        $it = @($job.items)[$i]
        $tmp = $null
        $script:warned = @{}
        try {
            $rg = $null
            # No range given: use what is selected in Excel right now, like copy/paste.
            if (-not $it.range -and $live) {
                try {
                    if ($xl.ActiveWorkbook.FullName -ieq $job.path) {
                        $sel = $xl.Selection
                        [void]$sel.Address(0, 0)
                        $rg = $sel
                    }
                } catch { $rg = $null }
            }
            if (-not $rg) {
                $sheetName = if ($it.range) { $it.sheet } else { $it.fallbackSheet }
                $address = if ($it.range) { $it.range } else { $it.fallbackRange }
                $ws = if ($sheetName) { $wb.Worksheets.Item($sheetName) } else { $wb.Worksheets.Item(1) }
                $rg = $ws.Range($address)
            }

            $copyRg = $rg
            if ($it.cmf) {
                $numCol = [int]$it.cmf.num; $nameCol = [int]$it.cmf.name
                $after = $false; $written = $null; $stop = $false; $pass = $false
                while ($true) {
                    $rows = Get-Rows $rg $numCol $nameCol
                    $msg = @{ ev = 'rows'; i = $i; sheet = $rg.Worksheet.Name; range = $rg.Address(0, 0); rows = $rows; after = $after }
                    if ($after) { $msg.write = $written }
                    Say $msg
                    $reply = [Console]::In.ReadLine()
                    if ($null -eq $reply -or $reply -eq 'STOP') { $stop = $true; break }
                    if ($reply -eq 'PASS') { $pass = $true; break }
                    if ($reply -eq 'SKIP') { break }
                    $plan = $reply | ConvertFrom-Json
                    if ($plan.write -and -not $after) {
                        # Sort the workbook itself, then ask again how to show the (now sorted) rows.
                        $written = Save-Ordered $xl $wb $rg $plan.write $numCol $live
                        $after = $true
                        continue
                    }
                    $xl.ScreenUpdating = $false
                    try { $built = Build-Ordered $xl $rg $plan $numCol } finally { $xl.ScreenUpdating = $true }
                    $tmp = $built.wb
                    $copyRg = $built.range
                    break
                }
                if ($stop) { break }
                if ($pass) { continue }
            }

            [void]$copyRg.Copy()
            Say @{ ev = 'copied'; i = $i; sheet = $rg.Worksheet.Name; range = $rg.Address(0, 0); ordered = ($null -ne $tmp);
                   warn = @($script:warned.Values) }
            $reply = [Console]::In.ReadLine()
            if ($reply -ne 'NEXT') { break }
        } catch {
            Say @{ ev = 'error'; i = $i; message = (Describe 'copy' $_) }
        } finally {
            if ($tmp) {
                try { $xl.CutCopyMode = $false } catch {}
                try { $tmp.Close($false) } catch {}
            }
        }
    }
} catch {
    Say @{ ev = 'fatal'; message = $_.Exception.Message }
} finally {
    if ($xl) { try { $xl.CutCopyMode = $false } catch {} }
    if ($own -and $wb) { try { $wb.Close($false) } catch {} }
    if ($own -and $xl) { try { $xl.Quit() } catch {} }
    if ($xl) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) } catch {} }
    Say @{ ev = 'end' }
}
