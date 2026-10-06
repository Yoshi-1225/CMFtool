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

# Long Excel work: tell the panel we're still alive (it stops the script after a long silence).
$script:lastBeat = [DateTime]::Now
function Beat {
    if (([DateTime]::Now - $script:lastBeat).TotalSeconds -ge 3) {
        Say @{ ev = 'busy' }
        $script:lastBeat = [DateTime]::Now
    }
}

# Values (Value2) or formulas of one column of the range in a single call, indexed 1..rows.
function Column-Values($rg, $col, [switch]$formula) {
    $n = [int]$rg.Rows.Count
    $part = $rg.Columns.Item($col)
    # Assign directly: an array produced inside "$x = if (...) { }" would be unrolled into a flat list.
    if ($formula) { $v = $part.Formula } else { $v = $part.Value2 }
    $out = New-Object object[] ($n + 1)
    if ($n -eq 1) { $out[1] = $v } else { for ($r = 1; $r -le $n; $r++) { $out[$r] = $v[$r, 1] } }
    return ,$out
}

# Serial text, object name, hidden flag and "part of a vertical merge" for every row of the range.
function Get-Rows($rg, $numCol, $nameCol) {
    $list = New-Object System.Collections.ArrayList
    $rows = [int]$rg.Rows.Count
    $nums = Column-Values $rg $numCol
    $names = Column-Values $rg $nameCol
    for ($r = 1; $r -le $rows; $r++) {
        Beat
        $row = $rg.Rows.Item($r)
        # Use the value for numbers: a narrow column shows "##" as its text.
        $v = $nums[$r]
        if ($v -is [double] -or $v -is [string] -or (Is-Empty $v)) { $s = [string]$v } else { $s = [string]$rg.Cells.Item($r, $numCol).Text }
        $nv = $names[$r]
        if ($nv -is [string] -or (Is-Empty $nv)) { $name = [string]$nv } else { $name = [string]$rg.Cells.Item($r, $nameCol).Text }
        $vm = $false
        if ($row.MergeCells -ne $false) {
            foreach ($cell in $row.Cells) {
                if ($cell.MergeCells -and $cell.MergeArea.Rows.Count -gt 1) { $vm = $true; break }
            }
        }
        [void]$list.Add(@{ s = $s; n = $name; h = [bool]$row.EntireRow.Hidden; v = $vm })
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

function Read-Border($b) { return ,@($b.LineStyle, $b.Weight, $b.Color) }

function Write-Border($b, $e) {
    if ($null -eq $e) { return }
    if ((Is-Empty $e[0]) -or $e[0] -eq -4142) { $b.LineStyle = -4142; return }
    $b.LineStyle = $e[0]
    if (-not (Is-Empty $e[1])) { $b.Weight = $e[1] }
    if (-not (Is-Empty $e[2])) { $b.Color = $e[2] }
}

# Top (8) or bottom (9) border of one row of a table: read once for the whole row when every cell
# has the same border (the usual case), otherwise cell by cell.
function Get-RowEdge($line, $cols, $idx) {
    try {
        $e = Read-Border ($line.Borders.Item($idx))
        if (-not ((Is-Empty $e[0]) -or (Is-Empty $e[1]) -or (Is-Empty $e[2]))) { return @{ all = $e } }
        $cells = New-Object object[] $cols
        for ($c = 1; $c -le $cols; $c++) {
            try { $cells[$c - 1] = Read-Border ($line.Cells.Item(1, $c).Borders.Item($idx)) } catch {}
        }
        return @{ cells = $cells }
    } catch { Warn 'borders' $_; return $null }
}

function Set-RowEdge($line, $cols, $idx, $rec) {
    if ($null -eq $rec) { return }
    try {
        if ($rec.ContainsKey('all')) { Write-Border ($line.Borders.Item($idx)) $rec.all; return }
        for ($c = 1; $c -le $cols; $c++) {
            if ($null -ne $rec.cells[$c - 1]) { Write-Border ($line.Cells.Item(1, $c).Borders.Item($idx)) $rec.cells[$c - 1] }
        }
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
                Beat
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
                Beat
                $k = [int]$order[$i] + 1
                try { if ($rg.Rows.Item($k).HasFormula -eq $false) { continue } } catch {}
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
            Beat
            $src = $rg.Rows.Item([int]$edge[$i] + 1)
            $dst = $ws.Range($ws.Cells.Item($i + 1, 1), $ws.Cells.Item($i + 1, $cols))
            Set-RowEdge $dst $cols 8 (Get-RowEdge $src $cols 8)    # xlEdgeTop
            Set-RowEdge $dst $cols 9 (Get-RowEdge $src $cols 9)    # xlEdgeBottom
        }

        try { $out = $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item($n, $cols)) } catch { throw (Describe 'range' $_) }
        return @{ wb = $tmp; range = $out }
    } catch {
        try { $tmp.Close($false) } catch {}
        throw
    }
}

function Merge-Row($sws, $row, $m) {
    try { [void]$sws.Range($sws.Cells.Item($row, $m.c), $sws.Cells.Item($row, $m.c + $m.n - 1)).Merge() }
    catch { Warn 'merge' $_ }
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
    $formulas = Column-Values $rg $numCol -formula
    for ($k = $head; $k -lt $rows; $k++) {
        Beat
        $row = $sws.Rows.Item($row0 + $k)
        $hidden[$k] = [bool]$row.Hidden
        if ($hidden[$k]) { $row.Hidden = $false }   # sorting skips hidden rows
        $height[$k] = Num $row.RowHeight 15
        $orig[$k] = $formulas[$k + 1]
        $line = $rg.Rows.Item($k + 1)
        $edges[$k] = @((Get-RowEdge $line $cols 8), (Get-RowEdge $line $cols 9))
    }

    # Merged cells in the data rows: Excel's Sort refuses merged cells of different sizes, so unmerge
    # them and merge them again at the rows' new positions after sorting.
    $merges = New-Object System.Collections.ArrayList
    for ($k = $head; $k -lt $rows; $k++) {
        Beat
        $line = $rg.Rows.Item($k + 1)
        if ($line.MergeCells -eq $false) { continue }
        $c = 0
        while ($c -lt $cols) {
            $cell = $sws.Cells.Item($row0 + $k, $col0 + $c)
            if ($cell.MergeCells -ne $true) { $c++; continue }
            $area = $cell.MergeArea
            $w = [int]$area.Columns.Count
            if ([int]$area.Rows.Count -gt 1) { throw "merge@0: row $($row0 + $k) has cells merged across rows" }
            if ([int]$area.Row -eq ($row0 + $k) -and [int]$area.Column -eq ($col0 + $c)) {
                [void]$merges.Add(@{ k = $k; c = $col0 + $c; n = $w })
                $c += [Math]::Max(1, $w)
            } else { $c++ }
        }
    }
    $data = $sws.Range($sws.Cells.Item($r1, $col0), $sws.Cells.Item($r2, $col0 + $cols - 1))

    try {
        # Sort key = target position, written into the serial column (replaced by the real serial below).
        for ($k = $head; $k -lt $rows; $k++) { $sws.Cells.Item($row0 + $k, $sc).Value2 = [double]($target[$k] + 1) }
        if ($merges.Count) { [void]$data.UnMerge() }
        $sort = $sws.Sort
        $sort.SortFields.Clear()
        [void]$sort.SortFields.Add($sws.Range($sws.Cells.Item($r1, $sc), $sws.Cells.Item($r2, $sc)), 0, 1)
        $sort.SetRange($data)
        $sort.Header = 2           # xlNo
        $sort.MatchCase = $false
        $sort.Orientation = 1      # xlTopToBottom
        $sort.Apply()
        $sort.SortFields.Clear()
    } catch {
        $err = $_
        # Nothing moved: put the serial column, merged cells and hidden rows back.
        for ($k = $head; $k -lt $rows; $k++) {
            try { $sws.Cells.Item($row0 + $k, $sc).Formula = $orig[$k] } catch {}
            if ($hidden[$k]) { try { $sws.Rows.Item($row0 + $k).Hidden = $true } catch {} }
        }
        foreach ($m in $merges) { Merge-Row $sws ($row0 + $m.k) $m }
        throw (Describe 'sort' $err)
    }

    foreach ($m in $merges) { Merge-Row $sws ($row0 + $target[$m.k]) $m }

    for ($j = $head; $j -lt $rows; $j++) {
        Beat
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
        $line = $rg.Rows.Item($j + 1)
        Set-RowEdge $line $cols 8 $edges[$j][0]
        Set-RowEdge $line $cols 9 $edges[$j][1]
    }
}

# Sort the workbook and save it. Returns saved | unsaved | readonly | error:<message>.
# A workbook the user has open with other unsaved edits is sorted but left for the user to save.
function Save-Ordered($xl, $wb, $rg, $plan, $numCol, $live) {
    if ($wb.ReadOnly) { return 'readonly' }
    $wasSaved = [bool]$wb.Saved
    $alerts = $xl.DisplayAlerts
    try {
        $xl.ScreenUpdating = $false
        $xl.DisplayAlerts = $false
        try { $null = Write-Ordered $rg $plan $numCol }
        catch { return 'error:' + (Describe 'sort' $_) }
        if ($live -and -not $wasSaved) { return 'unsaved' }
        try { $wb.Save(); return 'saved' }
        catch { return 'error:' + (Describe 'save' $_) }
    } finally {
        try { $xl.ScreenUpdating = $true } catch {}
        try { $xl.DisplayAlerts = $alerts } catch {}
    }
}

# The workbook if it is already open in this Excel. Files synced by OneDrive / SharePoint report a URL
# (https://...) as FullName, so match those by file name: one Excel can't open two workbooks with the same name.
function Find-Open($xl, $file) {
    $name = ([string]$file -split '[\\/]')[-1]
    foreach ($w in $xl.Workbooks) {
        $full = [string]$w.FullName
        if ($full -ieq $file) { return $w }
        if ($full -match '^https?://' -and ([string]$w.Name) -ieq $name) { return $w }
    }
    return $null
}

$job = [IO.File]::ReadAllText($jobFile, [Text.Encoding]::UTF8) | ConvertFrom-Json
$xl = $null; $wb = $null; $own = $false; $live = $false

try {
    # If the workbook is already open in Excel, use that instance (includes unsaved edits).
    try { $xl = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application') } catch { $xl = $null }
    if ($xl) { $wb = Find-Open $xl $job.path }
    if ($wb) {
        $live = $true
    } else {
        $xl = New-Object -ComObject Excel.Application
        $own = $true
        $xl.Visible = $false
        $xl.DisplayAlerts = $false
        if ($job.write) { try { $wb = $xl.Workbooks.Open($job.path, 0, $false) } catch { $wb = $null } }
        # In use somewhere else: open read-only. The table can still be shown, the file just isn't sorted.
        if (-not $wb) { $wb = $xl.Workbooks.Open($job.path, 0, $true) }
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
                    if (([string]$xl.ActiveWorkbook.Name) -ieq ([string]$wb.Name)) {
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
