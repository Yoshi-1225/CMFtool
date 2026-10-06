# Excel Sync - copy ranges from Excel to the clipboard, one at a time.
# Protocol (stdout = JSON lines, stdin = one reply line per question):
#   {"ev":"ready"}
#   for each item:
#     item.cmf only: {"ev":"rows","i":0,"rows":[...]}  <- SKIP | PASS | STOP | {"head":..,"order":[..],"serial":[..],"edge":[..]}
#     {"ev":"copied","i":0,...}                        <- NEXT | STOP
#   {"ev":"end"}
# item.cmf = { num, name }: 1-based columns (inside the range) of the serial number and the object name.
# The reply to "rows" says how to reorder the rows (see cmfPlan in js/table.js). The reordered table is
# built in a temporary workbook, so the user's workbook is never modified.
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

function Copy-Edge($src, $dst, $idx) {
    try {
        $s = $src.Borders.Item($idx)
        $d = $dst.Borders.Item($idx)
        $ls = $s.LineStyle
        if ((Is-Empty $ls) -or $ls -eq -4142) { $d.LineStyle = -4142; return }
        $d.LineStyle = $ls
        $d.Weight = $s.Weight
        $d.Color = $s.Color
    } catch {}
}

# Rebuild the range in a temporary workbook with its rows in the planned order.
function Build-Ordered($xl, $rg, $plan, $numCol) {
    $cols = $rg.Columns.Count
    $order = @($plan.order); $serial = @($plan.serial); $edge = @($plan.edge)
    $n = $order.Count
    $head = [int]$plan.head

    $tmp = $xl.Workbooks.Add()
    try { $tmp.Windows.Item(1).Visible = $false } catch {}
    $ws = $tmp.Worksheets.Item(1)

    try {
        # ColumnWidth depends on the workbook's default font, so adjust until the widths in points match.
        for ($c = 1; $c -le $cols; $c++) {
            $src = $rg.Columns.Item($c)
            $dst = $ws.Columns.Item($c)
            if ($src.EntireColumn.Hidden) { $dst.Hidden = $true; continue }
            $dst.ColumnWidth = $src.ColumnWidth
            for ($k = 0; $k -lt 3 -and $dst.Width -gt 0 -and [Math]::Abs($dst.Width - $src.Width) -gt 0.3; $k++) {
                $dst.ColumnWidth = [Math]::Min(255, $dst.ColumnWidth * $src.Width / $dst.Width)
            }
        }

        # Header rows stay on top; copy them as one block so merged cells survive.
        if ($head -gt 0) {
            $blk = $rg.Worksheet.Range($rg.Cells.Item(1, 1), $rg.Cells.Item($head, $cols))
            [void]$blk.Copy($ws.Cells.Item(1, 1))
        }
        for ($i = 0; $i -lt $n; $i++) {
            $srcRow = $rg.Rows.Item([int]$order[$i] + 1)
            if ($i -ge $head) { [void]$srcRow.Copy($ws.Cells.Item($i + 1, 1)) }
            $dstRow = $ws.Rows.Item($i + 1)
            if ($srcRow.EntireRow.Hidden) { $dstRow.Hidden = $true } else { $dstRow.RowHeight = $srcRow.RowHeight }
        }

        # Formulas may point at rows that moved: keep the values Excel shows instead.
        if ($rg.HasFormula -ne $false) {
            for ($i = 0; $i -lt $n; $i++) {
                $k = [int]$order[$i] + 1
                for ($c = 1; $c -le $cols; $c++) {
                    $sc = $rg.Cells.Item($k, $c)
                    if ($sc.HasFormula -ne $true) { continue }
                    $dc = $ws.Cells.Item($i + 1, $c)
                    $v = $sc.Value2
                    try {
                        if ($v -is [int]) { $dc.Value2 = [string]$sc.Text }   # error values such as #N/A
                        elseif (Is-Empty $v) { $dc.Value2 = '' }
                        else { $dc.Value2 = $v }
                    } catch {}
                }
            }
        }

        # Serial number column = callout numbers. The cell keeps its number format (e.g. 00 -> 01).
        for ($i = 0; $i -lt $n; $i++) {
            $sv = $serial[$i]
            if (Is-Empty $sv) { continue }
            $dc = $ws.Cells.Item($i + 1, $numCol)
            if ([string]$sv -eq '') { $dc.Value2 = '' } else { $dc.Value2 = [double]$sv }
        }

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

        return @{ wb = $tmp; range = $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item($n, $cols)) }
    } catch {
        try { $tmp.Close($false) } catch {}
        throw
    }
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
        $wb = $xl.Workbooks.Open($job.path, 0, $true)
    }
    Say @{ ev = 'ready'; live = $live }

    for ($i = 0; $i -lt @($job.items).Count; $i++) {
        $it = @($job.items)[$i]
        $tmp = $null
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
                $rows = Get-Rows $rg $numCol $nameCol
                Say @{ ev = 'rows'; i = $i; sheet = $rg.Worksheet.Name; range = $rg.Address(0, 0); rows = $rows }
                $reply = [Console]::In.ReadLine()
                if ($null -eq $reply -or $reply -eq 'STOP') { break }
                if ($reply -eq 'PASS') { continue }
                if ($reply -ne 'SKIP') {
                    $plan = $reply | ConvertFrom-Json
                    $xl.ScreenUpdating = $false
                    try { $built = Build-Ordered $xl $rg $plan $numCol } finally { $xl.ScreenUpdating = $true }
                    $tmp = $built.wb
                    $copyRg = $built.range
                }
            }

            [void]$copyRg.Copy()
            Say @{ ev = 'copied'; i = $i; sheet = $rg.Worksheet.Name; range = $rg.Address(0, 0); ordered = [bool]$tmp }
            $reply = [Console]::In.ReadLine()
            if ($reply -ne 'NEXT') { break }
        } catch {
            Say @{ ev = 'error'; i = $i; message = $_.Exception.Message }
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
