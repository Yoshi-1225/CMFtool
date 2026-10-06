# Excel Sync - copy ranges from Excel to the clipboard, one at a time.
# Protocol (stdout = JSON lines, stdin = "NEXT" after each paste):
#   {"ev":"ready"} -> {"ev":"copied","i":0,...} <- NEXT -> ... -> {"ev":"end"}
param([string]$jobFile)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8

function Say($obj) {
    [Console]::Out.WriteLine(($obj | ConvertTo-Json -Compress))
    [Console]::Out.Flush()
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
            [void]$rg.Copy()
            Say @{ ev = 'copied'; i = $i; sheet = $rg.Worksheet.Name; range = $rg.Address(0, 0) }
            $reply = [Console]::In.ReadLine()
            if ($reply -ne 'NEXT') { break }
        } catch {
            Say @{ ev = 'error'; i = $i; message = $_.Exception.Message }
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
