# Script to generate the complete Excel Workbook
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

function HexToOle($hex) {
    $hex = $hex.TrimStart('#')
    $r = [Convert]::ToInt32($hex.Substring(0,2), 16)
    $g = [Convert]::ToInt32($hex.Substring(2,2), 16)
    $b = [Convert]::ToInt32($hex.Substring(4,2), 16)
    return $r + ($g * 256) + ($b * 65536)
}

$cGreenHeader = HexToOle "38761D"
$cJamesCyan   = HexToOle "00C0F0"
$cOGGreen     = HexToOle "92D050"
$cArainSlate  = HexToOle "B4C6E7"
$cYellowTotal = HexToOle "FFC000"
$cPinkX       = HexToOle "FCE4D6"
$cRedTextX    = HexToOle "C00000"
$cDarkText    = HexToOle "000000"
$cWhite       = HexToOle "FFFFFF"
$cLightGrey   = HexToOle "F2F2F2"
$cBorder      = HexToOle "B0BEC5"
$cPrevPurple  = HexToOle "F3E5F5"
$cDarkBlack   = HexToOle "1A1A1A"

$wb = $excel.Workbooks.Add()
$wsDesktop = $wb.Worksheets.Add()
$wsSummary = $wb.Worksheets.Add()
$wsMobile = $wb.Worksheets.Item(1)

$wsMobile.Name = "Mobile Data Entry"
$wsSummary.Name = "Group Summary"
$wsDesktop.Name = "Desktop Grid View"

Write-Output "Populating Sheet 1: Mobile Data Entry..."
$wsMobile.Columns.Item(1).ColumnWidth = 5
$wsMobile.Columns.Item(2).ColumnWidth = 14
$wsMobile.Columns.Item(3).ColumnWidth = 8
$wsMobile.Columns.Item(4).ColumnWidth = 8
$wsMobile.Columns.Item(5).ColumnWidth = 8
$wsMobile.Columns.Item(6).ColumnWidth = 12
$wsMobile.Columns.Item(7).ColumnWidth = 12

# Top Banner
$banner = $wsMobile.Range("A1:G1")
$banner.Merge()
$banner.Value2 = "3rd report"
$banner.Font.Bold = $true
$banner.Font.Size = 14
$banner.Font.Color = $cWhite
$banner.Interior.Color = $cGreenHeader
$banner.HorizontalAlignment = -4108
$banner.VerticalAlignment = -4108
$wsMobile.Rows.Item(1).RowHeight = 32

$note = $wsMobile.Range("A2:G2")
$note.Merge()
$note.Value2 = "Mobile-friendly vertical entry: Tap any cell to enter 1C, 2C, 3C or 'x'. Enter 'x' for 0. Totals update automatically."
$note.Font.Size = 9
$note.Font.Italic = $true
$note.Interior.Color = $cLightGrey
$note.HorizontalAlignment = -4108
$wsMobile.Rows.Item(2).RowHeight = 20

$jamesMembers = @(
    @("Lucifer", "x", "x", "1", 8),
    @("Flash", "x", "x", "x", 7),
    @("Viking", "1", "1", "1", 8),
    @("M4", "x", "x", "x", 9),
    @("Baobei", "x", "x", "x", 8),
    @("Panda", "1", "1", "x", 9),
    @("Andrea", "x", "x", "x", 5),
    @("Gin", "x", "x", "x", 7),
    @("Jhapali", "x", "1", "x", 9),
    @("Beast", "x", "x", "1", 6),
    @("Django", "2", "x", "x", 9),
    @("Badda", "1", "1", "1", 9),
    @("Alex", "x", "x", "x", 10),
    @("Susu", "x", "x", "x", 10),
    @("Andy", "x", "x", "x", 10)
)

$ogMembers = @(
    @("Princess", "1", "x", "x", 11),
    @("Moon", "x", "x", "x", 9),
    @("Nairobi", "x", "1", "1", 9),
    @("Beloha", "4", "2", "2", 7),
    @("Devil", "1", "1", "x", 12),
    @("Terrie", "1", "x", "x", 13),
    @("Ven", "1", "x", "x", 7),
    @("Daenerys", "x", "1", "2", 7),
    @("Sunil", "x", "x", "x", 8),
    @("Parker", "x", "x", "x", 9)
)

$arainMembers = @(
    @("Arain", "1", "1", "x", 12),
    @("Jimmy", "x", "x", "x", 9),
    @("Delowar", "x", "1", "x", 8)
)

function WriteGroupMobile($ws, $startRow, $groupName, $headerColor, $members) {
    # Title
    $title = $ws.Range("A" + $startRow + ":G" + $startRow)
    $title.Merge()
    $title.Value2 = "GROUP: " + $groupName
    $title.Font.Bold = $true
    $title.Font.Size = 12
    $title.Font.Color = $cDarkBlack
    $title.Interior.Color = $headerColor
    $title.HorizontalAlignment = -4108
    $ws.Rows.Item($startRow).RowHeight = 26

    # Headers
    $headerRow = $startRow + 1
    $headers = @("#", "Name", "1C", "2C", "3C", "Prev Total", "Current Sum")
    for ($i = 0; $i -lt $headers.Count; $i++) {
        $c = $ws.Cells.Item($headerRow, $i + 1)
        $c.Value2 = $headers[$i]
        $c.Font.Bold = $true
        $c.Font.Size = 10
        $c.HorizontalAlignment = -4108
        $c.Interior.Color = HexToOle "ECEFF1"
    }
    $ws.Rows.Item($headerRow).RowHeight = 22

    $firstData = $headerRow + 1
    $cur = $firstData
    $idx = 1
    foreach ($m in $members) {
        $ws.Rows.Item($cur).RowHeight = 24

        $c0 = $ws.Cells.Item($cur, 1)
        $c0.Value2 = [string]$idx
        $c0.HorizontalAlignment = -4108
        $c0.Font.Bold = $true

        $cName = $ws.Cells.Item($cur, 2)
        $cName.Value2 = $m[0]
        $cName.HorizontalAlignment = -4131
        $cName.Font.Bold = $true

        # 1C, 2C, 3C
        for ($col = 0; $col -lt 3; $col++) {
            $val = $m[$col + 1]
            $cell = $ws.Cells.Item($cur, $col + 3)
            $cell.Value2 = $val
            $cell.HorizontalAlignment = -4108
            $cell.Font.Bold = $true
            if ($val -eq "x") {
                $cell.Interior.Color = $cPinkX
                $cell.Font.Color = $cRedTextX
            } else {
                $cell.Interior.Color = $cWhite
                $cell.Font.Color = HexToOle "0D47A1"
            }
        }

        # Prev Total
        $cPrev = $ws.Cells.Item($cur, 6)
        $cPrev.Value2 = [int]$m[4]
        $cPrev.HorizontalAlignment = -4108
        $cPrev.Font.Bold = $true
        $cPrev.Interior.Color = $cPrevPurple

        # Member Current Sum
        $cSum = $ws.Cells.Item($cur, 7)
        $cSum.Formula = "=SUM(C" + $cur + ":E" + $cur + ")"
        $cSum.HorizontalAlignment = -4108
        $cSum.Font.Bold = $true
        $cSum.Font.Color = HexToOle "0070C0"

        $cur++
        $idx++
    }
    $lastData = $cur - 1

    # Group Total Row
    $totRow = $cur
    $ws.Rows.Item($totRow).RowHeight = 26
    $ws.Cells.Item($totRow, 2).Value2 = $groupName + " TOTAL"
    $ws.Cells.Item($totRow, 2).Font.Bold = $true
    $ws.Cells.Item($totRow, 2).HorizontalAlignment = -4152

    $ws.Cells.Item($totRow, 3).Formula = "=SUM(C" + $firstData + ":C" + $lastData + ")"
    $ws.Cells.Item($totRow, 4).Formula = "=SUM(D" + $firstData + ":D" + $lastData + ")"
    $ws.Cells.Item($totRow, 5).Formula = "=SUM(E" + $firstData + ":E" + $lastData + ")"
    $ws.Cells.Item($totRow, 6).Formula = "=SUM(F" + $firstData + ":F" + $lastData + ")"
    $ws.Cells.Item($totRow, 7).Formula = "=SUM(G" + $firstData + ":G" + $lastData + ")"

    $tRange = $ws.Range("A" + $totRow + ":G" + $totRow)
    $tRange.Interior.Color = $cYellowTotal
    $tRange.Font.Bold = $true
    $tRange.Font.Size = 11
    $tRange.HorizontalAlignment = -4108
    $ws.Cells.Item($totRow, 2).HorizontalAlignment = -4152

    $grid = $ws.Range("A" + $headerRow + ":G" + $totRow)
    $grid.Borders.LineStyle = 1
    $grid.Borders.Color = $cBorder

    return @{ TotalRow = $totRow; FirstRow = $firstData; LastRow = $lastData }
}

$jamesRes = WriteGroupMobile $wsMobile 4 "JAMES" $cJamesCyan $jamesMembers
$ogRes = WriteGroupMobile $wsMobile ($jamesRes.TotalRow + 2) "OG" $cOGGreen $ogMembers
$arainRes = WriteGroupMobile $wsMobile ($ogRes.TotalRow + 2) "ARAIN" $cArainSlate $arainMembers

Write-Output "Populating Sheet 2: Group Summary Report..."
$wsSummary.Columns.Item(1).ColumnWidth = 4
$wsSummary.Columns.Item(2).ColumnWidth = 16
$wsSummary.Columns.Item(3).ColumnWidth = 12
$wsSummary.Columns.Item(4).ColumnWidth = 12
$wsSummary.Columns.Item(5).ColumnWidth = 12
$wsSummary.Columns.Item(6).ColumnWidth = 16
$wsSummary.Columns.Item(7).ColumnWidth = 16

# Summary Header Banner
$sBanner = $wsSummary.Range("B2:G2")
$sBanner.Merge()
$sBanner.Value2 = "3rd report - Group Summary Report"
$sBanner.Font.Bold = $true
$sBanner.Font.Size = 14
$sBanner.Font.Color = $cWhite
$sBanner.Interior.Color = $cGreenHeader
$sBanner.HorizontalAlignment = -4108
$wsSummary.Rows.Item(2).RowHeight = 32

$sNote = $wsSummary.Range("B3:G3")
$sNote.Merge()
$sNote.Value2 = "Automated summary table. Fully linked to mobile entries."
$sNote.Font.Size = 9
$sNote.Font.Italic = $true
$sNote.Interior.Color = $cLightGrey
$sNote.HorizontalAlignment = -4108
$wsSummary.Rows.Item(3).RowHeight = 20

# Summary Headers
$sumHeaders = @("Group Name", "Total 1C", "Total 2C", "Total 3C", "Current Total", "Previous Total")
for ($i = 0; $i -lt $sumHeaders.Count; $i++) {
    $c = $wsSummary.Cells.Item(5, $i + 2)
    $c.Value2 = $sumHeaders[$i]
    $c.Font.Bold = $true
    $c.Font.Size = 10
    $c.Font.Color = $cWhite
    $c.Interior.Color = HexToOle "263238"
    $c.HorizontalAlignment = -4108
}
$wsSummary.Rows.Item(5).RowHeight = 26

$groupsData = @(
    @{ Name = "JAMES"; Row = 6; SrcTot = $jamesRes.TotalRow; Color = $cJamesCyan },
    @{ Name = "OG";    Row = 7; SrcTot = $ogRes.TotalRow;    Color = $cOGGreen },
    @{ Name = "ARAIN"; Row = 8; SrcTot = $arainRes.TotalRow; Color = $cArainSlate }
)

foreach ($g in $groupsData) {
    $r = $g.Row
    $src = $g.SrcTot

    $cName = $wsSummary.Cells.Item($r, 2)
    $cName.Value2 = $g.Name
    $cName.Font.Bold = $true
    $cName.Interior.Color = $g.Color
    $cName.HorizontalAlignment = -4108

    $c1 = $wsSummary.Cells.Item($r, 3)
    $c1.Formula = "='Mobile Data Entry'!C" + $src
    $c1.Font.Bold = $true
    $c1.HorizontalAlignment = -4108

    $c2 = $wsSummary.Cells.Item($r, 4)
    $c2.Formula = "='Mobile Data Entry'!D" + $src
    $c2.Font.Bold = $true
    $c2.HorizontalAlignment = -4108

    $c3 = $wsSummary.Cells.Item($r, 5)
    $c3.Formula = "='Mobile Data Entry'!E" + $src
    $c3.Font.Bold = $true
    $c3.HorizontalAlignment = -4108

    $cTot = $wsSummary.Cells.Item($r, 6)
    $cTot.Formula = "=C" + $r + " + D" + $r + " + E" + $r
    $cTot.Font.Bold = $true
    $cTot.Interior.Color = HexToOle "E1F5FE"
    $cTot.Font.Color = HexToOle "01579B"
    $cTot.HorizontalAlignment = -4108

    $cPrev = $wsSummary.Cells.Item($r, 7)
    $cPrev.Formula = "='Mobile Data Entry'!F" + $src
    $cPrev.Font.Bold = $true
    $cPrev.Interior.Color = $cPrevPurple
    $cPrev.HorizontalAlignment = -4108

    $wsSummary.Rows.Item($r).RowHeight = 24
}

# Grand Total Row
$gRow = 9
$cGrand = $wsSummary.Cells.Item($gRow, 2)
$cGrand.Value2 = "GRAND TOTAL"
$cGrand.Font.Bold = $true
$cGrand.HorizontalAlignment = -4108

$wsSummary.Cells.Item($gRow, 3).Formula = "=SUM(C6:C8)"
$wsSummary.Cells.Item($gRow, 4).Formula = "=SUM(D6:D8)"
$wsSummary.Cells.Item($gRow, 5).Formula = "=SUM(E6:E8)"
$wsSummary.Cells.Item($gRow, 6).Formula = "=SUM(F6:F8)"
$wsSummary.Cells.Item($gRow, 7).Formula = "=SUM(G6:G8)"

$grandRange = $wsSummary.Range("B" + $gRow + ":G" + $gRow)
$grandRange.Interior.Color = $cYellowTotal
$grandRange.Font.Bold = $true
$grandRange.Font.Size = 11
$grandRange.HorizontalAlignment = -4108
$wsSummary.Rows.Item($gRow).RowHeight = 28

$sumTable = $wsSummary.Range("B5:G" + $gRow)
$sumTable.Borders.LineStyle = 1
$sumTable.Borders.Color = $cBorder

Write-Output "Populating Sheet 3: Desktop Grid View (Replicating exact screenshot)..."
$dBanner = $wsDesktop.Range("B1:S1")
$dBanner.Merge()
$dBanner.Value2 = "3rd report"
$dBanner.Font.Bold = $true
$dBanner.Font.Size = 13
$dBanner.Font.Color = $cWhite
$dBanner.Interior.Color = $cGreenHeader
$dBanner.HorizontalAlignment = -4108
$wsDesktop.Rows.Item(1).RowHeight = 26

function WriteGroupDesktop($ws, $startCol, $groupName, $headerColor, $members) {
    $cEnd = $startCol + 5
    $gRange = $ws.Range($ws.Cells.Item(2, $startCol), $ws.Cells.Item(2, $cEnd))
    $gRange.Merge()
    $gRange.Value2 = $groupName
    $gRange.Font.Bold = $true
    $gRange.Font.Size = 11
    $gRange.Interior.Color = $headerColor
    $gRange.HorizontalAlignment = -4108

    $colTitles = @("#", "Name", "1C", "2C", "3C", "Total")
    for ($i = 0; $i -lt $colTitles.Count; $i++) {
        $c = $ws.Cells.Item(3, $startCol + $i)
        $c.Value2 = $colTitles[$i]
        $c.Font.Bold = $true
        $c.HorizontalAlignment = -4108
        if ($i -eq 0) {
            $c.Interior.Color = HexToOle "000000"
            $c.Font.Color = $cWhite
        } else {
            $c.Interior.Color = HexToOle "D9D9D9"
        }
    }
    $ws.Rows.Item(3).RowHeight = 20

    $maxRows = 15
    for ($mIdx = 0; $mIdx -lt $maxRows; $mIdx++) {
        $r = 4 + $mIdx
        $ws.Rows.Item($r).RowHeight = 19

        $cNo = $wsDesktop.Cells.Item($r, $startCol)
        $cNo.Value2 = [string]($mIdx + 1)
        $cNo.Font.Bold = $true
        $cNo.HorizontalAlignment = -4108
        $cNo.Interior.Color = HexToOle "000000"
        $cNo.Font.Color = $cWhite

        if ($mIdx -lt $members.Count) {
            $m = $members[$mIdx]
            $cName = $ws.Cells.Item($r, $startCol + 1)
            $cName.Value2 = $m[0]
            $cName.Font.Bold = $true
            $cName.HorizontalAlignment = -4131

            for ($col = 0; $col -lt 3; $col++) {
                $val = $m[$col + 1]
                $cell = $ws.Cells.Item($r, $startCol + 2 + $col)
                $cell.Value2 = $val
                $cell.HorizontalAlignment = -4108
                $cell.Font.Bold = $true
                if ($val -eq "x") {
                    $cell.Interior.Color = $cPinkX
                    $cell.Font.Color = $cRedTextX
                }
            }

            $cTot = $ws.Cells.Item($r, $startCol + 5)
            $cTot.Value2 = [int]$m[4]
            $cTot.Font.Bold = $true
            $cTot.HorizontalAlignment = -4108
            $cTot.Interior.Color = $cPrevPurple
        } else {
            for ($k = 1; $k -le 5; $k++) {
                $ws.Cells.Item($r, $startCol + $k).Value2 = ""
            }
        }
    }

    $totRow = 19
    $ws.Rows.Item($totRow).RowHeight = 22

    $col1Letter = [char](64 + $startCol + 2)
    $col2Letter = [char](64 + $startCol + 3)
    $col3Letter = [char](64 + $startCol + 4)
    $colTotLetter = [char](64 + $startCol + 5)

    $ws.Cells.Item($totRow, $startCol + 2).Formula = "=SUM(" + $col1Letter + "4:" + $col1Letter + "18)"
    $ws.Cells.Item($totRow, $startCol + 3).Formula = "=SUM(" + $col2Letter + "4:" + $col2Letter + "18)"
    $ws.Cells.Item($totRow, $startCol + 4).Formula = "=SUM(" + $col3Letter + "4:" + $col3Letter + "18)"
    $ws.Cells.Item($totRow, $startCol + 5).Formula = "=SUM(" + $colTotLetter + "4:" + $colTotLetter + "18)"

    $tRange = $ws.Range($ws.Cells.Item($totRow, $startCol), $ws.Cells.Item($totRow, $cEnd))
    $tRange.Interior.Color = $cYellowTotal
    $tRange.Font.Bold = $true
    $tRange.Font.Size = 11
    $tRange.HorizontalAlignment = -4108

    $fullTable = $ws.Range($ws.Cells.Item(3, $startCol), $ws.Cells.Item($totRow, $cEnd))
    $fullTable.Borders.LineStyle = 1
    $fullTable.Borders.Color = HexToOle "7F7F7F"

    $ws.Columns.Item($startCol).ColumnWidth = 4
    $ws.Columns.Item($startCol + 1).ColumnWidth = 11
    $ws.Columns.Item($startCol + 2).ColumnWidth = 5
    $ws.Columns.Item($startCol + 3).ColumnWidth = 5
    $ws.Columns.Item($startCol + 4).ColumnWidth = 5
    $ws.Columns.Item($startCol + 5).ColumnWidth = 7
}

WriteGroupDesktop $wsDesktop 2 "JAMES" $cJamesCyan $jamesMembers
$wsDesktop.Columns.Item(8).ColumnWidth = 2
WriteGroupDesktop $wsDesktop 9 "OG" $cOGGreen $ogMembers
$wsDesktop.Columns.Item(15).ColumnWidth = 2
WriteGroupDesktop $wsDesktop 16 "ARAIN" $cArainSlate $arainMembers

$wsMobile.Activate()

$outputPath = Join-Path (Get-Location) "Report_Data_Entry_and_Summary.xlsx"
if (Test-Path $outputPath) { Remove-Item $outputPath -Force }

$wb.SaveAs($outputPath, 51)
$wb.Close($false)
$excel.Quit()

[System.Runtime.InteropServices.Marshal]::ReleaseComObject($wsMobile) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($wsSummary) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($wsDesktop) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

Write-Output "SUCCESS: Workbook generated at: $outputPath"
