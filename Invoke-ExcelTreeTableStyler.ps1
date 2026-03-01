
[CmdletBinding()]
param()

class Rectangle {
    # 0-origin
    [int]$Row
    [int]$Column

    [int]$RowsCount
    [int]$ColumnsCount

    Rectangle([int]$Row, [int]$Column, [int]$RowsCount, [int]$ColumnsCount) {
        $this.Row = $Row
        $this.Column = $Column
        $this.RowsCount = $RowsCount
        $this.ColumnsCount = $ColumnsCount
    }

    [string] ToString() {
        return "<Rectangle: {Row=$($this.Row); Column=$($this.Column); RowsCount=$($this.RowsCount); ColumnsCount=$($this.ColumnsCount)}>"
    }
}

function Invoke-DumpBooleanTable {
    param(
        [Parameter(Mandatory=$true)]
        [bool[,]]$Table,

        [string]$TableName
    )

    $header = if ($TableName) { "Table Name: ${TableName}" } else { "Table:" }
    Write-Host $header
    for ($rowIndex = 0; $rowIndex -lt $Table.GetLength(0); $rowIndex++) {
        $sb = [System.Text.StringBuilder]::new()
        $sb.AppendFormat("  Row {0:00}: ", $rowIndex) | Out-Null
        for ($colIndex = 0; $colIndex -lt $Table.GetLength(1); $colIndex++) {
            if ($Table[$rowIndex, $colIndex]) {
                $sb.Append("t") | Out-Null
            } else {
                $sb.Append("_") | Out-Null
            }
        }
        Write-Host $sb.ToString()
    }
}

function Get-TreeRectangleList {
    param(
        [Parameter(Mandatory=$true)]
        [bool[,]]$Table
    )

    $treeRectangleList = [System.Collections.Generic.List[Object]]::new()

    $tableRowsCount = $Table.GetLength(0)
    $tableColumnsCount = $Table.GetLength(1)

    $rowIndex = 0
    $rowsCount = 0

    # First row is always the top-left of a tree rectangle.
    $treeRectangle = [Rectangle]::new($rowIndex, 0, -1, $tableColumnsCount)
    $rowIndex++
    $rowsCount++

    while ($rowIndex -lt $tableRowsCount) {
        if ($Table[$rowIndex, 0]) {
            $treeRectangle.RowsCount = $rowsCount
            $treeRectangleList.Add($treeRectangle)

            $rowsCount = 0
            $treeRectangle = [Rectangle]::new($rowIndex, 0, -1, $tableColumnsCount)
        }

        $rowIndex++
        $rowsCount++
    }
    $treeRectangle.RowsCount = $rowsCount
    $treeRectangleList.Add($treeRectangle)

    return $treeRectangleList
}

class Tree {
    [int]$Id
    [int]$RootTreeIndex
    [Tree]$ParentTree
    [int]$Depth
    [Rectangle]$TreeRectangle
    [System.Collections.Generic.List[Object]]$Children

    Tree([int]$RootTreeIndex, [Tree]$ParentTree, [Rectangle]$TreeRectangle, [int]$Depth) {
        $this.RootTreeIndex = $RootTreeIndex
        $this.ParentTree = $ParentTree
        $this.TreeRectangle = $TreeRectangle
        $this.Depth = $Depth
        $this.Children = [System.Collections.Generic.List[Object]]::new()
    }

    AddChild([Tree]$child) {
        $this.Children.Add($child)
    }

    [string] ToString() {
        return "<Tree: {RootTreeIndex=$($this.RootTreeIndex); Depth=$($this.Depth); TreeRectangle=$($this.TreeRectangle.ToString()); ChildrenCount=$($this.Children.Count)}>"
    }
}

function Invoke-DumpTree {
    param(
        [Parameter(Mandatory=$true)]
        [Tree]$Tree,

        [string]$TreeName
    )

    if ($TreeName) {
        $header = "Tree Name: ${TreeName}"
        Write-Host $header
    }

    $indent = " " * ($Tree.Depth + 1)
    $line = $indent + $Tree.ToString()
    Write-Host $line
    foreach ($childTree in $Tree.Children) {
        Invoke-DumpTree -Tree $childTree
    }
}

class TreeFactory {
    [bool[,]]$Table
    TreeFactory([bool[,]]$Table) {
        $this.Table = $Table
    }

    [bool] FindNextSiblingTreeTopLeft([Rectangle]$TreeRectangle, [ref]$RowIndexRef, $ColumnIndex) {
        $rowIndex = $RowIndexRef.Value
        while ($rowIndex -lt $TreeRectangle.RowsCount) {
            $tableRowIndex = $TreeRectangle.Row + $rowIndex
            $tableColumnIndex = $TreeRectangle.Column + $ColumnIndex
            if ($this.Table[$tableRowIndex, $tableColumnIndex]) {
                $RowIndexRef.Value = $rowIndex
                return $true
            }
            $rowIndex++
        }
        $RowIndexRef.Value = -1
        return $false
    }

    [bool] FindNextTreeTopLeft([Rectangle]$TreeRectangle, [int]$HeaderColumnsCount, [ref]$RowIndexRef, [ref]$ColumnIndexRef) {
        $rowIndex = $RowIndexRef.Value
        $columnIndex = $ColumnIndexRef.Value
        while ($columnIndex -lt $HeaderColumnsCount) {
            while ($rowIndex -lt $TreeRectangle.RowsCount) {
                $tableRowIndex = $TreeRectangle.Row + $rowIndex
                $tableColumnIndex = $TreeRectangle.Column + $columnIndex
                if ($this.Table[$tableRowIndex, $tableColumnIndex]) {
                    $RowIndexRef.Value = $rowIndex
                    $ColumnIndexRef.Value = $columnIndex
                    return $true
                }
                $rowIndex++
            }
            $rowIndex = 0
            $columnIndex++
        }
        $RowIndexRef.Value = $rowIndex
        $ColumnIndexRef.Value = $columnIndex
        return $false
    }

    [bool] FindFirstTreeTopLeft([Rectangle]$TreeRectangle, [int]$HeaderColumnsCount, [ref]$RowIndexRef, [ref]$ColumnIndexRef) {
        $RowIndexRef.Value = 0
        $ColumnIndexRef.Value = 1
        return $this.FindNextTreeTopLeft($TreeRectangle, $HeaderColumnsCount, $RowIndexRef, $ColumnIndexRef)
    }

    [Tree] CreateTree([int]$RootTreeIndex, [Tree]$ParentTree, [Rectangle]$TreeRectangle, [int]$Depth, [int]$HeaderColumnsCount) {
        $Tree = [Tree]::new($RootTreeIndex, $ParentTree, $TreeRectangle, $Depth)

        $rowIndexRef = [ref]$null
        $columnIndexRef = [ref]$null
        if ($this.FindFirstTreeTopLeft($TreeRectangle, $HeaderColumnsCount, $rowIndexRef, $columnIndexRef)) {
            $rowIndex1  = $rowIndexRef.Value
            $columnIndex1 = $columnIndexRef.Value

            $childTreeHeaderColumnsCount = $HeaderColumnsCount - $columnIndex1

            $tableRowIndex = $TreeRectangle.Row + $rowIndex1
            $tableColumnIndex = $TreeRectangle.Column + $columnIndex1
            $columnsCount = $TreeRectangle.ColumnsCount - $columnIndex1
            $childTreeRectangle = [Rectangle]::new($tableRowIndex, $tableColumnIndex, -1, $columnsCount)

            $rowIndexRef = [ref]($rowIndex1 + 1)
            while ($this.FindNextSiblingTreeTopLeft($TreeRectangle, $rowIndexRef, $columnIndex1)) {
                $rowIndex2 = $rowIndexRef.Value

                $rowsCount = $rowIndex2 - $rowIndex1
                $childTreeRectangle.RowsCount = $rowsCount
                $childTree = $this.CreateTree($RootTreeIndex, $Tree, $childTreeRectangle, $Depth + 1, $childTreeHeaderColumnsCount)
                $Tree.AddChild($childTree)

                $tableRowIndex = $TreeRectangle.Row + $rowIndex2
                $tableColumnIndex = $TreeRectangle.Column + $columnIndex1
                $columnsCount = $TreeRectangle.ColumnsCount - $columnIndex1
                $childTreeRectangle = [Rectangle]::new($tableRowIndex, $tableColumnIndex, -1, $columnsCount)

                $rowIndex1 = $rowIndexRef.Value
                $rowIndexRef = [ref]($rowIndex1 + 1)
            }

            $rowsCount = ($TreeRectangle.Row + $TreeRectangle.RowsCount) - $childTreeRectangle.Row
            $childTreeRectangle.RowsCount = $rowsCount
            $childTree = $this.CreateTree($RootTreeIndex, $Tree, $childTreeRectangle, $Depth + 1, $childTreeHeaderColumnsCount)
            $Tree.AddChild($childTree)
        }

        return $Tree
    }
}

class StylerStrategy {
    StylerStrategy() {}
    BeforeStyle() {}
    Style([Tree]$Tree, [Rectangle]$Rectangle) {
        throw "NotImplementedException"
    }
}

class PrintStrategy : StylerStrategy {
    PrintStrategy() : base() {}
    Style([Tree]$Tree, [Rectangle]$Rectangle) {
        Write-Host "Style(): Tree=$($Tree.ToString()), Rectangle=$($Rectangle.ToString())"
    }
}

class ExcelTreeTableStylerException : System.Exception {
    ExcelTreeTableStylerException([string]$Message) : base($Message) {}
}

class FillSectionAndDrawBordersStrategy : StylerStrategy {
    $ExcelApplication
    $Workbook
    $Worksheet
    $HeaderColumnsCount
    $SectionDepthMax

    $SectionHeaderColor = "#BDD7EE"
    $HeaderColor = "#DDEBF7"

    FillSectionAndDrawBordersStrategy($ExcelApplication, $Workbook, $Worksheet, [int]$HeaderColumnsCount, [int]$SectionDepthMax) : base() {
        $this.ExcelApplication = $ExcelApplication
        $this.Workbook = $Workbook
        $this.Worksheet = $Worksheet
        $this.HeaderColumnsCount = $HeaderColumnsCount
        $this.SectionDepthMax = $SectionDepthMax
    }
    DrawRectangleBorder($Range) {
        # Constants for Excel's Borders.Item() method
        $xlEdgeLeft = 7
        $xlEdgeTop = 8
        $xlEdgeBottom = 9
        $xlEdgeRight = 10
        $edgeIndices = @($xlEdgeLeft, $xlEdgeTop, $xlEdgeBottom, $xlEdgeRight)

        # Line style
        $xlContinuous = 1

        # Line weight
        $xlThin = 2

        foreach ($edgeIndex in $edgeIndices) {
            $Range.Borders.Item($edgeIndex).LineStyle = $xlContinuous
            $Range.Borders.Item($edgeIndex).Weight = $xlThin
        }
    }
    DrawRectangleGrid($Range) {
        # Constants for Excel's Borders.Item() method
        $xlEdgeLeft = 7
        $xlEdgeTop = 8
        $xlEdgeBottom = 9
        $xlEdgeRight = 10
        $xlInsideVertical = 11
        $xlInsideHorizontal = 12
        $edgeIndices = @($xlEdgeLeft, $xlEdgeTop, $xlEdgeBottom, $xlEdgeRight, $xlInsideVertical, $xlInsideHorizontal)

        # Line style
        $xlContinuous = 1

        # Line weight
        $xlThin = 2

        foreach ($edgeIndex in $edgeIndices) {
            $Range.Borders.Item($edgeIndex).LineStyle = $xlContinuous
            $Range.Borders.Item($edgeIndex).Weight = $xlThin
        }
    }
    [int] RGB($r, $g, $b) {
        return $r + ($g * 256) + ($b * 256 * 256)
    }
    [int] HexToRGB([string]$hex) {
        if ($hex[0] -ne "#") {
            throw "Invalid hex color format. Expected format: #RRGGBB"
        }
        if ($hex.Length -ne 7) {
            throw "Invalid hex color format. Expected format: #RRGGBB"
        }

        $rString = $hex.Substring(1, 2)
        $gString = $hex.Substring(3, 2)
        $bString = $hex.Substring(5, 2)

        $r = [Convert]::ToInt32($rString, 16)
        $g = [Convert]::ToInt32($gString, 16)
        $b = [Convert]::ToInt32($bString, 16)

        return $this.RGB($r, $g, $b)
    }
    FillRectangle($Range, $Color) {
        # Fill pattern
        $xlSolid = 1
        $Range.Interior.Pattern = $xlSolid

        # Color
        $Range.Interior.Color = $Color
    }
    ClearFillRectangle($Range) {
        $xlPatternNone = -4142

        $Range.Interior.Pattern = $xlPatternNone
    }
    ClearFillAndBorders($Range) {
        $this.ClearFillRectangle($Range)

        # Constants for Excel's Borders.Item() method
        $xlEdgeLeft = 7
        $xlEdgeTop = 8
        $xlEdgeBottom = 9
        $xlEdgeRight = 10
        $xlInsideVertical = 11
        $xlInsideHorizontal = 12
        $xlLineStyleNone = -4142
        $edgeIndices = @($xlEdgeLeft, $xlEdgeTop, $xlEdgeBottom, $xlEdgeRight, $xlInsideVertical, $xlInsideHorizontal)

        foreach ($edgeIndex in $edgeIndices) {
            $Range.Borders.Item($edgeIndex).LineStyle = $xlLineStyleNone
        }
    }
    [Tree] GetRootTree([Tree]$Tree) {
        $currentTree = $Tree
        while ($null -ne $currentTree.ParentTree) {
            $currentTree = $currentTree.ParentTree
        }
        return $currentTree
    }
    BeforeStyle([Tree]$Tree, [Rectangle]$Rectangle) {
        Write-Debug "BeforeStyle(): Tree=$($Tree.ToString()), Rectangle=$($Rectangle.ToString())"

        # Calculate the rectangle for the header columns.
        $cellTop = $Rectangle.Row
        $cellLeft = $Rectangle.Column
        $cellBottom = $Rectangle.Row + $Rectangle.RowsCount - 1
        $cellRight = $Rectangle.Column + $Rectangle.ColumnsCount - 1

        # Get the range.
        $range1 = $this.CreateRange($cellTop, $cellLeft, $cellBottom, $cellRight)

        # Draw the border of the rectangle.
        $this.ClearFillAndBorders($range1)
    }
    [object] CreateRange([int]$Top, [int]$Left, [int]$Bottom, [int]$Right) {
        if ($Top -gt $Bottom) {
            throw [ExcelTreeTableStylerException]::new("Invalid rectangle: Top ($Top) is greater than Bottom ($Bottom).")
        }
        if ($Left -gt $Right) {
            throw [ExcelTreeTableStylerException]::new("Invalid rectangle: Left ($Left) is greater than Right ($Right).")
        }
        $cellTopLeft = $this.Worksheet.Cells($Top, $Left)
        $cellBottomRight = $this.Worksheet.Cells($Bottom, $Right)
        return $this.Worksheet.Range($cellTopLeft, $cellBottomRight)
    }
    Style([Tree]$Tree, [Rectangle]$Rectangle) {
        Write-Debug "Style(): Tree=$($Tree.ToString()), Rectangle=$($Rectangle.ToString())"
        $headerColumunsWidthInRectanble = [Math]::Max(0, $this.HeaderColumnsCount - $Tree.TreeRectangle.Column)

        # Calculate the rectangle for the header columns.
        $cellTop = $Rectangle.Row
        $cellLeft = $Rectangle.Column
        $cellBottom = $Rectangle.Row + $Rectangle.RowsCount - 1
        $cellRight = $Rectangle.Column + $Rectangle.ColumnsCount - 1

        # Get the range.
        $range1 = $this.CreateRange($cellTop, $cellLeft, $cellBottom, $cellRight)

        # Draw the border of the rectangle.
        $this.DrawRectangleBorder($range1)

        if ($Tree.Depth -le $this.SectionDepthMax) {
            # If the tree depth is less than or equal to the section depth max, treat it as a section.

            # Fill the rectangle with the section header color.
            $color = $this.HexToRGB($this.SectionHeaderColor)
            $this.FillRectangle($range1, $color)
        } else {
            # If not a section, treat it as a normal tree.

            # Fill the rectangle with the header color.
            $color = $this.HexToRGB($this.HeaderColor)
            $this.FillRectangle($range1, $color)

            # Draw border in the header area.
            $this.DrawRectangleBorder($range1)

            # Draw grid in the table body area.
            $cellBodyLeft = $cellLeft + $headerColumunsWidthInRectanble
            $hasBodyArea = $cellBodyLeft -le $cellRight
            if ($hasBodyArea) {
                $range2 = $this.CreateRange($cellTop, $cellBodyLeft, $cellBottom, $cellRight)

                # Draw the border of the body area.
                $this.DrawRectangleGrid($range2)

                # Clear the fill of the table body area.
                # This is necessary because the body area may be filled with the header color, and we want to keep it white.
                $this.ClearFillRectangle($range2)
            }
        }
    }
}

function Invoke-ExcelTreeTableStyler {
    param(
        [Parameter(Mandatory=$true)]
        $RowFromTableToExcelOffset,

        [Parameter(Mandatory=$true)]
        $ColumnFromTableToExcelOffset,

        [Parameter(Mandatory=$true)]
        [Tree]$Tree,

        [Parameter(Mandatory=$true)]
        [StylerStrategy]$StylerStrategy
    )

    $jobQueue = [System.Collections.Generic.Queue[Object]]::new()
    $jobQueue.Enqueue($Tree)

    $isFirstJob = $true

    while ($jobQueue.Count -gt 0) {
        $currentTree = $jobQueue.Dequeue()
        Write-Host "Processing tree: $($currentTree.ToString())"

        $excelRow = $currentTree.TreeRectangle.Row + $RowFromTableToExcelOffset
        $excelColumn = $currentTree.TreeRectangle.Column + $ColumnFromTableToExcelOffset

        [Rectangle]$rectangle = [Rectangle]::new(
            $excelRow, $excelColumn,
            $currentTree.TreeRectangle.RowsCount, $currentTree.TreeRectangle.ColumnsCount)

        if ($isFirstJob) {
            $isFirstJob = $false
            $StylerStrategy.BeforeStyle($currentTree, $rectangle)
        }

        $StylerStrategy.Style($currentTree, $rectangle)

        foreach ($childTree in $currentTree.Children) {
            $jobQueue.Enqueue($childTree)
        }
    }
}

if ($MyInvocation.InvocationName -ne '.') {
    # not dot-sourced, execute main process

    function Invoke-TestExcelProcessesIsLessThanTwo {
        # Check if there are more or equal to two Excel processes running.
        # If there are, prompt the user to stop them and exit the script.
        $excelProcesses = @(Get-Process -Name "Excel" -ErrorAction SilentlyContinue)
        if ($excelProcesses.Count -ge 2) {
            Write-Host "Error: Excel processes found."
            Write-Host ""
            Write-Host "Multiple processes are running; therefore, the operation will be terminated."
            Write-Host "Please run it again with only one process."
            Write-Host ""
            Write-Host "If you wish to force-close all processes, please execute the following."
            Write-Host "WARNING: DATA WILL NOT BE SAVED." -ForegroundColor Red
            Write-Host ""
            Write-Host "  PS > Get-Process -Name `"Excel`" -ErrorAction SilentlyContinue | Stop-Process -Force"
            Write-Host ""
            exit 1
        }
    }

    function Invoke-AttachExcelOrCreateExcel {
        param()

        Write-Host "Trying to get an active Excel application..."
        try {
            $excel = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
            Write-Host "Active Excel application found."
            return $excel
        } catch {
            Write-Host "No active Excel application found. Trying to create a new one..."
        }

        Write-Host "Creating a new Excel application..."
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $true
        return $excel
    }

    function Test-CombiningTwoRectanglesResultsInARectangle {
    param(
        [Parameter(Mandatory=$true)]
        [object]$Area1,

        [Parameter(Mandatory=$true)]
        [object]$Area2
    )
    # Check if the top-left rows of the two areas are the same.
    $cond1 = $Area1.Row -eq $Area2.Row

    # Check if the number of rows in the two areas is the same.
    $cond2 = $Area1.Rows.Count -eq $Area2.Rows.Count

    # Check if the right edge of the first area connects to the left edge of the second area.
    $cond3 = $Area1.Column + $Area1.Columns.Count -eq $Area2.Column

    $cond1 -and $cond2 -and $cond3
    }

    function Invoke-CreateTableFromExcelRange {
        param(
            [Parameter(Mandatory=$true)]
            [object]$Range,

            [Parameter(Mandatory=$true)]
            [ref]$TableRef
        )

        # Declare a 2D array to hold the boolean values of the cells in the range $Range.
        $table = New-Object 'bool[,]' $Range.Rows.Count, $Range.Columns.Count

        # Fill the $table with the boolean values of the cells in the range $Range.
        $cells = $Range.Cells
        for ($rowIndex = 0; $rowIndex -lt $Range.Rows.Count; $rowIndex++) {
            for ($colIndex = 0; $colIndex -lt $Range.Columns.Count; $colIndex++) {
                $cell = $cells.Item($rowIndex + 1, $colIndex + 1)
                $value = $cell.Value2
                $table[$rowIndex, $colIndex] = [bool]$value
            }
        }

        $TableRef.Value = $table
    }

    function Invoke-ConfirmUserInput {
        param(
            [Parameter(Mandatory=$true)]
            [System.Collections.Generic.List[Object]]$TreeList
        )

        # Confirm user input before styling.
        Write-Host ""
        Write-Host "The following trees will be styled:" -ForegroundColor Green
        foreach ($Tree in $TreeList) {
            Invoke-DumpTree -Tree $Tree
        }
        Write-Host ""
        $confirmation = Read-Host "Do you want to proceed? (yes/no)"
        if ($confirmation -cne "yes") {
            Write-Host "Operation cancelled by the user."
            exit 0
        }
    }

    function Get-RangeFromAreas {
        param(
            [Parameter(Mandatory=$true)]
            [object]$Areas,

            [Parameter(Mandatory=$true)]
            [ref]$RangeRef
        )

        if (($areas.Count -eq 1) -and ($area1.Rows.Count -eq 1) -and ($area1.Columns.Count -eq 1)) {
            Write-Host "The selected area is only one cell. Please select a range and re-run the script."
            exit 1
        }

        if ($areas.Count -eq 2) {
            $area2 = $areas[2]
            Write-Host "Area 2: $($area2.Address($false, $false))"
            if (-not (Test-CombiningTwoRectanglesResultsInARectangle $area1 $area2)) {
                Write-Host "The two areas cannot be combined into a rectangle."
                exit 1
            }

            # Get the range that combines the two areas.
            $c1 = $worksheet.Cells($area1.Row, $area1.Column)
            $c2 = $worksheet.Cells($area2.Row + $area2.Rows.Count - 1, $area2.Column + $area2.Columns.Count - 1)
            $range = $worksheet.Range($c1, $c2)
        } else {
            Write-Host "Area 2: (not selected)"
            $range = $area1
        }
        $RangeRef.Value = $range
    }

    Invoke-TestExcelProcessesIsLessThanTwo

    $excel = Invoke-AttachExcelOrCreateExcel

    $workbook  = $excel.ActiveWorkbook
    if ($null -eq $workbook) {
        Write-Host "No active workbook found. Open '*.xlsx' file in Excel and re-run the script."
        exit 1
    }

    $workbookPath = $workbook.Name
    Write-Host "Active workbook found: $($workbookPath)"

    $worksheet = $excel.ActiveSheet
    if ($null -eq $worksheet) {
        Write-Host "No active worksheet found."
        exit 1
    }
    $worksheetName = $worksheet.Name
    Write-Host "Active worksheet found: $($worksheetName)"

    $selection = $excel.Selection
    $areas = $selection.Areas
    if ($areas.Count -eq 0) {
        Write-Host "No selection found. Please select a range and re-run the script."
        exit 1
    } elseif ($areas.Count -gt 2) {
        Write-Host "Too many areas selected. : $($areas.Count)"
        Write-Host "Please select one or two ranges and re-run the script."
        exit 1
    }

    $area1 = $areas[1]
    Write-Host "Area 1: $($area1.Address($false, $false))"

    # Get the number of columns in the header area.
    $headerColumnsCount = $area1.Columns.Count

    # Get the row and column offsets from the top-left of the table to the top-left of the Excel range.
    $RowFromTableToExcelOffset = $area1.Row
    $ColumnFromTableToExcelOffset = $area1.Column

    $RangeRef = [ref]$null
    Get-RangeFromAreas -Areas $areas -RangeRef $RangeRef
    $r1 = $RangeRef.Value

    $TableRef = [ref]$null
    Invoke-CreateTableFromExcelRange -Range $r1 -TableRef $TableRef
    $table = $TableRef.Value

    # Get the list of rectangles for the root trees from the 2D boolean array.
    $rootTreeRectangleList = Get-TreeRectangleList $table
    $treeFactory = [TreeFactory]::new($table)
    $TreeList = [System.Collections.Generic.List[Object]]::new()
    for ($treeIndex = 0; $treeIndex -lt $rootTreeRectangleList.Count; $treeIndex++) {
        $rootTreeRectangle = $rootTreeRectangleList[$treeIndex]
        $Tree = $treeFactory.CreateTree($treeIndex, $null, $rootTreeRectangle, 0, $headerColumnsCount)
        $TreeList.Add($Tree)
    }

    # The section depth max is set to 0, which means that only the root trees will be treated as sections.
    $sectionDepthMax = 1

    # Set the styler strategy.
    $stylerStrategy = [FillSectionAndDrawBordersStrategy]::new($excel, $workbook, $worksheet, $headerColumnsCount, $sectionDepthMax)

    # Confirm user input before styling.
    Invoke-ConfirmUserInput -TreeList $TreeList

    # For each tree in the tree list, invoke the Excel tree table styler.
    foreach ($Tree in $TreeList) {
        Invoke-ExcelTreeTableStyler `
            -RowFromTableToExcelOffset $RowFromTableToExcelOffset `
            -ColumnFromTableToExcelOffset $ColumnFromTableToExcelOffset `
            -Tree $Tree -StylerStrategy $stylerStrategy
    }
}
