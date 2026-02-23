BeforeAll {
    . $PSCommandPath.Replace('.Tests.ps1', '.ps1')

    # Helper function to create a 2D boolean array from a jagged array pattern
    function New-TestTable {
        param(
            [Parameter(Mandatory=$true)]
            [object[]]$Rows,

            [ref]$TableRef
        )

        $rowCount = $Rows.Count
        $colCount = $Rows[0].Length
        $table = New-Object 'bool[,]' $rowCount, $colCount

        for ($r = 0; $r -lt $rowCount; $r++) {
            $row = $Rows[$r]
            for ($c = 0; $c -lt $colCount; $c++) {
                $table[$r, $c] = [bool]($row[$c])
            }
        }

        $TableRef.Value = $table
    }
}

Describe "Get-TreeRectangleList" {
    It "Should return single rectangle for table with one tree (no additional tree markers)" {
        # Tree with no additional markers in column 0 - all rows belong to one tree
        $tableRef = [ref]$null
        New-TestTable @(
            @(1, 1, 0),  # Row 0: tree start
            @(0, 1, 1),  # Row 1: part of tree 1
            @(0, 0, 1),  # Row 2: part of tree 1
            @(0, 1, 0)   # Row 3: part of tree 1
        ) $tableRef
        $table = $tableRef.Value

        $treeRectangleList = Get-TreeRectangleList $table

        $treeRectangleList.Count | Should -Be 1
        $treeRectangleList[0].Row | Should -Be 0
        $treeRectangleList[0].Column | Should -Be 0
        $treeRectangleList[0].RowsCount | Should -Be 4
        $treeRectangleList[0].ColumnsCount | Should -Be 3
    }

    It "Should return two rectangles for table with two trees separated by marker row" {
        # Two trees: rows 0-1 (tree 1), then row 2 (tree 2 marker), then rows 2-3 (tree 2)
        $tableRef = [ref]$null
        New-TestTable @(
            @(1, 1, 0),  # Row 0: tree 1 start
            @(0, 1, 1),  # Row 1: part of tree 1
            @(1, 0, 1),  # Row 2: tree 2 start (marker in column 0)
            @(0, 1, 0)   # Row 3: part of tree 2
        ) $tableRef
        $table = $tableRef.Value

        $treeRectangleList = Get-TreeRectangleList $table

        $treeRectangleList.Count | Should -Be 2

        # First tree
        $treeRectangleList[0].Row | Should -Be 0
        $treeRectangleList[0].RowsCount | Should -Be 2
        $treeRectangleList[0].ColumnsCount | Should -Be 3

        # Second tree
        $treeRectangleList[1].Row | Should -Be 2
        $treeRectangleList[1].RowsCount | Should -Be 2
        $treeRectangleList[1].ColumnsCount | Should -Be 3
    }

    It "Should return three rectangles for table with three trees" {
        # Three trees with marker rows
        $tableRef = [ref]$null
        New-TestTable @(
            @(1, 1, 0),  # Row 0: tree 1 start
            @(0, 1, 1),  # Row 1: part of tree 1
            @(1, 0, 1),  # Row 2: tree 2 start
            @(0, 1, 0),  # Row 3: part of tree 2
            @(1, 1, 1),  # Row 4: tree 3 start
            @(0, 0, 1)   # Row 5: part of tree 3
        ) $tableRef
        $table = $tableRef.Value

        $treeRectangleList = Get-TreeRectangleList $table

        $treeRectangleList.Count | Should -Be 3
        # First tree
        $treeRectangleList[0].Row | Should -Be 0
        $treeRectangleList[0].RowsCount | Should -Be 2
        # Second tree
        $treeRectangleList[1].Row | Should -Be 2
        $treeRectangleList[1].RowsCount | Should -Be 2
        # Third tree
        $treeRectangleList[2].Row | Should -Be 4
        $treeRectangleList[2].RowsCount | Should -Be 2
    }

    It "Should handle tree with single row" {
        # Tree 1: 1 row, Tree 2: 2 rows, Tree 3: 1 row
        $tableRef = [ref]$null
        New-TestTable @(
            @(1, 1, 0),  # Row 0: tree 1 (single row)
            @(1, 0, 1),  # Row 1: tree 2 start
            @(0, 1, 0),  # Row 2: part of tree 2
            @(1, 1, 1)   # Row 3: tree 3 (single row)
        ) $tableRef
        $table = $tableRef.Value

        $treeRectangleList = Get-TreeRectangleList $table

        $treeRectangleList.Count | Should -Be 3
        # Tree 1: single row
        $treeRectangleList[0].RowsCount | Should -Be 1
        # Tree 2: two rows
        $treeRectangleList[1].RowsCount | Should -Be 2
        # Tree 3: single row
        $treeRectangleList[2].RowsCount | Should -Be 1
    }

    It "Should always set Column to 0 for all rectangles" {
        $tableRef = [ref]$null
        New-TestTable @(
            @(1, 1, 1, 1),
            @(0, 0, 0, 0),
            @(1, 1, 1, 1)
        ) $tableRef
        $table = $tableRef.Value

        $treeRectangleList = Get-TreeRectangleList $table

        foreach ($rect in $treeRectangleList) {
            $rect.Column | Should -Be 0
        }
    }

    It "Should set ColumnsCount to table width for all rectangles" {
        $tableRef = [ref]$null
        New-TestTable @(
            @(1, 1, 1, 1, 1),
            @(0, 0, 1, 0, 0),
            @(1, 1, 1, 1, 1)
        ) $tableRef
        $table = $tableRef.Value

        $treeRectangleList = Get-TreeRectangleList $table

        foreach ($rect in $treeRectangleList) {
            $rect.ColumnsCount | Should -Be 5
        }
    }

    It "Should handle consecutive tree markers (trees with no rows between them)" {
        # Three trees: each starts immediately after previous
        $tableRef = [ref]$null
        New-TestTable @(
            @(1, 1, 0),  # Row 0: tree 1 start (1 row)
            @(1, 0, 1),  # Row 1: tree 2 start (1 row)
            @(1, 1, 1)   # Row 2: tree 3 start (1 row)
        ) $tableRef
        $table = $tableRef.Value

        $treeRectangleList = Get-TreeRectangleList $table

        $treeRectangleList.Count | Should -Be 3
        $treeRectangleList[0].RowsCount | Should -Be 1
        $treeRectangleList[1].RowsCount | Should -Be 1
        $treeRectangleList[2].RowsCount | Should -Be 1
    }

    It "Should preserve Row index correctly across multiple trees" {
        $tableRef = [ref]$null
        New-TestTable @(
            @(1, 1, 0),  # Row 0
            @(0, 1, 1),  # Row 1
            @(0, 0, 1),  # Row 2
            @(1, 1, 0),  # Row 3 (tree 2 starts here)
            @(0, 1, 0),  # Row 4
            @(0, 1, 1)   # Row 5
        ) $tableRef
        $table = $tableRef.Value

        $treeRectangleList = Get-TreeRectangleList $table

        $treeRectangleList.Count | Should -Be 2
        $treeRectangleList[0].Row | Should -Be 0
        $treeRectangleList[1].Row | Should -Be 3
    }
}

Describe "TreeFactory" {
    It "FindFirstTreeTopLeft should find the first tree top-left position" {
        # Test table with tree markers in different columns
        $tableRef = [ref]$null
        New-TestTable @(
            @(1, 1, 0, 0),  # Row 0: root, first tree marker in column 1
            @(0, 1, 0, 0),  # Row 1: continuation
            @(0, 0, 1, 0),  # Row 2: second tree marker in column 2
            @(0, 0, 1, 1)   # Row 3: continuation
        ) $tableRef
        $table = $tableRef.Value

        $factory = [TreeFactory]::new($table)
        $rect = [Rectangle]::new(0, 0, 4, 4)

        $rowIndexRef = [ref]$null
        $columnIndexRef = [ref]$null
        $result = $factory.FindFirstTreeTopLeft($rect, 3, $rowIndexRef, $columnIndexRef)

        $result | Should -Be $true
        $rowIndexRef.Value | Should -Be 0
        $columnIndexRef.Value | Should -Be 1
    }

    It "FindFirstTreeTopLeft should return false when no tree marker found" {
        # Test table with no tree markers beyond column 0
        $tableRef = [ref]$null
        New-TestTable @(
            @(1, 0, 0, 0),
            @(0, 0, 0, 0),
            @(0, 0, 0, 0)
        ) $tableRef
        $table = $tableRef.Value

        $factory = [TreeFactory]::new($table)
        $rect = [Rectangle]::new(0, 0, 3, 4)

        $rowIndexRef = [ref]$null
        $columnIndexRef = [ref]$null
        $result = $factory.FindFirstTreeTopLeft($rect, 3, $rowIndexRef, $columnIndexRef)

        $result | Should -Be $false
    }

    It "FindNextSiblingTreeTopLeft should find sibling tree markers in same column" {
        # Test table with multiple tree markers in column 1
        $tableRef = [ref]$null
        New-TestTable @(
            @(1, 1, 0),  # Row 0: first tree marker in column 1
            @(0, 0, 1),  # Row 1: no marker in column 1
            @(0, 1, 0),  # Row 2: second tree marker in column 1 (sibling)
            @(0, 0, 1),  # Row 3: no marker in column 1
            @(0, 1, 1)   # Row 4: third tree marker in column 1 (sibling)
        ) $tableRef
        $table = $tableRef.Value

        $factory = [TreeFactory]::new($table)
        $rect = [Rectangle]::new(0, 0, 5, 3)

        # Find first sibling after row 0
        $rowIndexRef = [ref]1
        $result = $factory.FindNextSiblingTreeTopLeft($rect, $rowIndexRef, 1)

        $result | Should -Be $true
        $rowIndexRef.Value | Should -Be 2
    }

    It "FindNextSiblingTreeTopLeft should return false when no more siblings" {
        # Test table with limited tree markers
        $tableRef = [ref]$null
        New-TestTable @(
            @(1, 1, 0),  # Row 0: tree marker in column 1
            @(0, 0, 1),  # Row 1: continuation
            @(0, 0, 0)   # Row 2: no marker
        ) $tableRef
        $table = $tableRef.Value

        $factory = [TreeFactory]::new($table)
        $rect = [Rectangle]::new(0, 0, 3, 3)

        $rowIndexRef = [ref]1
        $result = $factory.FindNextSiblingTreeTopLeft($rect, $rowIndexRef, 1)

        $result | Should -Be $false
        $rowIndexRef.Value | Should -Be -1
    }

    It "FindNextTreeTopLeft should advance through rows and columns" {
        # Test table with tree markers in different positions
        $tableRef = [ref]$null
        New-TestTable @(
            @(1, 0, 0, 0),  # Row 0: no marker beyond column 0
            @(0, 0, 0, 0),  # Row 1: no marker
            @(0, 1, 0, 0),  # Row 2: marker in column 1
            @(0, 0, 1, 0)   # Row 3: marker in column 2
        ) $tableRef
        $table = $tableRef.Value

        $factory = [TreeFactory]::new($table)
        $rect = [Rectangle]::new(0, 0, 4, 4)

        # Find starting from row 0, column 1
        $rowIndexRef = [ref]0
        $columnIndexRef = [ref]1
        $result = $factory.FindNextTreeTopLeft($rect, 3, $rowIndexRef, $columnIndexRef)

        $result | Should -Be $true
        $rowIndexRef.Value | Should -Be 2
        $columnIndexRef.Value | Should -Be 1
    }

    It "CreateTree should build tree structure with no children when no tree markers found" {
        # Simple tree with no child trees
        $tableRef = [ref]$null
        New-TestTable @(
            @(1, 0, 0),  # Row 0: root marker
            @(0, 0, 0),  # Row 1: no child markers
            @(0, 0, 0)   # Row 2: no child markers
        ) $tableRef
        $table = $tableRef.Value

        $factory = [TreeFactory]::new($table)
        $rect = [Rectangle]::new(0, 0, 3, 3)

        $tree = $factory.CreateTree(0, $null, $rect, 0, 3)

        $tree | Should -Not -BeNullOrEmpty
        $tree.Depth | Should -Be 0
        $tree.RootTreeIndex | Should -Be 0
        $tree.Children.Count | Should -Be 0
    }

    It "CreateTree should build tree with one child" {
        # Tree with one child branch
        $tableRef = [ref]$null
        New-TestTable @(
            @(1, 0, 0),  # Row 0: root marker
            @(0, 1, 0),  # Row 1: child marker in column 1
            @(0, 0, 0)   # Row 2: no marker
        ) $tableRef
        $table = $tableRef.Value

        $factory = [TreeFactory]::new($table)
        $rect = [Rectangle]::new(0, 0, 3, 3)

        $tree = $factory.CreateTree(0, $null, $rect, 0, 3)

        $tree | Should -Not -BeNullOrEmpty
        $tree.Children.Count | Should -Be 1
        $tree.Children[0].Depth | Should -Be 1
    }

    It "CreateTree should build tree with multiple children" {
        # Tree with two sibling child branches at same level
        $tableRef = [ref]$null
        New-TestTable @(
            @(1, 0, 0),  # Row 0: root marker
            @(0, 1, 0),  # Row 1: first child marker in column 1
            @(0, 0, 0),  # Row 2: middle row
            @(0, 1, 0),  # Row 3: second child marker in column 1 (sibling)
            @(0, 0, 0)   # Row 4: final row
        ) $tableRef
        $table = $tableRef.Value

        $factory = [TreeFactory]::new($table)
        $rect = [Rectangle]::new(0, 0, 5, 3)

        $tree = $factory.CreateTree(0, $null, $rect, 0, 3)

        $tree | Should -Not -BeNullOrEmpty
        $tree.Children.Count | Should -Be 2
        $tree.Children[0].Depth | Should -Be 1
        $tree.Children[1].Depth | Should -Be 1
    }

    It "CreateTree should build tree with nested children" {
        # Tree with nested hierarchy: root -> level 1 -> level 2
        $tableRef = [ref]$null
        New-TestTable @(
            @(1, 0, 0, 0),  # Row 0: root marker
            @(0, 1, 0, 0),  # Row 1: level 1 child marker in column 1
            @(0, 0, 1, 0),  # Row 2: level 2 child marker in column 2
            @(0, 0, 0, 0)   # Row 3: continuation
        ) $tableRef
        $table = $tableRef.Value

        $factory = [TreeFactory]::new($table)
        $rect = [Rectangle]::new(0, 0, 4, 4)

        $tree = $factory.CreateTree(0, $null, $rect, 0, 4)

        $tree | Should -Not -BeNullOrEmpty
        $tree.Children.Count | Should -Be 1
        $tree.Children[0].Depth | Should -Be 1
        $tree.Children[0].Children.Count | Should -Be 1
        $tree.Children[0].Children[0].Depth | Should -Be 2
    }

    It "CreateTree should set ParentTree correctly" {
        # Verify parent-child relationship
        $tableRef = [ref]$null
        New-TestTable @(
            @(1, 0, 0),  # Row 0: root marker
            @(0, 1, 0),  # Row 1: child marker
            @(0, 0, 0)   # Row 2: continuation
        ) $tableRef
        $table = $tableRef.Value

        $factory = [TreeFactory]::new($table)
        $rect = [Rectangle]::new(0, 0, 3, 3)

        $rootTree = $factory.CreateTree(0, $null, $rect, 0, 3)

        $rootTree.ParentTree | Should -BeNullOrEmpty
        $rootTree.Children[0].ParentTree | Should -Be $rootTree
    }

    It "CreateTree should preserve RootTreeIndex for all nodes" {
        # Verify RootTreeIndex is same for entire tree hierarchy
        $tableRef = [ref]$null
        New-TestTable @(
            @(1, 0, 0, 0),  # Row 0: root marker
            @(0, 1, 0, 0),  # Row 1: child marker
            @(0, 0, 1, 0),  # Row 2: grandchild marker
            @(0, 0, 0, 0)   # Row 3: continuation
        ) $tableRef
        $table = $tableRef.Value

        $factory = [TreeFactory]::new($table)
        $rect = [Rectangle]::new(0, 0, 4, 4)
        $rootTreeIndex = 42

        $rootTree = $factory.CreateTree($rootTreeIndex, $null, $rect, 0, 4)

        $rootTree.RootTreeIndex | Should -Be $rootTreeIndex
        $rootTree.Children[0].RootTreeIndex | Should -Be $rootTreeIndex
        $rootTree.Children[0].Children[0].RootTreeIndex | Should -Be $rootTreeIndex
    }
}

