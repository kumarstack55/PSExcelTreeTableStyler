# PSExcelTreeTableStyler

![demo](images/demo.gif)

PSExcelTreeTableStyler is a PowerShell script that styles Excel tables in a tree structure.
It allows you to easily create a tree table in Excel by selecting multiple ranges and applying styles.

## Requirements

- Windows 11+
- Microsoft Excel
- Windows PowerShell 5.1+

## Usage

```powershell
# powershell

.\Invoke-ExcelTreeTableStyler.ps1
```

## Demonstration

- First, launch Excel and create a new workbook. Then, fill in some data.
- Then, select two ranges, B3:D14 and E3:F14, and run the Excel Tree Table Styler. To select two ranges, click and hold on cell B3, drag to D14, then hold down the Ctrl key and click and hold on cell E3, drag to F14.

![before](images/before.png)

- Finally, run Invoke-ExcelTreeTableStyler.ps1 to style the selected ranges as a tree table.
- You can see that the two ranges are treated as one table, and the tree structure is created based on the data. The cells with the same value in the same column are grouped together, and the borders are drawn to show the tree structure.

![after](images/after.png)

If you want to try it yourself, you can run the following code in PowerShell to launch Excel and fill in some data.

```powershell
# powershell

git clone https://github.com/kumarstack55/PSExcelTreeTableStyler.git
Set-Location .\PSExcelTreeTableStyler\

# Launch Excel and create a new workbook.
$application = New-Object -ComObject Excel.Application
$application.Visible = $true
$workbook = $application.Workbooks.Add()
$worksheet = $workbook.Worksheets.Item(1)

$range1 = $worksheet.Range("B2:F14")
$range1.Select()
$range1.Clear()

# Fill in some data
$worksheet.Range("B3").Value2 = "header1"
$worksheet.Range("C4").Value2 = "header2"
$worksheet.Range("C5").Value2 = "header3"
$worksheet.Range("D6").Value2 = "header4"
$worksheet.Range("D7").Value2 = "header5"
$worksheet.Range("C8").Value2 = "header6"
$worksheet.Range("D9").Value2 = "header7"
$worksheet.Range("B10").Value2 = "header8"
$worksheet.Range("C11").Value2 = "header9"
$worksheet.Range("C13").Value2 = "header10"
$worksheet.Range("D14").Value2 = "header11"
$worksheet.Range("E2").Value2 = "header12"
$worksheet.Range("F2").Value2 = "header13"

# Select two ranges.
$range2 = $worksheet.Range("B3:D14,E3:F14")
$range2.Select()

# Run the Excel Tree Table Styler.
.\Invoke-ExcelTreeTableStyler.ps1
```

If you want to close the Excel application after testing, you can run the following code in PowerShell.
Note that we need to release the COM object and run garbage collection to ensure that the Excel process is closed properly.

```powershell
# powershell

$workbook.Close($false)
$workbook = $null
$worksheet = $null

# Close Excel application.
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
[System.GC]::Collect()
$application.Quit()

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($application)
$application = $null

[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
[System.GC]::Collect()

# Check if Excel process is still running or not.
# We should check if COM objects are released, not just the process. But we won't.
Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue
```

> Note: This project abandons the release of COM objects. For example:
>
> - The code does not release COM objects that are implicitly created, such as Range in the above code.
> - Generally, foreach loops generate an enumerator objects, but the code does not release it.
>
> If you are using a client OS that can be restarted occasionally, it should not be a problem at all.
>
> If you are a perfectionist who pursues ideals, I recommend that you spend time researching the release of COM objects. However, life is short.

## Development

```powershell
# powershell

# Run tests.
Invoke-Pester
```

## LICENSE

MIT
