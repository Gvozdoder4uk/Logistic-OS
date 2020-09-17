Param(
	[Parameter(Mandatory=$true)] [string]$FilePath
)

    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $true
    $Workbooks = $Excel.Workbooks.Open($FilePath)
    $ClearSheet = $WorkBooks.Worksheets.Item(1)


    $Filler = [System.Type]::Missing
    $UsedRange = $ClearSheet.UsedRange
    $UsedRange.EntireColumn.AutoFit() | Out-Null
    $T = "S" + $UsedRange.Rows.Count
    $Sorting_Space = $ClearSheet.range("S2:$T" )
    $UsedRange.Sort($Sorting_Space,1,$Filler,$Filler,$Filler,$Filler,$Filler,1)


    $UsedRange = $ClearSheet.UsedRange
    $UsedRange.Interior.ColorIndex = 0

    $T = "U" + $UsedRange.Rows.Count
    $Sorting_Space = $ClearSheet.range("U2:$T" )
    $Sorting_Space.Clear()
    $Workbooks.Save()
    $Workbooks.Close()
    
