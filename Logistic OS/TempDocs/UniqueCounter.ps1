﻿$FilePath = "C:\Logistic OS\Templates\Templates_config.xlsx"

$Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $true
    $Workbooks = $Excel.Workbooks.Open($FilePath)
    $ClearSheet = $WorkBooks.Worksheets.Item('Unique')

    $UsedRange = $ClearSheet.UsedRange
    $UsedRange.EntireColumn.AutoFit() | Out-Null
    $LAST = "A" + $UsedRange.Rows.Count
    $IMO_ALL = $ClearSheet.range("A2:$LAST" )
    


    $Massive = $IMO_ALL.Value2 | Sort-Object | Group-Object -NoElement
    $Start = 2
    foreach($IMO in $Massive){
        $ClearSheet.Cells.Item($Start,2) = $IMO.Name
        $ClearSheet.Cells.Item($Start,3) = $IMO.Count
        $Start++
    }


    

   