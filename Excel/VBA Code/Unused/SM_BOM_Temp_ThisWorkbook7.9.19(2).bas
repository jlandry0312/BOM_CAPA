Sub addRow()
Dim objList As ListObject
Dim oListRow As ListRow

'Temp Disable - Sheet Protection / Screen Updating / Auto Calc
    'Sheets("BOM").Unprotect Password:="Elliot19"
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

'Create ListObjects object and reference data within SMDATAModel table body
    Set objList = ThisWorkbook.Sheets("BOM").ListObjects("SMDataModel")

'Add a new row above last within SMDATAModel
'Set Status to default value P
    
        With objList
            Set oListRow = .ListRows.Add(1)
            oListRow.Range.Locked = False
            oListRow.Range.Cells(8).Value = "P"
        End With
    
        
        
    'With objList
        'Set oListRow = .ListRows.Add(1)
        'oListRow.Range.Cells(8).Value = "P"
    'End With

'Re-enable - Sheet Protection / Screen Updating / Auto Calc
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
'Sheets("BOM").Protect Password:="Elliot19", _
    UserInterfaceOnly:=True, _
    DrawingObjects:=True, _
    Contents:=True, _
    Scenarios:=True, _
    AllowFormattingCells:=True, _
    AllowFormattingColumns:=True, _
    AllowFormattingRows:=True, _
    AllowInsertingColumns:=False, _
    AllowInsertingRows:=False, _
    AllowInsertingHyperlinks:=False, _
    AllowDeletingColumns:=False, _
    AllowDeletingRows:=True, _
    AllowSorting:=True, _
    AllowFiltering:=True, _
    AllowUsingPivotTables:=False



End Sub
