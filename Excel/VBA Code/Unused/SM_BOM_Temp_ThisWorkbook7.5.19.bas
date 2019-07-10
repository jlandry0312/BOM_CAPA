Private Sub Workbook_Open()
Sheets("BOM").Protect Password:="Elliot19", _
    UserInterfaceOnly:=True, _
    DrawingObjects:=False, _
    Contents:=True, _
    Scenarios:=False, _
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


Private Sub Auto_Close()
    
    Dim DestFolder As String
    Dim MyFileName As String
    Dim CurrentWB As Workbook, TempWB As Workbook
    Application.ScreenUpdating = False
    
    DestFolder = "X:\DataDump"
   
    Set CurrentWB = ActiveWorkbook
    ActiveWorkbook.Sheets("BOM").ListObjects("SMDataModel").Range.Copy

    Set TempWB = Application.Workbooks.Add(1)
    With TempWB.Sheets(1).Range("A1")
        .PasteSpecial xlPasteValues
        .PasteSpecial xlPasteFormats
    End With

    'Dim Change below to "- 4"  to become compatible with .xls files
    MyFileName = DestFolder & "\" & Left(CurrentWB.Name, Len(CurrentWB.Name) - 5) & ".csv"

    Application.DisplayAlerts = False
    TempWB.SaveAs filename:=MyFileName, FileFormat:=xlCSV, CreateBackup:=False, Local:=True
    TempWB.Close SaveChanges:=False
    ThisWorkbook.Saved = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub




