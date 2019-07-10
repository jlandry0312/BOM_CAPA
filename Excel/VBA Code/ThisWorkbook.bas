

Private Sub Workbook_Open()

'Set Workbook Security on Open
With Sheets("BOM")
    .Protect Password:="Elliot19", _
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
    .Range("A1").Calculate
    .Range("A2").Calculate
    .Range("D6").Calculate
End With


End Sub


Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
   'Export BOM data as .csv to X:\DataDump
    
    Dim DestFolder As String
    Dim MyFileName As String
    Dim CurrentWB As Workbook, TempWB As Workbook
    Set CurrentWB = ActiveWorkbook
    
    DestFolder = "X:\DataDump"
    Application.ScreenUpdating = False
       
    dataRow = CurrentWB.Sheets("BOM").Range("SMDataModel").Rows.Count
    dataRow = (dataRow + 8)
    
    CurrentWB.Sheets("BOM").ListObjects("SMDataModel").Range.Copy

    Set TempWB = Application.Workbooks.Add(1)
    With TempWB.Sheets(1).Range("A1")
        .PasteSpecial xlPasteValues
        .PasteSpecial xlPasteFormats
    End With

   
    MyFileName = DestFolder & "\" & Left(CurrentWB.Name, Len(CurrentWB.Name) - 5) & ".csv"

    Application.DisplayAlerts = False
    TempWB.SaveAs filename:=MyFileName, FileFormat:=xlCSV, CreateBackup:=False, Local:=True
    TempWB.Close SaveChanges:=False
    ThisWorkbook.Saved = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
End Sub






