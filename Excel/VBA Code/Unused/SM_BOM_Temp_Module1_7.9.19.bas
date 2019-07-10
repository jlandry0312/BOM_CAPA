Sub addRow()
Dim objList As ListObject
Dim oListRow As ListRow

'Temp Disable - Sheet Protection / Screen Updating / Auto Calc
    ActiveSheet.Unprotect Password:="Elliot19"
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

'Create ListObjects object and reference data within SMDATAModel table body
    Set objList = ActiveSheet.ListObjects("SMDATAModel")

'Add a new row above last within SMDATAModel
'Set Status to default value P
    With objList
        Set oListRow = .ListRows.Add(1)
        oListRow.Range.Cells(8).Value = "P"
    End With

'Re-enable - Sheet Protection / Screen Updating / Auto Calc
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
ActiveSheet.Protect Password:="Elliot19"



End Sub

