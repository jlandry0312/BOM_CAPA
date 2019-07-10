Sub Add_Row()

Dim sht As Worksheet
Dim LastRow As Long
Dim i As Integer
Dim stat As String
Dim r As Range
stat = "P"
Set sht = ActiveSheet

Application.ScreenUpdating = False
LastRow = sht.Range("SMDataModel").Rows.Count
LastRow = (LastRow + 8)

Rows(LastRow).Select
Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone

Set r = Range("I" & LastRow)
r.Value = stat




Application.ScreenUpdating = True


End Sub