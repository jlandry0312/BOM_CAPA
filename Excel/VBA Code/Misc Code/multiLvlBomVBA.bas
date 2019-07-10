ub AutoGroupBOM()
    'Define Variables
    Dim StartingCell As Range
    Dim StartingRow As Integer
    Dim LevelCol As Integer
    Dim LastRow As Integer
    Dim CurrentLevel As Integer
    Dim i As Integer
    Dim j As Integer
    
    'Ask user to select the starting row
    Set StartingCell = Application.InputBox("Select Top Left Cell of BOM, Must contain level", Type:=8)
    StartingRow = StartingCell.Row
    LevelCol = StartingCell.Column
    LastRow = ActiveSheet.UsedRange.Rows.Count
    
    'Walk down the bom lines and group items
    For i = StartingRow To LastRow
        CurrentLevel = Cells(i, LevelCol)
        If CurrentLevel > 1 Then
            For j = 1 To CurrentLevel - 1
                Rows(i).Select
                Selection.Rows.Group
            Next j
        End If
    Next i
End Sub