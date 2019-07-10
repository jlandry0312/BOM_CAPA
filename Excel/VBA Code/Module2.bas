Attribute VB_Name = "Module2"
Function IsCellInTable(cell As Range) As Boolean
'PURPOSE: Determine if a cell is within an Excel Table
'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault

IsCellInTable = False

On Error Resume Next
  IsCellInTable = (cell.ListObject.Name <> "")
On Error GoTo 0
   
End Function



Private Sub AddTableRows()
'PURPOSE: Add table row based on user's selection
'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault

Dim rng As Range
Dim InsertRows As Long
Dim StartRow As Long
Dim InsideTable As Boolean
Dim RowToBottom As Boolean
Dim ReProtect As Boolean
Dim Password As String
Dim area As Range

'Optimize Code
  Application.ScreenUpdating = False


  Password = "Elliot19"

'Set Range Variable
  On Error GoTo InvalidSelection
    Set rng = Selection
  On Error GoTo 0

'Unprotect Worksheet
  With ActiveSheet
    If .ProtectContents Or .ProtectDrawingObjects Or .ProtectScenarios Then
      On Error GoTo InvalidPassword
      .Unprotect Password
      ReProtect = True
      On Error GoTo 0
    End If
  End With

'Loop Through each Area in Selection
  For Each area In rng.Areas

    'Is selected Cell within a table?
      InsideTable = IsCellInTable(area.Cells(1, 1))
      
    'Is selected cell 1 row under a table?
      RowToBottom = IsCellInTable(area.Cells(1, 1).Offset(-1))
    
    'How Many Rows In Selection?
      InsertRows = area.Rows.Count
    
    'Selection Not Within Table?
      If Not InsideTable And Not RowToBottom Then GoTo InvalidSelection
    
    'Add Rows To Table
      If InsideTable Then
      
        'Which Row in Table is selected?
          With area.Cells(1, 1)
            x = .Row
            y = .ListObject.DataBodyRange.Row
            Z = .ListObject.DataBodyRange.Rows.Count
          End With
          
          StartRow = Z - ((y + Z - 1) - x)
          
        'Insert rows based on how many rows are currently selected
          For x = 1 To InsertRows
            
            With area
                .ListObject.ListRows.Add (StartRow)
                .ListObject.DataBodyRange.Cells(StartRow, 8).Value = "P"
            'area.ListObject.ListRows.Add (StartRow)
            'area.ListObject.DataBodyRange.Cells(x, 8).Value = "P"
            End With
            
          Next x
      ElseIf RowToBottom Then
        For x = 1 To InsertRows
          With area.Cells(1, 1)
          .Offset(-1).ListObject.ListRows.Add AlwaysInsert:=True
          .Offset(-1, 7).Value = "P"
          End With
          'area.Cells(1, 1).Offset(-1).ListObject.ListRows.Add AlwaysInsert:=True
          'area.Cells(1, 1).Offset(-1).ListObject.DataBodyRange.Cells(8).Value = "P"
        Next x
      End If

  Next area

'Protect Worksheet
  If ReProtect = True Then Sheets("BOM").Protect Password:="Elliot19", _
  UserInterfaceOnly:=True, _
    DrawingObjects:=False, _
    Contents:=True, _
    Scenarios:=False, _
    AllowFormattingCells:=True, _
    AllowFormattingColumns:=True, _
    AllowFormattingRows:=True, _
    AllowInsertingColumns:=False, _
    AllowInsertingRows:=True, _
    AllowInsertingHyperlinks:=False, _
    AllowDeletingColumns:=False, _
    AllowDeletingRows:=True, _
    AllowSorting:=True, _
    AllowFiltering:=True, _
    AllowUsingPivotTables:=False
  
  

Exit Sub

'ERROR HANDLERS
InvalidSelection:
  MsgBox "You must select a cell within or directly below the BOM Table"
  If ReProtect = True Then Sheets("BOM").Protect Password:="Elliot19", _
  UserInterfaceOnly:=True, _
    DrawingObjects:=False, _
    Contents:=True, _
    Scenarios:=False, _
    AllowFormattingCells:=True, _
    AllowFormattingColumns:=True, _
    AllowFormattingRows:=True, _
    AllowInsertingColumns:=False, _
    AllowInsertingRows:=True, _
    AllowInsertingHyperlinks:=False, _
    AllowDeletingColumns:=False, _
    AllowDeletingRows:=True, _
    AllowSorting:=True, _
    AllowFiltering:=True, _
    AllowUsingPivotTables:=False
  Exit Sub

InvalidPassword:
  MsgBox "Failed to unlock password with the following password: " & Password
  Exit Sub
  
End Sub

Private Sub DeleteTableRows()
'PURPOSE: Delete table row based on user's selection
'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault

Dim rng As Range
Dim DeleteRng As Range
Dim cell As Range
Dim TempRng As Range
Dim Answer As Variant
Dim Password As String
Dim area As Range
Dim ReProtect As Boolean

'What is the worksheet password?
  Password = "Elliot19"

'Set Range Variable
  On Error GoTo InvalidSelection
    Set rng = Selection
  On Error GoTo 0

'Unprotect Worksheet
  With ActiveSheet
    If ProtectContents Or ProtectDrawingObjects Or ProtectScenarios Then
      On Error GoTo InvalidPassword
      .Unprotect Password
      ReProtect = True
      On Error GoTo 0
    End If
  End With

'Loop Through each Area in Selection
  For Each area In rng.Areas
    For Each cell In area.Cells.Columns(1)
  
      'Is selected Cell within a table?
        InsideTable = IsCellInTable(cell)
    
      'Gather rows to delete
        If InsideTable Then
          On Error GoTo InvalidActiveCell
          Set TempRng = Intersect(cell.EntireRow, ActiveCell.ListObject.DataBodyRange)
          On Error GoTo 0
          
          If DeleteRng Is Nothing Then
            Set DeleteRng = TempRng
          Else
            Set DeleteRng = Union(TempRng, DeleteRng)
          End If
    
        End If
        
    Next cell
  Next area
  
'Error Handling
  If DeleteRng Is Nothing Then GoTo InvalidSelection
  If DeleteRng.Address = ActiveCell.ListObject.DataBodyRange.Address Then GoTo DeleteAllRows
  If ActiveCell.ListObject.DataBodyRange.Rows.Count = 1 Then GoTo DeleteOnlyRow
  
'Ask User To confirm delete (since this cannot be undone)
    DeleteRng.Select
    
    If DeleteRng.Rows.Count = 1 And DeleteRng.Areas.Count = 1 Then
      Answer = MsgBox("Are you sure you want to delete the currently selected table row? " & _
      " This cannot be undone...", vbYesNo, "Delete Row?")
    Else
      Answer = MsgBox("Are you sure you want to delete the currently selected table rows? " & _
       " This cannot be undone...", vbYesNo, "Delete Rows?")
    End If
      
'Delete row (if wanted)
  If Answer = vbYes Then DeleteRng.Delete xlShiftUp

'Protect Worksheet
  If ReProtect = True Then Sheets("BOM").Protect Password:="Elliot19", _
  UserInterfaceOnly:=True, _
    DrawingObjects:=False, _
    Contents:=True, _
    Scenarios:=False, _
    AllowFormattingCells:=True, _
    AllowFormattingColumns:=True, _
    AllowFormattingRows:=True, _
    AllowInsertingColumns:=False, _
    AllowInsertingRows:=True, _
    AllowInsertingHyperlinks:=False, _
    AllowDeletingColumns:=False, _
    AllowDeletingRows:=True, _
    AllowSorting:=True, _
    AllowFiltering:=True, _
    AllowUsingPivotTables:=False

Exit Sub

'Error Handlers
InvalidActiveCell:
  MsgBox "The first cell you select must be inside an Excel Table. " & _
   "The first cell you selected was cell " & ActiveCell.Address, vbCritical, "Invalid Selection!"
  If ReProtect = True Then Sheets("BOM").Protect Password:="Elliot19", _
  UserInterfaceOnly:=True, _
    DrawingObjects:=False, _
    Contents:=True, _
    Scenarios:=False, _
    AllowFormattingCells:=True, _
    AllowFormattingColumns:=True, _
    AllowFormattingRows:=True, _
    AllowInsertingColumns:=False, _
    AllowInsertingRows:=True, _
    AllowInsertingHyperlinks:=False, _
    AllowDeletingColumns:=False, _
    AllowDeletingRows:=True, _
    AllowSorting:=True, _
    AllowFiltering:=True, _
    AllowUsingPivotTables:=False
  Exit Sub

InvalidSelection:
  MsgBox "You must select a cell within an Excel table", vbCritical, "Invalid Selection!"
  If ReProtect = True Then Sheets("BOM").Protect Password:="Elliot19", _
  UserInterfaceOnly:=True, _
    DrawingObjects:=False, _
    Contents:=True, _
    Scenarios:=False, _
    AllowFormattingCells:=True, _
    AllowFormattingColumns:=True, _
    AllowFormattingRows:=True, _
    AllowInsertingColumns:=False, _
    AllowInsertingRows:=True, _
    AllowInsertingHyperlinks:=False, _
    AllowDeletingColumns:=False, _
    AllowDeletingRows:=True, _
    AllowSorting:=True, _
    AllowFiltering:=True, _
    AllowUsingPivotTables:=False
  Exit Sub

DeleteAllRows:
  MsgBox "You cannot delete all the rows in the table. " & _
   "You must leave at least one row existing in a table", vbCritical, "Cannot Delete!"
  If ReProtect = True Then Sheets("BOM").Protect Password:="Elliot19", _
  UserInterfaceOnly:=True, _
    DrawingObjects:=False, _
    Contents:=True, _
    Scenarios:=False, _
    AllowFormattingCells:=True, _
    AllowFormattingColumns:=True, _
    AllowFormattingRows:=True, _
    AllowInsertingColumns:=False, _
    AllowInsertingRows:=True, _
    AllowInsertingHyperlinks:=False, _
    AllowDeletingColumns:=False, _
    AllowDeletingRows:=True, _
    AllowSorting:=True, _
    AllowFiltering:=True, _
    AllowUsingPivotTables:=False
  Exit Sub

DeleteOnlyRow:
  MsgBox "You cannot delete the only row in the table.", vbCritical, "Cannot Delete!"
  If ReProtect = True Then Sheets("BOM").Protect Password:="Elliot19", _
  UserInterfaceOnly:=True, _
    DrawingObjects:=False, _
    Contents:=True, _
    Scenarios:=False, _
    AllowFormattingCells:=True, _
    AllowFormattingColumns:=True, _
    AllowFormattingRows:=True, _
    AllowInsertingColumns:=False, _
    AllowInsertingRows:=True, _
    AllowInsertingHyperlinks:=False, _
    AllowDeletingColumns:=False, _
    AllowDeletingRows:=True, _
    AllowSorting:=True, _
    AllowFiltering:=True, _
    AllowUsingPivotTables:=False
  Exit Sub
  
InvalidPassword:
  MsgBox "Failed to unlock password with the following password: " & Password
  Exit Sub

End Sub

