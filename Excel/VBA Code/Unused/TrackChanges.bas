Private Sub Worksheet_Change(ByVal Target As Range)
        Dim r As Integer
        Dim c As Integer
        Dim arr(1 To 1, 1 To 12)
        If Not Intersect(Target, Range("B10:N90")) Is Nothing Then
            r = Target.Row
            For c = 1 To 12
            arr(1, c) = Cells(r, c).Value
            Next
                Dim rowCount As Integer
                rowCount = Sheet2.UsedRange.Rows.Count
            
                Dim rowToUse As Integer
                rowToUse = rowCount + 1
            
                Sheet2.Cells(rowToUse, 15).Value = Format(Now(), "dd/mm/yyyy hh:mm AM/PM")
                Sheet2.Cells(rowToUse, 14).Value = Environ("USER")
        
            With Sheet2
                .Range(.Cells(rowToUse, 1), .Cells(rowToUse, 16)) = arr
            End With
        End If




Option Compare Text

Sub checkStatus()
    Dim SrchRng As Range
    Dim cel As Range
    
    
    
    Set SrchRng = Range("A10:A100")
    For Each cel In SrchRng
        If InStr(1, cel.Value, "Released to PM") > 0 Then
            Sheet1.Select
            Range(2, 13).Copy
            Sheets("ChangeLog").Select
            Range("A1").Select
            ActiveSheet.Paste
        End If
    Next cel
End Sub

Set KeyCells = Range("A1:A500")
    
    
    
    If Not Application.Intersect(Target, Range("A1:A500")) Is Nothing Then
       If
            Sheet1.Select
            Range(Target, 13).Copy
            Sheets("ChangeLog").Select
            Range("A1").Select
            ActiveSheet.Paste
        End If
    End If
           
       
    
End Sub

            


End Sub



Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    Dim r As Integer
    Dim c As Integer
    Dim arr(1 To 1, 1 To 13)
    

    If Not Intersect(Target, Range("A")) Is Nothing Then
        If InStr(Target.Value, "RELEASED TO PM") > 0 Then
         r = Target.Row
         
         For c = 1 To 13
            arr(1, c) = Cells(r, c).Value
        Next
            Dim rowCount As Integer
            rowCount = Sheets("ChangeLog").UsedRange.Rows.Count
            
            Dim rowToUse As Integer
            rowToUse = rowCount + 1
            
        With Sheets("ChangeLog")
            .Range(.Cells(rowToUse, 1), .Cells(rowToUse, 16)) = arr
        End With
    End If
End If

        
End Sub
    
    