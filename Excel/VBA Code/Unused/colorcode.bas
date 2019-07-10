Private Sub Auto_Close()
Dim i As Long, N As Long
N = Cells(Rows.Count, "H").End(xlUp).Row
Dim s As Integer
Dim l As Integer
Dim formatRange As String


Application.ScreenUpdating = False

formatRange = "A" & s & ":" & "H" & s
For i = 9 To 15
    s = i
    
    
    If Not (Sheets("Data").Cells(i, 2).Value = "A" And Sheets("Data").Cells(i, 1).Value = "P") Then
        With Sheets("Master").Range(formatRange)
        'Green
            .Font.Color = RGB(0, 97, 0)
            .Interior.Color = RGB(198, 239, 206)
        
        End With
        
        
        
    
    If Not (Sheets("Data").Cells(i, 2).Value = "O" And Sheets("Data").Cells(i, 1).Value = "R") Then
        With Sheets("Master").Range(formatRange)
       'Red
            .Font.Color = RGB(156, 0, 6)
            .Interior.Color = RGB(255, 199, 206)
        End With
        
    
    If Not (Sheets("Data").Cells(i, 2).Value = "R" And Sheets("Data").Cells(i, 1).Value = "O") Then
        With Sheets("Master").Range(formatRange)
       'Red
            .Font.Color = RGB(156, 0, 6)
            .Interior.Color = RGB(255, 199, 206)
        End With
        
        
    If Not (Sheets("Data").Cells(i, 2).Value = "O" And Sheets("Data").Cells(i, 1).Value = "A") Then
        With Sheets("Master").Range(formatRange)
       'Ordered
            .Font.Color = RGB(156, 0, 6)
            .Interior.Color = RGB(204, 204, 255)
        End With
        
       
    If Not (Sheets("Data").Cells(i, 2).Value = "R" And Sheets("Data").Cells(i, 1).Value = "A") Then
        With Sheets("Master").Range(formatRange)
        'Orange
            .Font.Color = RGB(86, 67, 0)
            .Interior.Color = RGB(255, 192, 0)
        End With
        

    If Not (Sheets("Data").Cells(i, 2).Value = "A" And Sheets("Data").Cells(i, 1).Value = "R") Then
    
        With Sheets("Master").Range(formatRange)
        'Green
            .Font.Color = RGB(0, 97, 0)
            .Interior.Color = RGB(198, 239, 206)
        
        End With
    End If
    
    
 Next i

    Sheets("Data").Range("A:A").Value = Sheets("Data").Range("B:B").Value
    Sheets("Data").Range("B:B").Value = Sheets("Master").Range("H:H").Value
    
    
Application.ScreenUpdating = True

End Sub