
Private Sub Workbook_Open()
Dim wb As Workbook

Dim DocTitle As Range
Dim checkCondition As Range
Dim test As Range

Set checkCondition = Range("H2")
Set DocTitle = Range("H1")



DocTitle.Value = Application.Caption

If checkCondition = "Bad" Then

MsgBox ("Revision Control Unavailable, use Excel with Autodesk Vault Add-in")

DocTitle.Value = ""
wb.Sheets(Sheet1).Visible = xlVeryHidden

 
End If

    
End Sub