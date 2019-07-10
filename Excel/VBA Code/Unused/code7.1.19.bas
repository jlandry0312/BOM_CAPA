Private Sub Titleblock()


Dim DocTitle As Range
Dim checkCondition As Range
Dim docName As String
Dim docRev As Range
Dim jobName As Range
Dim ws As Worksheet



Set checkCondition = Range("H2")
Set DocTitle = Range("H1")



DocTitle.Value = Application.Caption

If checkCondition = "Bad" Then

MsgBox ("Revision Control Unavailable, use Excel with Autodesk Vault Add-in")
ws.Visible = xlSheetVeryHidden

 
End If



Set jobName = Range("D2")
Set docRev = Range("E2")
    
    
'docName = ActiveWindow.Caption
    

jobName.Value = DocTitle
    
jobName = Replace(jobName, ".xlsm", "")
jobName = Replace(jobName, "[Checked Out]", "")
docRev.Value = jobName
    
    
jobPosition = InStr(1, jobName, "[", vbBinaryCompare)
docRev = Right(docRev, Len(docRev) - (jobPosition - 1))
jobName.Value = Left(jobName, Len(jobName) - (jobPosition + 1))

End Sub
   
