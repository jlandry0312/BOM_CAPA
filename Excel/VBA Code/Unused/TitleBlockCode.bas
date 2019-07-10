Sub TitleBlock()
    Dim docName As String
    Dim docRev As Range
    Dim jobName As Range
    Dim StartingPosition As Integer
    
    Set jobName = Range("D2")
    Set docRev = Range("E2")
    docName = ActiveWindow.Caption
    

    jobName.Value = docName
    
    jobName = Replace(jobName, ".xlsm", "")
    jobName = Replace(jobName, "[Checked Out]", "")
    docRev.Value = jobName
    
    
    jobPosition = InStr(1, jobName, "[", vbBinaryCompare)
    docRev = Right(docRev, Len(docRev) - (jobPosition - 1))
    jobName.Value = Left(jobName, Len(jobName) - (jobPosition + 1))
    
    
End Sub