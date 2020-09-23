Attribute VB_Name = "modCommonCode"
Public Function ReadFile(sFilename As String) As String
    Dim fhFile As Integer
    fhFile = FreeFile
    
    On Error Resume Next
    Open sFilename For Input As #fhFile
    If Err Then
        MsgBox "No such file: '" & sFilename & "'" & vbCrLf & Err.Description, vbCritical, "Error"
    End If
    Close #fhFile
    On Error GoTo 0
    
    Open sFilename For Binary As #fhFile
    ReadFile = Space$(LOF(fhFile))
    Get #fhFile, , ReadFile
    Close #fhFile
End Function

Public Function PathFromFile(sPathAndFile As String, sFilename) As String
    Dim nPos As Integer
    'Extract drive
    sFilename = sPathAndFile
    nPos = InStr(1, sFilename, ":")
    If nPos <> 0 Then
        PathFromFile = PathFromFile & Left(sFilename, nPos)
        sFilename = Right(sFilename, Len(sFilename) - nPos)
    End If
    'Extract path
    Do
        nPos = InStr(1, sFilename, "\")
        If nPos = 0 Then Exit Do
        PathFromFile = PathFromFile & Left(sFilename, nPos)
        sFilename = Right(sFilename, Len(sFilename) - nPos)
    Loop
End Function

