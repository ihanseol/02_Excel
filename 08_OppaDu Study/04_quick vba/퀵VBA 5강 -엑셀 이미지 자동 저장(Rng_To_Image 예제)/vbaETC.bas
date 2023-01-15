Public Function FileExists(ByVal path_ As String) As Boolean
    
    FileExists = (Dir(path_, vbDirectory) <> "")
    
End Function

Sub 파일존재여부확인(Path   As String)
    
    '##################################################################################
    'Path로 입력된 경로에 파일이 존재하는지 여부를 확인하여 안내메세지를 띄웁니다.
    '##################################################################################
    
    If FileExists(Path) = TRUE Then
        MsgBox "해당 파일이 이미 존재합니다." & vbNewLine & _
               "파일경로 : " & vbNewLine & _
               Path
    Else
        MsgBox "해당 파일이 존재하지 않습니다."
    End If
    
End Sub

Public Function GetDesktopPath(Optional BackSlash As Boolean = True) As String
    
    '// 사용자의 바탕화면 경로를 출력합니다.
    
    Dim oWSHShell   As Object
    
    Set oWSHShell = CreateObject("WScript.Shell")
    
    If BackSlash = TRUE Then
        GetDesktopPath = oWSHShell.SpecialFolders("Desktop") & "\"
    Else
        GetDesktopPath = oWSHShell.SpecialFolders("Desktop")
    End If
    
    Set oWSHShell = Nothing
    
End Function

Sub 바탕화면경로_()
    
    MsgBox "현재 사용중인 컴퓨터의 바탕화면 경로는" & vbNewLine & _
           "[[ " & GetDesktopPath & " ]]" & vbNewLine & _
           "입니다."
    
End Sub

Function FileSequence(FilePath As String, Optional Sequence As Long = 1) As String
    
    Dim Ext         As String: Dim Path As String: Dim newPath As String
    Dim Pnt         As Long
    
    Pnt = InStrRev(FilePath, ".")
    Path = Left(FilePath, Pnt - 1)
    Ext = Right(FilePath, Len(FilePath) - Pnt + 1)
    
    newPath = Path & Sequence & Ext
    
    Do Until FileExists(newPath) = FALSE
        Sequence = Sequence + 1
        newPath = Path & Sequence & Ext
    Loop
    
    FileSequence = newPath
    
End Function

Sub 파일저장()
    
    Dim strPath     As String
    Dim newPath     As String
    
    strPath = GetDesktopPath & "복사본.xlsm"
    newPath = FileSequence(strPath, 1)
    
    ThisWorkbook.SaveAs newPath
    
    MsgBox newPath & "경로로 파일저장을 완료하였습니다."
    
End Sub

Function ValidFileName(ByVal FileName As String) As Boolean
    
    Dim Arr         As Variant: Dim Val As Variant
    Dim Pnt         As Long
    
    Arr = Array("/", "\", ":", "*", "?", """", "<", ">", "|")
    
    If InStr(1, FileName, ":\") > 0 Then
        Pnt = InStrRev(FileName, "\")
        FileName = Right(FileName, Len(FileName) - Pnt)
        Debug.Print FileName
    End If
    
    ValidFileName = TRUE
    
    For Each Val In Arr
        If InStr(1, FileName, Val) > 0 Then ValidFileName = False: Exit Function
    Next
    
End Function

Sub 파일이름체크()
    
    If ValidFileName(Sheet1.Range("F2").Value) = TRUE Then
        MsgBox "사용가능한 파일이름입니다."
    Else
        MsgBox "사용 불가한 파일이름입니다.", vbCritical
    End If
    
End Sub