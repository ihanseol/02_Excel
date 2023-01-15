Attribute VB_Name = "z_SubModule"
Option Explicit

Public Function FileExists(ByVal path_ As String) As Boolean
    
    FileExists = (Dir(path_, vbDirectory) <> "")

End Function

Public Function GetDesktopPath(Optional BackSlash As Boolean = True)

Dim oWSHShell As Object

Set oWSHShell = CreateObject("WScript.Shell")

If BackSlash = True Then
    GetDesktopPath = oWSHShell.SpecialFolders("Desktop") & "\"
Else
    GetDesktopPath = oWSHShell.SpecialFolders("Desktop")
End If

Set oWSHShell = Nothing

End Function

Function FileSequence(FilePath As String, Optional Sequence As Long = 1)

Dim Ext As String: Dim Path As String: Dim newPath As String
Dim Pnt As Long

Pnt = InStrRev(FilePath, ".")
Path = Left(FilePath, Pnt - 1)
Ext = Right(FilePath, Len(FilePath) - Pnt + 1)

newPath = Path & Sequence & Ext

Do Until FileExists(newPath) = False
    Sequence = Sequence + 1
    newPath = Path & Sequence & Ext
Loop

FileSequence = newPath

End Function

Function ValidFileName(ByVal FileName As String) As Boolean

Dim Arr As Variant: Dim Val As Variant
Dim Pnt As Long

Arr = Array("/", "\", ":", "*", "?", """", "<", ">", "|")

If InStr(1, FileName, ":\") > 0 Then
    Pnt = InStrRev(FileName, "\")
    FileName = Right(FileName, Len(FileName) - Pnt)
    Debug.Print FileName
End If

ValidFileName = True

For Each Val In Arr
    If InStr(1, FileName, Val) > 0 Then ValidFileName = False: Exit Function
Next

End Function

