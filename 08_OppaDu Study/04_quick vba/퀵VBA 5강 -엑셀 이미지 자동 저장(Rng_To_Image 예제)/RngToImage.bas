Option Explicit
 
Sub Test()
 
Rng_To_Image Selection
 
End Sub
 
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
 
Sub Rng_To_Image(rngSelection As Range, _
                Optional FileName As String = "엑셀이미지", _
                Optional SavePath As String = "", _
                Optional AddSequence As Boolean = True)
 
 
Dim NewWs As Worksheet
Dim picRange As Object: Dim MyObj As Chart
Dim PicH As Double: Dim PicW As Double
Dim FilePath As String
 
If ValidFileName(FileName) = False Then MsgBox "올바른 파일명을 사용하세요.": End
 
If SavePath = "" Then SavePath = GetDesktopPath
FilePath = SavePath & FileName & ".png"
 
rngSelection.CopyPicture xlScreen, xlPicture
 
Set NewWs = ActiveWorkbook.Sheets.Add
NewWs.Paste
 
Set picRange = NewWs.Shapes.Item(1)
With picRange
    PicH = .Height
    PicW = .Width
    .Delete
End With
 
With NewWs.Shapes.AddChart2
    .Height = PicH
    .Width = PicW
End With
 
Set MyObj = NewWs.Shapes.Item(1).Chart
 
MyObj.ChartArea.Select
MyObj.Paste
 
If AddSequence = True Then
    FilePath = FileSequence(FilePath, 1)
End If
 
MyObj.Export FilePath, "PNG"
 
Application.DisplayAlerts = False
NewWs.Delete
Application.DisplayAlerts = True
 
End Sub
 
Public Function FileExists(ByVal path_ As String) As Boolean
 
    FileExists = (Dir(path_, vbDirectory) <> "")
 
End Function
 
Public Function GetDesktopPath(Optional BackSlash As Boolean = True) As String
 
Dim oWSHShell As Object
 
Set oWSHShell = CreateObject("WScript.Shell")
 
If BackSlash = True Then
    GetDesktopPath = oWSHShell.SpecialFolders("Desktop") & "\"
Else
    GetDesktopPath = oWSHShell.SpecialFolders("Desktop")
End If
 
Set oWSHShell = Nothing
 
End Function
 
Function FileSequence(FilePath As String, Optional Sequence As Long = 1) As String
 
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