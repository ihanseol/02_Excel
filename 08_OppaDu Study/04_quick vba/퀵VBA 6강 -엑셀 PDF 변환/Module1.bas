Attribute VB_Name = "Module1"
Option Explicit

Sub Test()

Dim FileName As String

FileName = Sheet4.Range("E3").Value & "-" & Sheet4.Range("H3").Value & "년 " & Sheet4.Range("H4").Value & "월"

Page_Setup Sheet4, FileName, HCenter:=False

Rng_To_Pdf Sheet4.UsedRange, FileName, AddSequence:=False

End Sub

Sub Rng_To_Pdf(rngSelect As Range, _
                Optional FileName As String = "pdf출력", _
                Optional SavePath As String = "", _
                Optional DocProperty As Boolean = True, _
                Optional PrintArea As Boolean = False, _
                Optional OpenPdf As Boolean = False, _
                Optional AddSequence As Boolean = True)

Dim WS As Worksheet
Dim FilePath As String

Set WS = rngSelect.Parent

If SavePath = "" Then SavePath = GetDesktopPath


FilePath = SavePath & FileName & ".pdf"

If ValidFileName(FilePath) = False Then MsgBox ("올바른 파일명을 사용하세요"): Exit Sub

If AddSequence = True Then
    FilePath = FileSequence(FilePath, 1)
End If

rngSelect.ExportAsFixedFormat xlTypePDF, FilePath, xlQualityStandard, DocProperty, PrintArea, , , OpenPdf


End Sub
