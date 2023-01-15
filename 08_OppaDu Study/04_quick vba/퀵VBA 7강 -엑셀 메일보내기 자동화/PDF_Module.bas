Attribute VB_Name = "PDF_Module"

'####################################################
'모듈을 추가한 뒤 복사/붙여넣기 하세요.
'####################################################

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

