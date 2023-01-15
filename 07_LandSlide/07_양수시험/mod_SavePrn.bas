Attribute VB_Name = "mod_SavePrn"
Global WB_NAME      As String

Public Function MyDocsPath() As String
    MyDocsPath = Environ$("USERPROFILE") & "\" & "Documents"
    Debug.Print MyDocsPath
End Function

Public Function WB_HEAD() As String
    WB_HEAD = MyDocsPath + "\" + Left(ThisWorkbook.name, 5)
    Debug.Print WB_HEAD
End Function

Sub janggi_01()
    ActiveWorkbook.SaveAs Filename:= _
                          WB_HEAD + "_janggi_01.prn", FileFormat _
                          :=xlTextPrinter, CreateBackup:=False
End Sub

Sub janggi_02()
    ActiveWorkbook.SaveAs Filename:= _
                          WB_HEAD + "_janggi_02.prn", FileFormat _
                          :=xlTextPrinter, CreateBackup:=False
End Sub

Sub recover_01()
    Debug.Print WB_HEAD
    ActiveWorkbook.SaveAs Filename:= _
                          WB_HEAD + "_recover_01.prn", FileFormat:= _
                          xlTextPrinter, CreateBackup:=False
End Sub

Sub step_01()
    Range("a1").Select
    
    ActiveWorkbook.SaveAs Filename:= _
                          WB_HEAD + "_step_01.prn", FileFormat:= _
                          xlTextPrinter, CreateBackup:=False
End Sub

Sub save_original()
    ActiveWorkbook.SaveAs Filename:=WB_HEAD + "_OriginalSaveFile", FileFormat:= _
                          xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
End Sub

