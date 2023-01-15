Attribute VB_Name = "RemoveError"
Option Explicit

  
    
Sub RemoveErrorFromAllSheet()

'
' Dim i, c As Integer
'
'    c = 1
'    For i = 1 To Worksheets.Count
'          Sheets(i).Activate
'           If Sheets(i).Tab.color = vbRed And c <= 3 Then
'                 SHEET_NUMBER(getSheetIndex(Sheets(i).Name)) = i
'                 c = c + 1
'           End If
'    Next i
    
    Dim Sht As Worksheet
         
    For Each Sht In Application.Worksheets
        Sht.Activate
        Call iRemoveError
        Range("m3").Select
        Application.CutCopyMode = False
    Next Sht
End Sub



Private Sub iRemoveError()
    
    Dim rngCell As Range, bError As Byte
    
    
    Range("D4").CurrentRegion.Select
    
    
    For Each rngCell In Selection.Cells
        For bError = 1 To 7 Step 1
            With rngCell
                If .Errors(bError).Value Then
                    .Errors(bError).Ignore = True
                End If
        End With
        Next bError
    Next rngCell

End Sub

Sub SetView_NormalView()
  Dim Sht As Worksheet
         
    For Each Sht In Application.Worksheets
        Sht.Activate
        ActiveWindow.View = xlNormalView
    Next Sht


End Sub

Sub SetView_PageBreakPreview()
  Dim Sht As Worksheet
         
    For Each Sht In Application.Worksheets
        Sht.Activate
        ActiveWindow.View = xlPageBreakPreview
    Next Sht
End Sub






Private Sub testr1c1()
Dim i, j As Integer

        For j = 1 To 10
            Cells(10 + j, 3).FormulaR1C1 = "=r[-1]c[0]"
        Next j


End Sub




Private Sub delete_selected_row()
'
    Dim row As Integer
    
    row = Selection.row
    
    Range(row & ":" & row).Select
    Selection.Delete Shift:=xlUp
End Sub





Private Sub aWorksheet_Change(ByVal target As Range)

    Dim N As Long
    N = target.row
    If Intersect(target, Range("G:G")) Is Nothing Then Exit Sub
    If target.Text <> "Done" Then Exit Sub
    ActiveSheet.Unprotect
      Range("A" & N & ":G" & N).Locked = True
    
    ActiveSheet.Protect

End Sub
















