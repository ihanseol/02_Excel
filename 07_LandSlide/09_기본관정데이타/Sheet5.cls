VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub CommandButton1_Click()
    Call CopyOneSheet
End Sub

'Jojung Button
'add new feature - correct border frame ...
Private Sub CommandButton2_Click()
    Dim nofWell As Integer

    nofWell = sheets_count()
    Call JojungSheetData
    Call make_wellstyle
    Call DecorateWellBorder(nofWell)
End Sub


Private Sub DecorateWellBorder(ByVal nofWell As Integer)
    Sheets("Well").Activate
    Range("A2:R" & CStr(nofWell + 3)).Select
    
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlDot
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlDot
        .Weight = xlThin
    End With
    
    Range("D15").Select
End Sub


Private Sub getDuoSolo(ByVal nofWell As Integer, ByRef nDuo As Integer, ByRef nSolo As Integer)
    Dim page, quotient, remainder As Integer
    
    quotient = WorksheetFunction.quotient(nofWell, 2)
    remainder = nofWell Mod 2
    
    If remainder = 0 Then
        nDuo = quotient
        nSolo = 0
    Else
        nDuo = quotient
        nSolo = 1
    End If

End Sub


Private Sub CommandButton4_Click()
    Sheets("Aggregate1").Visible = True
    Sheets("Aggregate1").Select
End Sub

Private Sub CommandButton5_Click()
    Sheets("Aggregate2").Visible = True
    Sheets("Aggregate2").Select
End Sub


'one button
'delete all well except for one ...

Private Sub CommandButton6_Click()
    Dim i, nofWell As Integer
    
    nofWell = GetNumberOfWell()
    
    For i = 2 To nofWell
         RemoveSheetIfExists (CStr(i))
    Next i
    
    Sheets("Well").Activate
    Rows("5:" & CStr(nofWell + 3)).Select
    Selection.Delete Shift:=xlUp
    
    For i = 1 To 12
        If Not RemoveSheetIfExists("p" & CStr(i)) Then Exit For
    Next i
    
    Call DecorateWellBorder(1)
    Range("A1").Select
End Sub

Function RemoveSheetIfExists(shname As String) As Boolean
    Dim ws As Worksheet
    Dim sheetExists As Boolean
    
    sheetExists = False

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(shname)
    If Not ws Is Nothing Then sheetExists = True
    On Error GoTo 0

    If sheetExists Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
        RemoveSheetIfExists = True
        Exit Function
    Else
        RemoveSheetIfExists = False
        Exit Function
    End If
End Function


Private Sub CommandButton7_Click()
    Sheets("aggWhpa").Visible = True
    Sheets("aggWhpa").Select
End Sub

Private Sub Worksheet_Activate()
    Call InitialSetColorValue
End Sub


