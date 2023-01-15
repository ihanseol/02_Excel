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

'Generate Simple Yangsoo
'지열공에서 여러공일때의 간이양수시험에 대한 내역을 만들때 ...
Private Sub CommandButton3_Click()
    Dim nofWell, i, final As Integer
    Dim page, quotient, remainder As Integer
    Dim nDuo As Integer
    Dim nSolo As Integer
    
      
    nofWell = GetNumberOfWell()
    
    quotient = WorksheetFunction.quotient(nofWell, 2)
    remainder = nofWell Mod 2
    page = quotient + remainder
    
    nDuo = 0: nSolo = 0
    
    Call getDuoSolo(nofWell, nDuo, nSolo)
     
    If Contains(Sheets, "a1") Then Exit Sub
    
    For i = 1 To nDuo
        Call makeTwoPage(i, nofWell)
    Next i
    
    If nSolo Then
        Call makeSinglePage(page, nofWell)
    End If
    
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


Private Sub makeTwoPage(ByVal page As Integer, ByVal nofWell As Integer)
    Sheets("Q1").Select
    Sheets("Q1").Copy After:=Sheets(10 + nofWell + page)
    Sheets("Q1 (2)").Select
    Sheets("Q1 (2)").Name = "a" & CStr(page)
    
    If (page > 1) Then
        changeDoubleWellNumber (page)
    End If
End Sub


Private Sub makeSinglePage(ByVal page As Integer, ByVal nofWell As Integer)
    Sheets("w1").Select
    Sheets("w1").Copy After:=Sheets(10 + nofWell + page)
    Sheets("w1 (2)").Select
    Sheets("w1 (2)").Name = "a" & CStr(page)
    
    Sheets("a" & CStr(page)).Range("D12").Select
    ActiveCell.FormulaR1C1 = "W-" & CStr(nofWell)
    Sheets("a" & CStr(page)).Range("H12").Select
    ActiveCell.FormulaR1C1 = "W-" & CStr(nofWell)
    Sheets("a" & CStr(page)).Range("L12").Select
    ActiveCell.FormulaR1C1 = "W-" & CStr(nofWell)
End Sub

Private Sub changeDoubleWellNumber(ByVal page As Integer)
    Sheets("a" & CStr(page)).Range("D12:F12").Select
    ActiveCell.FormulaR1C1 = "W-" & CStr(page * 2 - 1)
    Sheets("a" & CStr(page)).Range("G12:I12").Select
    ActiveCell.FormulaR1C1 = "W-" & CStr(page * 2)
    

    Sheets("a" & CStr(page)).Range("D29:F29").Select
    ActiveCell.FormulaR1C1 = "W-" & CStr(page * 2 - 1)
    Sheets("a" & CStr(page)).Range("H29:J29").Select
    ActiveCell.FormulaR1C1 = "W-" & CStr(page * 2)
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
        Sheets(CStr(i)).Select
        ActiveWindow.SelectedSheets.Delete
    Next i
    
    Sheets("Well").Activate
    Rows("5:" & CStr(nofWell + 3)).Select
    Selection.Delete Shift:=xlUp
    Call DecorateWellBorder(1)
    
    Range("A1").Select
End Sub

Private Sub Worksheet_Activate()
    Call InitialSetColorValue
End Sub


