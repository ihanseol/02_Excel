Attribute VB_Name = "ZZ_2020_C_and_C"
Option Explicit


Dim nARROW As Long


Sub Main_Function()
    
    Dim original_sheet, new_sheet As String
    
    Dim da As String
    Dim db As String
    Dim dc As String
    Dim dd As String
    Dim de As Long
    
    Dim n5 As Long
    Dim n1 As Long
    
    Dim i As Long
    
    nARROW = 1
   
    original_sheet = ActiveSheet.Name
    Sheets.Add
    new_sheet = ActiveSheet.Name
    Sheets(original_sheet).Activate
        
    For i = 1 To 50000
          
        Call read_one_line(original_sheet, i, da, db, dc, dd, de)
        If is_end(de) Then Exit For
        
        Call get_answer(de, n5, n1)
        
        If n5 = 0 Then
            Call write_single(new_sheet, nARROW, da, db, dc, dd, de)
            nARROW = nARROW + 1
        Else
            Call write_multi(new_sheet, n5, n1, da, db, dc, dd, de)
        End If
    
    Next i
    
    Sheets(new_sheet).Activate
    Call decorate
    Call decorate_number
    
   
End Sub

Sub Set_Number()
    
    Dim original_sheet, new_sheet As String
     
    Dim last_row, i  As Long
    
    
    
    original_sheet = ActiveSheet.Name
    last_row = lastRowByKey("A1")
          
    For i = 1 To last_row
        Cells(i, "a").Value = i
    Next i
    
    Columns("A:A").Select
    Selection.Font.Bold = True

End Sub



Private Function lastRowByKey(cell As String) As Long

    lastRowByKey = Range(cell).End(xlDown).Row

End Function


Private Function lastRowByFind() As Long
    Dim LASTROW As Long
    
    LASTROW = Cells.Find("*", searchorder:=xlByRows, searchdirection:=xlPrevious).Row
    
    lastRowByFind = LASTROW
End Function



Private Function is_end(de As Long) As Boolean

    If de = 0 Then
        is_end = True
    Else
        is_end = False
    End If

End Function



Private Sub read_one_line(ByVal sht As String, ByVal nrow As Long, ByRef da As String, ByRef db As String, ByRef dc As String, ByRef dd As String, ByRef de As Long)

    'Range("E21").Formula = "=Well!" & Cells(nsheet, "I").Address
    
    da = Sheets(sht).Cells(nrow, "a").Value
    db = Sheets(sht).Cells(nrow, "b").Value
    dc = Sheets(sht).Cells(nrow, "c").Value
    dd = Sheets(sht).Cells(nrow, "d").Value
    de = Sheets(sht).Cells(nrow, "e").Value

End Sub


Private Sub write_single(ByVal sht As String, ByVal nrow As Long, ByVal da As String, ByVal db As String, ByVal dc As String, ByVal dd As String, ByVal de As Long)

    Sheets(sht).Cells(nrow, "a").Value = da
    Sheets(sht).Cells(nrow, "b").Value = db
    Sheets(sht).Cells(nrow, "c").Value = dc
    Sheets(sht).Cells(nrow, "d").Value = dd
    Sheets(sht).Cells(nrow, "e").Value = de
    

End Sub

Private Sub write_multi(ByVal sht As String, ByVal n5 As Long, ByVal n1 As Long, ByVal da As String, ByVal db As String, ByVal dc As String, ByVal dd As String, ByVal de As Long)
    
    Dim i As Integer
    
    For i = 1 To n5
        Call write_single(sht, nARROW, da, db, dc, dd, 50000)
        nARROW = nARROW + 1
    Next i
    
    If n1 <> 0 Then
        Call write_single(sht, nARROW, da, db, dc, dd, n1)
        nARROW = nARROW + 1
    End If

End Sub



Private Sub get_answer(ByRef nval As Long, ByRef n5 As Long, ByRef n1 As Long)

    'Dim n5, n1 As Long
    
    n5 = Int(nval / 50000)
    n1 = Int(nval Mod 50000)

End Sub

Private Sub decorate_number()

    Range("A:A").NumberFormat = "00000"

End Sub


Private Sub decorate()

    Range("E1").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.Font.Bold = True
    Selection.Style = "Comma [0]"
  With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    
    Range("B1").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("H15").Select
    
    Columns("E:E").ColumnWidth = 14.25
End Sub

