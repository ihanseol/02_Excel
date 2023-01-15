Attribute VB_Name = "modAutomation"
Option Explicit

Private Function lastRowByKey(cell As String) As Long

    lastRowByKey = Sheets("DB").Range(cell).End(xlDown).row

End Function


Public Sub TestLinest()

    Dim x
    Dim i, row As Long
    Dim evalString As String
    Dim sheetDisplayName As String
    Dim polyOrder As String
    
    sheetDisplayName = Sheet1.name
    i = Range(sheetDisplayName & "!A12").End(xlDown).row
    polyOrder = "{1,2,3}"
    
    evalString = "=linest(" & "J12:J" & i & ", " & "A12:A" & i & "^" & polyOrder & ")"
    x = Application.Evaluate(evalString)
    
    
    row = lastRowByKey("C3") + 1
    
    If row > 10000 Then row = 4
    
    Sheets("DB").Cells(row, "c") = x(1)
    Sheets("DB").Cells(row, "d") = x(2)
    Sheets("DB").Cells(row, "e") = x(3)
    Sheets("DB").Cells(row, "f") = x(4)

End Sub


Function gen_rand(x1 As Double, x2 As Double)

   Dim LRandomNumber As Double

   Randomize
   LRandomNumber = ((x2 - x1 + 1) * Rnd + x1)
   
   gen_rand = LRandomNumber

End Function

Function Power(x As Double, y As Double) As Double
    
    Power = Application.WorksheetFunction.Power(x, y)

End Function

Function water_level(ByVal x As Double) As Double

    Dim a3, a2, a1 As Double
    
'    a3 = 3.5375 * 0.00000001
'    a2 = 2.17349 * 0.00001
'    a1 = 4.41364 * 0.001
    
    a3 = Range("j8").Value * 0.00000001
    a2 = Range("k8").Value * 0.00001
    a1 = Range("l8").Value * 0.001
    
    water_level = (-a3) * Power(x, 3) + a2 * Power(x, 2) - a1 * x + 0.995173

End Function


Function water_level2(ByVal x As Double) As Double

    water_level2 = Exp(-0.0034 * x)

End Function


Sub water_drawdown()
    
    Dim a(5) As Long
    Dim i As Long
    Dim st As Long
    
    st = Range("n5").Value
    ' * gen_rand(1.11, 1.13)
    
    a(0) = Int(gen_rand(st - 50, st + 50))
    ' a(1) = a(0) + Int(gen_rand(55, 75))
    a(1) = a(0) + Int(gen_rand(55, 75) * gen_rand(1.11, 1.13))
    
    a(2) = a(1) + Int(gen_rand(35, 55))
    a(3) = a(2) + Int(gen_rand(25, 35))
    a(4) = a(3) + Int(gen_rand(25, 30))
    
    
    For i = 0 To 4
          Range("c" & (13 + i)).Value = a(i)
    Next

End Sub

Sub reset_formula()

    Range("K13").Select
    ActiveCell.FormulaR1C1 = "=water_level(RC[-10])"
    Range("K14").Select
    ActiveCell.FormulaR1C1 = "=water_level(RC[-10])"
    Range("K15").Select
    ActiveCell.FormulaR1C1 = "=water_level(RC[-10])"
    Range("K16").Select
    ActiveCell.FormulaR1C1 = "=water_level(RC[-10])"
    Range("K17").Select
    ActiveCell.FormulaR1C1 = "=water_level(RC[-10])"
    Range("K18").Select
    
    
    Range("L13:L17").Select
    Selection.Copy
    Range("C13").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("L20").Select
    Application.CutCopyMode = False

End Sub

Sub set_abc()

    Range("j8").Value = gen_rand(2.885, 3.548)
    Range("k8").Value = gen_rand(1.95, 2.17)
    Range("l8").Value = gen_rand(4.313, 4.513)

End Sub

Sub main()

    Dim i As Integer
    
    For i = 1 To 10
        Debug.Print genrand(52962, 56000)
    Next

End Sub
