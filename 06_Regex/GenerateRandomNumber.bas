Attribute VB_Name = "GenerateRandomNumber"
Option Explicit

Sub GenerateRandomNumber()

    Dim i As Integer
    Dim result(1 To 10) As Integer
    
    Randomize
    For i = 1 To 10
        Debug.Print Int(Rnd * 3) + 1
    Next i

End Sub


Sub GenerateSoilHardness()

    Call fillData(Range("B35"))
    Call fillData(Range("E35"))
    Call fillData(Range("H35"))

End Sub


Sub fillData(ByVal rg As Range)
    
    Dim a As Variant, i As Integer, j As Integer
    Dim targetNumber As Double, t As Integer
    Dim r(1 To 10) As Double, x As Double, sum As Integer
    
    a = ProduceUniqRandom: sum = 0
    targetNumber = rg.Value
    
    't = Application.WorksheetFunction.RoundDown(targetNumber, 0)
    
    t = Fix(targetNumber)
        
    For i = 1 To 3
        j = a(i)
        r(j) = t
        sum = sum + t
    Next i
    
    Randomize
    For i = 4 To 6
        j = a(i)
        r(j) = t + (Int(Rnd * 3) + 1)
        sum = sum + r(j)
    Next i
    
    Randomize
    For i = 7 To 8
        j = a(i)
        r(j) = t + (Int(Rnd * 2) + 1)
        sum = sum + r(j)
    Next i
        
    x = (targetNumber * 10 - CDbl(sum)) / 2
    
    If isTheRest(x) Then
        j = a(9)
        r(j) = WorksheetFunction.RoundUp(x, 0)
        j = a(10)
        r(j) = WorksheetFunction.RoundUp(x, 0) - 1
        
        Cells(31, rg.Column).Value = "in"
        Cells(31, rg.Column + 1).Value = x
    Else
        For i = 9 To 10
            j = a(i)
            r(j) = x
        Next i
        
        Cells(31, rg.Column).Value = "out"
        Cells(31, rg.Column + 1).Value = x
    End If
    
    Call resultOut(rg, r)
End Sub


Sub resultOut(ByVal rg As Range, r As Variant)
    Dim i As Integer
        
    For i = 1 To 10
        Cells(20 + i, rg.Column).Value = r(i)
    Next i
End Sub

Function isTheRest(x As Double) As Boolean

    Dim r As Double
    
    r = (x - Fix(x)) * 10
         
    If CInt(r) Then
        isTheRest = True
    Else
        isTheRest = False
    End If

End Function


Function getAverage(r As Variant) As Double

    Dim i As Integer, c As Integer, sum As Double
    
    sum = 0: c = 0
    For i = LBound(r) To UBound(r)
        sum = sum + r(i)
        c = c + 1
    Next i
    
    getAverage = sum / c

End Function


Function ProduceUniqRandom() As Variant

    Dim myStart As Long, myEnd As Long, i As Long
    Dim a()
    Dim sh As Worksheet
    
    Set sh = ActiveSheet
    myStart = 1: myEnd = 10
    
    ReDim a(1 To myEnd - myStart + 1)
        
    With CreateObject("System.Collections.SortedList")
        Randomize
        For i = myStart To myEnd
            .item(Rnd) = i
        Next i
        For i = 1 To .Count
             a(i) = .GetByIndex(i - 1)
        Next
    End With
    
    'sh.Range("A5").Resize(UBound(a) + 1).Value = Application.Transpose(a)
    ProduceUniqRandom = a
    
End Function



