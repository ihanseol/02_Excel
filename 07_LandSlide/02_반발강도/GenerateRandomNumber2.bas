Attribute VB_Name = "GenerateRandomNumber2"
Option Explicit


#If VBA7 Then ' Excel 2010 or later
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)
#Else ' Excel 2007 or earlier
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As Long)
#End If

#If Win64 Then
    Public Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
#Else
    Public Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
#End If

Sub GenerateR1()
    
    Dim rg1, rg2, rg3, rg4 As Range
    Dim r As Variant

    Set rg1 = Range("B39")
    Set rg2 = Range("E39")
    Set rg3 = Range("H39")
    Set rg4 = Range("K39")
    
    r = FillData(rg1)
    Call fittingData(rg1, r)
    
    r = FillData(rg2)
    Call fittingData(rg2, r)
    
    r = FillData(rg3)
    Call fittingData(rg3, r)
    
    r = FillData(rg4)
    Call fittingData(rg4, r)
    
    
End Sub

Sub GenerateR2()

    Dim rg1, rg2, rg3, rg4 As Range
    Dim r As Variant

    Set rg1 = Range("B39")
    Set rg2 = Range("E39")
    Set rg3 = Range("H39")
    Set rg4 = Range("K39")
    
    r = FillData(rg1)
    Call fittingDataSecond(rg1, r)
    
    r = FillData(rg2)
    Call fittingDataSecond(rg2, r)
    
    r = FillData(rg3)
    Call fittingDataSecond(rg3, r)
    
    r = FillData(rg4)
    Call fittingDataSecond(rg4, r)
    
End Sub


Function Sign() As Integer
    Dim i As Integer
    
    If Int(Rnd * 10) Mod 2 Then
        Sign = 1
    Else
        Sign = -1
    End If

End Function

Sub WasteTime(Finish As Long)
    Dim NowTick As Long
    Dim EndTick As Long
 
    EndTick = GetTickCount + (Finish * 1000)
    Do
        NowTick = GetTickCount
        DoEvents
    Loop Until NowTick >= EndTick
End Sub


Sub fittingData(ByVal rg As Range, r As Variant)

    Dim targetNumber As Integer, goalNumber As Integer
    Dim diff As Integer, i As Integer
    
    targetNumber = rg.Value
    goalNumber = Cells(34, rg.Column).Value
    
    diff = goalNumber - targetNumber
            
    For i = 1 To 20
        r(i) = r(i) - diff
    Next
    
    Call ResultOut(rg, r)
    
End Sub


Sub fittingDataSecond(ByVal rg As Range, r As Variant)

    Dim targetNumber As Integer, goalNumber As Integer
    Dim diff As Integer, i As Integer, sum As Long: sum = 0
    Dim ar_temp As Variant
    
    targetNumber = rg.Value
    
    ar_temp = ArrayListSort(r)
                
    For i = 0 To 9
        sum = sum + ar_temp(i)
    Next i
    
    goalNumber = sum / 10
    
    diff = goalNumber - targetNumber
            
    For i = LBound(r) To UBound(r)
        r(i) = r(i) - diff
    Next
    
    Call ResultOut(rg, r)
    Cells(34, rg.Column).Value = targetNumber
    
End Sub


Function FillData(ByVal rg As Range) As Variant
    
    Dim a As Variant, i As Integer, j As Integer
    Dim targetNumber As Double, t As Integer
    Dim r(1 To 20) As Double, x As Double
    Dim r2 As Variant
    
    a = ProduceUniqRandom: targetNumber = rg.Value
    t = Fix(targetNumber)
     
    Randomize
    For i = 1 To 10
        j = a(i)
        r(j) = t + Int(Rnd * 13) * Sign
    Next i
    
    For i = 11 To 20
        j = a(i)
        r(j) = t + Int(Rnd * 13) * Sign
    Next i

    Call ResultOut(rg, r)
    FillData = r
    
End Function


'If i Mod 2 <> 0 Then debug.print "Odd" Else: debug.print "Even"
Sub ResultOut(ByVal rg As Range, r As Variant)
    Dim i As Integer, check As Integer
        
    If LBound(r) Then check = 0 Else: check = 1
        
    For i = LBound(r) To UBound(r)
        Cells(8 + i + check, rg.Column).Value = r(i)
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


Function ProduceUniqRandom() As Variant

    Dim myStart As Long, myEnd As Long, i As Long
    Dim a()
    Dim sh As Worksheet
    
    Set sh = ActiveSheet
    myStart = 1: myEnd = 20
    
    ReDim a(1 To myEnd - myStart + 1)
    
    With CreateObject("System.Collections.SortedList")
        Randomize
        
        For i = myStart To myEnd
            .Item(Rnd) = i
        Next i
        
        For i = 1 To .Count
             a(i) = .GetByIndex(i - 1)
        Next
    End With
    
    'sh.Range("A5").Resize(UBound(a) + 1).Value = Application.Transpose(a)
    ProduceUniqRandom = a
    
End Function

Function ArrayListSort(ByVal SortArray As Variant) As Variant
'https://stackoverflow.com/questions/152319/vba-array-sort-function

    Static ArrayListObj As Object
    Dim i As Long
    Dim LBnd As Long
    Dim UBnd As Long

    LBnd = LBound(SortArray)
    UBnd = UBound(SortArray)

    If ArrayListObj Is Nothing Then
        Set ArrayListObj = CreateObject("System.Collections.ArrayList")
    Else
        ArrayListObj.Clear
    End If
    
    For i = LBnd To UBnd
        ArrayListObj.Add (SortArray(i))
    Next i

    ArrayListObj.Sort
    ArrayListObj.Reverse

    ArrayListSort = ArrayListObj.ToArray
    
End Function





