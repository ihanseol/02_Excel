
Private Sub CommandButton1_Click()
'Hide Aggregate

    Sheets("YangSoo").Visible = False
    Sheets("Well").Select
End Sub

Private Sub CommandButton2_Click()
'Collect Data
    
    Call GetBaseDataFromYangSoo(999, False)
End Sub

Private Sub EraseCellData(str_range As String)
    With Range(str_range)
        .value = ""
    End With
End Sub



Private Sub CommandButton4_Click()
'single well import

Dim singleWell  As Integer
Dim WB_NAME As String


WB_NAME = GetOtherFileName
'MsgBox WB_NAME
    
'If Workbook Is Nothing Then
'    GetOtherFileName = "Empty"
'Else
'    GetOtherFileName = Workbook.name
'End If
    
If WB_NAME = "Empty" Then
    MsgBox "WorkBook is Empty"
    Exit Sub
Else
    singleWell = CInt(ExtractNumberFromString(WB_NAME))
'   MsgBox (SingleWell)
End If

Call GetBaseDataFromYangSoo(singleWell, True)

End Sub

'<><><><><>><><><><><><><>><><><><><><><>><><><><><><><>><><><><><><><>><><><><><><><>><><><><><><><>><><><><><><><>><><><><><><><>><><><><><><><>><><>
' Code Refactor by OpenAI
'


Sub GetBaseDataFromYangSoo(ByVal singleWell As Integer, ByVal isSingleWellImport As Boolean)
    Dim nofwell As Integer
    Dim i As Integer
    Dim rngString As String

    ' Arrays to store data
    Dim dataArrays As Variant
    dataArrays = Array("natural", "stable", "recover", "delta_h", "Sw", "radius", _
                       "Rw", "well_depth", "casing", "Q", "delta_s", "hp", _
                       "daeSoo", "T1", "T2", "TA", "S1", "S2", "K", "time_", _
                       "shultze", "webber", "jacob", "skin", "er", "ER1", _
                       "ER2", "ER3", "qh", "qg", "sd1", "sd2", "q1", "C", _
                       "B", "ratio", "T0", "S0", "ER_MODE")

    ' Check if all well data should be imported
    nofwell = GetNumberOfWell()
    If Not isSingleWellImport And singleWell = 999 Then
        rngString = "A5:AN" & (nofwell + 5 - 1)
        Call EraseCellData(rngString)
    End If

    ' Loop through each well
    For i = 1 To nofwell
        ' Import data for all wells or only for the specified single well
        If Not isSingleWellImport Or (isSingleWellImport And i = singleWell) Then
            ImportDataForWell i, dataArrays
        End If
    Next i
End Sub

Sub ImportDataForWell(ByVal wellIndex As Integer, ByVal dataArrays As Variant)
    Dim fName As String
    Dim wb As Workbook
    Dim wsInput As Worksheet
    Dim wsSkinFactor As Worksheet
    Dim wsSafeYield As Worksheet
    Dim dataIdx As Integer
    Dim cellOffset As Integer
    Dim dataCell As Range

    ' Open the workbook
    fName = "A" & CStr(wellIndex) & "_ge_OriginalSaveFile.xlsm"
    If Not IsWorkBookOpen(fName) Then
        MsgBox "Please open the yangsoo data! " & fName
        Exit Sub
    End If
    Set wb = Workbooks(fName)

    ' Loop through data arrays and import values
    For dataIdx = LBound(dataArrays) To UBound(dataArrays)
        SetDataArrayValues wb, wellIndex, dataArrays(dataIdx)
    Next dataIdx
    
    ' Close workbook
    ' wb.Close SaveChanges:=False
End Sub

Sub SetDataArrayValues(ByVal wb As Workbook, ByVal wellIndex As Integer, ByVal dataArrayName As String)
    Dim wsInput As Worksheet
    Dim wsSkinFactor As Worksheet
    Dim wsSafeYield As Worksheet
    Dim dataCell As Range
    Dim value As Variant

    Set wsInput = wb.Worksheets("Input")
    Set wsSkinFactor = wb.Worksheets("SkinFactor")
    Set wsSafeYield = wb.Worksheets("SafeYield")

    Select Case dataArrayName
        Case "Q"
            Set dataCell = wsInput.Range("m51")
        Case "hp"
            Set dataCell = wsInput.Range("i48")
        
        
        Case "natural"
            Set dataCell = wsInput.Range("m48")
        Case "stable"
            Set dataCell = wsInput.Range("m49")
        Case "radius"
            Set dataCell = wsInput.Range("m44")
        Case "Rw"
            Set dataCell = wsSkinFactor.Range("e4")
        
        Case "well_depth"
            Set dataCell = wsInput.Range("m45")
        Case "casing"
            Set dataCell = wsInput.Range("i52")
        
        Case "C"
            Set dataCell = wsInput.Range("A31")
         Case "B"
            Set dataCell = wsInput.Range("B31")
        
        
        Case "recover"
            Set dataCell = wsSkinFactor.Range("c10")
        Case "Sw"
            Set dataCell = wsSkinFactor.Range("c11")
        
        Case "delta_h"
            Set dataCell = wsSkinFactor.Range("b16")
        Case "delta_s"
            Set dataCell = wsSkinFactor.Range("b4")
    
        Case "daeSoo"
            Set dataCell = wsSkinFactor.Range("c16")
            
  '--------------------------------------------------------------
  
       Case "T0"
            Set dataCell = wsSkinFactor.Range("d4")
        Case "S0"
            Set dataCell = wsSkinFactor.Range("f4")
       Case "ER_MODE"
            Set dataCell = wsSkinFactor.Range("h10")
                  
        Case "T1"
            Set dataCell = wsSkinFactor.Range("d5")
        Case "T2"
            Set dataCell = wsSkinFactor.Range("h13")
        Case "TA"
            Set dataCell = wsSkinFactor.Range("d16")
            
       Case "S1"
            Set dataCell = wsSkinFactor.Range("e10")
        Case "S2"
            Set dataCell = wsSkinFactor.Range("i16")
        
        Case "K"
            Set dataCell = wsSkinFactor.Range("e16")
        Case "time_"
            Set dataCell = wsSkinFactor.Range("h16")
            
        Case "shultze"
            Set dataCell = wsSkinFactor.Range("c13")
        Case "webber"
            Set dataCell = wsSkinFactor.Range("c18")
        Case "jacob"
            Set dataCell = wsSkinFactor.Range("c23")
                    
                        
       Case "skin"
            Set dataCell = wsSkinFactor.Range("g6")
        Case "er"
            Set dataCell = wsSkinFactor.Range("c8")
            
        Case "ER1"
            Set dataCell = wsSkinFactor.Range("k8")
        Case "ER2"
            Set dataCell = wsSkinFactor.Range("k9")
        Case "ER3"
            Set dataCell = wsSkinFactor.Range("k10")


        Case "qh"
            Set dataCell = wsSafeYield.Range("b13")
        Case "qg"
            Set dataCell = wsSafeYield.Range("b7")
            
        Case "sd1"
            Set dataCell = wsSafeYield.Range("b3")
        Case "sd2"
            Set dataCell = wsSafeYield.Range("b4")
        Case "q1"
            Set dataCell = wsSafeYield.Range("b2")
        Case "ratio"
            Set dataCell = wsSafeYield.Range("b11")
    End Select

    SetCellValueForWell wellIndex, dataCell, dataArrayName
End Sub

Sub SetCellValueForWell(ByVal wellIndex As Integer, ByVal dataCell As Range, ByVal dataArrayName As String)
    Dim wellData As Variant

    wellData = dataCell.value
    
    
    Cells(4 + wellIndex, 1).value = "W-" & wellIndex
    Cells(4 + wellIndex, GetColumnIndex(dataArrayName)).value = wellData
    
    If dataArrayName = "recover" Or dataArrayName = "Sw" Then
        Cells(4 + wellIndex, GetColumnIndex(dataArrayName)).NumberFormat = "0.00"
    ElseIf dataArrayName = "S2" Then
        Cells(4 + wellIndex, GetColumnIndex(dataArrayName)).NumberFormat = "0.0000000"
    ElseIf dataArrayName = "T1" Or dataArrayName = "T2" Or dataArrayName = "TA" Then
        Cells(4 + wellIndex, GetColumnIndex(dataArrayName)).NumberFormat = "0.0000"
    ElseIf dataArrayName = "qh" Then
        Cells(4 + wellIndex, GetColumnIndex(dataArrayName)).NumberFormat = "0."
    ElseIf dataArrayName = "qg" Then
        Cells(4 + wellIndex, GetColumnIndex(dataArrayName)).NumberFormat = "0.00"
    ElseIf dataArrayName = "q1" Or dataArrayName = "sd1" Or dataArrayName = "sd2" Then
        Cells(4 + wellIndex, GetColumnIndex(dataArrayName)).NumberFormat = "0.00"
    ElseIf dataArrayName = "skin" Then
        Cells(4 + wellIndex, GetColumnIndex(dataArrayName)).value = Format(wellData, "0.0000")
    ElseIf dataArrayName = "er" Then
        Cells(4 + wellIndex, GetColumnIndex(dataArrayName)).NumberFormat = "0.0000"
    ElseIf dataArrayName = "ratio" Then
        Cells(4 + wellIndex, GetColumnIndex(dataArrayName)).NumberFormat = "0.0%"
    ElseIf dataArrayName = "T0" Or dataArrayName = "S0" Then
        Cells(4 + wellIndex, GetColumnIndex(dataArrayName)).NumberFormat = "0.0000"
    End If
End Sub

Function GetColumnIndex(ByVal columnName As String) As Integer
    Dim colIndex As Integer

    Select Case columnName
        Case "Q"
            colIndex = 11
        Case "hp"
            colIndex = 13
        
        
        Case "natural"
            colIndex = 2
        Case "stable"
            colIndex = 3
        Case "radius"
            colIndex = 7
        Case "Rw"
            colIndex = 8
        
        Case "well_depth"
            colIndex = 9
        Case "casing"
           colIndex = 10
        
        Case "C"
           colIndex = 32
         Case "B"
            colIndex = 33
        
        
        Case "recover"
            colIndex = 4
        Case "Sw"
            colIndex = 5
        
        Case "delta_h"
            colIndex = 6
        Case "delta_s"
            colIndex = 12
    
        Case "daeSoo"
           colIndex = 14
            
  '--------------------------------------------------------------
  
       Case "T0"
           colIndex = 35
        Case "S0"
           colIndex = 36
       Case "ER_MODE"
           colIndex = 37
                  
        Case "T1"
           colIndex = 15
        Case "T2"
            colIndex = 16
        Case "TA"
           colIndex = 17
            
       Case "S1"
           colIndex = 18
        Case "S2"
            colIndex = 19
        
        Case "K"
           colIndex = 20
        Case "time_"
            colIndex = 21
            
        Case "shultze"
           colIndex = 22
        Case "webber"
            colIndex = 23
        Case "jacob"
            colIndex = 24
                    
                        
       Case "skin"
            colIndex = 25
        Case "er"
            colIndex = 26
            
        Case "ER1"
            colIndex = 38
        Case "ER2"
            colIndex = 39
        Case "ER3"
            colIndex = 40

        Case "qh"
            colIndex = 27
        Case "qg"
            colIndex = 28
            
        Case "sd1"
            colIndex = 30
        Case "sd2"
            colIndex = 31
        Case "q1"
            colIndex = 29
        Case "ratio"
            colIndex = 34
    End Select

    GetColumnIndex = colIndex
End Function


' in here by refctor by  openai
' replace GetBaseDataFromYangSoo Module
'
'<><><><><>><><><><><><><>><><><><><><><>><><><><><><><>><><><><><><><>><><><><><><><>><><><><><><><>><><><><><><><>><><><><><><><>><><><><><><><>><><>

Public Sub MyDebug(sPrintStr As String, Optional bClear As Boolean = False)
   If bClear = True Then
      Application.SendKeys "^g^{END}", True

      DoEvents '  !!! DoEvents is VERY IMPORTANT here !!!

      Debug.Print String(30, vbCrLf)
   End If

   Debug.Print sPrintStr
End Sub


'0 : skin factor, cell, C8
'1 : Re1,         cell, E8
'2 : Re2,         cell, H8
'3 : Re3,         cell, G10

Function DetermineEffectiveRadius(ERMode As String) As Integer
    Dim er, r As String
    
    er = ERMode
    'MsgBox er
    r = Mid(er, 5, 1)
    
    If r = "F" Then
        DetermineEffectiveRadius = erRE0
    Else
        DetermineEffectiveRadius = val(r)
    End If
End Function



Private Sub CommandButton3_Click()
' Formula Button

Dim formula1, formula2 As String
Dim nofwell As String
Dim i As Integer
Dim T, Q, radius, skin, er As Double
Dim T0, S0 As Double
Dim ERMode As String
Dim ER1, ER2, ER3, B, S1 As Double

' Save array to a file
Dim filePath As String
Dim FileNum As Integer



nofwell = GetNumberOfWell()
Sheets("YangSoo").Select


filePath = ThisWorkbook.Path & "\" & ActiveSheet.name & ".csv"
FileNum = FreeFile

Open filePath For Output As FileNum

    
Call MyDebug("Formula SkinFactor ... ", True)

Debug.Print "************************************************************************************************************************************************************************************************"
Print #FileNum, "************************************************************************************************************************************************************************************************"

For i = 1 To nofwell

    T = Format(Cells(4 + i, "o").value, "0.0000")
    Q = Cells(4 + i, "k").value
    
    T0 = Format(Cells(4 + i, "AI").value, "0.0000")
    S0 = Format(Cells(4 + i, "AJ").value, "0.0000")
    S1 = Cells(4 + i, "R").value
            
    delta_s = Format(Cells(4 + i, "l").value, "0.00")
    radius = Format(Cells(4 + i, "h").value, "0.000")
    skin = Cells(4 + i, "y").value
    er = Cells(4 + i, "z").value
    
    
    B = Format(Cells(4 + i, "AG").value, "0.0000")
    ER1 = Cells(4 + i, "AL").value
    ER2 = Cells(4 + i, "AM").value
    ER3 = Cells(4 + i, "AN").value
    
    
    Select Case DetermineEffectiveRadius(Cells(4 + i, "AK").value)
    ' 경험식 1번
    Case erRE1
        formula1 = "W-" & i & "호공~~r _{e-" & i & "} `=~ sqrt {{2.25 TIMES  " & T0 & " TIMES  0.0833333} over {" & S0 & " TIMES  10 ^{5.46 TIMES  " & T & " TIMES  " & B & "}}} `=~" & ER1 & "m"
        formula2 = "erRE1, 경험식 1번"
        
    ' 경험식 2번
    Case erRE2
        formula1 = "W-" & i & "호공~~r _{e-" & i & "} `=~ sqrt {{2.25 TIMES  " & T0 & " TIMES  0.0833333} over {" & S0 & " TIMES  10 ^{4 pi TIMES " & T & " TIMES  " & B & "}}} `=~" & ER2 & "m"
        formula2 = "erRE2, 경험식 2번"
    ' 경험식 3번
    Case erRE3
        formula1 = "W-" & i & "호공~~r _{e-" & i & "} `=~" & radius & " TIMES  sqrt {{" & S1 & "} over {" & S0 & "}} `=~" & ER3 & "m"
        formula2 = "erRE3, 경험식 3번"
        
    Case Else
        ' 스킨계수
        formula1 = "W-" & i & "호공~~ sigma  _{w-" & i & "} = {2 pi  TIMES  " & T & " TIMES  " & delta_s & " } over {" & Q & "} -1.15 TIMES  log {2.25 TIMES  " & T & " TIMES  (1/1440)} over {0.0005 TIMES  (" & radius & " TIMES  " & radius & ")} =`" & skin
        ' 유효우물반경
        formula2 = "W-" & i & "호공~~r _{e-" & i & "} `=~r _{w} e ^{- sigma  _{w-" & i & "}} =" & radius & " TIMES e ^{-(" & skin & ")} =" & er & "m"
    End Select
    
    Debug.Print formula1
    Debug.Print formula2
    Debug.Print "-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    
    Print #FileNum, formula1
    Print #FileNum, formula2
    Print #FileNum, "-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    

Next i

Debug.Print "************************************************************************************************************************************************************************************************"
Print #FileNum, "************************************************************************************************************************************************************************************************"


Call FormulaChwiSoo(FileNum)
' 3-7, 적정취수량

Call FormulaRadiusOfInfluence(FileNum)
' 양수영향반경
End Sub


Sub FormulaChwiSoo(FileNum As Integer)
' 3-7, 적정취수량

    Dim formula As String
    Dim nofwell As String
    Dim i As Integer
    Dim q1, S1, S2, res As Double
    
    nofwell = GetNumberOfWell()
    Sheets("YangSoo").Select
        
        
    For i = 1 To 3
        Debug.Print " "
        Print #FileNum, " "
    Next i
    
    Debug.Print "************************************************************************************************************************************************************************************************"
    Print #FileNum, "************************************************************************************************************************************************************************************************"
    
    For i = 1 To nofwell
        q1 = Cells(4 + i, "ac").value
 
        S1 = Format(Cells(4 + i, "ad").value, "0.00")
        S2 = Format(Cells(4 + i, "ae").value, "0.00")
        
        res = Cells(4 + i, "ab").value
        formula = "W-" & i & "호공~~Q _{ & 2} `＝" & q1 & "` TIMES  `(` {" & S2 & "} over {" & S1 & "} `) ^{2/3} `＝" & res & "㎥/일"
        
        Debug.Print formula
        Debug.Print "-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        
        Print #FileNum, formula
        Print #FileNum, "-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        
    Next i
    
    Debug.Print "************************************************************************************************************************************************************************************************"
    Print #FileNum, "************************************************************************************************************************************************************************************************"
    
    
End Sub


Sub FormulaRadiusOfInfluence(FileNum As Integer)
' 양수영향반경

    Dim formula1, formula2, formula3 As String
    ' 슐츠, 웨버, 제이콥
    
    Dim nofwell As String
    Dim i As Integer
    Dim shultze, webber, jacob, T, K, S, time_, delta_h As String
    
    nofwell = GetNumberOfWell()
    Sheets("YangSoo").Select
        
        
    For i = 1 To 3
        Debug.Print " "
        Print #FileNum, " "
    Next i
    
    Debug.Print "************************************************************************************************************************************************************************************************"
    Print #FileNum, "************************************************************************************************************************************************************************************************"
    
    
    For i = 1 To nofwell
        shultze = CStr(Cells(4 + i, "v").value)
        webber = CStr(Cells(4 + i, "w").value)
        jacob = CStr(Cells(4 + i, "x").value)
        
        T = CStr(Format(Cells(4 + i, "q").value, "0.0000"))
        S = CStr(Format(Cells(4 + i, "s").value, "0.0000000"))
        K = CStr(Format(Cells(4 + i, "t").value, "0.0000"))
    
        delta_h = CStr(Cells(4 + i, "f").value)
        time_ = CStr(Format(Cells(4 + i, "u").value, "0.0000"))
        
        
        ' Cells(4 + i, "y").value = Format(skin(i), "0.0000")
        
        formula1 = "W-" & i & "호공~~R _{W-" & i & "} ``=`` sqrt {6 TIMES  " & delta_h & " TIMES  " & K & " TIMES  " & time_ & "/" & S & "} ``=~" & shultze & "m"
        formula2 = "W-" & i & "호공~~R _{W-" & i & "} ``=3`` sqrt {" & delta_h & " TIMES " & K & " TIMES " & time_ & "/" & S & "} `=`" & webber & "`m"
        formula3 = "W-" & i & "호공~~r _{0(W-" & i & ")} `=~ sqrt {{2.25 TIMES  " & T & " TIMES  " & time_ & "} over {" & S & "}} `=~" & jacob & "m"
        
        Debug.Print ""
        Debug.Print formula1
        Debug.Print formula2
        Debug.Print formula3
        Debug.Print ""
        
        Debug.Print "-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        
        
         Print #FileNum, " "
         Print #FileNum, formula1
         Print #FileNum, formula2
         Print #FileNum, formula3
         Print #FileNum, "-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        
    Next i
    
    Debug.Print "************************************************************************************************************************************************************************************************"
    Print #FileNum, "-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    
End Sub








