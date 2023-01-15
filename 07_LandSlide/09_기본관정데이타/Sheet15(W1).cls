VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
UserFormTS.Show
End Sub


'Get Water Spec from YanSoo ilbo
Private Sub CommandButton2_Click()
  Dim thisname, fname As String
  Dim cell  As String
  Dim time As Date
  Dim bTemp, ec1, ph1 As Double
  
  
  cell = Range("d12").value
  
  thisname = ActiveWorkbook.Name
  fname = "A" & GetNumeric2(cell) & "_ge_OriginalSaveFile.xlsm"
 
  If Not IsWorkBookOpen(fname) Then
    MsgBox "Please open the yangsoo data ! " & fname
    Exit Sub
  End If
  
  ' Range("k2") = fname
   
  '------------------------------------------------------------------------
  time = Workbooks(fname).Worksheets("w1").Range("c6").value
  bTemp = Workbooks(fname).Worksheets("w1").Range("c7").value
  
  ec1 = Workbooks(fname).Worksheets("w1").Range("c8").value
  ph1 = Workbooks(fname).Worksheets("w1").Range("c9").value
  
  '------------------------------------------------------------------------
  
  Range("c6").value = time
  Range("c7").value = bTemp
  Range("c8").value = ec1
  Range("c9").value = ph1
    
  Call getModDataFromYangSooSingle(thisname, fname)
End Sub


Sub getModDataFromYangSooSingle(ByVal thisname As String, ByVal fname As String)
    Windows(fname).Activate
    Sheets("w1").Activate
    Sheets("w1").Range("H14:J23").Select
    Selection.Copy
    
    Windows(thisname).Activate
    Range("h14").Select
   
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub

' Get(Ec, Ph, Temp) Range - 지열공에서 통계내는 함수 ....
' Ph, EC, Temp statistics, find range
' data gathering function in EarthThermal test ...
Private Sub CommandButton3_Click()
    Dim nofWell, i As Integer
    
    Dim lowEC() As Double
    Dim hiEC() As Double
    Dim lowPH() As Double
    Dim hiPH() As Double
    Dim lowTEMP() As Double
    Dim hiTEMP() As Double

    nofWell = sheets_count()
    
    If nofWell < 2 Or Not Contains(Sheets, "a1") Then
        MsgBox "first Generate Simple YangSoo"
        Exit Sub
    End If
    
    ReDim lowPH(1 To nofWell)
    ReDim hiPH(1 To nofWell)
    
    ReDim lowEC(1 To nofWell)
    ReDim hiEC(1 To nofWell)
    
    ReDim lowTEMP(1 To nofWell)
    ReDim hiTEMP(1 To nofWell)
    
    For i = 1 To nofWell
        lowEC(i) = getEC(cellLOW, i)
        hiEC(i) = getEC(cellHI, i)
        
        lowPH(i) = getPH(cellLOW, i)
        hiPH(i) = getPH(cellHI, i)
        
        lowTEMP(i) = getTEMP(cellLOW, i)
        hiTEMP(i) = getTEMP(cellHI, i)
    Next i
    
    Debug.Print String(3, vbCrLf)
    
    Debug.Print "--Temp----------------------------------------"
    Debug.Print "low : " & Application.min(lowTEMP), Application.max(lowTEMP)
    Debug.Print "hi  : " & Application.min(hiTEMP), Application.max(hiTEMP)
    Debug.Print "----------------------------------------------"
    
    Debug.Print "--PH------------------------------------------"
    Debug.Print "low : " & Application.min(lowPH), Application.max(lowPH)
    Debug.Print "hi  : " & Application.min(hiPH), Application.max(hiPH)
    Debug.Print "----------------------------------------------"
       
    Debug.Print "--EC------------------------------------------"
    Debug.Print "low : " & Application.min(lowEC), Application.max(lowEC)
    Debug.Print "hi  : " & Application.min(hiEC), Application.max(hiEC)
    Debug.Print "----------------------------------------------"
End Sub

Function getEC(ByVal LOWHI As Integer, ByVal well As Integer)
    Dim page, quo, remainder As Integer
    
    quo = WorksheetFunction.quotient(well, 2)
    remainder = well Mod 2
    page = quo + remainder
    
    Sheets("a" & CStr(page)).Activate
    
    If remainder = 1 Then
        If LOWHI = cellLOW Then
            getEC = Sheets("a" & CStr(page)).Range("e25").value
        Else
            getEC = Sheets("a" & CStr(page)).Range("e24").value
        End If
    Else
        If LOWHI = cellLOW Then
            getEC = Sheets("a" & CStr(page)).Range("h25").value
        Else
            getEC = Sheets("a" & CStr(page)).Range("h24").value
        End If
    End If
End Function

Function getPH(ByVal LOWHI As Integer, ByVal well As Integer)
    Dim page, quo, remainder As Integer
    
    quo = WorksheetFunction.quotient(well, 2)
    remainder = well Mod 2
    page = quo + remainder
    
    Sheets("a" & CStr(page)).Activate
    
    If remainder = 1 Then
        If LOWHI = cellLOW Then
            getPH = Sheets("a" & CStr(page)).Range("f25").value
        Else
            getPH = Sheets("a" & CStr(page)).Range("f24").value
        End If
    Else
        If LOWHI = cellLOW Then
            getPH = Sheets("a" & CStr(page)).Range("i25").value
        Else
            getPH = Sheets("a" & CStr(page)).Range("i24").value
        End If
    End If
End Function

Function getTEMP(ByVal LOWHI As Integer, ByVal well As Integer)
    Dim page, quo, remainder As Integer
    
    quo = WorksheetFunction.quotient(well, 2)
    remainder = well Mod 2
    page = quo + remainder
    
    Sheets("a" & CStr(page)).Activate
    
    If remainder = 1 Then
        If LOWHI = cellLOW Then
            getTEMP = Sheets("a" & CStr(page)).Range("d25").value
        Else
            getTEMP = Sheets("a" & CStr(page)).Range("d24").value
        End If
    Else
        If LOWHI = cellLOW Then
            getTEMP = Sheets("a" & CStr(page)).Range("g25").value
        Else
            getTEMP = Sheets("a" & CStr(page)).Range("g24").value
        End If
    End If
End Function

































