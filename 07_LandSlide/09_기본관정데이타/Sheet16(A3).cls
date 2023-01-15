VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
UserFormTS.Show
End Sub


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
    Sheets("w1").Range("H14:J23").Select
    Selection.Copy
    
    Windows(thisname).Activate
    Range("h14").Select
   
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
End Sub

' Ph, EC, Temp statistics, find range

Private Sub CommandButton3_Click()

    Dim nOfWell, i As Integer
    
    Dim lowEC() As Double
    Dim hiEC() As Double
    
    Dim lowPH() As Double
    Dim hiPH() As Double
    
    Dim lowTEMP() As Double
    Dim hiTEMP() As Double
    
    
    Sheets("Well").Select
    Sheets("Well").Range("A30").Select
    Selection.End(xlUp).Select
    
    nOfWell = GetNumeric2(Selection.value)
    
    
    ReDim lowPH(1 To nOfWell)
    ReDim hiPH(1 To nOfWell)
    
    ReDim lowEC(1 To nOfWell)
    ReDim hiEC(1 To nOfWell)
    
    ReDim lowTEMP(1 To nOfWell)
    ReDim hiTEMP(1 To nOfWell)
    
    
    For i = 1 To nOfWell
        lowEC(i) = getEC(cellLOW, i)
        hiEC(i) = getEC(cellHI, i)
        
        lowPH(i) = getPH(cellLOW, i)
        hiPH(i) = getPH(cellHI, i)
        
        lowTEMP(i) = getTEMP(cellLOW, i)
        hiTEMP(i) = getTEMP(cellHI, i)
    Next i
    
    
    Debug.Print "--Temp----------------------------------------"
    Debug.Print "low : " & Application.Min(lowTEMP), Application.Max(lowTEMP)
    Debug.Print "hi  : " & Application.Min(hiTEMP), Application.Max(hiTEMP)
    Debug.Print "----------------------------------------------"
    
    Debug.Print "--PH------------------------------------------"
    Debug.Print "low : " & Application.Min(lowPH), Application.Max(lowPH)
    Debug.Print "hi  : " & Application.Min(hiPH), Application.Max(hiPH)
    Debug.Print "----------------------------------------------"
       
    Debug.Print "--EC------------------------------------------"
    Debug.Print "low : " & Application.Min(lowEC), Application.Max(lowEC)
    Debug.Print "hi  : " & Application.Min(hiEC), Application.Max(hiEC)
    Debug.Print "----------------------------------------------"
    
   
   
End Sub

Function getEC(ByVal LOWHI As Integer, ByVal well As Integer)
    Dim page, quo, remainder As Integer
    
    quo = WorksheetFunction.Quotient(well, 2)
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
    
    quo = WorksheetFunction.Quotient(well, 2)
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
    
    quo = WorksheetFunction.Quotient(well, 2)
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




























