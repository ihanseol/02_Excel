VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet_Q1"
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
  
  thisname = ActiveWorkbook.name
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

Function IsSheet(shname As String) As Boolean
    Dim ws As Worksheet
    
    On Error GoTo ErrorHandler
    
    Set ws = ThisWorkbook.Worksheets(shname)
    Application.DisplayAlerts = False
    Application.DisplayAlerts = True
    
    IsSheet = True
    Exit Function
    
ErrorHandler:
    ' MsgBox "An error occurred while trying to delete the worksheet."
    
    IsSheet = False
    Application.DisplayAlerts = True
End Function


' Get(Ec, Ph, Temp) Range - 지열공에서 통계내는 함수 ....
' Ph, EC, Temp statistics, find range
' data gathering function in EarthThermal test ...
Private Sub CommandButton3_Click()
    Dim nofwell, i As Integer
    
    Dim lowEC() As Double
    Dim hiEC() As Double
    Dim lowPH() As Double
    Dim hiPH() As Double
    Dim lowTEMP() As Double
    Dim hiTEMP() As Double

    nofwell = sheets_count()
    
'    If nofwell < 2 Or Not Contains(Sheets, "a1") Then
'        MsgBox "first Generate Simple YangSoo"
'        Exit Sub
'    End If
    
    If Not IsSheet("p1") Then
        MsgBox "First Make Summary Page"
        Exit Sub
    End If
    
 
    ReDim lowPH(1 To nofwell)
    ReDim hiPH(1 To nofwell)
    
    ReDim lowEC(1 To nofwell)
    ReDim hiEC(1 To nofwell)
    
    ReDim lowTEMP(1 To nofwell)
    ReDim hiTEMP(1 To nofwell)
    
    For i = 1 To nofwell
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
    Sheets("p" & well).Activate
    
    If LOWHI = cellLOW Then
        getEC = Sheets("p" & CStr(well)).Range("e25").value
    Else
        getEC = Sheets("p" & CStr(well)).Range("e24").value
    End If
End Function

Function getPH(ByVal LOWHI As Integer, ByVal well As Integer)
    Sheets("p" & CStr(well)).Activate
    
    If LOWHI = cellLOW Then
        getPH = Sheets("p" & CStr(well)).Range("f25").value
    Else
        getPH = Sheets("p" & CStr(well)).Range("f24").value
    End If
    
End Function

Function getTEMP(ByVal LOWHI As Integer, ByVal well As Integer)
    Sheets("p" & CStr(well)).Activate

    If LOWHI = cellLOW Then
        getTEMP = Sheets("p" & CStr(well)).Range("d25").value
    Else
        getTEMP = Sheets("p" & CStr(well)).Range("d24").value
    End If
End Function


Sub DuplicateQ1Page(ByVal n As Integer)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Q1")
    
    
    ws.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)
    ActiveSheet.name = "p" & n
    
    With ActiveSheet.Tab
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = 0
    End With
    
    Call SetWellPropertyQ1(n)
    
End Sub

Sub SetWellPropertyQ1(ByVal i As Integer)
    ActiveSheet.Range("C4") = "W-" & CStr(i)
    ActiveSheet.Range("D12") = "W-" & CStr(i)
    ActiveSheet.Range("H12") = "W-" & CStr(i)
    
    ActiveSheet.Shapes.Range(Array("CommandButton3")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("CommandButton4")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("CommandButton5")).Select
    Selection.Delete
End Sub


Private Sub DeleteWorksheet(shname As String)
    Dim ws As Worksheet
    
    On Error GoTo ErrorHandler
    
    Set ws = ThisWorkbook.Worksheets(shname)
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
    
    Exit Sub
    
ErrorHandler:
    ' MsgBox "An error occurred while trying to delete the worksheet."
    Application.DisplayAlerts = True
End Sub


' make summary page
Private Sub CommandButton4_Click()
    Dim nofwell As Integer
    Dim i As Integer
    
    nofwell = GetNumberOfWell()
    
    For i = 1 To nofwell
        DuplicateQ1Page (i)
    Next i
End Sub


' delete all summary page
Private Sub CommandButton5_Click()

    Dim nofwell As Integer
    Dim i As Integer
    
    nofwell = GetNumberOfWell()
    
    For i = 1 To nofwell
        DeleteWorksheet ("p" & i)
    Next i
    
    Sheets("Q1").Activate

End Sub






















