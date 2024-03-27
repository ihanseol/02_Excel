Private Sub CommandButton1_Click()
UserFormTS.Show
End Sub

Function IsSheet(shname As String) As Boolean
    Dim ws As Worksheet
    
    On Error GoTo ErrorHandler
    
    Set ws = ThisWorkbook.Worksheets(shname)
    
    Application.DisplayAlerts = False
    IsSheet = True
    Application.DisplayAlerts = True
    
    Exit Function
    
ErrorHandler:
    ' MsgBox "An error occurred while trying to delete the worksheet."
    
    IsSheet = False
    Application.DisplayAlerts = True
End Function


Private Sub CommandButton5_Click()
' Get(Ec, Ph, Temp) Range - 지열공에서 통계내는 함수 ....
' Ph, EC, Temp statistics, find range
' data gathering function in EarthThermal test ...

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


Function getEC(ByVal LOWHI As Integer, ByVal Well As Integer)
    Dim page, quo, remainder As Integer
    
    quo = WorksheetFunction.quotient(Well, 2)
    remainder = Well Mod 2
    page = quo + remainder
    
    Sheets("p" & CStr(page)).Activate
    
    If remainder = 1 Then
        If LOWHI = cellLOW Then
            getEC = Sheets("p" & CStr(page)).Range("e25").value
        Else
            getEC = Sheets("p" & CStr(page)).Range("e24").value
        End If
    Else
        If LOWHI = cellLOW Then
            getEC = Sheets("p" & CStr(page)).Range("h25").value
        Else
            getEC = Sheets("p" & CStr(page)).Range("h24").value
        End If
    End If
End Function

Function getPH(ByVal LOWHI As Integer, ByVal Well As Integer)
    Dim page, quo, remainder As Integer
    
    quo = WorksheetFunction.quotient(Well, 2)
    remainder = Well Mod 2
    page = quo + remainder
    
    Sheets("p" & CStr(page)).Activate
    
    If remainder = 1 Then
        If LOWHI = cellLOW Then
            getPH = Sheets("p" & CStr(page)).Range("f25").value
        Else
            getPH = Sheets("p" & CStr(page)).Range("f24").value
        End If
    Else
        If LOWHI = cellLOW Then
            getPH = Sheets("p" & CStr(page)).Range("i25").value
        Else
            getPH = Sheets("p" & CStr(page)).Range("i24").value
        End If
    End If
End Function

Function getTEMP(ByVal LOWHI As Integer, ByVal Well As Integer)
    Dim page, quo, remainder As Integer
    
    quo = WorksheetFunction.quotient(Well, 2)
    remainder = Well Mod 2
    page = quo + remainder
    
    Sheets("p" & CStr(page)).Activate
    
    If remainder = 1 Then
        If LOWHI = cellLOW Then
            getTEMP = Sheets("p" & CStr(page)).Range("d25").value
        Else
            getTEMP = Sheets("p" & CStr(page)).Range("d24").value
        End If
    Else
        If LOWHI = cellLOW Then
            getTEMP = Sheets("p" & CStr(page)).Range("g25").value
        Else
            getTEMP = Sheets("p" & CStr(page)).Range("g24").value
        End If
    End If
End Function






Sub DuplicateQ2Page(ByVal n As Integer)
' n : Q2 page 복사할 회수
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Q2")
    
    For i = 1 To n
        ws.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)
        ActiveSheet.name = "p" & i
        
        With ActiveSheet.Tab
            .ThemeColor = xlThemeColorAccent3
            .TintAndShade = 0
        End With
        
        Call SetWellPropertyQ2(i)
    Next i
End Sub


Sub SetWellPropertyQ2(ByVal i As Integer)
' i : index of well

    ActiveSheet.Range("D12") = "W-" & CStr((i - 1) * 2 + 1)
    ActiveSheet.Range("D29") = "W-" & CStr((i - 1) * 2 + 1)
    
    ActiveSheet.Range("G12") = "W-" & CStr((i - 1) * 2 + 2)
    ActiveSheet.Range("H29") = "W-" & CStr((i - 1) * 2 + 2)
    
    ActiveSheet.Shapes.Range(Array("CommandButton3")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("CommandButton4")).Select
    Selection.Delete
End Sub


Sub SetWellPropertyRest(ByVal w2page As Integer)
    Dim firstwell As Integer
      
    firstwell = 2 * w2page + 1
    
    ActiveSheet.Range("D12") = "W-" & CStr(firstwell)
    ActiveSheet.Range("H12") = "W-" & CStr(firstwell)
    
    ActiveSheet.Shapes.Range(Array("CommandButton3")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("CommandButton4")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("CommandButton5")).Select
    Selection.Delete
End Sub


Sub DuplicateRest(ByVal w2page As Integer)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Q1")
    
    ws.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)
    ActiveSheet.name = "p" & CStr(w2page + 1)
    
    With ActiveSheet.Tab
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = 0
    End With
        
    Call SetWellPropertyRest(w2page)
    
End Sub


Function DivideWellsBy2(ByVal numberOfWells As Integer) As Integer()
    Dim quotient As Integer
    Dim remainder As Integer
    Dim result(1) As Integer
    
    quotient = (numberOfWells - 1) \ 2
    remainder = numberOfWells Mod 2
    
    
    If remainder = 0 Then
        result(0) = quotient + 1
    Else
        result(0) = quotient
    End If
    
    result(1) = remainder
    
    DivideWellsBy2 = result
End Function


Private Sub CommandButton3_Click()
' make summary page

    Dim result() As Integer
    Dim w2page, wselect, restpage As Integer
    'wselect = 1 --> only w1
       
    If IsSheet("p1") Then
        MsgBox "Sheet P1 Exist .... Delete First ... ", vbOKOnly
        Exit Sub
    End If
       
       
       
    result = DivideWellsBy2(sheets_count())
    
    ' result(0) = quotient
    ' result(1) = remainder
    
    w2page = result(0)
    restpage = result(1)
    
    Call DuplicateQ2Page(w2page)
    
    If restpage = 0 Then
        Exit Sub
    Else
        Call DuplicateRest(w2page)
    End If

End Sub


Private Sub CommandButton2_Click()
' get waterspec from yangsoo
  
  Dim thisname, fname1, fname2 As String
  Dim cell1, cell2 As String
  Dim time1 As Date
  Dim bTemp1, bTemp2, ec1, ec2, ph1, ph2 As Double
  
  
  
  cell1 = Range("d12").value
  cell2 = Range("g12").value
  
  thisname = ActiveWorkbook.name
  fname1 = "A" & GetNumeric2(cell1) & "_ge_OriginalSaveFile.xlsm"
  fname2 = "A" & GetNumeric2(cell2) & "_ge_OriginalSaveFile.xlsm"
  
  If Not IsWorkBookOpen(fname1) Then
    MsgBox "Please open the yangsoo data ! " & fname1
    Exit Sub
  End If
  
  If Not IsWorkBookOpen(fname2) Then
    MsgBox "Please open the yangsoo data ! " & fname2
    Exit Sub
  End If
  
  ' Range("k2") = fname1
  ' Range("k3") = fname2
  
  '------------------------------------------------------------------------
  time1 = Workbooks(fname1).Worksheets("w1").Range("c6").value
  bTemp1 = Workbooks(fname1).Worksheets("w1").Range("c7").value
  bTemp2 = Workbooks(fname2).Worksheets("w1").Range("c7").value
  
  ec1 = Workbooks(fname1).Worksheets("w1").Range("c8").value
  ec2 = Workbooks(fname2).Worksheets("w1").Range("c8").value
  
  ph1 = Workbooks(fname1).Worksheets("w1").Range("c9").value
  ph2 = Workbooks(fname2).Worksheets("w1").Range("c9").value
  
  '------------------------------------------------------------------------
  
  
  Range("c6").value = time1
  Range("c7").value = bTemp1
  Range("d7").value = bTemp2
  Range("c8").value = ec1
  Range("c9").value = ph1
  
  Range("d8").value = ec2
  Range("d9").value = ph2
  
  
  Call getModDataFromYangSooDual(thisname, fname1)
  Call getModDataFromYangSooDual(thisname, fname2)
  
  
End Sub


Sub getModDataFromYangSooDual(ByVal thisname As String, ByVal fName As String)

    Dim f As Integer

    f = CInt(GetNumeric2(fName)) Mod 2

    Windows(fName).Activate
    Sheets("w1").Activate
    Sheets("w1").Range("H14:J23").Select
    Selection.Copy
    
    Windows(thisname).Activate
    
    If f = 0 Then
        Range("h31").Select
    Else
        Range("d31").Select
    End If
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
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



Private Sub CommandButton4_Click()
' delete all summary page

    Dim i, nofwell, pn As Integer
    Dim response As VbMsgBoxResult
        
    nofwell = GetNumberOfWell()
    pn = (nofwell \ 2) + (nofwell Mod 2)
    
    Sheets("Q2").Activate
    
    response = MsgBox("Do you deletel all water well?", vbYesNo)
    If response = vbYes Then
        For i = 1 To pn
            Call DeleteWorksheet("p" & i)
        Next i
    End If
    
    Sheets("Q2").Activate
    
End Sub

