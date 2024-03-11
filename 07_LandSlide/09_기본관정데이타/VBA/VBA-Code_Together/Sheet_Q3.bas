Private Sub CommandButton1_Click()
UserFormTS.Show
End Sub


Function DivideWellsBy3(ByVal numberOfWells As Integer) As Integer()

    Dim quotient As Integer
    Dim remainder As Integer
    Dim result(1) As Integer
    
    quotient = numberOfWells \ 3
    remainder = numberOfWells Mod 3
    
    result(0) = quotient
    result(1) = remainder
    
    DivideWellsBy3 = result
    
End Function

Sub SetWellPropertyQ3(ByVal i As Integer)
    
    ActiveSheet.Range("D12") = "W-" & CStr((i - 1) * 3 + 1)
    ActiveSheet.Range("D29") = "W-" & CStr((i - 1) * 3 + 1)
    
    ActiveSheet.Range("G12") = "W-" & CStr((i - 1) * 3 + 2)
    ActiveSheet.Range("H29") = "W-" & CStr((i - 1) * 3 + 2)
    
    ActiveSheet.Range("J12") = "W-" & CStr((i - 1) * 3 + 3)
    ActiveSheet.Range("L29") = "W-" & CStr((i - 1) * 3 + 3)
    
    ActiveSheet.Shapes.Range(Array("CommandButton3")).Select
    Selection.Delete
    
    ActiveSheet.Shapes.Range(Array("CommandButton4")).Select
    Selection.Delete
    
End Sub

Sub SetWellPropertyRest(ByVal wselect As Integer, ByVal w3page As Integer)
    Dim firstwell As Integer
      
    firstwell = 3 * w3page + 1
    
    If wselect = 2 Then
        ActiveSheet.Shapes.Range(Array("CommandButton3")).Select
        Selection.Delete
        ActiveSheet.Shapes.Range(Array("CommandButton4")).Select
        Selection.Delete
        
        
        ActiveSheet.Range("D12") = "W-" & CStr(firstwell)
        ActiveSheet.Range("D29") = "W-" & CStr(firstwell)
        
        ActiveSheet.Range("G12") = "W-" & CStr(firstwell + 1)
        ActiveSheet.Range("H29") = "W-" & CStr(firstwell + 1)
    Else
    
        ActiveSheet.Shapes.Range(Array("CommandButton4")).Select
        Selection.Delete
        ActiveSheet.Shapes.Range(Array("CommandButton5")).Select
        Selection.Delete
    
        ActiveSheet.Range("D12") = "W-" & CStr(firstwell)
        ActiveSheet.Range("H12") = "W-" & CStr(firstwell)
    End If
End Sub

Sub DuplicateQ3Page(ByVal n As Integer)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Q3")
    
    For i = 1 To n
        ws.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)
        ActiveSheet.name = "p" & i
        
        With ActiveSheet.Tab
            .ThemeColor = xlThemeColorAccent3
            .TintAndShade = 0
        End With
        
        Call SetWellPropertyQ3(i)
    Next i
        
End Sub

Sub DuplicateRest(ByVal wselect As Integer, ByVal w3page As Integer)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Q" & CStr(wselect))
    
    ws.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)
    ActiveSheet.name = "p" & CStr(w3page + 1)
    
    With ActiveSheet.Tab
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = 0
    End With
        
    Call SetWellPropertyRest(wselect, w3page)
    
End Sub

'
' 2023/3/15, make summary page
' i think, this procedure made before .. but ican't the source excel file ..
' so i make again
' get page from pagenum
' quotient, remainder
' pagenum - 7 : 7/3 - 2, 1 (if remainder 1 = w1, remain = 2, w2)
'
Private Sub CommandButton3_Click()
' make summary page

    Dim n_sheets As Integer
    Dim result() As Integer
    Dim w3page, wselect, restpage As Integer
    'wselect = 1 --> only w1
    'wselect = 2 --> w1, w2
    
    
    If IsSheet("p1") Then
        MsgBox "Sheet P1 Exist .... Delete First ... ", vbOKOnly
        Exit Sub
    End If
       
    n_sheets = sheets_count()
    result = DivideWellsBy3(n_sheets)
    
    
    ' result(0) = quotient
    ' result(1) = remainder
    w3page = result(0)
    
    Select Case result(1)
        Case 0
            restpage = 0
            wselect = 0
            
        Case 1
            restpage = 1
            wselect = 1
            
        Case 2
            restpage = 1
            wselect = 2
    End Select
    
    
    Call DuplicateQ3Page(w3page)
    
    If restpage = 0 Then
        Exit Sub
    Else
        Call DuplicateRest(wselect, w3page)
    End If
    
End Sub

'
' 2024/2/28, Get Water Spec From YangSoo
'

Private Sub CommandButton2_Click()
  Dim thisname, fname1, fname2, fname3 As String
  Dim cell1, cell2, cell3 As String
  Dim time1 As Date
  Dim bTemp, bTemp2, bTemp3, ec1, ec2, ec3, ph1, ph2, ph3 As Double
  
  cell1 = Range("d12").value
  cell2 = Range("g12").value
  cell3 = Range("j12").value
  
  
  thisname = ActiveWorkbook.name
  fname1 = "A" & GetNumeric2(cell1) & "_ge_OriginalSaveFile.xlsm"
  fname2 = "A" & GetNumeric2(cell2) & "_ge_OriginalSaveFile.xlsm"
  fname3 = "A" & GetNumeric2(cell3) & "_ge_OriginalSaveFile.xlsm"
   
  If Not IsWorkBookOpen(fname1) Then
    MsgBox "Please open the yangsoo data ! " & fname1
    Exit Sub
  End If
  
  If Not IsWorkBookOpen(fname2) Then
    MsgBox "Please open the yangsoo data ! " & fname2
    Exit Sub
  End If
  
  If Not IsWorkBookOpen(fname3) Then
    MsgBox "Please open the yangsoo data ! " & fname3
    Exit Sub
  End If
  
  'Range("k2") = fname1
  'Range("k3") = fname2
  'Range("k4") = fname3
  
  '------------------------------------------------------------------------
  time1 = Workbooks(fname1).Worksheets("w1").Range("c6").value
  
  bTemp = Workbooks(fname1).Worksheets("w1").Range("c7").value
  ec1 = Workbooks(fname1).Worksheets("w1").Range("c8").value
  ph1 = Workbooks(fname1).Worksheets("w1").Range("c9").value
  
  
  bTemp2 = Workbooks(fname2).Worksheets("w1").Range("c7").value
  ec2 = Workbooks(fname2).Worksheets("w1").Range("c8").value
  ph2 = Workbooks(fname2).Worksheets("w1").Range("c9").value
  
  bTemp3 = Workbooks(fname3).Worksheets("w1").Range("c7").value
  ec3 = Workbooks(fname3).Worksheets("w1").Range("c8").value
  ph3 = Workbooks(fname3).Worksheets("w1").Range("c9").value
  '------------------------------------------------------------------------
  
  
  Range("c6").value = time1
  Range("c7").value = bTemp
  Range("d7").value = bTemp2
  Range("e7").value = bTemp3
  
  Range("c8").value = ec1
  Range("c9").value = ph1
  
  Range("d8").value = ec2
  Range("d9").value = ph2
  
  Range("e8").value = ec3
  Range("e9").value = ph3
  
  
  Call getModDataFromYangSooTripple(thisname, fname1)
  Call getModDataFromYangSooTripple(thisname, fname2)
  Call getModDataFromYangSooTripple(thisname, fname3)
End Sub


Sub getModDataFromYangSooTripple(ByVal thisname As String, ByVal fName As String)

    Dim f As Integer

    f = CInt(GetNumeric2(fName)) Mod 3

    Windows(fName).Activate
    Sheets("w1").Activate
    Sheets("w1").Range("H14:J23").Select
    Selection.Copy
    
    Windows(thisname).Activate
    
    If f = 0 Then
        Range("l31").Select
    ElseIf f = 1 Then
        Range("d31").Select
    Else
        Range("h31").Select
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

Private Function IsSheet(shname As String) As Boolean
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


Private Sub CommandButton4_Click()
' delete all summary page

    Dim i, nofwell, pn As Integer
    Dim response As VbMsgBoxResult
        
    nofwell = GetNumberOfWell()
    pn = (nofwell \ 3) + (nofwell Mod 3)
    
    Sheets("Q3").Activate
    
    response = MsgBox("Do you deletel all water well?", vbYesNo)
    If response = vbYes Then
        For i = 1 To pn
            Call DeleteWorksheet("p" & i)
        Next i
    End If
    
    Sheets("Q3").Activate
    
End Sub

Private Sub CommandButton5_Click()
' get ec, ph, temp

    Call DataAnalysis
End Sub


' Get(Ec, Ph, Temp) Range - 지열공에서 통계내는 함수 ....
' Ph, EC, Temp statistics, find range
' data gathering function in EarthThermal test ...
Private Sub DataAnalysis()
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



' 1, 2, 3 --> p1
' 4, 5, 6 --> p2

Function getEC(ByVal LOWHI As Integer, ByVal well As Integer)
    Dim page, quo, remainder As Integer
    
    quo = WorksheetFunction.quotient(well - 1, 3)
    remainder = well Mod 3
    page = quo + 1
       
    Select Case remainder
        Case 0
            If LOWHI = cellLOW Then
                getEC = Sheets("p" & CStr(page)).Range("k25").value
            Else
                getEC = Sheets("p" & CStr(page)).Range("k24").value
            End If
            
        Case 1
            If LOWHI = cellLOW Then
                getEC = Sheets("p" & CStr(page)).Range("e25").value
            Else
                getEC = Sheets("p" & CStr(page)).Range("e24").value
            End If
        
        Case 2
            If LOWHI = cellLOW Then
                getEC = Sheets("p" & CStr(page)).Range("h25").value
            Else
                getEC = Sheets("p" & CStr(page)).Range("h24").value
            End If
    End Select
End Function

Function getPH(ByVal LOWHI As Integer, ByVal well As Integer)
    Dim page, quo, remainder As Integer
    
    quo = WorksheetFunction.quotient(well - 1, 3)
    remainder = well Mod 3
    page = quo + 1
       
    Select Case remainder
        Case 0
            If LOWHI = cellLOW Then
                getPH = Sheets("p" & CStr(page)).Range("l25").value
            Else
                getPH = Sheets("p" & CStr(page)).Range("l24").value
            End If
            
        Case 1
            If LOWHI = cellLOW Then
                getPH = Sheets("p" & CStr(page)).Range("f25").value
            Else
                getPH = Sheets("p" & CStr(page)).Range("f24").value
            End If
        
        Case 2
            If LOWHI = cellLOW Then
                getPH = Sheets("p" & CStr(page)).Range("i25").value
            Else
                getPH = Sheets("p" & CStr(page)).Range("i24").value
            End If
    End Select
End Function

Function getTEMP(ByVal LOWHI As Integer, ByVal well As Integer)
    Dim page, quo, remainder As Integer
    
    quo = WorksheetFunction.quotient(well - 1, 3)
    remainder = well Mod 3
    page = quo + 1
       
    Select Case remainder
        Case 0
            If LOWHI = cellLOW Then
                getTEMP = Sheets("p" & CStr(page)).Range("J25").value
            Else
                getTEMP = Sheets("p" & CStr(page)).Range("J24").value
            End If
            
        Case 1
            If LOWHI = cellLOW Then
                getTEMP = Sheets("p" & CStr(page)).Range("d25").value
            Else
                getTEMP = Sheets("p" & CStr(page)).Range("d24").value
            End If
        
        Case 2
            If LOWHI = cellLOW Then
                getTEMP = Sheets("p" & CStr(page)).Range("g25").value
            Else
                getTEMP = Sheets("p" & CStr(page)).Range("g24").value
            End If
    End Select
End Function


