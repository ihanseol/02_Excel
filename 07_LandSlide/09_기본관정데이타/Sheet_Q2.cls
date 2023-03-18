VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet_Q2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
UserFormTS.Show
End Sub

Function DivideWellsBy2(ByVal numberOfWells As Integer) As Integer()
    Dim quotient As Integer
    Dim remainder As Integer
    Dim result(1) As Integer
    
    quotient = numberOfWells \ 2
    remainder = numberOfWells Mod 2
    
    result(0) = quotient
    result(1) = remainder
    
    DivideWellsBy2 = result
End Function


Sub SetWellPropertyQ2(ByVal i As Integer)
    ActiveSheet.Range("D12") = "W-" & CStr((i - 1) * 2 + 1)
    ActiveSheet.Range("D29") = "W-" & CStr((i - 1) * 2 + 1)
    
    ActiveSheet.Range("G12") = "W-" & CStr((i - 1) * 2 + 2)
    ActiveSheet.Range("H29") = "W-" & CStr((i - 1) * 2 + 2)
    
    ActiveSheet.Shapes.Range(Array("CommandButton3")).Select
    ActiveSheet.Shapes.Range(Array("CommandButton4")).Select
    Selection.Delete
End Sub



Sub DuplicateQ2Page(ByVal n As Integer)
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

Sub SetWellPropertyRest(ByVal w2page As Integer)
    Dim firstwell As Integer
      
    firstwell = 2 * w2page + 1
    
    ActiveSheet.Range("D12") = "W-" & CStr(firstwell)
    ActiveSheet.Range("H12") = "W-" & CStr(firstwell)
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


Private Sub CommandButton3_Click()
' make summary page

    Dim result() As Integer
    Dim w2page, wselect, restpage As Integer
    'wselect = 1 --> only w1
       
    result = DivideWellsBy2(sheets_count())
    
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
  Dim bTemp, ec1, ec2, ph1, ph2 As Double
  
  
  
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
  bTemp = Workbooks(fname1).Worksheets("w1").Range("c7").value
  
  ec1 = Workbooks(fname1).Worksheets("w1").Range("c8").value
  ph1 = Workbooks(fname1).Worksheets("w1").Range("c9").value
  
  ec2 = Workbooks(fname2).Worksheets("w1").Range("c8").value
  ph2 = Workbooks(fname2).Worksheets("w1").Range("c9").value
  
  '------------------------------------------------------------------------
  
  
  Range("c6").value = time1
  Range("c7").value = bTemp
  Range("c8").value = ec1
  Range("c9").value = ph1
  
  Range("d8").value = ec2
  Range("d9").value = ph2
  
  
  Call getModDataFromYangSooDual(thisname, fname1)
  Call getModDataFromYangSooDual(thisname, fname2)
  
  
End Sub


Sub getModDataFromYangSooDual(ByVal thisname As String, ByVal fname As String)

    Dim f As Integer

    f = CInt(GetNumeric2(fname)) Mod 2

    Windows(fname).Activate
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
    
    response = MsgBox("Do you deletel all water well?", vbYesNo)
    If response = vbYes Then
        For i = 1 To pn
            Call DeleteWorksheet("p" & i)
        Next i
    End If
    
    Sheets("Q2").Activate
    
End Sub
