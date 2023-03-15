VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet24"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
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
    
End Sub

Sub SetWellPropertyRest(ByVal wselect As Integer, ByVal w3page As Integer)
    Dim firstwell As Integer
      
    firstwell = 3 * w3page + 1
    
    If wselect = 2 Then
        ActiveSheet.Shapes.Range(Array("CommandButton3")).Select
        Selection.Delete
        
        ActiveSheet.Range("D12") = "W-" & CStr(firstwell)
        ActiveSheet.Range("D29") = "W-" & CStr(firstwell)
        
        ActiveSheet.Range("G12") = "W-" & CStr(firstwell + 1)
        ActiveSheet.Range("H29") = "W-" & CStr(firstwell + 1)
    Else
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
    Dim n_sheets As Integer
    Dim result() As Integer
    Dim w3page, wselect, restpage As Integer
    'wselect = 1 --> only w1
    'wselect = 2 --> w1, w2
    
    n_sheets = sheets_count()
    result = DivideWellsBy3(n_sheets)
    
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


Private Sub CommandButton2_Click()
  
  Dim thisname, fname1, fname2, fname3 As String
  Dim cell1, cell2, cell3 As String
  Dim time1 As Date
  Dim bTemp, ec1, ec2, ec3, ph1, ph2, ph3 As Double
  
  
  
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
  
  ec2 = Workbooks(fname2).Worksheets("w1").Range("c8").value
  ph2 = Workbooks(fname2).Worksheets("w1").Range("c9").value
  
  ec3 = Workbooks(fname3).Worksheets("w1").Range("c8").value
  ph3 = Workbooks(fname3).Worksheets("w1").Range("c9").value
  '------------------------------------------------------------------------
  
  
  Range("c6").value = time1
  Range("c7").value = bTemp
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


Sub getModDataFromYangSooTripple(ByVal thisname As String, ByVal fname As String)

    Dim f As Integer

    f = CInt(GetNumeric2(fname)) Mod 3

    Windows(fname).Activate
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



