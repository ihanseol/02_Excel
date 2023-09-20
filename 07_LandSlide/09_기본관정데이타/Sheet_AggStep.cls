VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet_AggStep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit



Private Sub CommandButton1_Click()
    Sheets("AggStep").Visible = False
    Sheets("Well").Select
End Sub


Private Sub CommandButton2_Click()
    If ActiveSheet.name <> "AggStep" Then Sheets("AggStep").Select
    Call WriteStepTestData
End Sub


Private Sub EraseCellData(str_range As String)
    With Range(str_range)
        .value = ""
    End With
End Sub

Private Sub WriteStepTestData()
    Dim fname As String
    Dim nofwell, i As Integer
    
    
    Dim a1() As String
    Dim a2() As String
    Dim a3() As String
    
    Dim Q() As String
    Dim h() As String
    Dim delta_h() As String
    Dim qsw() As String
    Dim swq() As String
    
    nofwell = GetNumberOfWell()
    ' --------------------------------------------------------------------------------------
    ReDim a1(1 To nofwell)
    ReDim a2(1 To nofwell)
    ReDim a3(1 To nofwell)
    
    ReDim Q(1 To nofwell)
    ReDim h(1 To nofwell)
    ReDim delta_h(1 To nofwell)
    ReDim qsw(1 To nofwell)
    ReDim swq(1 To nofwell)
    
    ' --------------------------------------------------------------------------------------
    
    If ActiveSheet.name <> "AggStep" Then Sheets("AggStep").Select
    
    For i = 1 To nofwell
    
        fname = "A" & CStr(i) & "_ge_OriginalSaveFile.xlsm"
        If Not IsWorkBookOpen(fname) Then
            MsgBox "Please open the yangsoo data ! " & fname
            Exit Sub
        End If
        
        Q(i) = Workbooks(fname).Worksheets("Input").Range("q64").value
        h(i) = Workbooks(fname).Worksheets("Input").Range("r64").value
        delta_h(i) = Workbooks(fname).Worksheets("Input").Range("s64").value
        qsw(i) = Workbooks(fname).Worksheets("Input").Range("t64").value
        swq(i) = Workbooks(fname).Worksheets("Input").Range("u64").value

        a1(i) = Workbooks(fname).Worksheets("Input").Range("v64").value
        a2(i) = Workbooks(fname).Worksheets("Input").Range("w64").value
        a3(i) = Workbooks(fname).Worksheets("Input").Range("x64").value
        
    Next i
    
    Call Write31_StepTestData(a1, a2, a3, Q, h, delta_h, qsw, swq, nofwell)
End Sub


Sub Write31_StepTestData(a1 As Variant, a2 As Variant, a3 As Variant, Q As Variant, h As Variant, delta_h As Variant, qsw As Variant, swq As Variant, nofwell As Variant)

    Dim i As Integer
        
    Call EraseCellData("d5:k19")
    
    For i = 1 To nofwell
        Cells(4 + i, "c").value = "W-" & CStr(i)
        
        Cells(4 + i, "d").value = a1(i)
        Cells(4 + i, "e").value = a2(i)
        Cells(4 + i, "f").value = a3(i)
    
        Cells(4 + i, "g").value = Q(i)
        Cells(4 + i, "h").value = h(i)
        Cells(4 + i, "i").value = delta_h(i)
        Cells(4 + i, "j").value = qsw(i)
        Cells(4 + i, "k").value = swq(i)
    Next i
End Sub


