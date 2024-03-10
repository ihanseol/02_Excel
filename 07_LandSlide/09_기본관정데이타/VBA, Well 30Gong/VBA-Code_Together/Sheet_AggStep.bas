Option Explicit



Private Sub CommandButton1_Click()
'Hide Aggregate

    Sheets("AggStep").Visible = False
    Sheets("Well").Select
End Sub


Private Sub CommandButton2_Click()
'Collect Data

    If ActiveSheet.name <> "AggStep" Then Sheets("AggStep").Select
    Call WriteStepTestData(999, False)
End Sub



Private Sub CommandButton3_Click()
'Single Well Import

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

Call WriteStepTestData(singleWell, True)

End Sub


Private Sub EraseCellData(str_range As String)
    With Range(str_range)
        .value = ""
    End With
End Sub

Private Sub WriteStepTestData(ByVal singleWell As Integer, ByVal isSingleWellImport As Boolean)
' isSingleWellImport = True ---> SingleWell Import
' isSingleWellImport = False ---> AllWell Import
'
' SingleWell --> ImportWell Number
' 999 & False --> 모든관정을 임포트
'


    Dim fName As String
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
    
        ' isSingleWellImport = True ---> SingleWell Import
        ' isSingleWellImport = False ---> AllWell Import
        
        If isSingleWellImport Then
            If i = singleWell Then
                GoTo SINGLE_ITERATION
            Else
                GoTo NEXT_ITERATION
            End If
        End If
        
    
SINGLE_ITERATION:

        fName = "A" & CStr(i) & "_ge_OriginalSaveFile.xlsm"
        If Not IsWorkBookOpen(fName) Then
            MsgBox "Please open the yangsoo data ! " & fName
            Exit Sub
        End If
        
        Q(i) = Workbooks(fName).Worksheets("Input").Range("q64").value
        h(i) = Workbooks(fName).Worksheets("Input").Range("r64").value
        delta_h(i) = Workbooks(fName).Worksheets("Input").Range("s64").value
        qsw(i) = Workbooks(fName).Worksheets("Input").Range("t64").value
        swq(i) = Workbooks(fName).Worksheets("Input").Range("u64").value

        a1(i) = Workbooks(fName).Worksheets("Input").Range("v64").value
        a2(i) = Workbooks(fName).Worksheets("Input").Range("w64").value
        a3(i) = Workbooks(fName).Worksheets("Input").Range("x64").value
        
        Call Write31_StepTestData_Single(a1(i), a2(i), a3(i), Q(i), h(i), delta_h(i), qsw(i), swq(i), i)

NEXT_ITERATION:

    Next i
    
    'Call Write31_StepTestData(a1, a2, a3, Q, h, delta_h, qsw, swq, nofwell)
End Sub


Sub Write31_StepTestData_Single(a1 As Variant, a2 As Variant, a3 As Variant, Q As Variant, h As Variant, delta_h As Variant, qsw As Variant, swq As Variant, i As Integer)
' i : well_index
    
    Call EraseCellData("C5:K36")
    
    Cells(4 + i, "c").value = "W-" & CStr(i)
    
    Cells(4 + i, "d").value = a1
    Cells(4 + i, "e").value = a2
    Cells(4 + i, "f").value = a3

    Cells(4 + i, "g").value = Q
    Cells(4 + i, "h").value = h
    Cells(4 + i, "i").value = delta_h
    Cells(4 + i, "j").value = qsw
    Cells(4 + i, "k").value = swq

End Sub

