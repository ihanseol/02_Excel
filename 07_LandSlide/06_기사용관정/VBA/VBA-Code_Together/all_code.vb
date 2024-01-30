
' ***************************************************************
' ThisWorkbook
'
' ***************************************************************


Private Sub Workbook_Open()
    Sheets("ss").Activate
    ISIT_FIRST = True
    
     Call clearRowA
     
End Sub

' ***************************************************************
' Sheet5_ss(ss)
'
' ***************************************************************




Private Sub combobox_initialize()

'    Dim tbl As ListObject
'    Dim tableNAME, shNAME As String
'
'    Dim cell As Range
'    Dim i As Integer
'    Dim isFirst As Boolean: isFirst = True
'
'
'    If ISIT_FIRST Then
'        comboAREA.Clear
'
'        If chkboxJIYEOL.Value = True Then
'            tableNAME = "tableJIYEOL"
'            shNAME = "ref1"
'        Else
'            tableNAME = "tableCNU"
'            shNAME = "ref"
'        End If
'
'        Set tbl = Sheets(shNAME).ListObjects(tableNAME)
'
'        i = 0
'        For Each cell In tbl.HeaderRowRange.Cells
'            If isFirst Then
'                isFirst = False
'                GoTo NEXT_ITER
'            End If
'
'             comboAREA.AddItem cell.Value
'NEXT_ITER:
'        Next cell
'    Else
'        ISIT_FIRST = False
'    End If
End Sub


Private Sub CommandButton5_Click()
    UserForm_survey.Show
End Sub


Private Sub Worksheet_Activate()
   
End Sub

'Private Sub chkboxJIYEOL_Click()
'    ISIT_FIRST = True
'    Call combobox_initialize
'    ISIT_FIRST = False
'End Sub


Private Sub comboAREA_DropButtonClick()
    'Call combobox_initialize
End Sub

Private Sub comboAREA_GotFocus()
   'Call combobox_initialize
End Sub


Private Sub comboAREA_Change()
    ' Dim selectedHeader As String
    ' selectedHeader = comboAREA.Value
    ' Range("S21").Value = selectedHeader
End Sub


Private Sub CommandButton1_Click()
    Call insertRow
End Sub

Private Sub CommandButton2_Click()
    ' 지하수 이용실태 현장조사표
    ' Groundwater Availability Field Survey Sheet
    
    Call MakeFieldList
    Sheets("ss").Activate
    
End Sub

Private Sub CommandButton3_Click()
    Call ComputeQ
    Sheets("ss").Activate
End Sub

Private Sub CommandButton4_Click()

   If Sheets("ref").Visible Then
        Sheets("ref").Visible = False
        Sheets("ref1").Visible = False
    Else
        Sheets("ref").Visible = True
        Sheets("ref1").Visible = True
    End If
    
End Sub

Private Sub CommandButtonCopy_Click()
    Call MainMoudleGenerateCopy
End Sub

Private Sub CommandButtonDelete_Click()
    Call SubModuleCleanCopySection
End Sub

Private Sub CommandButtonInitialClear_Click()
 Call SubModuleInitialClear
End Sub


Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    
End Sub
' ***************************************************************
' Sheet2_aa(aa)
'
' ***************************************************************



Private Sub CommandButton1_Click()
    Call MainMoudleGenerateCopy
End Sub

Private Sub CommandButton2_Click()
    Call SubModuleCleanCopySection
End Sub

Private Sub CommandButton3_Click()
    Call insertRow
End Sub

Private Sub CommandButton4_Click()
    Call ComputeQ
    Sheets("aa").Activate
End Sub

Private Sub CommandButton5_Click()
    ' 지하수 이용실태 현장조사표
    ' Groundwater Availability Field Survey Sheet
    
    Call MakeFieldList
    Sheets("aa").Activate
End Sub

Private Sub CommandButtonInitialClear_Click()
 Call SubModuleInitialClear
End Sub


Private Sub Worksheet_Activate()
    
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    
End Sub
' ***************************************************************
' Sheet3_ii(ii)
'
' ***************************************************************



Private Sub CommandButton1_Click()
    Call MainMoudleGenerateCopy
End Sub


Private Sub CommandButton2_Click()
    Call SubModuleCleanCopySection
End Sub

Private Sub CommandButton3_Click()
    Call insertRow
End Sub

Private Sub CommandButton4_Click()
 Call SubModuleInitialClear
End Sub

Private Sub Worksheet_Activate()
    
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    
End Sub
' ***************************************************************
' water_q
'
' ***************************************************************


Public SS(1 To 5, 1 To 2) As Double
Public AA(1 To 6, 1 To 2) As Double

Public SS_CITY As Double
Public ISIT_FIRST As Boolean

Public Enum SS_VALUE
    svGAJUNG = 1
    svILBAN = 2
    svSCHOOL = 3
    svGONGDONG = 4
    svMAEUL = 5
End Enum

Public Enum AA_VALUE
    avJEONJAK = 1
    avDAPJAK = 2
    avWONYE = 3
    avCOW = 4
    avPIG = 5
    avCHICKEN = 6
End Enum

Function CheckBoxFind(objNAME As String) As MSForms.CheckBox
    Dim ws As Worksheet
    Dim obj As OLEObject
    Dim myCheckBox As MSForms.CheckBox
    
    
    Set ws = ThisWorkbook.Worksheets("ss")
    Set myCheckBox = Nothing
    
    For Each obj In ws.OLEObjects
        If TypeOf obj.Object Is MSForms.CheckBox Then
            If obj.Name = objNAME Then
                Set myCheckBox = obj.Object
                Exit For
            End If
        End If
    Next obj
    
    If Not (myCheckBox Is Nothing) Then
        ' found
        Set CheckBoxFind = myCheckBox
    Else
        ' not found
        Set CheckBoxFind = Nothing
    End If
End Function

Function ComboBoxFind(objNAME As String) As MSForms.ComboBox
    Dim ws As Worksheet
    Dim obj As OLEObject
    Dim myComboBox As MSForms.ComboBox
    
    
    Set ws = ThisWorkbook.Worksheets("ss")
    Set myComboBox = Nothing
    
    For Each obj In ws.OLEObjects
        If TypeOf obj.Object Is MSForms.ComboBox Then
            If obj.Name = objNAME Then
                Set myComboBox = obj.Object
                Exit For
            End If
        End If
    Next obj
    
    If Not (myComboBox Is Nothing) Then
        ' found
        Set ComboBoxFind = myComboBox
    Else
        ' not found
        Set ComboBoxFind = Nothing
    End If
End Function


Function TextBoxFind(objNAME As String) As MSForms.TextBox
    Dim ws As Worksheet
    Dim obj As OLEObject
    Dim myTextBox As MSForms.TextBox
    
    
    Set ws = ThisWorkbook.Worksheets("ss")
    Set myTextBox = Nothing
    
    For Each obj In ws.OLEObjects
        If TypeOf obj.Object Is MSForms.TextBox Then
            If obj.Name = objNAME Then
                Set myTextBox = obj.Object
                Exit For
            End If
        End If
    Next obj
    
    If Not (myTextBox Is Nothing) Then
        ' found
        Set TextBoxFind = myTextBox
    Else
        ' not found
        Set TextBoxFind = Nothing
    End If
End Function



Function is_Jiyeol(ByVal area As String) As Boolean
    Dim tbl As ListObject
    Dim headerRowArray() As Variant
    
    Set tbl = Sheets("ref1").ListObjects("tableJIYEOL")
    
    headerRowArray = tbl.HeaderRowRange.Value
    
    Dim i As Integer
    
    For i = LBound(headerRowArray, 2) To UBound(headerRowArray, 2)
        If headerRowArray(1, i) = area Then
            is_Jiyeol = True
            Exit Function
        End If
    Next i
    
    is_Jiyeol = False
End Function


Sub initialize()
    Dim TextBox_AREA As MSForms.TextBox

    Set TextBox_AREA = TextBoxFind("TextBox_AREA")
        
    If is_Jiyeol(TextBox_AREA.Value) Then
        Call initialize_JIYEOL(TextBox_AREA.Value)
    Else
        Call initialize_CNU(TextBox_AREA.Value)
    End If
       
End Sub


Private Function lastRowByKey(cell As String) As Long
    lastRowByKey = Range(cell).End(xlDown).row
End Function


' 물량계산
Sub ComputeQ()
    Dim i As Integer
    Dim lastRow As Long

    Call initialize
    
    Sheets("ss").Activate
    lastRow = lastRowByKey("A1")
    
    For i = 2 To lastRow
        Cells(i, "L").Value = ss_water(Range("I" & CStr(i)).Value, Range("K" & CStr(i)).Value, 100)
    Next i
    
    Sheets("aa").Activate
    lastRow = lastRowByKey("A1")
    
    For i = 2 To lastRow
        Cells(i, "L").Value = aa_water(Range("I" & CStr(i)).Value, Range("K" & CStr(i)).Value, 100)
    Next i
End Sub


Function ss_water(ByVal qhp As Integer, ByVal strPurpose As String, Optional ByVal npopulation As Integer = 60) As Double
    '지열 냉난방
    If CheckSubstring(strPurpose, "냉") Then
        ss_water = qhp * 0.01
        Exit Function
    End If
    
    ' 일반용
    If CheckSubstring(strPurpose, "일") Then
        ss_water = Round(SS(svILBAN, 1) + qhp * SS(svILBAN, 2), 2)
        Exit Function
    End If
    
    
    ' 가정용
    If CheckSubstring(strPurpose, "가") Then
        ss_water = Round(SS(svGAJUNG, 1) + SS_CITY * SS(svGAJUNG, 2), 2)
        Exit Function
    End If
    
    ' 기타
    If CheckSubstring(strPurpose, "기") Then
        ss_water = Round(SS(svGAJUNG, 1) + SS_CITY * SS(svGAJUNG, 2), 2)
        Exit Function
    End If
    
    ' 농생활겸용
    If CheckSubstring(strPurpose, "농") Then
        ss_water = Round(SS(svILBAN, 1) + qhp * SS(svILBAN, 2), 2)
        Exit Function
    End If
    
    ' 청소용
    If CheckSubstring(strPurpose, "청") Then
        ss_water = Round(SS(svGAJUNG, 1) + SS_CITY * SS(svGAJUNG, 2), 2)
        Exit Function
    End If
    
    '간이상수도
    If CheckSubstring(strPurpose, "상") Then
        ss_water = Round(SS(svMAEUL, 1) + npopulation * SS(svMAEUL, 2), 2)
        Exit Function
    End If
    
    ' 공사용
    If CheckSubstring(strPurpose, "공사") Then
        ss_water = Round(SS(svILBAN, 1) + qhp * SS(svILBAN, 2), 2)
        Exit Function
    End If
    
    ' 공동주택용
    If CheckSubstring(strPurpose, "공동") Then
        ss_water = Round(SS(svGONGDONG, 1) + npopulation * SS(svGONGDONG, 2), 2)
        Exit Function
    End If
        
    ' 민방위용
    If CheckSubstring(strPurpose, "민방") Then
        ss_water = Round(SS(svILBAN, 1) + qhp * SS(svILBAN, 2), 2)
        Exit Function
    End If
    
    ' 학교용
    If CheckSubstring(strPurpose, "학교") Then
        ss_water = Round(SS(svSCHOOL, 1) + npopulation * SS(svSCHOOL, 2), 2)
        Exit Function
    End If
    
   ss_water = 900
End Function




Function aa_water(qhp As Integer, strPurpose As String, Optional ByVal nhead As Integer = 30) As Double
    'nhead - 축산업의 두수 ....

    ' 전작용
    If CheckSubstring(strPurpose, "전") Then
        aa_water = Round(AA(avJEONJAK, 1) + qhp * AA(avJEONJAK, 2), 2)
        Exit Function
    End If
    
    ' 답작용
    If CheckSubstring(strPurpose, "답") Then
        aa_water = Round(AA(avDAPJAK, 1) + qhp * AA(avDAPJAK, 2), 2)
        Exit Function
    End If
    
    
    ' 원예용
    If CheckSubstring(strPurpose, "원") Then
        aa_water = Round(AA(avWONYE, 1) + qhp * AA(avWONYE, 2), 2)
        Exit Function
    End If
    
    ' 농생활겸용
    If CheckSubstring(strPurpose, "농") Then
        aa_water = Round(AA(avJEONJAK, 1) + qhp * AA(avJEONJAK, 2), 2)
        Exit Function
    End If
    
    ' 양계장용
    If CheckSubstring(strPurpose, "양") Then
        aa_water = Round(AA(avDAPJAK, 1) + qhp * AA(avDAPJAK, 2), 2)
        Exit Function
    End If
    
    '축산용
    If CheckSubstring(strPurpose, "축") Then
        aa_water = Round(AA(avCOW, 1) + nhead * AA(avCOW, 2), 2)
        Exit Function
    End If
    
    ' 기타
    If CheckSubstring(strPurpose, "기타") Then
        aa_water = Round(AA(avDAPJAK, 1) + nhead * AA(avDAPJAK, 2), 2)
        Exit Function
    End If
    
   aa_water = 900
End Function











' ***************************************************************
' water_GenerationCopy
'
' ***************************************************************




Option Explicit

Private Function lastRowByKey(cell As String) As Long
    lastRowByKey = Range(cell).End(xlDown).row
End Function


Private Function lastRowByRowsCount(cell As String) As Long
    lastRowByRowsCount = Cells(Rows.Count, cell).End(xlUp).row
End Function

Public Sub clearRowA()
    
End Sub

Private Function lastRowByFind() As Long
    Dim lastRow As Long
    
    lastRow = Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
    
    lastRowByFind = lastRow
End Function

Private Sub DoCopy(lastRow As Long)
    Range("F2:H" & lastRow).Select
    Selection.Copy
    
    Range("n2").Select
    ActiveSheet.Paste
    
    
    ' 물량
    Range("L2:L" & lastRow).Select
    Selection.Copy
    
    Range("q2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Range("k2:k" & lastRow).Select
    Selection.Copy
    
    Range("r2").Select
    ActiveSheet.Paste
    
    Range("N14").Select
    Application.CutCopyMode = False
End Sub



' return letter of range ...
Function Alpha_Column(Cell_Add As Range) As String
    Dim No_of_Rows As Integer
    Dim No_of_Cols As Integer
    Dim Num_Column As Integer
    
    No_of_Rows = Cell_Add.Rows.Count
    No_of_Cols = Cell_Add.Columns.Count
    
    If ((No_of_Rows <> 1) Or (No_of_Cols <> 1)) Then
        Alpha_Column = ""
        Exit Function
    End If
    
    Num_Column = Cell_Add.Column
    If Num_Column < 26 Then
        Alpha_Column = Chr(64 + Num_Column)
    Else
        Alpha_Column = Chr(Int(Num_Column / 26) + 64) & Chr((Num_Column Mod 26) + 64)
    End If
End Function


' Ctrl+D , Toggle OX, Toggle SINGO, HEOGA
Sub ToggleOX()
    Dim activeCellColumn, activeCellRow As String
    Dim row As Long
    Dim col As Long
    Dim lastRow As Long
    Dim cp, fillRange As String
    

    activeCellColumn = Split(ActiveCell.address, "$")(1)
    activeCellRow = Split(ActiveCell.address, "$")(2)
  
    row = ActiveCell.row
    col = ActiveCell.Column
    
    Debug.Print Alpha_Column(ActiveCell)
    
    If activeCellColumn = "S" Then
        If ActiveCell.Value = "O" Then
            ActiveCell.Value = "X"
        Else
            ActiveCell.Value = "O"
        End If
    End If
    

    If activeCellColumn = "B" Then
        If ActiveCell.Value = "신고공" Then
            ActiveCell.Value = "허가공"
            With Selection.Font
                .Color = -16776961
                .TintAndShade = 0
            End With
            Selection.Font.Bold = True
        Else
            ActiveCell.Value = "신고공"
             With Selection.Font
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0
            End With
            Selection.Font.Bold = False
        End If
    End If
    
    If activeCellColumn = "D" Then
        cp = Replace(ActiveCell.address, "$", "")
        lastRow = lastRowByKey(ActiveCell.address)
        
        fillRange = "D" & Range(cp).row & ":D" & lastRow
        
        Range(cp).Select
        Selection.AutoFill Destination:=Range(fillRange)
        
        Range(cp).Select
    End If
    
    If ActiveSheet.Name = "ss" And activeCellColumn = "K" Then
        UserForm_SS.Show
    End If
    
    If ActiveSheet.Name = "aa" And activeCellColumn = "K" Then
        UserForm_AA.Show
    End If
    
End Sub


Sub MainMoudleGenerateCopy()
    Dim lastRow As Long
        
    lastRow = lastRowByKey("A1")
    Call DoCopy(lastRow)
End Sub


Sub SubModuleInitialClear()
    Dim lastRow As Long
    Dim userChoice As VbMsgBoxResult
    
    lastRow = lastRowByKey("A1")
  
    userChoice = MsgBox("Do you want to continue?", vbOKCancel, "Confirmation")

    If userChoice <> vbOK Then
        Exit Sub
    End If
    
    Range("e2:j" & lastRow).Select
    Selection.ClearContents
    Range("n2:r" & lastRow).Select
    Selection.ClearContents
    Range("P14").Select
    
    
    If lastRow >= 23 Then
        Rows("23:" & lastRow).Select
        Selection.Delete Shift:=xlUp
    End If
    
    
    Range("m2").Select

End Sub

Sub SubModuleCleanCopySection()
    Dim lastRow As Long
        
    lastRow = lastRowByKey("A1")
    Range("n2:r" & lastRow).Select
    Selection.ClearContents
    Range("P14").Select
End Sub


' 2023/4/19 - copy modify

Sub insertRow()
    Dim lastRow As Long, i As Long, j As Long
    Dim selection_origin, selection_target As String
    Dim AddingRowCount As Long
    
    'lastRow = lastRowByKey("A1")

    AddingRowCount = 10

    lastRow = lastRowByRowsCount("A")
    
    Rows(CStr(lastRow + 1) & ":" & CStr(lastRow + AddingRowCount)).Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
    
    i = lastRowByKey("A1")
    j = i + AddingRowCount
    
    selection_origin = "A" & i & ":D" & i
    selection_target = "A" & i & ":D" & j
    
    Range(selection_origin).Select
    Selection.AutoFill Destination:=Range(selection_target), Type:=xlFillDefault
 
    selection_origin = "K" & i & ":M" & i
    selection_target = "K" & i & ":M" & j

    Range(selection_origin).Select
    Selection.AutoFill Destination:=Range(selection_target), Type:=xlFillDefault
    
    Range("S" & i).Select
    Selection.AutoFill Destination:=Range("S" & i & ":S" & j), Type:=xlFillDefault
    
    Application.CutCopyMode = False
    
    ActiveWindow.LargeScroll Down:=-1
    ActiveWindow.LargeScroll Down:=-1
    ActiveWindow.LargeScroll Down:=-1
    ActiveWindow.LargeScroll Down:=-1
End Sub





' ***************************************************************
' mod_initialize_setting
'
' ***************************************************************


Option Explicit



Private Sub init_nonsan()

   
    SS(svGAJUNG, 1) = 0.173
    SS(svGAJUNG, 2) = 0.21
    SS_CITY = 2.42
    
    SS(svILBAN, 1) = 3.154
    SS(svILBAN, 2) = 0.023
    
    SS(svSCHOOL, 1) = 7.986
    SS(svSCHOOL, 2) = 0.005
    
    SS(svGONGDONG, 1) = 0.173
    SS(svGONGDONG, 2) = 0.21
    
    SS(svMAEUL, 1) = 7.13
    SS(svMAEUL, 2) = 0.001
    
'----------------------------------------

    AA(avJEONJAK, 1) = 5.66
    AA(avJEONJAK, 2) = 0.014
    
    AA(avDAPJAK, 1) = 1.98
    AA(avDAPJAK, 2) = 0.044
    
    AA(avWONYE, 1) = 2.789
    AA(avWONYE, 2) = 0.011
    
    AA(avCOW, 1) = 3.48
    AA(avCOW, 2) = 0.009
    
    AA(avPIG, 1) = 4.719
    AA(avPIG, 2) = 0.001
    
    AA(avCHICKEN, 1) = 5.492
    AA(avCHICKEN, 2) = 0.041
    
End Sub

Private Sub init_daejeon()

   
    SS(svGAJUNG, 1) = 0.173
    SS(svGAJUNG, 2) = 0.21
    SS_CITY = 2.43

    SS(svILBAN, 1) = 2.119
    SS(svILBAN, 2) = 0.021
    
    SS(svSCHOOL, 1) = 7.986
    SS(svSCHOOL, 2) = 0.005
    
    SS(svGONGDONG, 1) = 7.13
    SS(svGONGDONG, 2) = 0.001
    
    SS(svMAEUL, 1) = 6.463
    SS(svMAEUL, 2) = 0.178
    
'----------------------------------------

    AA(avJEONJAK, 1) = 5.66
    AA(avJEONJAK, 2) = 0.014
    
    AA(avDAPJAK, 1) = 1.98
    AA(avDAPJAK, 2) = 0.044
    
    AA(avWONYE, 1) = 2.789
    AA(avWONYE, 2) = 0.011
    
    AA(avCOW, 1) = 3.48
    AA(avCOW, 2) = 0.009
    
    AA(avPIG, 1) = 4.719
    AA(avPIG, 2) = 0.001
    
    AA(avCHICKEN, 1) = 5.492
    AA(avCHICKEN, 2) = 0.041

End Sub


Private Sub init_yeonki()

   
    SS(svGAJUNG, 1) = 0.265
    SS(svGAJUNG, 2) = 0.181
    SS_CITY = 2.75

    SS(svILBAN, 1) = 3.521
    SS(svILBAN, 2) = 0.011
    
    SS(svSCHOOL, 1) = 11.687
    SS(svSCHOOL, 2) = 0.007
    
    SS(svGONGDONG, 1) = 0.265
    SS(svGONGDONG, 2) = 0.181
    
    SS(svMAEUL, 1) = 7.287
    SS(svMAEUL, 2) = 0.001
    
'----------------------------------------

    AA(avJEONJAK, 1) = 5.66
    AA(avJEONJAK, 2) = 0.014
    
    AA(avDAPJAK, 1) = 1.98
    AA(avDAPJAK, 2) = 0.044
    
    AA(avWONYE, 1) = 2.789
    AA(avWONYE, 2) = 0.011
    
    AA(avCOW, 1) = 3.48
    AA(avCOW, 2) = 0.009
    
    AA(avPIG, 1) = 4.719
    AA(avPIG, 2) = 0.001
    
    AA(avCHICKEN, 1) = 5.492
    AA(avCHICKEN, 2) = 0.041

End Sub

Private Sub init_boryoung()

   
    SS(svGAJUNG, 1) = 0.173
    SS(svGAJUNG, 2) = 0.21
    SS_CITY = 2.36
    
    SS(svILBAN, 1) = 3.154
    SS(svILBAN, 2) = 0.023
    
    SS(svSCHOOL, 1) = 7.986
    SS(svSCHOOL, 2) = 0.005
    
    SS(svGONGDONG, 1) = 0.173
    SS(svGONGDONG, 2) = 0.21
    
    SS(svMAEUL, 1) = 7.13
    SS(svMAEUL, 2) = 0.001
    
'----------------------------------------

    AA(avJEONJAK, 1) = 6.964
    AA(avJEONJAK, 2) = 0.013
    
    AA(avDAPJAK, 1) = 2.089
    AA(avDAPJAK, 2) = 0.043
    
    AA(avWONYE, 1) = 2.789
    AA(avWONYE, 2) = 0.011
    
    AA(avCOW, 1) = 3.48
    AA(avCOW, 2) = 0.009
    
    AA(avPIG, 1) = 4.719
    AA(avPIG, 2) = 0.001
    
    AA(avCHICKEN, 1) = 5.492
    AA(avCHICKEN, 2) = 0.041
    
End Sub

Private Sub init_dangjin()

   
    SS(svGAJUNG, 1) = 0.173
    SS(svGAJUNG, 2) = 0.21
    SS_CITY = 2.59
    
    SS(svILBAN, 1) = 3.154
    SS(svILBAN, 2) = 0.023
    
    SS(svSCHOOL, 1) = 7.986
    SS(svSCHOOL, 2) = 0.005
    
    SS(svGONGDONG, 1) = 0.173
    SS(svGONGDONG, 2) = 0.21
    
    SS(svMAEUL, 1) = 7.13
    SS(svMAEUL, 2) = 0.001
    
'----------------------------------------

    AA(avJEONJAK, 1) = 6.964
    AA(avJEONJAK, 2) = 0.013
    
    AA(avDAPJAK, 1) = 2.089
    AA(avDAPJAK, 2) = 0.043
    
    AA(avWONYE, 1) = 2.789
    AA(avWONYE, 2) = 0.011
    
    AA(avCOW, 1) = 3.48
    AA(avCOW, 2) = 0.009
    
    AA(avPIG, 1) = 4.719
    AA(avPIG, 2) = 0.001
    
    AA(avCHICKEN, 1) = 5.492
    AA(avCHICKEN, 2) = 0.041
    
End Sub

Private Sub init_yesan()

   
    SS(svGAJUNG, 1) = 0.265
    SS(svGAJUNG, 2) = 0.181
    SS_CITY = 2.34
    
    SS(svILBAN, 1) = 3.521
    SS(svILBAN, 2) = 0.011
    
    SS(svSCHOOL, 1) = 11.687
    SS(svSCHOOL, 2) = 0.007
    
    SS(svGONGDONG, 1) = 0.265
    SS(svGONGDONG, 2) = 0.181
    
    SS(svMAEUL, 1) = 7.287
    SS(svMAEUL, 2) = 0.001
    
'----------------------------------------

    AA(avJEONJAK, 1) = 6.964
    AA(avJEONJAK, 2) = 0.013
    
    AA(avDAPJAK, 1) = 2.089
    AA(avDAPJAK, 2) = 0.043
    
    AA(avWONYE, 1) = 2.789
    AA(avWONYE, 2) = 0.011
    
    AA(avCOW, 1) = 3.48
    AA(avCOW, 2) = 0.009
    
    AA(avPIG, 1) = 4.719
    AA(avPIG, 2) = 0.001
    
    AA(avCHICKEN, 1) = 5.492
    AA(avCHICKEN, 2) = 0.041
    
End Sub



Private Sub init_sejong()

   
    SS(svGAJUNG, 1) = 0.173
    SS(svGAJUNG, 2) = 0.21
    SS_CITY = 2.57

    SS(svILBAN, 1) = 3.154
    SS(svILBAN, 2) = 0.023
    
    SS(svSCHOOL, 1) = 7.986
    SS(svSCHOOL, 2) = 0.005
    
    SS(svGONGDONG, 1) = 7.13
    SS(svGONGDONG, 2) = 0.001
    
    SS(svMAEUL, 1) = 6.463
    SS(svMAEUL, 2) = 0.178
    
'----------------------------------------

    AA(avJEONJAK, 1) = 5.66
    AA(avJEONJAK, 2) = 0.014
    
    AA(avDAPJAK, 1) = 1.98
    AA(avDAPJAK, 2) = 0.044
    
    AA(avWONYE, 1) = 2.789
    AA(avWONYE, 2) = 0.011
    
    AA(avCOW, 1) = 3.48
    AA(avCOW, 2) = 0.009
    
    AA(avPIG, 1) = 4.719
    AA(avPIG, 2) = 0.001
    
    AA(avCHICKEN, 1) = 5.492
    AA(avCHICKEN, 2) = 0.041

End Sub


Public Enum LC_COMBOBOX
    lcDAEJEON = 1
    lcJIYEOL = 2
End Enum

Public IS_FIRST_LOAD As Boolean

Private Sub OptionButton_DAEJEON_Click()

    If IS_FIRST_LOAD Then
        Call LoadComboBox
        IS_FIRST_LOAD = False
    Else
        Call LoadComboBox
        ComboBox_AREA.Value = "default"
        IS_FIRST_LOAD = False
    End If
    
End Sub


Private Sub OptionButton_JIYEOL_Click()
    
    If IS_FIRST_LOAD Then
        Call LoadComboBox
        IS_FIRST_LOAD = False
    Else
        Call LoadComboBox
        ComboBox_AREA.Value = "default"
        IS_FIRST_LOAD = False
    End If
    
End Sub

Sub PutDataToASheet(ByVal sh As String, ByVal table As String, ByVal area As String, SurveyData As Variant)
    Dim tbl As ListObject
    Dim cell As Range
    Dim i As Integer: i = 1
    
    Set tbl = Sheets(sh).ListObjects(table)
        
    For Each cell In tbl.ListColumns(area).DataBodyRange.Cells
        cell.Value = SurveyData(i)
        i = i + 1
    Next cell
    
End Sub

Function getSurveyData() As Variant
    Dim values() As Variant
    Dim i As Integer: i = 1
    Dim ctl As Control

    ReDim values(1 To 23)
    
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "TextBox" Then
            values(i) = ctrl.Value
            i = i + 1
        End If
    Next ctrl
    
    getSurveyData = values
    
End Function



Private Sub CommandButton_Insert_Click()

    Dim values As Variant
    Dim area As String
    
    values = getSurveyData()

    area = ComboBox_AREA.Value
    
    If area = "" Then
        area = Default
    End If
    
    
    If OptionButton_JIYEOL.Value Then
        Call PutDataToASheet("ref1", "tableJIYEOL", area, values)
    Else
        Call PutDataToASheet("ref", "tableCNU", area, values)
    End If
    
    Call PutText(area)
    
    Unload Me
    
End Sub


Sub PutText(area As String)
    Dim TextBox_AREA As MSForms.TextBox

    Set TextBox_AREA = TextBoxFind("TextBox_AREA")
    TextBox_AREA.Value = area
End Sub

' ***************************************************************
' UserForm_Survey
'
' ***************************************************************



Private Sub CommandButton_LOAD_Click()
    Call LoadSurveyData(ComboBox_AREA.Value)
End Sub


Private Sub ComboBox_AREA_Change()

End Sub

Sub Initialize_Setting()
    Dim TextBox_AREA As MSForms.TextBox

    Set TextBox_AREA = TextBoxFind("TextBox_AREA")
    Debug.Print "TextBox_AREA.Value", "'" & TextBox_AREA.Value & "'"
    
    If is_Jiyeol(TextBox_AREA.Value) Then
        OptionButton_JIYEOL.Value = True
    Else
        OptionButton_DAEJEON.Value = True
    End If
    
    ' Call LoadComboBox
    ' OptionButton.Value = True set is triggered clicked event
    
    ComboBox_AREA.Value = TextBox_AREA.Value
    LoadSurveyData (TextBox_AREA.Value)
    
End Sub


Sub LoadSurveyData(area As String)
    Dim tbl As ListObject
    Dim values() As Variant
    
    
    If OptionButton_JIYEOL.Value Then
        Set tbl = Sheets("ref1").ListObjects("tableJIYEOL")
    Else
        Set tbl = Sheets("ref").ListObjects("tableCNU")
    End If
    
    
    If area = "" Then
        area = Default
    End If
    
    values = tbl.ListColumns(area).DataBodyRange.Value
    Dim i As Integer: i = 1
        
    Dim ctl As Control

    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "TextBox" Then
            ' MsgBox "Found a TextBox with the name: " & ctrl.NAME
            ctrl.Value = values(i, 1)
            i = i + 1
        End If
    Next ctrl
    
End Sub


Sub LoadComboBox()
    Dim tbl As ListObject
    Dim tableNAME, shNAME As String
    Dim headerRowArray() As Variant
    
    ComboBox_AREA.Clear
    
    If OptionButton_JIYEOL.Value Then
        tableNAME = "tableJIYEOL"
        shNAME = "ref1"
    Else
        tableNAME = "tableCNU"
        shNAME = "ref"
    End If
    
    Set tbl = Sheets(shNAME).ListObjects(tableNAME)

    headerRowArray = tbl.HeaderRowRange.Value
    
    Dim i As Integer
    Dim isFirst As Boolean: isFirst = True
    
    
    For i = LBound(headerRowArray, 2) To UBound(headerRowArray, 2)
        If isFirst Then
            isFirst = False
            GoTo NEXT_LOOP
        End If
        
        ComboBox_AREA.AddItem headerRowArray(1, i)
        
NEXT_LOOP:
    Next i
End Sub



Function is_Jiyeol(ByVal area As String) As Boolean

    Dim tbl As ListObject
    Dim headerRowArray() As Variant
    
    Set tbl = Sheets("ref1").ListObjects("tableJIYEOL")
    
    headerRowArray = tbl.HeaderRowRange.Value
    
    Dim i As Integer
    
    For i = LBound(headerRowArray, 2) To UBound(headerRowArray, 2)
        If headerRowArray(1, i) = area Then
            is_Jiyeol = True
            Exit Function
        End If
    Next i
    
    is_Jiyeol = False

End Function


Private Sub UserForm_Initialize()
    Dim i As Integer

    IS_FIRST_LOAD = True
    Call Initialize_Setting
    
    
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    
End Sub

Private Sub CommandButton_Cancel_Click()
    Unload Me
End Sub




' ***************************************************************
' mod_MakeFieldList
'
' ***************************************************************


Option Explicit

' 이곳에다가 기본적인 설정값을 세팅해준다.
' 파일이름과, 조사일같은것들을 ...

Const EXPORT_DATE As String = "2023-12-20"
Const EXPORT_ADDR_HEADER As String = "서울특별시 "
Const EXPORT_FILE_NAME As String = "d:\05_Send\iyong_template.xlsx"
        
' 1인 1일당 급수량, 엑셀파일을 보고 검사
' 서울특별시 강북구
Const ONEMAN_WATER_SUPPLY As Double = 265.16
        
Public Enum ALLOW_TYPE_VALUE
     at_HEOGA = 0
     at_SINGO = 1
End Enum


Sub delay(ti As Integer)
    Application.Wait Now + TimeSerial(0, 0, ti)
End Sub


Sub MakeFieldList()
    Call make_datamid
    Call Delete_Outside_Boundary
    Call ExportData
End Sub


Sub ExportData()
    ' data_mid 에서, 중간과정으로 만들어진 데이타를 불러와서, 파이썬처리용 엑셀쉬트를 만든다.
    Call Make_DataOut
    Call ExportCurrentWorksheet("data_out")
End Sub

Sub ExportCurrentWorksheet(sh As String)
    Dim filePath As String
    
    If Not ActivateSheet(sh) Then
        Debug.Print "ActivateSheet Error, maybe sheet does not exist ...."
        Exit Sub
    End If
        
    'filePath = Application.GetSaveAsFilename(FileFilter:="Excel Files (*.xlsx), *.xlsx")
    ' filePath = "d:\05_Send\aaa.xlsx"
    
    filePath = EXPORT_FILE_NAME
    
    If VarType(filePath) = vbString Then
    
        If Dir(filePath) <> "" Then
            ' Delete the file
            Kill filePath
    
'            If MsgBox("The file " & filePath & " already exists. Do you want to overwrite it?", _
'                      vbQuestion + vbYesNo, "Confirm Overwrite") = vbNo Then
'                Exit Sub
'            End If
        End If
    
    
        If Sheets(sh).Visible = False Then
            Sheets(sh).Visible = True
        End If
        
        Sheets(sh).Activate
        ActiveSheet.Copy
        ActiveWorkbook.SaveAs fileName:=filePath, FileFormat:=xlOpenXMLWorkbook, ConflictResolution:=xlLocalSessionChanges
        ActiveWorkbook.Close savechanges:=False
    End If
End Sub


Private Sub DeleteFile(filePath As String)
    ' Check if the file exists before attempting to delete
    If Dir(filePath) <> "" Then
        ' Delete the file
        Kill filePath
        ' MsgBox "File deleted successfully.", vbInformation
    Else
        ' MsgBox "File not found.", vbExclamation
    End If
End Sub



Function ActivateSheet(sh As String) As Boolean
    On Error GoTo ErrorHandler
    Sheets(sh).Activate
    ActivateSheet = True
    Exit Function
    
ErrorHandler:
'    MsgBox "An error occurred while trying to activate the sheet." & vbNewLine & _
'           "Please check that the sheet name is correct and try again.", _
'           vbExclamation, "Error"

    ActivateSheet = False
End Function

Sub Make_DataOut()
    Dim str_, address, id, purpose As String
    Dim allowType, i, lastRow  As Integer
    Dim simdo, diameter, hp, capacity, tochool, Q As Double
    Dim setting As String
    
    Dim ag_start, ag_end, ag_year  As String
    Dim sayong_gagu, sayong_ingu, sayong_ilin_geupsoo As Double
    Dim usage_day, usage_month, usage_year As Double
    
    str_ = ChrW(&H2714)
    
    
    If Not Sheets("data_mid").Visible Then
        Sheets("data_mid").Visible = True
    End If
    
    Sheets("data_mid").Activate
    
    Call initialize
    lastRow = getlastrow()
    
    For i = 2 To lastRow
    
        Call GetDataFromSheet(i, id, address, allowType, simdo, diameter, hp, capacity, tochool, purpose, Q)
        
        If allowType = at_HEOGA Then
            setting = setting & "b,"
            ' 허가시설
        Else
            setting = setting & "c,"
            ' 신고시설
        End If
        
'       충적관정인지, 암반관정인지를 검사해서 추가해줌 ...
'       If (diameter >= 150) And (hp >= 1#) Then
'            setting = setting & "aq,"
'       Else
'            setting = setting & "ap,"
'       End If

        setting = setting & IIf(diameter >= 150 And hp >= 1#, "aq,", "ap,")

       
        Select Case LCase(Left(id, 1))
            Case "s"
                setting = setting & "f,"
                setting = setting & SS_StringCheck(purpose)
                setting = setting & SS_PublicCheck(purpose)
            
            Case "a"
                setting = setting & "u,"
                setting = setting & AA_StringCheck(purpose)
                
                If allowType = at_HEOGA Then
                    setting = setting & "ab,"
                Else
                    setting = setting & AA_PublicCheck(purpose)
                End If
                                            
            Case "i"
                setting = setting & "o,"
                setting = setting & II_StringCheck(purpose)
                setting = setting & II_PublicCheck(purpose)
                
                
        End Select
        
        ' 음용수 인가 , 먹을수있는 물인가 ?
        If IsDrinking(purpose) Then
            setting = setting & "ah,"
        Else
            setting = setting & "ai,"
        End If
        
        
        
        ' ad = 연중사용
        Select Case LCase(Left(id, 1))
            Case "s"
                setting = setting & "ad,"
                ag_start = "1"
                ag_end = "12"
                ag_year = "365"
            
            Case "a"
                '농업용 : 3 ~ 11월까지
                ag_start = "3"
                ag_end = "11"
                ag_year = "274"
            
            
            Case "i"
                ' 공업용 - 연중사용
                setting = setting & "ad,"
                ag_start = "1"
                ag_end = "12"
                ag_year = "365"
                
        End Select
        
        
        '음용수, 사용가구, 사용인구, 일인급수량이 결정됨
        If IsDrinking(purpose) Then
                 ' 용도가, 가정용일때 ...
                 If CheckSubstring(purpose, "가정") Then
                        sayong_gagu = 1
                        sayong_ingu = SS_CITY
                        sayong_ilin_geupsoo = Q / SS_CITY
                 End If
                
                 ' https://kosis.kr/statHtml/statHtml.do?orgId=110&tblId=DT_11001N_2013_A055
                 ' 용도가 간이상수도 일때 ...
                 If CheckSubstring(purpose, "간이") Then
                        sayong_gagu = 30
                        sayong_ingu = 90
                        sayong_ilin_geupsoo = ONEMAN_WATER_SUPPLY
                End If
        Else
            sayong_gagu = 0
            sayong_ingu = 0
            sayong_ilin_geupsoo = 0
        End If
                
         
        ' 일사용량 계산
        usage_day = Q
        usage_month = Q * 29
        
        If LCase(Left(id, 1)) = "s" Then
            usage_year = usage_month * 12
        Else
            usage_year = usage_month * 8
        End If
        
        
        '허가공 -  av,aw,ay,az,ba,
        
        ' 관정현황 체크
        Select Case LCase(Left(id, 1))
            Case "s"
                If allowType = at_SINGO Then ' 신고시설이면
                    If CheckSubstring(purpose, "일반") Then setting = setting & "aw,ay,"
                    If CheckSubstring(purpose, "간이") Then setting = setting & "av,aw,ax,ay,az,ba,"
                    If CheckSubstring(purpose, "공동") Then setting = setting & "av,aw,ay,"
                    If CheckSubstring(purpose, "민방") Then setting = setting & "av,aw,ay,"
                    If CheckSubstring(purpose, "학교") Then setting = setting & "av,aw,ay,"
                    If CheckSubstring(purpose, "청소") Then setting = setting & "av,aw,ay,"
                    If CheckSubstring(purpose, "공사") Then setting = setting & "av,aw,ay,"
                    If CheckSubstring(purpose, "겸용") Then setting = setting & "av,aw,ay,"
                Else ' 허가시설이면
                    setting = setting & "av,aw,ax,ay,az,ba"
                End If
            
            Case "a"
                If allowType = at_SINGO Then ' 신고시설이면
                    If CheckSubstring(purpose, "전작") Then setting = setting & "aw,ay,"
                    If CheckSubstring(purpose, "답작") Then setting = setting & "aw,ay,"
                    If CheckSubstring(purpose, "원예") Then setting = setting & "aw,ay,"
                    If CheckSubstring(purpose, "겸용") Then setting = setting & "aw,ay,"
                    If CheckSubstring(purpose, "양어장") Then setting = setting & "aw,ay,"
                    If CheckSubstring(purpose, "축산") Then setting = setting & "aw,ay,"
                    If CheckSubstring(purpose, "기타") Then setting = setting & "aw,ay,"
                Else ' 허가시설이면
                    setting = setting & "av,aw,ax,ay,az,ba"
                End If
            
            
            Case "i"
                ' 공업용 - 연중사용
                setting = setting & "ad,"
                If allowType = at_SINGO Then
                    ' 신고시설이면
                    setting = setting & "aw,ay,"
                    
                Else
                    ' 허가시설이면
                    setting = setting & "av,aw,ax,ay,az,ba"
                End If
                
        End Select
        
        
        
        
        Debug.Print "**********************************"
        Debug.Print setting
        
        Call PutDataSheetOut(i, setting, address, simdo, diameter, hp, capacity, tochool, Q, ag_start, ag_end, ag_year, _
                             sayong_gagu, sayong_ingu, sayong_ilin_geupsoo, usage_day, usage_month, usage_year)
        
        
        setting = ""
    
    Next i

' =INDEX(itable[value], MATCH("d1", itable[key], 0))

End Sub

Sub PutDataSheetOut(ii As Variant, setting As Variant, address As Variant, simdo As Variant, diameter As Variant, hp As Variant, _
                    capacity As Variant, tochool As Variant, Q As Variant, _
                    ag_start As Variant, ag_end As Variant, ag_year As Variant, _
                    sayong_gagu As Variant, sayong_ingu As Variant, sayong_ilin_geupsoo As Variant, _
                    usage_day As Variant, usage_month As Variant, usage_year As Variant)

    Dim out() As String
    Dim i As Integer
    Dim index, str_, setting_1 As String
    
    Sheets("data_out").Activate
    
    With Range("A" & CStr(ii) & ":BB" & CStr(ii))
        .Value = " "
    End With

    str_ = ChrW(&H2714)
    
    
    setting_1 = DeepCopyString(CStr(setting))
    
    out = FilterString(setting_1)
    
    For i = LBound(out) To UBound(out)
        index = out(i)
        Sheets("data_out").Cells(ii, index).Value = str_
    Next i
    
    '  myString = Format(myDate, "yyyy-mm-dd")
    Sheets("data_out").Cells(ii, "a").Value = " " & Format(EXPORT_DATE, "yyyy-mm-dd") & "."
    Sheets("data_out").Cells(ii, "e").Value = address
    Sheets("data_out").Cells(ii, "ar").Value = simdo
    Sheets("data_out").Cells(ii, "as").Value = diameter
    Sheets("data_out").Cells(ii, "at").Value = hp
    Sheets("data_out").Cells(ii, "au").Value = capacity
    Sheets("data_out").Cells(ii, "av").Value = tochool
    
    Sheets("data_out").Cells(ii, "ae").Value = ag_start
    Sheets("data_out").Cells(ii, "af").Value = ag_end
    Sheets("data_out").Cells(ii, "ag").Value = ag_year
    
    ' 음용수 일때만, 사용가구, 사용인구, 1인급수 세팅
    If Sheets("data_out").Cells(ii, "ah").Value = ChrW(&H2714) Then
        Sheets("data_out").Cells(ii, "aj").Value = CStr(Format(sayong_gagu, "0.00"))
        Sheets("data_out").Cells(ii, "ak").Value = CStr(Format(sayong_ingu, "0.00"))
        Sheets("data_out").Cells(ii, "al").Value = CStr(Format(sayong_ilin_geupsoo, "0.00"))
    End If
    
    Sheets("data_out").Cells(ii, "am").Value = CStr(Format(usage_day, "0.00"))
    Sheets("data_out").Cells(ii, "an").Value = CStr(Format(usage_month, "#,##0"))
    Sheets("data_out").Cells(ii, "ao").Value = CStr(Format(usage_year, "#,##0"))
    

End Sub
                             
                          
' GetDataFromSheet(i, id, address, allowType, simdo, diameter, hp, capacity, tochool, purpose, Q)
Sub GetDataFromSheet(i As Variant, id As Variant, address As Variant, allowType As Variant, _
                     simdo As Variant, diameter As Variant, hp As Variant, capacity As Variant, tochool As Variant, _
                     purpose As Variant, Q As Variant)
    
    id = Sheets("data_mid").Cells(i, "a").Value
    address = Sheets("data_mid").Cells(i, "b").Value
    allowType = Sheets("data_mid").Cells(i, "c").Value
    simdo = Sheets("data_mid").Cells(i, "d").Value
    diameter = Sheets("data_mid").Cells(i, "e").Value
    hp = Sheets("data_mid").Cells(i, "f").Value
    capacity = Sheets("data_mid").Cells(i, "g").Value
    tochool = Sheets("data_mid").Cells(i, "h").Value
    purpose = Sheets("data_mid").Cells(i, "i").Value
    Q = Sheets("data_mid").Cells(i, "j").Value
    
End Sub


Function getlastrow() As Integer
    ' ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
    getlastrow = ActiveSheet.Range("A3333").End(xlUp).row
End Function


' 2024-1-11 , modify last cell check
' using cell reference SUM_SS, SUM_AA, SUM_II

Sub LastRowFindAll(row_ss As Variant, row_aa As Variant, row_ii As Variant)
    
    If Range("SUM_SS").Value = 0 Then
        row_ss = 0
    Else
        Sheets("ss").Activate
        row_ss = getlastrow() - 1
    End If
           
    If Range("SUM_AA").Value = 0 Then
        row_aa = 0
    Else
        Sheets("aa").Activate
        row_aa = getlastrow() - 1
    End If
      
    
    If Range("SUM_II").Value = 0 Then
        row_ii = 0
        Exit Sub
    Else
        Sheets("ii").Activate
        row_ii = getlastrow() - 1
    End If
    
End Sub

Sub EraseSheetData()
    Worksheets("data_mid").Range("A2:J1000").Delete
    Worksheets("data_out").Range("A2:BD1000").Delete
End Sub


' allowType = 1 - 신고공
' allowType = 0 - 허가공
Public Sub make_datamid()
    Dim i, j, row_end As Integer
    Dim newAddress, id, purpose As String
    Dim allowType As Integer
    Dim well_data(1 To 5) As Double
    Dim Q As Double
    Dim boundary As String
    Dim row_ss, row_aa, row_ii As Integer
    
    Call LastRowFindAll(row_ss, row_aa, row_ii)
    Call EraseSheetData
    
    Sheets("ss").Activate
    ' Debug.Print row_end
    For i = 1 To row_ss
        id = Cells(i + 1, "a").Value
        ' 주소헤더, 지역에 따라 값을 다시 설정해주어야 한다.
        newAddress = EXPORT_ADDR_HEADER & Cells(i + 1, "c") & " " & Cells(i + 1, "d") & " " & Cells(i + 1, "e") & " , " & id
        
        If Cells(i + 1, "b").Value = "신고공" Then
            allowType = 1
        Else
            allowType = 0
        End If
        
        ' Debug.Print allowType, newAddress
        
        For j = 1 To 5
            well_data(j) = Cells(i + 1, Chr(Asc("f") + j - 1)).Value
        Next j
        
        purpose = Cells(i + 1, "k").Value
        Q = Cells(i + 1, "l").Value
        boundary = Cells(i + 1, "s").Value
        
        Call putdata(i, id, newAddress, allowType, well_data, purpose, Q, boundary)
    Next i
    
    
    Sheets("aa").Activate
    ' Debug.Print row_end
    For i = 1 To row_aa
    
        id = Cells(i + 1, "a").Value
        newAddress = EXPORT_ADDR_HEADER & Cells(i + 1, "c") & " " & Cells(i + 1, "d") & " " & Cells(i + 1, "e") & " , " & id
        
        If Cells(i + 1, "b").Value = "신고공" Then
            allowType = 1
        Else
            allowType = 0
        End If
        
        ' Debug.Print allowType, newAddress
        
        For j = 1 To 5
            well_data(j) = Cells(i + 1, Chr(Asc("f") + j - 1)).Value
        Next j
        
        purpose = Cells(i + 1, "k").Value
        Q = Cells(i + 1, "l").Value
        boundary = Cells(i + 1, "s").Value
        
        Call putdata(i + row_ss, id, newAddress, allowType, well_data, purpose, Q, boundary)
    Next i
    
    Sheets("ii").Activate
    ' Debug.Print row_end
    
    For i = 1 To row_ii
    
        id = Cells(i + 1, "a").Value
        newAddress = EXPORT_ADDR_HEADER & Cells(i + 1, "c") & " " & Cells(i + 1, "d") & " " & Cells(i + 1, "e") & " , " & id
        
        If Cells(i + 1, "b").Value = "신고공" Then
            allowType = 1
        Else
            allowType = 0
        End If
        
        ' Debug.Print allowType, newAddress
        
        For j = 1 To 5
            well_data(j) = Cells(i + 1, Chr(Asc("f") + j - 1)).Value
        Next j
        
        purpose = Cells(i + 1, "k").Value
        Q = Cells(i + 1, "l").Value
        boundary = Cells(i + 1, "s").Value
        
        Call putdata(i + row_ss + row_aa, id, newAddress, allowType, well_data, purpose, Q, boundary)
    Next i
    
End Sub

' 2024-1-11
' delete outside boundary

Private Sub Delete_Outside_Boundary()

    Dim row_ss, row_aa, row_ii As Integer
    Dim i, j As Integer
        
    j = 2
    Sheets("data_mid").Activate
    
    For i = 1 To getlastrow()
        
        If Cells(j, "K").Value = "O" Then
            j = j + 1
        Else
            Rows(j & ":" & j).Select
            Selection.Delete Shift:=xlUp
        End If
    
    Next i

End Sub

Sub putdata(i As Variant, id As Variant, newAddress As Variant, allowType As Variant, well_data As Variant, _
            purpose As Variant, Q As Variant, boundary As Variant)
        
    Sheets("data_mid").Cells(i + 1, "a").Value = id
    Sheets("data_mid").Cells(i + 1, "b").Value = newAddress
    Sheets("data_mid").Cells(i + 1, "c").Value = allowType
    Sheets("data_mid").Cells(i + 1, "d").Value = well_data(1)
    Sheets("data_mid").Cells(i + 1, "e").Value = well_data(2)
    Sheets("data_mid").Cells(i + 1, "f").Value = well_data(3)
    Sheets("data_mid").Cells(i + 1, "g").Value = well_data(4)
    Sheets("data_mid").Cells(i + 1, "h").Value = well_data(5)
    Sheets("data_mid").Cells(i + 1, "i").Value = purpose
    Sheets("data_mid").Cells(i + 1, "j").Value = Q
    Sheets("data_mid").Cells(i + 1, "k").Value = boundary
    
End Sub













' ***************************************************************
' Sheet1_index(index)
'
' ***************************************************************



Option Explicit

' ***************************************************************
' Sheet4_data_out(data_out)
'
' ***************************************************************



Option Explicit

' ***************************************************************
' Sheet6_data_mid(data_mid)
'
' ***************************************************************



Option Explicit

' ***************************************************************
' mod_UniCOde
'
' ***************************************************************


Option Explicit

Sub GetCheckMarkCode()
    Dim checkMark As String
    Dim code As Long

    checkMark = "?" ' the check mark symbol

    code = AscW(checkMark)

    Debug.Print "The Unicode code point for " & checkMark & " is " & code
End Sub

Sub InsertCheckMark()
    Dim checkMark As String

    checkMark = ChrW(&H2714) ' &H2714 is the Unicode code point for the check mark symbol

    Range("A1").Value = checkMark ' Replace "A1" with the cell where you want to insert the check mark symbol
End Sub

Sub TestUniCode()
    Dim i As Integer
    Dim str_check As String
    Dim code As Long
    Dim index As Variant
        
    str_check = Sheets("index").Range("a1").Value
    code = AscW(str_check)
    
    index = Array("a", "b", "c", "f", "k")
    
    For i = LBound(index) To UBound(index)
        Cells(33, index(i)).Value = ChrW(&H2714)   ' str_check
    Next i
    
    Debug.Print "strcheck", code
End Sub



' ***************************************************************
' mod_FilterString
'
' ***************************************************************


Option Explicit

Function FilterString(str As String) As Variant
    Dim elements() As String
    Dim element As Variant
    Dim out() As String
    Dim i As Long
    
    elements = Split(str, ",")
    For Each element In elements
        If element = "" Then Exit For
        ReDim Preserve out(i)
        out(i) = Trim(element)
        i = i + 1
    Next element

    FilterString = out
End Function


'Function DeepCopyString(originalStr As String) As String
'    Dim copiedStr As String
'
'    copiedStr = StrConv(originalStr, vbFromUnicode)
'    DeepCopyString = copiedStr
'
'    ' Debug.Print "Original string: " & originalStr
'    ' Debug.Print "Copied string: " & copiedStr
'End Function


Function DeepCopyString(originalStr As String) As String
    
    DeepCopyString = Left$(originalStr, Len(originalStr))
End Function


Sub TestFilterString()
    Dim out() As Variant
    Dim i As Integer
    
    out = FilterString("a,b, c,d,af, ae, k, x, ag")
    
    Debug.Print "***************************"
    
    For i = LBound(out) To UBound(out)
        
        Debug.Print out(i)
        Debug.Print "***************************"
    
    Next i
End Sub

' ***************************************************************
' mod_CheckString
'
' ***************************************************************


Option Explicit


Function CheckSubstring(str As String, chk As String) As Boolean
    If InStr(str, chk) > 0 Then
        ' The string contains "chk"
        CheckSubstring = True
    Else
        ' The string does not contain "chk"
        CheckSubstring = False
    End If
End Function


Function IsDrinking(str As String) As Boolean
   ' 가정용 - 사설
    If CheckSubstring(str, "가정") Then
            IsDrinking = True
            Exit Function
    End If
    
    ' 일반용 - 사설
    If CheckSubstring(str, "일반") Then
           IsDrinking = False
            Exit Function
    End If
    
    ' 학교용 - 공공
    If CheckSubstring(str, "학교") Then
             IsDrinking = True
            Exit Function
    End If
        
    ' 민방위용 - 공공
    If CheckSubstring(str, "민방") Then
             IsDrinking = False
            Exit Function
    End If
    
    ' 공동주택용 - 사설
    If CheckSubstring(str, "공동") Then
             IsDrinking = True
            Exit Function
    End If
    
    ' 간이상수도 - 공공
    If CheckSubstring(str, "간이") Then
             IsDrinking = True
            Exit Function
    End If
    
    ' 농생활겸용 - 사설
    If CheckSubstring(str, "겸용") Then
             IsDrinking = False
            Exit Function
    End If
    
    ' 기타 - 사설
    If CheckSubstring(str, "기타") Then
             IsDrinking = False
            Exit Function
    End If
    
    IsDrinking = False
End Function



Function SS_StringCheck(str As String) As String
    ' 가정용 - 사설
    If CheckSubstring(str, "가정") Then
            SS_StringCheck = "g,"
            Exit Function
    End If
    
    ' 일반용 - 사설
    If CheckSubstring(str, "일반") Then
            SS_StringCheck = "h,"
            Exit Function
    End If
    
    ' 학교용 - 공공
    If CheckSubstring(str, "학교") Then
            SS_StringCheck = "i,"
            Exit Function
    End If
        
    ' 민방위용 - 공공
    If CheckSubstring(str, "민방") Then
            SS_StringCheck = "j,"
            Exit Function
    End If
    
    ' 공동주택용 - 사설
    If CheckSubstring(str, "공동") Then
            SS_StringCheck = "k,"
            Exit Function
    End If
    
    ' 간이상수도 - 공공
    If CheckSubstring(str, "간이") Then
            SS_StringCheck = "l,"
            Exit Function
    End If
    
    ' 농생활겸용 - 사설
    If CheckSubstring(str, "겸용") Then
            SS_StringCheck = "m,"
            Exit Function
    End If
    
    ' 기타 - 사설
    If CheckSubstring(str, "기타") Then
            SS_StringCheck = "n,"
            Exit Function
    End If
    
    SS_StringCheck = "n,"
End Function

Function AA_StringCheck(str As String) As String
    
    ' 농업용은 전부 사설, 이중 허가공 - 공공
    If CheckSubstring(str, "전작") Then
            AA_StringCheck = "v,"
            Exit Function
    End If
    
    If CheckSubstring(str, "답작") Then
            AA_StringCheck = "w,"
            Exit Function
    End If
    
    If CheckSubstring(str, "원예") Then
            AA_StringCheck = "x,"
            Exit Function
    End If
    
    If CheckSubstring(str, "축산") Then
            AA_StringCheck = "y,"
            Exit Function
    End If
    
    If CheckSubstring(str, "양어") Then
            AA_StringCheck = "z,"
            Exit Function
    End If
    
    If CheckSubstring(str, "기타") Then
            AA_StringCheck = "aa,"
            Exit Function
    End If
    
    AA_StringCheck = "aa,"
End Function


Function II_StringCheck(str As String) As String
    ' 극가, 지방, 농공 - 공공
    If CheckSubstring(str, "국가") Then
            II_StringCheck = "p,"
            Exit Function
    End If
    
    If CheckSubstring(str, "지방") Then
            II_StringCheck = "q,"
            Exit Function
    End If
    
    If CheckSubstring(str, "농공") Then
            II_StringCheck = "r,"
            Exit Function
    End If
    
    ' 자유입지, 기타 - 사설
    If CheckSubstring(str, "자유입지") Then
            II_StringCheck = "s,"
            Exit Function
    End If
    
    If CheckSubstring(str, "기타") Then
            II_StringCheck = "t,"
            Exit Function
    End If

    II_StringCheck = "t,"
End Function



Function SS_PublicCheck(str As String) As String
    ' 가정용 - 사설
    If CheckSubstring(str, "가정") Then
            SS_PublicCheck = "ac,"
            Exit Function
    End If
    
    ' 일반용 - 사설
    If CheckSubstring(str, "일반") Then
            SS_PublicCheck = "ac,"
            Exit Function
    End If
    
    ' 학교용 - 공공
    If CheckSubstring(str, "학교") Then
            SS_PublicCheck = "ab,"
            Exit Function
    End If
        
    ' 민방위용 - 공공
    If CheckSubstring(str, "민방") Then
            SS_PublicCheck = "ab,"
            Exit Function
    End If
    
    ' 공동주택용 - 사설
    If CheckSubstring(str, "공동") Then
            SS_PublicCheck = "ac,"
            Exit Function
    End If
    
    ' 간이상수도 - 공공
    If CheckSubstring(str, "간이") Then
            SS_PublicCheck = "ab,"
            Exit Function
    End If
    
    ' 농생활겸용 - 사설
    If CheckSubstring(str, "겸용") Then
            SS_PublicCheck = "ac,"
            Exit Function
    End If
    
    ' 기타 - 사설
    If CheckSubstring(str, "기타") Then
            SS_PublicCheck = "ac,"
            Exit Function
    End If
    
    SS_PublicCheck = "ac,"
End Function

Function AA_PublicCheck(str As String) As String
    ' 농업용은 전부 사설, 이중 허가공 - 공공
    If CheckSubstring(str, "전작") Then
            AA_PublicCheck = "ac,"
            Exit Function
    End If
    
    If CheckSubstring(str, "답작") Then
            AA_PublicCheck = "ac,"
            Exit Function
    End If
    
    If CheckSubstring(str, "원예") Then
            AA_PublicCheck = "ac,"
            Exit Function
    End If
    
    If CheckSubstring(str, "축산") Then
            AA_PublicCheck = "ac,"
            Exit Function
    End If
    
    If CheckSubstring(str, "양어") Then
            AA_PublicCheck = "ac,"
            Exit Function
    End If
    
    If CheckSubstring(str, "기타") Then
            AA_PublicCheck = "ac,"
            Exit Function
    End If
    
    AA_PublicCheck = "ac,"
End Function


Function II_PublicCheck(str As String) As String
    ' 극가, 지방, 농공 - 공공
    If CheckSubstring(str, "국가") Then
            II_PublicCheck = "ab,"
            Exit Function
    End If
    
    If CheckSubstring(str, "지방") Then
            II_PublicCheck = "ab,"
            Exit Function
    End If
    
    If CheckSubstring(str, "농공") Then
            II_PublicCheck = "ab,"
            Exit Function
    End If
    
    ' 자유입지, 기타 - 사설
    If CheckSubstring(str, "자유입지") Then
            II_PublicCheck = "ac,"
            Exit Function
    End If
    
    If CheckSubstring(str, "기타") Then
            II_PublicCheck = "ac,"
            Exit Function
    End If

    II_PublicCheck = "ac,"
End Function


' ***************************************************************
' Sheet1(ref)
'
' ***************************************************************


Option Explicit

Private Sub CommandButton1_Click()
    ActiveSheet.Visible = False
End Sub
' ***************************************************************
' Sheet2(ref1)
'
' ***************************************************************




Option Explicit

Private Sub CommandButton1_Click()
    ActiveSheet.Visible = False
End Sub
' ***************************************************************
' modTable
'
' ***************************************************************


Option Explicit

Sub test_tableindex()
    Dim tbl As ListObject
    Set tbl = ActiveSheet.ListObjects("tableCNU")
    
    Dim values() As Variant
    values = tbl.ListColumns("nonsan").DataBodyRange.Value
    
    Dim i As Long
    For i = 1 To UBound(values, 1)
        ActiveSheet.Cells(29, Chr(Asc("A") + i)).Value = values(i, 1)
    Next i
End Sub


Sub initialize_CNU(area As String)
    Dim tbl As ListObject
    Set tbl = Sheets("ref").ListObjects("tableCNU")
    
    Dim values() As Variant
    
    If area = "" Then
        area = "default"
    End If
    
    values = tbl.ListColumns(area).DataBodyRange.Value
         
    '전라남도, 목포시, 2020 환경부 지하수업무수행지침
    SS(svGAJUNG, 1) = values(1, 1)
    SS(svGAJUNG, 2) = values(2, 1)
    SS_CITY = values(3, 1)
    
    SS(svILBAN, 1) = values(4, 1)
    SS(svILBAN, 2) = values(5, 1)
    
    SS(svSCHOOL, 1) = values(6, 1)
    SS(svSCHOOL, 2) = values(7, 1)
    
    SS(svGONGDONG, 1) = values(8, 1)
    SS(svGONGDONG, 2) = values(9, 1)
    
    SS(svMAEUL, 1) = values(10, 1)
    SS(svMAEUL, 2) = values(11, 1)
    
'----------------------------------------

    AA(avJEONJAK, 1) = values(12, 1)
    AA(avJEONJAK, 2) = values(13, 1)
    
    AA(avDAPJAK, 1) = values(14, 1)
    AA(avDAPJAK, 2) = values(15, 1)
    
    AA(avWONYE, 1) = values(16, 1)
    AA(avWONYE, 2) = values(17, 1)
    
    AA(avCOW, 1) = values(18, 1)
    AA(avCOW, 2) = values(19, 1)
    
    AA(avPIG, 1) = values(20, 1)
    AA(avPIG, 2) = values(21, 1)
    
    AA(avCHICKEN, 1) = values(22, 1)
    AA(avCHICKEN, 2) = values(23, 1)
    
End Sub


Sub initialize_JIYEOL(area As String)
    Dim tbl As ListObject
    Set tbl = Sheets("ref1").ListObjects("tableJIYEOL")
    
    Dim values() As Variant
    
    
    If (area = "") Then
        area = "default"
    End If
    
    values = tbl.ListColumns(area).DataBodyRange.Value
           
    '전라남도, 목포시, 2020 환경부 지하수업무수행지침
    SS(svGAJUNG, 1) = values(1, 1)
    SS(svGAJUNG, 2) = values(2, 1)
    SS_CITY = values(3, 1)
    
    SS(svILBAN, 1) = values(4, 1)
    SS(svILBAN, 2) = values(5, 1)
    
    SS(svSCHOOL, 1) = values(6, 1)
    SS(svSCHOOL, 2) = values(7, 1)
    
    SS(svGONGDONG, 1) = values(8, 1)
    SS(svGONGDONG, 2) = values(9, 1)
    
    SS(svMAEUL, 1) = values(10, 1)
    SS(svMAEUL, 2) = values(11, 1)
    
'----------------------------------------

    AA(avJEONJAK, 1) = values(12, 1)
    AA(avJEONJAK, 2) = values(13, 1)
    
    AA(avDAPJAK, 1) = values(14, 1)
    AA(avDAPJAK, 2) = values(15, 1)
    
    AA(avWONYE, 1) = values(16, 1)
    AA(avWONYE, 2) = values(17, 1)
    
    AA(avCOW, 1) = values(18, 1)
    AA(avCOW, 2) = values(19, 1)
    
    AA(avPIG, 1) = values(20, 1)
    AA(avPIG, 2) = values(21, 1)
    
    AA(avCHICKEN, 1) = values(22, 1)
    AA(avCHICKEN, 2) = values(23, 1)
End Sub




' ***************************************************************
' UserForm_SS
'
' ***************************************************************



' Optionbutton1 - 가정용
' Optionbutton2 - 일반용
' Optionbutton3 - 청소용
' Optionbutton4 - 민방위용
' Optionbutton5 - 학교용
' Optionbutton6 - 공동주택용
' Optionbutton7 - 간이상수도
' Optionbutton8 - 농생활겸용
' Optionbutton9 - 기타


Private Sub CommandButton1_Click()
    Dim i As Integer
    Dim options() As Variant
    Dim selectedOption As String
    
    ' Assign captions to an array
    options = Array("가정용", "일반용", "청소용", "민방위용", "학교용", "공동주택용", "간이상수도", "농생활겸용", "기타")
    
    ' Loop through OptionButtons to find the selected one
    For i = 0 To 8
        If Controls("OptionButton" & i + 1).Value Then
            selectedOption = options(i)
            Exit For
        End If
    Next i
    
    ' Set the value of the active cell
    If selectedOption <> "" Then
        ActiveCell.Value = selectedOption
        Unload Me
    Else
        MsgBox "Please select an option."
    End If
End Sub

Private Sub CommandButton2_Click()
  Unload Me
End Sub

Private Sub UserForm_Initialize()
    Dim i As Integer
    
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    
   OptionButton1.Value = True
    
End Sub




' ***************************************************************
' UserForm_AA
'
' ***************************************************************


' Optionbutton1 - 답작용
' Optionbutton2 - 전작용
' Optionbutton3 - 원예용
' Optionbutton4 - 축산용
' Optionbutton5 - 양어장용
' Optionbutton6 - 기타


Private Sub CommandButton1_Click()
    Dim i As Integer
    Dim options() As Variant
    Dim selectedOption As String
    
    ' Assign captions to an array
    options = Array("답작용", "전작용", "원예용", "축산업", "양어장용", "기타")
    
    ' Loop through OptionButtons to find the selected one
    For i = 0 To 5
        If Controls("OptionButton" & i + 1).Value Then
            selectedOption = options(i)
            Exit For
        End If
    Next i
    
    ' Set the value of the active cell
    If selectedOption <> "" Then
        ActiveCell.Value = selectedOption
        Unload Me
    Else
        MsgBox "Please select an option."
    End If
End Sub


Private Sub CommandButton2_Click()
  Unload Me
End Sub


Private Sub UserForm_Initialize()
    Dim i As Integer
    
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    
   OptionButton1.Value = True
    
End Sub




' ***************************************************************
' vbGit
'
' ***************************************************************


Sub GitSave()
    
    DeleteAndMake
    ExportModules
    PrintAllCode
    PrintAllContainers
    
End Sub

Sub DeleteAndMake()
        
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim parentFolder As String: parentFolder = ThisWorkbook.Path & "\VBA"
    Dim childA As String: childA = parentFolder & "\VBA-Code_Together"
    Dim childB As String: childB = parentFolder & "\VBA-Code_By_Modules"
        
    On Error Resume Next
    fso.DeleteFolder parentFolder
    On Error GoTo 0
    
    MkDir parentFolder
    MkDir childA
    MkDir childB
    
End Sub

Sub PrintAllCode()
    
    Dim item  As Variant
    Dim textToPrint As String
    Dim lineToPrint As String
    
    For Each item In ThisWorkbook.VBProject.VBComponents
        lineToPrint = item.codeModule.Lines(1, item.codeModule.CountOfLines)
        Debug.Print lineToPrint
        textToPrint = textToPrint & vbCrLf & lineToPrint
    Next item
    
    Dim pathToExport As String: pathToExport = ThisWorkbook.Path & "\VBA\VBA-Code_Together\"
    If Dir(pathToExport) <> "" Then Kill pathToExport & "*.*"
    SaveTextToFile textToPrint, pathToExport & "all_code.vb"
    
End Sub

Sub PrintAllContainers()
    
    Dim item  As Variant
    Dim textToPrint As String
    Dim lineToPrint As String
    
    For Each item In ThisWorkbook.VBProject.VBComponents
        lineToPrint = item.Name
        Debug.Print lineToPrint
        textToPrint = textToPrint & vbCrLf & lineToPrint
    Next item
    
    Dim pathToExport As String: pathToExport = ThisWorkbook.Path & "\VBA\VBA-Code_Together\"
    SaveTextToFile textToPrint, pathToExport & "all_modules.vb"
    
End Sub

Sub ExportModules()
       
    Dim pathToExport As String: pathToExport = ThisWorkbook.Path & "\VBA\VBA-Code_By_Modules\"
    
    If Dir(pathToExport) <> "" Then
        Kill pathToExport & "*.*"
    End If
     
    Dim wkb As Workbook: Set wkb = Excel.Workbooks(ThisWorkbook.Name)
    
    Dim unitsCount As Long
    Dim filePath As String
    Dim component As VBIDE.VBComponent
    Dim tryExport As Boolean

    For Each component In wkb.VBProject.VBComponents
        tryExport = True
        filePath = component.Name
       
        Select Case component.Type
            Case vbext_ct_ClassModule
                filePath = filePath & ".cls"
            Case vbext_ct_MSForm
                filePath = filePath & ".frm"
            Case vbext_ct_StdModule
                filePath = filePath & ".bas"
            Case vbext_ct_Document
                tryExport = False
        End Select
        
        If tryExport Then
            Debug.Print unitsCount & " exporting " & filePath
            component.Export pathToExport & "\" & filePath
        End If
    Next

    Debug.Print "Exported at " & pathToExport
    
End Sub

Sub SaveTextToFile(dataToPrint As String, pathToExport As String)
    
    Dim fileSystem As Object
    Dim textObject As Object
    Dim fileName As String
    Dim newFile  As String
    Dim shellPath  As String
    
    If Dir(ThisWorkbook.Path & newFile, vbDirectory) = vbNullString Then MkDir ThisWorkbook.Path & newFile
    
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Set textObject = fileSystem.CreateTextFile(pathToExport, True)
    
    textObject.WriteLine dataToPrint
    textObject.Close
        
    On Error GoTo 0
    Exit Sub

CreateLogFile_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CreateLogFile of Sub mod_TDD_Export"

End Sub
