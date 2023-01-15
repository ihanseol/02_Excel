Attribute VB_Name = "MainMoudule"
Option Explicit

Public Enum CellCopyMode
    ccmValue = 0
    ccmFormula = 1
End Enum

Public SHEET_NUMBER(1 To 3) As Integer
Public ARRAY_CELL_COPY_MODE(1 To 26) As Integer


Private Sub delay(sec As Integer)

    Application.Wait Now() + TimeSerial(0, 0, sec)

End Sub


Private Sub TurnOffStuff()

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    

End Sub

Private Sub TurnOnStuff()

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
End Sub



Private Function get_first_row() As Integer
    
    Dim firstRow, lasRow As Long
    
    If ActiveWindow.FreezePanes Then
        firstRow = getFreezingRow()
    Else
        firstRow = 1
    End If
    
    get_first_row = firstRow

End Function



Private Function getSheetIndex(strSheetName As String) As Integer
    Dim CutString As String
    Dim ireturn, i As Integer
    
    CutString = Right(strSheetName, 3)
    
    Select Case CutString
        Case "°è»ê¼­"
            ireturn = 1
        Case "ÃÑ°ýÇ¥"
            ireturn = 2
        Case "³»¿ª¼­"
            ireturn = 3
        Case Else
            ireturn = 0
    End Select
       
   If InStr(strSheetName, "ÃÑ°ý") <> 0 Then ireturn = 1
   If InStr(strSheetName, "¿µÇâÁ¶»ç") <> 0 Then ireturn = 2
   If InStr(strSheetName, "»çÈÄ°ü¸®") <> 0 Then ireturn = 3
    
   getSheetIndex = ireturn
    
End Function

Private Sub search_red_tab()

    Dim i, c As Integer
    
    c = 1
    For i = 1 To Worksheets.Count
          Sheets(i).Activate
           If (Sheets(i).Tab.color = vbRed Or Sheets(i).Tab.color = 192) And c <= 3 Then
                 SHEET_NUMBER(getSheetIndex(Sheets(i).Name)) = i
                 c = c + 1
           End If
    Next i

End Sub


Private Function getFreezingRow() As Integer

    Dim lngRowNumber As Long, _
        lngColNumber As Long
    Dim strColLetter As String
    Dim rngFreezeWindow As Range
    
    With ActiveWindow
        If .SplitRow = 0 And .SplitColumn = 0 Then
            MsgBox "There are no Rows or Columns frozen.", vbExclamation, "Frozen Cell Address Editor"
            Exit Function
        Else
            lngRowNumber = .SplitRow + 1
            lngColNumber = .SplitColumn + 1
        End If
    End With
    
    'Code to convert a Column Number to a Column String has been adapted from: _
    http://www.freevbcode.com/ShowCode.asp?ID=4303
    If lngColNumber > 26 Then
        
        strColLetter = Chr(Int((lngColNumber - 1) / 26) + 64) & _
            Chr(((lngColNumber - 1) Mod 26) + 65)
    Else
        'Columns A-Z
        strColLetter = Chr(lngColNumber + 64)
    End If
    
    Set rngFreezeWindow = Range(strColLetter & lngRowNumber)
    'Debug.Print rngFreezeWindow.Address
    
    getFreezingRow = rngFreezeWindow.row
    
    
End Function


Private Function lastRowByKey(cell As String) As Long

    lastRowByKey = Range(cell).End(xlDown).row

End Function


Private Function lastRowByFind() As Long
    Dim LASTROW As Long
    
    LASTROW = Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
    
    lastRowByFind = LASTROW
End Function



'cell copy mode
' 0 : copy by value
' 1 : copy by formula

Private Function getIndex(inx As String) As Integer
    getIndex = Asc(UCase(inx)) - 64
End Function

'initial setting
Private Sub preProcessGlobal(nsheet As Integer)
    Dim i As Integer
    
    '¿ø°¡°è»ê¼­
    If nsheet = 1 Then
        For i = 1 To 26
            ARRAY_CELL_COPY_MODE(i) = ccmFormula
        Next i
        
        ARRAY_CELL_COPY_MODE(getIndex("g")) = ccmValue
        ARRAY_CELL_COPY_MODE(getIndex("h")) = ccmValue
        
        Exit Sub
    End If
    
    
    '³»¿ª¼­ÃÑ°ýÇ¥
    If nsheet = 2 Then
        For i = 1 To 26
            ARRAY_CELL_COPY_MODE(i) = ccmFormula
        Next i
        
        ARRAY_CELL_COPY_MODE(getIndex("e")) = ccmValue
        ARRAY_CELL_COPY_MODE(getIndex("f")) = ccmValue
        
        Exit Sub
    End If
        
    
    '³»¿ª¼­
    If nsheet = 3 Then
        For i = 1 To 26
            ARRAY_CELL_COPY_MODE(i) = ccmFormula
        Next i
        
        ARRAY_CELL_COPY_MODE(getIndex("e")) = ccmValue
        ARRAY_CELL_COPY_MODE(getIndex("f")) = ccmValue
        ARRAY_CELL_COPY_MODE(getIndex("i")) = ccmValue
        ARRAY_CELL_COPY_MODE(getIndex("k")) = ccmValue
        ARRAY_CELL_COPY_MODE(getIndex("m")) = ccmValue
        
        Exit Sub
    End If
    
End Sub


Private Sub change_interior(color As Integer)

     With Selection.Font
        .Name = "¸¼Àº °íµñ"
        .Size = 9
    End With
    Selection.Font.Bold = True
    With Selection.Font
        .color = color
    End With
    
End Sub


Private Sub set_oneline_interior(ByVal nsheet As Integer, ByVal i As Integer, ByVal n0 As String, ByVal n1 As String)

    Range(Cells(i, n0), Cells(i, n1)).Select
    change_interior (vbRed)
    Range(Cells(i + 1, n0), Cells(i + 1, n1)).Select
    change_interior (vbBlack)
    
End Sub


Private Sub copy_one_cell_v2(nsheet As Integer, row As Integer, col As Integer)

    If Cells(row, col).Value = "" Then
        Exit Sub
    End If
    
    If ARRAY_CELL_COPY_MODE(col) = ccmValue Then
        Cells(row, col).Copy
        Cells(row + 1, col).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    Else
        Cells(row, col).Copy
        Cells(row + 1, col).PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    End If
    
End Sub



'n0 -- start column
'n1 -- end column
'v2 -- is more faster version

Private Sub copy_oneline_action(ByVal nsheet As Integer, ByVal row As Integer, ByVal n0 As String, ByVal n1 As String)

        Dim a0, a1, i As Integer
        
        a0 = getIndex(n0)
        a1 = getIndex(n1)
        
        For i = a0 To a1
            Call copy_one_cell_v2(nsheet, row, i)
        Next i
        
    
End Sub



'1bun sheet -- ¿ø°¡°è»ê¼­ - 1
'2bun sheet -- ³»¿ª¼­ÃÑ°ýÇ¥ - 2
'3bun sheet -- ³»¿ª¼­ - 3


Private Sub do_copy(nsheet As Integer, n0 As String, n1 As String)
    
    Dim firstRow, lasRow, rp As Long
    Dim i As Long
    
    Dim dTime As Double
    
    dTime = Now()
    
    firstRow = get_first_row()
    lasRow = lastRowByFind()
    preProcessGlobal (nsheet)
    
    For i = firstRow To lasRow Step 2
       Call copy_oneline_action(nsheet, i, n0, n1)
       Call set_oneline_interior(nsheet, i, n0, n1)
    Next i
        
    Debug.Print "DoCopy (sheet:" & nsheet & ") Time is : " & (Now() - dTime) * 1000

End Sub


Private Sub merge_two_rows(ByVal nsheet As Integer, ByVal nrow As Integer, ByVal c1 As String, ByVal c2 As String)

    Dim col_start, col_end, i, r  As Integer
    
    col_start = getIndex(c1)
    col_end = getIndex(c2)
    
    For i = col_start To col_end
        
        r = nrow
        Range(Cells(r, i), Cells(r + 1, i)).Select
        
        With Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        Selection.Merge
    Next i
    
End Sub


Private Sub merge_action(ByVal nsheet As Integer, ByVal c1 As String, ByVal c2 As String)

    Dim firstRow, lasRow, rp As Long
    Dim i As Long
    
    firstRow = get_first_row()
    lasRow = lastRowByFind()
    
    For i = firstRow To lasRow Step 2
            Call merge_two_rows(nsheet, i, c1, c2)
    Next i
        
End Sub


Private Sub merge_cell(ByVal nsheet As Integer)

    If nsheet = 1 Then Call merge_action(nsheet, "d", "e")
    If nsheet = 2 Then Call merge_action(nsheet, "c", "d")
    If nsheet = 3 Then Call merge_action(nsheet, "c", "d")

End Sub



Private Sub generate_mergecell()
    Dim i As Integer
    Dim dTime As Double
    
    Call TurnOffStuff
    
    dTime = Now()
    
    For i = 1 To 3
        Sheets(SHEET_NUMBER(i)).Activate
        Call merge_cell(i)
    Next i
    
    Debug.Print " Merge Timer : " & (Now() - dTime) * 1000
    
    Call TurnOnStuff
End Sub



Private Sub duplicate_builder_test()
    Dim i As Integer

    Call search_red_tab
    
     For i = 1 To 3
        Sheets(SHEET_NUMBER(i)).Activate
        'Call delay(2)
        Call insert_in_a_sheet
    Next i
    

End Sub


Sub duplicate_builder()
    Dim i As Integer
    
    Call TurnOffStuff
    
    Call search_red_tab
    
     For i = 1 To 3
        Sheets(SHEET_NUMBER(i)).Activate
        Call insert_in_a_sheet
    Next i
    
    
    Sheets(SHEET_NUMBER(1)).Activate
    Call do_copy(1, "f", "h")
    
    Sheets(SHEET_NUMBER(2)).Activate
    Call do_copy(2, "e", "j")
    
    Sheets(SHEET_NUMBER(3)).Activate
    Call do_copy(3, "e", "n")
    
    Call TurnOnStuff
    
    'merge_cell(1)

End Sub




Private Sub insert_in_a_sheet()
' rp -- row position

    Dim rg As Range
    Dim firstRow, lasRow, rp As Long
    Dim i As Integer
    
    firstRow = get_first_row()
    lasRow = lastRowByFind()
    
        
    rp = firstRow
    
    For i = 1 To lasRow - firstRow + 1
        
        Call insert_row2(rp)
        rp = Selection.row + 1
        
    Next i
    
End Sub


Private Sub insert_row()
    Dim nrow As Long
    nrow = Selection.row + 1
    
    Range(nrow & ":" & nrow).Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
End Sub


Private Sub insert_row2(ByVal nrow As Long)
    nrow = nrow + 1
    Range(nrow & ":" & nrow).Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
End Sub


