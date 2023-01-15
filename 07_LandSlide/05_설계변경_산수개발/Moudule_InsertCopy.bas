Attribute VB_Name = "Moudule_InsertCopy"
Option Explicit

Dim NSNE(1 To 3, 0 To 1) As String

'용역비 총괄 - 1
'영향조사 명세서 - 2
'사후관리 명세서 - 3

Private Function getID_from_sheet() As Integer
    Dim strSheetName As String
    Dim ireturn As Integer
    
   strSheetName = ActiveSheet.Name
    
   ireturn = 999
   If InStr(strSheetName, "총괄") <> 0 Then ireturn = 1
   If InStr(strSheetName, "영향조사") <> 0 Then ireturn = 2
   If InStr(strSheetName, "사후관리") <> 0 Then ireturn = 3
    
   getID_from_sheet = ireturn
End Function

Private Sub initialize_multidimensional_array()
    Dim myArray
    myArray = Evaluate("{1,""a"";2,""b"";3,""c"";4,""d"";5,""e"";6,""f"";7,""g"";8,""h"";9,""i"";9,""j"";10,""k"";11,""l""}")
End Sub


Private Sub preprocess()
    
    NSNE(1, 0) = "C"
    NSNE(1, 1) = "E"
    
    NSNE(2, 0) = "C"
    NSNE(2, 1) = "I"
    
    NSNE(3, 0) = "C"
    NSNE(3, 1) = "L"

End Sub


Private Sub change_interior(color As Integer)

     With Selection.Font
        .Name = "맑은 고딕"
        .Size = 9
        .Bold = True
        .color = color
    End With
    
End Sub


Private Sub set_oneline_interior(ByVal i As Integer, ByVal n0 As String, ByVal n1 As String)

    Range(Cells(i, n0), Cells(i, n1)).Select
    change_interior (vbRed)
    Range(Cells(i + 1, n0), Cells(i + 1, n1)).Select
    change_interior (vbBlack)
    
End Sub


Sub make_copy()
    Dim val As Variant
    Dim rows, rowe, i, sheetID As Long

    Call preprocess
    sheetID = getID_from_sheet
    If sheetID = 999 Then Exit Sub
          
    Call ReplaceSUMtoEachCell
    val = insert_in_a_sheet_by_input()
    rows = val(0)
    rowe = val(1)
    
    For i = rows To rowe Step 2
        Call MakeCopyCurrentRow(i, NSNE(sheetID, 0), NSNE(sheetID, 1))
    Next i
    
    For i = rows To rowe Step 2
        Call set_oneline_interior(i, NSNE(sheetID, 0), NSNE(sheetID, 1))
    Next i
    
End Sub


Function insert_in_a_sheet_by_input() As Variant
    Dim Sel As Range
    Dim strReturn As String
    Dim retVal As Variant
    
    Dim ns, ne As Long
    Dim rp, i As Long
    Dim val(0 To 1) As Long
    
    Set Sel = Application.InputBox("원하는 영역에 값적용", "범위선택", Type:=8)
       
    retVal = ExtractStartEnd_FromRange(Sel)
    
    ns = CLng(retVal(0))
    ne = CLng(retVal(1))
           
    rp = ns
    val(0) = ns
    For i = ns To ne
        retVal = Application.Run("MainMoudule.insert_row2", rp)
        Debug.Print retVal
        rp = Selection.row + 1
    Next i
    val(1) = rp - 1
    
    insert_in_a_sheet_by_input = val
End Function


Private Sub Relative2Absolute()
    For Each c In Selection
        If c.HasFormula = True Then
            c.Formula = Application.ConvertFormula(c.Formula, xlA1, xlA1, xlAbsolute)
        End If
    Next c
End Sub


Public Sub convertAbsolute()
    Dim Sel As Range
    Dim retVal As Variant
    
    Set Sel = Application.InputBox("원하는 영역에 값적용", "범위선택", Type:=8)
    retVal = ConvertAbsoluteAddress(Sel)
End Sub


Private Function ConvertAbsoluteAddress(rng As Range) As String
    Dim cell As Range
            
    For Each cell In rng
        If cell.HasFormula = True Then
                cell.Formula = Application.ConvertFormula(cell.Formula, xlA1, xlA1, xlAbsolute)
        End If
    Next cell
    'ConvertAbsoluteAddress = cell.Formula
End Function


Private Sub test()
    Dim strText As String
        
    strText = ConvertAbsoluteAddress(Range("C5"))
    Debug.Print strText
End Sub


'nrow - copy down row
'ns - start column
'ne - end column

Private Sub MakeCopyCurrentRow(ByVal nrow As Long, Optional ByVal ns As String = "C", Optional ByVal ne As String = "I")
    
    Range(ns & CStr(nrow) & ":" & ne & CStr(nrow)).Select
    Selection.AutoFill Destination:=Range(ns & CStr(nrow) & ":" & ne & CStr(nrow + 1)), Type:=xlFillDefault
    
End Sub

Private Sub callingPrivateSubTest()

    Dim i As Integer

    i = Application.Run("MainMoudule.getSheetIndex", "계산서")
    Debug.Print "hello", i
    
End Sub






