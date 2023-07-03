Attribute VB_Name = "BaseData_ETC"
Option Explicit

'------------------------------------------------------------------------------------------
' 2022/6/11

Public Enum cellLowHi
    cellLOW = 0
    cellHI = 1
End Enum

Function GetNumberOfWell() As Integer
    Dim save_name As String
    Dim n As Integer
    
    save_name = ActiveSheet.Name
    Sheets("Well").Activate
    Sheets("Well").Range("A30").Select
    Selection.End(xlUp).Select
    n = CInt(GetNumeric2(Selection.value))
    
    GetNumberOfWell = n
End Function

Public Function sheets_count() As Long
    Dim i, nSheetsCount, nWell  As Integer
    Dim strSheetsName(50) As String
    
    nSheetsCount = ThisWorkbook.Sheets.count
    nWell = 0
    
    For i = 1 To nSheetsCount
        strSheetsName(i) = ThisWorkbook.Sheets(i).Name
        'MsgBox (strSheetsName(i))
        If (ConvertToLongInteger(strSheetsName(i)) <> 0) Then
            nWell = nWell + 1
        End If
    Next
    
    'MsgBox (CStr(nWell))
    sheets_count = nWell
End Function


Function GetNumeric2(ByVal CellRef As String)
    Dim StringLength, i  As Integer
    Dim result      As String
    
    StringLength = Len(CellRef)
    For i = 1 To StringLength
        If IsNumeric(Mid(CellRef, i, 1)) Then result = result & Mid(CellRef, i, 1)
    Next i
    GetNumeric2 = result
End Function

'********************************************************************************************************************************************************************************
'Function Name                    : IsWorkBookOpen(ByVal OWB As String)
'Function Description             : Function to check whether specified workbook is open
'Data Parameters                  : OWB:- Specify name or path to the workbook. eg: "Book1.xlsx" or "C:\Users\Kannan.S\Desktop\Book1.xlsm"

'********************************************************************************************************************************************************************************
Function IsWorkBookOpen(ByVal OWB As String) As Boolean
    IsWorkBookOpen = False
    Dim WB          As Excel.Workbook
    Dim WBName      As String
    Dim WBPath      As String
    Dim OWBArray    As Variant
    
    Err.Clear
    
    On Error Resume Next
    OWBArray = Split(OWB, Application.PathSeparator)
    Set WB = Application.Workbooks(OWBArray(UBound(OWBArray)))
    WBName = OWBArray(UBound(OWBArray))
    WBPath = WB.Path & Application.PathSeparator & WBName
    
    If Not WB Is Nothing Then
        If UBound(OWBArray) > 0 Then
            If LCase(WBPath) = LCase(OWB) Then IsWorkBookOpen = True
        Else
            IsWorkBookOpen = True
        End If
    End If
    Err.Clear
    
End Function

'------------------------------------------------------------------------------------------

Public Function GetLengthByColor(ByVal tabColor As Variant) As Integer
    Dim n_sheets, i, j, nTab As Integer
    n_sheets = sheets_count()
    
    nTab = 0
    
    For i = 1 To n_sheets
        If (Sheets(CStr(i)).Tab.Color = tabColor) Then
            nTab = nTab + 1
        End If
    Next i
    
    GetLengthByColor = nTab
End Function

Private Sub get_tabsize_by_well(ByRef nof_sheets As Integer, ByRef nof_unique_tab As Variant, ByRef n_tabcolors As Variant)
    ' n_tabcolors : return value
    ' nof_unique_tab : return value
    
    Dim n_sheets, i, j As Integer
    Dim limit()     As Integer
    Dim arr_tabcolors(), new_tabcolors() As Variant
    
    n_sheets = sheets_count()
    
    ReDim arr_tabcolors(1 To n_sheets)
    ReDim new_tabcolors(1 To n_sheets)
    ReDim limit(0 To n_sheets)
    
    For i = 1 To n_sheets
        arr_tabcolors(i) = Sheets(CStr(i)).Tab.Color
    Next i
    
    new_tabcolors = getUnique(arr_tabcolors)
    
    For i = 0 To UBound(new_tabcolors)
        limit(i) = GetLengthByColor(new_tabcolors(i))
    Next i
    
    nof_sheets = n_sheets
    nof_unique_tab = limit
    n_tabcolors = new_tabcolors
End Sub
