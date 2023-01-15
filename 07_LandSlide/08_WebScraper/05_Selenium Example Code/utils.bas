Attribute VB_Name = "utils"

Private Declare PtrSafe Function FindWindowExA Lib "user32.dll" ( _
  ByVal hwndParent As LongPtr, _
  ByVal hwndChildAfter As LongPtr, _
  ByVal lpszClass As String, _
  ByVal lpszWindow As String) As Long

Private Declare PtrSafe Function AccessibleObjectFromWindow Lib "oleacc.dll" ( _
  ByVal hwnd As LongPtr, _
  ByVal dwId As Long, _
  ByRef riid As Any, _
  ByRef ppvObject As IAccessible) As Long


''
' Loads a translation table in a dictionary from a worksheet
' The first column is the result and second is the input
' Usage:
'  Set dict = LoadTranslation([Sheet5])
'  Debug.Print = dict("Cancel")
''
Public Function LoadTranslation(sheet As Worksheet) As Collection
  Dim values(), translation$
  Set LoadTranslation = New Collection
  values = sheet.Cells.CurrentRegion.Value2
  For r = LBound(values) To UBound(values)
    If Not IsEmpty(values(r, 1)) Then translation = values(r, 1)
    LoadTranslation.Add translation, values(r, 2)
  Next
End Function

''
' Returns all the active instances of Excel
''
Public Function GetExcelInstances() As Collection
  Dim guid&(0 To 4), app As Object, hwnd
  guid(0) = &H20400
  guid(1) = &H0
  guid(2) = &HC0
  guid(3) = &H46000000
  
  Set GetExcelInstances = New Collection
  Do
    hwnd = FindWindowExA(0, hwnd, "XLMAIN", vbNullString)
    If hwnd = 0 Then Exit Do
    hwnd = FindWindowExA(hwnd, 0, "XLDESK", vbNullString)
    If hwnd Then
      hwnd = FindWindowExA(hwnd, 0, "EXCEL7", vbNullString)
      If hwnd Then
        If AccessibleObjectFromWindow(hwnd, &HFFFFFFF0, guid(0), app) = 0 Then
          GetExcelInstances.Add app.Application
        End If
      End If
    End If
  Loop
End Function

''
' Returns True if a file is locked by another application, False otherwise
''
Public Function IsFileLocked(file_path As String) As Boolean
  Dim num As Long
  
  On Error Resume Next
  Name file_path As file_path
  num = Err.Number
  On Error GoTo 0
  
  If num <> 0 And num <> 75 Then Error num
  IsFileLocked = num <> 0
End Function


''
' A simple hash function
''
Public Function HashFnv$(str As String)
  Dim bytes() As Byte, i&, lo&, hi&
  lo = &H9DC5&
  hi = &H11C&
  bytes = str
  For i = 0 To UBound(bytes) Step 2
    lo = 31& * ((bytes(i) + bytes(i + 1) * 256&) Xor (lo And 65535))
    hi = 31& * hi + lo \ 65536 And 65535
  Next
  lo = (lo And 65535) + (hi And 32767) * 65536 Or (&H80000000 And -(hi And 32768))
  HashFnv = Hex(lo)
End Function

