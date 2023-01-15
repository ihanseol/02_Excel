Attribute VB_Name = "mod_util"
Option Explicit


Function ConvertToLongInteger(ByVal stValue As String) As Long
    On Error GoTo ConversionFailureHandler
    ConvertToLongInteger = CLng(stValue)         'TRY to convert to an Integer value
    Exit Function                                'If we reach this point, then we succeeded so exit

ConversionFailureHandler:
    'IF we've reached this point, then we did not succeed in conversion
    'If the error is type-mismatch, clear the error and return numeric 0 from the function
    'Otherwise, disable the error handler, and re-run the code to allow the system to
    'display the error
    If Err.Number = 13 Then                      'error # 13 is Type mismatch
        Err.Clear
        ConvertToLongInteger = 0
        Exit Function
    Else
        On Error GoTo 0
        Resume
    End If

End Function

Function sheets_count() As Long

    Dim i, nSheetsCount, nWell  As Integer
    Dim strSheetsName(50) As String
    
    
    nSheetsCount = ThisWorkbook.Sheets.Count
    nWell = 0
      
    
    For i = 1 To nSheetsCount
        strSheetsName(i) = ThisWorkbook.Sheets(i).name
        'MsgBox (strSheetsName(i))
        If (ConvertToLongInteger(strSheetsName(i)) <> 0) Then
            nWell = nWell + 1
        End If
    Next
    
    'MsgBox (CStr(nWell))
    sheets_count = nWell

End Function




