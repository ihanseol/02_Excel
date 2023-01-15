Attribute VB_Name = "mod_util"
Option Explicit

Function Contains(objCollection As Object, strName As String) As Boolean
    Dim o           As Object
    
    On Error Resume Next
    Set o = objCollection(strName)
    Contains = (Err.Number = 0)
    Err.Clear
End Function

Function ConvertToLongInteger(ByVal stValue As String) As Long
    On Error GoTo ConversionFailureHandler
    ConvertToLongInteger = CLng(stValue)        'TRY to convert to an Integer value
    Exit Function        'If we reach this point, then we succeeded so exit
    
ConversionFailureHandler:
    'IF we've reached this point, then we did not succeed in conversion
    'If the error is type-mismatch, clear the error and return numeric 0 from the function
    'Otherwise, disable the error handler, and re-run the code to allow the system to
    'display the error
    If Err.Number = 13 Then        'error # 13 is Type mismatch
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

' https://www.google.com/search?q=excel+vba+how+to+get+number+from+string&oq=excel+vba+how+to+get+number+from+string&aqs=chrome..69i57&sourceid=chrome&ie=UTF-8
' https://stackoverflow.com/questions/28771802/extract-number-from-string-in-vba

Function GetNumbers(str As String) As Long
    Dim regex       As Object
    Dim matches     As Variant
    
    Set regex = CreateObject("vbscript.regexp")
    
    regex.Pattern = "(\d+)"
    regex.Global = True
    
    Set matches = regex.Execute(str)
    GetNumbers = matches(0)
End Function

Function CleanString(strIn As String) As String
    Dim objRegex
    Set objRegex = CreateObject("vbscript.regexp")
    With objRegex
        .Global = True
        .Pattern = "[^\d]+"
        CleanString = .Replace(strIn, vbNullString)
    End With
End Function

'https://stackoverflow.com/questions/40365573/excel-vba-extract-numeric-value-in-string
'Requires a reference to Microsoft VBScript Regular Expressions X.X

Public Function ExtractNumber(inValue As String) As Double
    With New regExp
        .Pattern = "(\d{1,3},?)+(\.\d{2})?"
        .Global = True
        If .test(inValue) Then
            ExtractNumber = CDbl(.Execute(inValue)(0))
        End If
    End With
End Function

'https://stackoverflow.com/questions/50994883/how-to-extract-numbers-from-a-text-string-in-vba

Sub ExtractNumbers()
    Dim str         As String, regex As regExp, matches As MatchCollection, match As match
    
    str = "ID CSys ID Set ID Set Value Set Title 7026..Plate Top MajorPrn Stress 7027..Plate Top MinorPrn Stress 7033..Plate Top VonMises Stress"
    
    Set regex = New regExp
    regex.Pattern = "\d+"        '~~~> Look for variable length numbers only
    regex.Global = True
    
    If (regex.test(str) = True) Then
        Set matches = regex.Execute(str)        '~~~> Execute search
        
        For Each match In matches
            Debug.Print match.Value        '~~~> Prints: 7026, 7027, 7033
        Next
    End If
End Sub
