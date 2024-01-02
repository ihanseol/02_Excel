Attribute VB_Name = "modDumpArrayMyData"

'2024-01-02
'이것이 파일로 세이브 하는 메인함수이다.
'이것으로 강수량 데이타를 세이브 할수있다.

Sub DumpRangeToArrayAndSaveTest()
Attribute DumpRangeToArrayAndSaveTest.VB_ProcData.VB_Invoke_Func = "d\n14"
    Dim myArray() As Variant
    Dim rng As Range
    Dim i As Integer, j As Integer
    Dim AREA_STR As String
    
    ' Set the range you want to dump to an array
    Set rng = ThisWorkbook.ActiveSheet.Range("B6:N35")
    
    ' Read the range into an array
    myArray = rng.Value
    
    ' Save array to a file
    Dim FilePath As String
    
    
    FilePath = ThisWorkbook.Path & "\" & ActiveSheet.name & ".csv"
    SaveArrayToFileByExcelForm myArray, FilePath
 
End Sub


Function getAreaName()
    Dim lookupValue As Variant
    Dim result As Variant

    Dim tableRange As Range
    Set tableRange = Range("tblAREAREF")

    If ActiveSheet.name = "main" Then
        lookupValue = Range("S8")
    Else
        lookupValue = ActiveSheet.name
    End If
    
    On Error Resume Next
    result = Application.VLookup(lookupValue, tableRange, 2, False)
    On Error GoTo 0
    
    If Not IsError(result) Then
        getAreaName = UCase(result)
    Else
        ' If no match is found, display an error message
        getAreaName = "MAIN"
    End If

End Function



Private Sub SaveArrayToFileByExcelForm(myArray As Variant, FilePath As String)
    Dim i As Integer, j As Integer
    Dim FileNum As Integer
    Dim AREA_STR As String
    
    FileNum = FreeFile
    AREA_STR = getAreaName()
    
    
    Open FilePath For Output As FileNum
    
    Print #FileNum, "Function data_" & AREA_STR & "() As Variant"
    Print #FileNum, ""
    Print #FileNum, "    Dim myArray() As Variant"
    Print #FileNum, "    ReDim myArray(1 To 30, 1 To 13)"
    Print #FileNum, ""
    
    For i = LBound(myArray, 1) To UBound(myArray, 1)
        For j = LBound(myArray, 2) To UBound(myArray, 2)
            Print #FileNum, "myArray(" & i & ", " & j & ") = ";
            
            ' Separate values with a comma (CSV format)
            If j <= UBound(myArray, 2) Then
                Print #FileNum, myArray(i, j);
            End If
            
            Print #FileNum, ""
        Next j
        ' Start a new line for each row
        Print #FileNum, ""
    Next i
    
    Print #FileNum, ""
    Print #FileNum, "    data_" & AREA_STR & "= myArray"
    Print #FileNum, ""
    Print #FileNum, "End Function"
    
    Close FileNum
End Sub


Private Sub importFromArray()
    Dim myArray As Variant
    Dim rng As Range
    
    indexString = "data_" & UCase(Range("s11").Value)
    
    On Error Resume Next
    myArray = Application.Run(indexString)
    On Error GoTo 0
    
    If Not IsArray(myArray) Then
        MsgBox "An error occurred while fetching data.", vbExclamation
        Exit Sub
    End If
    
    
    Set rng = ThisWorkbook.ActiveSheet.Range("B6:N35")
    rng.Value = myArray
       
End Sub





