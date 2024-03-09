Attribute VB_Name = "modDumpArrayMyData"

Sub DumpRangeToArrayAndSaveTest()
Attribute DumpRangeToArrayAndSaveTest.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim myArray() As Variant
    Dim rng As Range
    Dim i As Integer, j As Integer
    
    ' Set the range you want to dump to an array
    Set rng = ThisWorkbook.ActiveSheet.Range("B6:N35")
    
    ' Read the range into an array
    myArray = rng.value
    
    ' Save array to a file
    Dim filePath As String
    filePath = ThisWorkbook.Path & "\" & ActiveSheet.name & ".csv"
    SaveArrayToFileByExcelForm myArray, filePath
    
End Sub



Sub SaveArrayToFileByExcelForm(myArray As Variant, filePath As String)
    Dim i As Integer, j As Integer
    Dim FileNum As Integer
    
    FileNum = FreeFile
    
    Open filePath For Output As FileNum
    
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
    
    Close FileNum
End Sub


Sub importFromArray()
    Dim myArray As Variant
    Dim rng As Range
    
    indexString = "data_" & UCase(Range("s11").value)
    
    myArray = Application.Run(indexString)
    
    
    Set rng = ThisWorkbook.ActiveSheet.Range("B6:N35")
    rng.value = myArray
       
End Sub





