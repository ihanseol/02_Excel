Private Sub DumpRangeToArray()
    Dim myArray As Variant
    Dim rng As Range
    Dim cell As Range
    Dim i As Integer, j As Integer

    ' Set the range you want to dump to an array
    Set rng = ThisWorkbook.ActiveSheet.Range("B6:N35")
    
    ' Read the range into an array
    myArray = rng.Value
    
    ' Loop through the array (for demonstration purposes)
    For i = LBound(myArray, 1) To UBound(myArray, 1)
        For j = LBound(myArray, 2) To UBound(myArray, 2)
            Debug.Print myArray(i, j)
        Next j
    Next i
End Sub


Private Sub DumpRangeToArrayAndSaveLoad()
    Dim myArray() As Variant
    Dim rng As Range
    Dim i As Integer, j As Integer
    
    ' Set the range you want to dump to an array
    Set rng = ThisWorkbook.ActiveSheet.Range("B6:N35")
    
    ' Read the range into an array
    myArray = rng.Value
    
    ' Save array to a file
    Dim filePath As String
    filePath = ThisWorkbook.Path & "\myArray.csv"
    SaveArrayToFile myArray, filePath
    
    ' Load array from file
    Dim loadedArray() As Variant
    Dim finalArray() As Variant
    
    loadedArray = LoadArrayFromFile(filePath)
    
    'ReDim loadedArray(1 To 30, 1 To 13)
    ' Check if the loaded array is the same as the original array
    
    For i = LBound(myArray, 1) To UBound(myArray, 1)
        For j = LBound(myArray, 2) To UBound(myArray, 2)
            If myArray(i, j) <> CDbl(loadedArray(i, j)) Then
                MsgBox "Arrays are not equal!" & "i :" & i & "  j : " & j
                Exit Sub
            End If
        Next j
    Next i
    
    MsgBox "Arrays are equal!"
End Sub


Private Sub SaveArrayToFile(myArray As Variant, filePath As String)
    Dim i As Integer, j As Integer
    Dim FileNum As Integer
    
    FileNum = FreeFile
    
    Open filePath For Output As FileNum
    
    For i = LBound(myArray, 1) To UBound(myArray, 1)
        For j = LBound(myArray, 2) To UBound(myArray, 2)
            Print #FileNum, myArray(i, j);
            
            ' Separate values with a comma (CSV format)
            If j < UBound(myArray, 2) Then
                Print #FileNum, ",";
            End If
        Next j
        ' Start a new line for each row
        Print #FileNum, ""
    Next i
    
    Close FileNum
End Sub

Private Function LoadArrayFromFile(filePath As String) As Variant
    Dim FileContent As String
    Dim Lines() As String
    Dim Values() As String
    Dim i As Integer, j As Integer
    Dim loadedArray() As Variant
    
    Open filePath For Input As #1
    FileContent = Input$(LOF(1), #1)
    Close #1
    
    Lines = Split(FileContent, vbCrLf)
    ReDim loadedArray(1 To UBound(Lines) + 1, 1 To UBound(Split(Lines(0), ",")) + 1)
    
    For i = LBound(Lines) To UBound(Lines)
        Values = Split(Lines(i), ",")
        For j = LBound(Values) To UBound(Values)
            loadedArray(i + 1, j + 1) = Values(j)
        Next j
    Next i
    
    LoadArrayFromFile = loadedArray
End Function
