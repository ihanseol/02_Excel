Sub ChangeFileNameInCurrentDir()
    Dim CurrentDir As String
    Dim OldFileName As String
    Dim NewFileName As String

    ' Get the current working directory
    CurrentDir = ThisWorkbook.Path & "\"

    ' Define the old and new file names
    OldFileName = "myArray.csv"
    NewFileName = ActiveSheet.name & ".csv"

    ' Check if the old file exists in the current directory
    If Dir(CurrentDir & OldFileName) <> "" Then
        ' Rename the file
        Name CurrentDir & OldFileName As CurrentDir & NewFileName
        MsgBox "File name changed successfully!"
    Else
        MsgBox "The old file does not exist in the current directory."
    End If
End Sub


Sub SaveRangeToFile()
    Dim ws As Worksheet
    Dim rng As Range
    Dim filePath As String
    
    ' Set the worksheet and range
    Set ws = ThisWorkbook.ActiveSheet
    Set rng = ws.Range("B6:N35")
    
    ' Generate the file path using the worksheet name
    filePath = ThisWorkbook.Path & "\" & ws.name & ".csv"
    
    ' Save the range to a CSV file
    rng.ExportAsFixedFormat Type:=xlCSV, fileName:=filePath, Quality:=xlQualityStandard
End Sub










