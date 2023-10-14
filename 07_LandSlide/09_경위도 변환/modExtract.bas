Attribute VB_Name = "modExtract"
Function SplitText(inputText As String) As Variant
    Dim splitArray() As String
    Dim length As Integer
    
    ' Check if the input string contains a comma
    If InStr(inputText, ",") > 0 Then
        ' If comma is found, proceed with splitting
        splitArray = Split(inputText, ",")
        length = UBound(splitArray) - LBound(splitArray) + 1
    Else
        ' If comma is not found, handle the error here
        ' For example, you can set both elements to "0"
        ReDim splitArray(0 To 1)
        splitArray(0) = "0"
        splitArray(1) = "0"
        length = 2
    End If
    
    ' Optional: Print the elements
    Debug.Print "Element 1: " & splitArray(0)
    Debug.Print "Element 2: " & splitArray(1)
    
    ' Return the array
    SplitText = splitArray
End Function




Sub SeperateVH()
Attribute SeperateVH.VB_ProcData.VB_Invoke_Func = "d\n14"
    Dim result() As String


    result = SplitText(ActiveCell.Value)
    
    
    Debug.Print result(0), result(1)
    
    Range("c" & ActiveCell.Row).Value = result(0)
    Range("d" & ActiveCell.Row).Value = result(1)


End Sub





