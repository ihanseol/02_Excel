Attribute VB_Name = "mod_FilterString"
Option Explicit



Function FilterString(str As String) As Variant
    Dim elements() As String
    Dim element As Variant
    Dim out() As String
    Dim i As Long
    
    elements = Split(str, ",")
    For Each element In elements
        If element = "" Then Exit For
        ReDim Preserve out(i)
        out(i) = Trim(element)
        i = i + 1
    Next element

    FilterString = out
End Function


'Function DeepCopyString(originalStr As String) As String
'    Dim copiedStr As String
'
'    copiedStr = StrConv(originalStr, vbFromUnicode)
'    DeepCopyString = copiedStr
'
'    ' Debug.Print "Original string: " & originalStr
'    ' Debug.Print "Copied string: " & copiedStr
'End Function


Function DeepCopyString(originalStr As String) As String
    
    DeepCopyString = Left$(originalStr, Len(originalStr))
End Function


Sub TestFilterString()
    Dim out() As Variant
    Dim i As Integer
    
    out = FilterString("a,b, c,d,af, ae, k, x, ag")
    
    Debug.Print "***************************"
    
    For i = LBound(out) To UBound(out)
        
        Debug.Print out(i)
        Debug.Print "***************************"
    
    Next i
End Sub

