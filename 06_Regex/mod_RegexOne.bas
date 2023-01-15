Attribute VB_Name = "mod_RegexOne"
Option Explicit

Function RegexReplace(ByVal text As String, _
                      ByVal replace_what As String, _
                      ByVal replace_with As String) As String

    Dim RE As Object
    Set RE = CreateObject("vbscript.regexp")
    
    RE.Pattern = replace_what
    RE.Global = True
    RegexReplace = RE.Replace(text, replace_with)

End Function

Private Sub simpleRegex()
    Dim strPattern As String: strPattern = "^[0-9]{1,2}"
    Dim strReplace As String: strReplace = ""
    Dim regEx As New RegExp
    Dim strInput As String
    Dim myRange As Range
    
    Set myRange = ActiveSheet.Range("A1")
    
    If strPattern <> "" Then
        strInput = myRange.Value
        With regEx
            .Global = True
            .MultiLine = True
            .IgnoreCase = False
            .Pattern = strPattern
        End With
        If regEx.test(strInput) Then
            MsgBox (regEx.Replace(strInput, strReplace))
        Else
            MsgBox ("Not matched")
        End If
    End If
End Sub

Function simpleCellRegex(myRange As Range) As String
    Dim regEx As New RegExp
    Dim strPattern As String
    Dim strInput As String
    Dim strReplace As String
    Dim strOutput As String
    
    strPattern = "^[0-9]{1,3}"
    
    If strPattern <> "" Then
        strInput = myRange.Value
        strReplace = ""
        
        With regEx
            .Global = True
            .MultiLine = True
            .IgnoreCase = False
            .Pattern = strPattern
        End With
        
        If regEx.test(strInput) Then
            simpleCellRegex = regEx.Replace(strInput, strReplace)
        Else
            simpleCellRegex = "Not matched"
        End If
    End If
End Function

Private Sub simpleRegex1()
    Dim strPattern As String: strPattern = "^[0-9]{1,2}"
    Dim strReplace As String: strReplace = ""
    Dim regEx As New RegExp
    Dim strInput As String
    Dim myRange As Range
    
    Set myRange = ActiveSheet.Range("A1:A5")
    
    For Each cell In myRange
        If strPattern <> "" Then
            strInput = cell.Value
            
            With regEx
                .Global = True
                .MultiLine = True
                .IgnoreCase = False
                .Pattern = strPattern
            End With
            
            If regEx.test(strInput) Then
                MsgBox (regEx.Replace(strInput, strReplace))
            Else
                MsgBox ("Not matched")
            End If
        End If
    Next
End Sub



Private Sub splitUpRegexPattern()
    Dim regEx As New RegExp
    Dim strPattern As String
    Dim strInput As String
    Dim myRange As Range
    
    Set myRange = ActiveSheet.Range("A1:A3")
    
    For Each C In myRange
        strPattern = "(^[0-9]{3})([a-zA-Z])([0-9]{4})"
        
        If strPattern <> "" Then
            strInput = C.Value
            
            With regEx
                .Global = True
                .MultiLine = True
                .IgnoreCase = False
                .Pattern = strPattern
            End With
            
            If regEx.test(strInput) Then
                C.Offset(0, 1) = regEx.Replace(strInput, "$1")
                C.Offset(0, 2) = regEx.Replace(strInput, "$2")
                C.Offset(0, 3) = regEx.Replace(strInput, "$3")
            Else
                C.Offset(0, 1) = "(Not matched)"
            End If
        End If
    Next

End Sub


