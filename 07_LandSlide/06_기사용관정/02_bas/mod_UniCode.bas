Attribute VB_Name = "mod_UniCode"
Option Explicit

Sub GetCheckMarkCode()
    Dim checkMark As String
    Dim code As Long

    checkMark = "?" ' the check mark symbol

    code = AscW(checkMark)

    Debug.Print "The Unicode code point for " & checkMark & " is " & code
End Sub

Sub InsertCheckMark()
    Dim checkMark As String

    checkMark = ChrW(&H2714) ' &H2714 is the Unicode code point for the check mark symbol

    Range("A1").Value = checkMark ' Replace "A1" with the cell where you want to insert the check mark symbol
End Sub

Sub TestUniCode()
    Dim i As Integer
    Dim str_check As String
    Dim code As Long
    Dim index As Variant
        
    str_check = Sheets("index").Range("a1").Value
    code = AscW(str_check)
    
    index = Array("a", "b", "c", "f", "k")
    
    For i = LBound(index) To UBound(index)
        Cells(33, index(i)).Value = ChrW(&H2714)   ' str_check
    Next i
    
    Debug.Print "strcheck", code
End Sub



