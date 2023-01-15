Attribute VB_Name = "Chapter10"
Sub zapthevowels()
    Application.Workbooks("excel2010powerprogrammingbasics.xlsm").Worksheets("10").Activate
    inputword = Range("A4").Value                'thequickbrowfoxjumpedoverthelazydog
    Range("B5").Value = removevowels(inputword)  'print thqckbrwfxjmpdvrthlzydg
End Sub

Function removevowels(text) As String
    Dim i As Long
    removevowels = ""
    For i = 1 To Len(text)
        'Mid function to return a single character from the input string and converts this character
        'to uppercase. That character is then compared to a list of characters by using VBA’s Like
        'operator. In other words, the If clause is true if the character isn’t A, E, I, O, or U.
        'In such a case, the character is appended to the RemoveVowels variable.
        If Not UCase(Mid(text, i, 1)) Like "[AEIOU]" Then
            removevowels = removevowels & Mid(text, i, 1)
        End If
    Next i
End Function

Function commission(sales)
    Application.Workbooks("excel2010powerprogrammingbasics.xlsm").Worksheets("10").Activate
    'vlookup is better
    Const tier1 = 0.08
    Const tier2 = 0.105
    Const tier3 = 0.12
    Const tier4 = 0.14
    Select Case sales
    Case 0 To 999.99: commission = sales * tier1
    Case 1000 To 19999.99: commission = sales * tier2
    Case 20000 To 39999.99: commission = sales * tier3
    Case Is >= 40000: commission = sales * tier4
    End Select
End Function

Function commission2(sales, years)
    Application.Workbooks("excel2010powerprogrammingbasics.xlsm").Worksheets("10").Activate
    'vlookup is better
    Const tier1 = 0.08
    Const tier2 = 0.105
    Const tier3 = 0.12
    Const tier4 = 0.14
    Select Case sales
    Case 0 To 999.99: commission2 = sales * tier1 * years
    Case 1000 To 19999.99: commission2 = sales * tier2 * years
    Case 20000 To 39999.99: commission2 = sales * tier3 * years
    Case Is >= 40000: commission2 = sales * tier4 * years
    End Select
End Function

Function commission3(years, Optional sales)
    Application.Workbooks("excel2010powerprogrammingbasics.xlsm").Worksheets("10").Activate
    Const tier1 = 0.08
    Const tier2 = 0.105
    Const tier3 = 0.12
    Const tier4 = 0.14
    Select Case sales
    Case 0 To 999.99: commission3 = sales * tier1 * years
    Case 1000 To 19999.99: commission3 = sales * tier2 * years
    Case 20000 To 39999.99: commission3 = sales * tier3 * years
    Case Is >= 40000: commission3 = sales * tier4 * years
    End Select
End Function

