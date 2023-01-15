Attribute VB_Name = "find"
Sub find()
    'pages 4-5 in Excel VBA Find - A Complete Guide - Excel Macro Mastery.pdf lists Find parameters
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("find").Activate
    Dim namerange As Range
    Set namerange = Range("A1:A5")
    Range("B1").Value = namerange.find("Jena")   'print Jena
    Range("B2").Value = namerange.find("Jena").Address 'print $A$4

    'Find function unfriendly not finding a match.  You need to handle error message.
    Dim eachnamerange As Range
    For Each eachnamerange In namerange
        If eachnamerange.find("John") Is Nothing Then 'not finding a match
            eachnamerange.offset(0, 2).Value = "Not here"
        Else
            eachnamerange.offset(0, 2).Value = "Here"
        End If
    Next eachnamerange
    For Each eachnamerange In namerange
        If eachnamerange.find("Jena") Is Nothing Then 'not finding a match
            eachnamerange.offset(0, 3).Value = "Not here"
        Else
            eachnamerange.offset(0, 3).Value = "Here"
        End If
    Next eachnamerange

    'find the second instance
    Range("B7").Value = Range("A7:A12").find("Rachal", after:=Range("A8")).Address 'print $A$12
    'If a match is not found then the search will “wrap around? This means it will go back to the start _
    of the range.
    Range("B8").Value = Range("A7:A12").find("Drucilla", after:=Range("A8")).Address 'print $A$7

    'find partial and find whole
    Range("C14") = Range("A14:A15").find("Apple", lookat:=xlPart).Address 'print $A$15
    Range("C15") = Range("A14:A15").find("Apple", lookat:=xlWhole).Address 'print $A$14
    'RM:  notice if Apple Orange was A14 and Apple was A15, then xlPart is $A$15.  I don't know why.

    'match case
    Dim matchcaserange As Range
    Set matchcaserange = Range("A17:A23")
    For Each f In matchcaserange
        If f.find("elli", MatchCase:=True) Is Nothing Then
            f.offset(0, 1) = "Not found elli"
        Else
            f.offset(0, 1) = "Found elli"
        End If
    Next f

    'SearchFormat.  Code below finds Elli bolded.
    Application.FindFormat.Font.Bold = True
    Dim matchformatrange As Range
    Set matchformatrange = Range("A25:A31")
    For Each g In matchformatrange
        If g.find("Elli", SearchFormat:=True) Is Nothing Then
            g.offset(0, 1) = "Not found elli"
        Else
            g.offset(0, 1) = "Found elli"
        End If
    Next g
    Application.FindFormat.Clear                 'When you set the FindFormat attributes they remain in _
                                                 place until you set them again.  It is a good idea to clear the format before you use it.

    'wild card.  Find "li" in text.
    Dim wildcardrange As Range
    Set wildcardrange = Range("A25:A31")
    For Each wild In wildcardrange
        If wild.find("*li*") Is Nothing Then
            wild.offset(0, 2) = "Wild card search found nothing"
        Else
            wild.offset(0, 2) = "Wild card search found something"
        End If
    Next wild

    'multiple searches.  Use Find and FindNext.
    'error message .FindNext
    'Range("B33").Value = Range("A33:A41").find("Elli").Address
    'Range("B34").Value = Range("A33:A41").FindNext("Elli").Address
    'error message .FindNext
    'Dim findnextrange As Range
    'Set findnextrange = Range("A33:A41")
    'Range("B33") = findnextrange.find("Elli").Address
    'Range("B34") = findnextrange.FindNext("Elli").Address
    'RM:  Use For Each Loop

    'find last cell containing data
    ''find last row containing data column A
    'lastrowcolumna = Cells(Rows.count, 1).End(xlUp).Row
    'MsgBox lastrowcolumna 'return 41
    ''find last row containing data column B
    'lastrowcolumnb = Cells(Rows.count, 2).End(xlUp).Row
    'MsgBox lastrowcolumnb 'return 33
    ''find last column contianing data row 1
    'lastcolumnrow1 = Cells(1, Columns.count).End(xlToLeft).Column
    'MsgBox lastcolumnrow1 'return 4
    ''find last column contianing data row 7
    'lastcolumnrow7 = Cells(7, Columns.count).End(xlToLeft).Column
    'MsgBox lastcolumnrow7 'return 2

    'find cells with patterns
    Dim wilda33a41, lettere As Range
    Set wilda33a41 = Range("A33:A41")
    For Each lettere In wilda33a41
        If lettere Like "[E]*" Then
            lettere.offset(0, 2).Value = "Wild card found name begins with E"
        Else
            lettere.offset(0, 2).Value = "Wild card not found name begins with E"
        End If
    Next lettere

    'find and replace is available.  Use the Replace function.  Replace function not taught here.
End Sub


